const fs = require('fs');
const path = require('path');
const https = require('https');

function getRefreshToken() {
  const home = process.env.HOME || process.env.USERPROFILE;
  const configPath = path.join(home, '.config', 'configstore', 'firebase-tools.json');
  if (!fs.existsSync(configPath)) {
    throw new Error('找不到 firebase-tools.json，請確認是否已登入');
  }
  const config = JSON.parse(fs.readFileSync(configPath, 'utf8'));
  const tokens = config.tokens || {};
  let refreshToken = tokens.refresh_token || (config.tokens && config.tokens.default && config.tokens.default.refresh_token);
  if (!refreshToken && config.tokens) {
    for (const key in config.tokens) {
      if (config.tokens[key] && config.tokens[key].refresh_token) {
        refreshToken = config.tokens[key].refresh_token;
        break;
      }
    }
  }
  if (!refreshToken) {
    throw new Error('未在設定檔中找到 refresh_token，請先執行 npx firebase-tools login');
  }
  return refreshToken;
}

function getAccessToken(refreshToken) {
  return new Promise((resolve, reject) => {
    const reqData = new URLSearchParams({
      client_id: '563577406560-3k6s4j563s11tkmh361s4j563s11tkmh.apps.googleusercontent.com',
      client_secret: 'v2e_gh747snm7l-N65g26uSM',
      grant_type: 'refresh_token',
      refresh_token: refreshToken
    }).toString();

    const options = {
      hostname: 'oauth2.googleapis.com',
      path: '/token',
      method: 'POST',
      headers: {
        'Content-Type': 'application/x-www-form-urlencoded',
        'Content-Length': Buffer.byteLength(reqData)
      }
    };

    const req = https.request(options, (res) => {
      let body = '';
      res.on('data', (chunk) => body += chunk);
      res.on('end', () => {
        try {
          const parsed = JSON.parse(body);
          if (parsed.access_token) {
            resolve(parsed.access_token);
          } else {
            reject(new Error('取得 Access Token 失敗：' + body));
          }
        } catch (e) {
          reject(e);
        }
      });
    });

    req.on('error', reject);
    req.write(reqData);
    req.end();
  });
}

// 查閱不設條件的所有 answerBatches
function runQueryAll(projectId, accessToken) {
  return new Promise((resolve, reject) => {
    const payload = JSON.stringify({
      structuredQuery: {
        from: [{ collectionId: 'answerBatches' }]
      }
    });

    const options = {
      hostname: 'firestore.googleapis.com',
      path: `/v1/projects/${projectId}/databases/(default)/documents:runQuery`,
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'Authorization': `Bearer ${accessToken}`,
        'Content-Length': Buffer.byteLength(payload)
      }
    };

    const req = https.request(options, (res) => {
      let body = '';
      res.on('data', (chunk) => body += chunk);
      res.on('end', () => {
        try {
          const parsed = JSON.parse(body);
          resolve(parsed);
        } catch (e) {
          reject(e);
        }
      });
    });

    req.on('error', reject);
    req.write(payload);
    req.end();
  });
}

function parseVal(val) {
  if (!val) return null;
  if (val.stringValue !== undefined) return val.stringValue;
  if (val.integerValue !== undefined) return parseInt(val.integerValue, 10);
  if (val.doubleValue !== undefined) return parseFloat(val.doubleValue);
  if (val.booleanValue !== undefined) return val.booleanValue;
  if (val.timestampValue !== undefined) return val.timestampValue;
  if (val.arrayValue !== undefined) {
    const list = val.arrayValue.values || [];
    return list.map(parseVal);
  }
  if (val.mapValue !== undefined) {
    const fields = val.mapValue.fields || {};
    const obj = {};
    for (const key in fields) {
      obj[key] = parseVal(fields[key]);
    }
    return obj;
  }
  return null;
}

function parseFields(fields) {
  const obj = {};
  for (const key in fields) {
    obj[key] = parseVal(fields[key]);
  }
  return obj;
}

async function main() {
  try {
    const token = getRefreshToken();
    const accessToken = await getAccessToken(token);
    
    console.log(`正在從 Firebase 下載所有 answerBatches (無任何過濾)...`);
    const results = await runQueryAll('ap-neuron', accessToken);
    
    console.log(`從 Firebase 取得的原始結果長度: ${results ? results.length : 0}`);
    
    let allDocs = [];
    if (Array.isArray(results)) {
      results.forEach((r) => {
        if (r.document && r.document.fields) {
          allDocs.push(parseFields(r.document.fields));
        }
      });
    }
    
    console.log(`解析成功文件數: ${allDocs.length}`);
    
    // 儲存為 JSON 備查
    fs.writeFileSync(path.join(__dirname, 'all_firebase_batches.json'), JSON.stringify(allDocs, null, 2), 'utf8');
    
    // 搜尋「張欣翰」
    const target = allDocs.filter(d => d.name && d.name.includes('張欣翰'));
    console.log(`\n🔍 搜尋「張欣翰」結果共 ${target.length} 筆：`);
    target.forEach((t, i) => {
      console.log(`[${i+1}] 學號: ${t.studentId}, 單元: ${t.topic}, 分數: ${t.score}, 時間: ${t.createdAt}`);
    });
    
    // 輸出最早與最晚的時間
    if (allDocs.length > 0) {
      const times = allDocs.map(d => d.createdAt).filter(Boolean).sort();
      console.log(`\n最早紀錄時間: ${times[0]}`);
      console.log(`最晚紀錄時間: ${times[times.length - 1]}`);
    }
    
  } catch (err) {
    console.error('執行錯誤:', err.message);
  }
}

main();
