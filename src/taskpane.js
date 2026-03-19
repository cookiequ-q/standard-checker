/**
 * Standard Checker Word Add-in (IE11 compatible, JSONP)
 */
import 'core-js/stable/promise';
import 'core-js/stable/object/assign';
import 'core-js/stable/array/index-of';

// ============ Config ============
var API_BASE = 'https://standard-checker-api.cookiequ747.workers.dev';
var CACHE_TTL = 30 * 24 * 3600 * 1000;

// ============ Regex ============
var STD_PATTERN = /(?:GB\/T|GB|GBZ|HG\/T|HG|AQ|TSG|SH\/T|SH|JB\/T|JB|DL\/T|DL|DB\d{2}\/T|DB\d{2})\s*[\d]+(?:\.[\d]+)*(?:-[\d]{4})?(?:\/(?:XG|T)\d*-\d{4})?/g;
var LAW_PATTERN = /\u300a([^\u300b]+)\u300b(?:\s*[\uff08(]([^\uff09)]+)[\uff09)])?/g;
var LAW_KEYWORDS = ['\u6cd5', '\u6761\u4f8b', '\u89c4\u5b9a', '\u529e\u6cd5', '\u89c4\u7a0b', '\u5bfc\u5219', '\u7ec6\u5219', '\u901a\u77e5', '\u610f\u89c1', '\u51b3\u5b9a', '\u4ee4', '\u516c\u544a'];

// ============ Cache ============
function getCached(code) {
  try {
    var key = 'sc_' + code;
    var raw = localStorage.getItem(key);
    if (!raw) return null;
    var data = JSON.parse(raw);
    if (Date.now() - data._cachedAt > CACHE_TTL) {
      localStorage.removeItem(key);
      return null;
    }
    return data;
  } catch (e) { return null; }
}

function setCache(code, data) {
  try {
    var key = 'sc_' + code;
    data._cachedAt = Date.now();
    localStorage.setItem(key, JSON.stringify(data));
  } catch (e) { }
}

// ============ JSONP (IE11 cross-origin) ============
var jsonpCounter = 0;
function jsonpGet(url) {
  return new Promise(function(resolve, reject) {
    var cbName = '_sc_cb_' + (++jsonpCounter) + '_' + Date.now();
    var timeout = setTimeout(function() {
      cleanup();
      reject(new Error('Timeout'));
    }, 120000);

    function cleanup() {
      clearTimeout(timeout);
      try { delete window[cbName]; } catch(e) { window[cbName] = undefined; }
      var el = document.getElementById(cbName);
      if (el && el.parentNode) el.parentNode.removeChild(el);
    }

    window[cbName] = function(data) {
      cleanup();
      resolve(data);
    };

    var sep = url.indexOf('?') === -1 ? '?' : '&';
    var script = document.createElement('script');
    script.id = cbName;
    script.src = url + sep + 'callback=' + cbName;
    script.onerror = function() {
      cleanup();
      reject(new Error('Network error'));
    };
    document.head.appendChild(script);
  });
}

// ============ API query with progress ============
function queryAPI(codes, onProgress) {
  var cached = [];
  var toQuery = [];
  var i;

  for (i = 0; i < codes.length; i++) {
    var c = getCached(codes[i]);
    if (c) {
      cached.push(c);
    } else {
      toQuery.push(codes[i]);
    }
  }

  if (toQuery.length === 0) {
    if (onProgress) onProgress(1);
    return Promise.resolve(cached);
  }

  var queried = [];
  var batchSize = 10;
  var batches = [];
  for (i = 0; i < toQuery.length; i += batchSize) {
    batches.push(toQuery.slice(i, i + batchSize));
  }

  var completed = cached.length;
  var total = codes.length;
  var chain = Promise.resolve();

  batches.forEach(function(batch) {
    chain = chain.then(function() {
      var url = API_BASE + '/api/check?codes=' + encodeURIComponent(batch.join(','));
      return jsonpGet(url).then(function(results) {
        if (!Array.isArray(results)) results = [results];
        for (var j = 0; j < results.length; j++) {
          setCache(results[j].code, results[j]);
          queried.push(results[j]);
        }
        completed += batch.length;
        if (onProgress) onProgress(completed / total);
      }).catch(function(err) {
        for (var j = 0; j < batch.length; j++) {
          queried.push({ code: batch[j], level: 'unknown', message: 'Query failed: ' + err.message });
        }
        completed += batch.length;
        if (onProgress) onProgress(completed / total);
      });
    });
  });

  return chain.then(function() {
    return cached.concat(queried);
  });
}

// ============ Document scan ============
function scanDocument() {
  var btnScan = document.getElementById('btn-scan');
  btnScan.disabled = true;
  btnScan.textContent = '扫描中...';
  showStatus('正在扫描文档...');

  Word.run(function(context) {
    var body = context.document.body;
    body.load('text');
    return context.sync().then(function() {
      var fullText = body.text;
      var stdCodes = [];
      var m;

      STD_PATTERN.lastIndex = 0;
      while ((m = STD_PATTERN.exec(fullText)) !== null) {
        var code = m[0].replace(/\s+/g, ' ').trim();
        if (stdCodes.indexOf(code) === -1) stdCodes.push(code);
      }

      var lawIssues = [];
      LAW_PATTERN.lastIndex = 0;
      while ((m = LAW_PATTERN.exec(fullText)) !== null) {
        var name = m[1];
        var docNumber = m[2] || '';
        var isLaw = false;
        for (var k = 0; k < LAW_KEYWORDS.length; k++) {
          if (name.indexOf(LAW_KEYWORDS[k]) !== -1) { isLaw = true; break; }
        }
        if (!isLaw) continue;

        if (!docNumber) {
          lawIssues.push({
            level: 'green', type: 'law', text: m[0], name: name, code: name,
            message: '缺少文号',
            suggestion: '格式：《法规名称》（文号）',
          });
        } else {
          lawIssues.push({
            level: 'ok', type: 'law', text: m[0], name: name, code: name,
            message: '格式正确',
          });
        }
      }

      showStatus('正在查询 ' + stdCodes.length + ' 条标准...');
      showProgress(0);

      var queryPromise = stdCodes.length > 0 ? queryAPI(stdCodes, function(pct) {
        showProgress(pct);
        showStatus('查询中... ' + Math.round(pct * 100) + '%');
      }) : Promise.resolve([]);

      return queryPromise.then(function(results) {
        var issues = [];
        for (var i = 0; i < results.length; i++) {
          var r = results[i];
          issues.push({
            level: r.level || 'ok',
            type: 'standard',
            text: r.code,
            code: r.code,
            message: r.message || '',
            suggestion: r.suggestion || (r.replacementCode ? '替代标准: ' + r.replacementCode : ''),
            name: r.name || '',
          });
        }
        issues = issues.concat(lawIssues);

        var highlightPromise = Promise.resolve();
        var COLORS = { red: 'Red', yellow: 'Yellow', green: 'BrightGreen', blue: 'Turquoise' };

        issues.forEach(function(issue) {
          if (issue.level === 'ok' || issue.level === 'unknown') return;
          var color = COLORS[issue.level];
          if (!color) return;

          highlightPromise = highlightPromise.then(function() {
            var searchResults = body.search(issue.text, { matchCase: true, matchWholeWord: false });
            searchResults.load('items');
            return context.sync().then(function() {
              for (var j = 0; j < searchResults.items.length; j++) {
                searchResults.items[j].font.highlightColor = color;
                try {
                  var comment = issue.message;
                  if (issue.suggestion) comment += '\n' + issue.suggestion;
                  searchResults.items[j].insertComment(comment);
                } catch (e) { }
              }
            });
          });
        });

        return highlightPromise.then(function() {
          return context.sync();
        }).then(function() {
          displayResults(issues);
          var problemCount = 0;
          for (var i = 0; i < issues.length; i++) {
            if (issues[i].level !== 'ok') problemCount++;
          }
          showStatus('检查完成！在 ' + issues.length + ' 个条目中发现 ' + problemCount + ' 个问题');
        });
      });
    });
  }).catch(function(error) {
    showStatus('错误: ' + error.message, true);
  }).then(function() {
    btnScan.disabled = false;
    btnScan.textContent = '开始检查';
  });
}

// ============ Clear highlights ============
function clearHighlights() {
  Word.run(function(context) {
    var paragraphs = context.document.body.paragraphs;
    paragraphs.load('items');
    return context.sync().then(function() {
      for (var i = 0; i < paragraphs.items.length; i++) {
        paragraphs.items[i].getRange().font.highlightColor = 'None';
      }
      return context.sync();
    }).then(function() {
      showStatus('已清除所有标记');
      document.getElementById('results').classList.add('hidden');
    });
  }).catch(function(error) {
    showStatus('错误: ' + error.message, true);
  });
}

// ============ UI ============
function showStatus(text, isError) {
  var bar = document.getElementById('status-bar');
  var textEl = document.getElementById('status-text');
  bar.className = 'status-bar' + (isError ? ' error' : '');
  textEl.textContent = text;
}

function showProgress(pct) {
  var bar = document.getElementById('progress-bar');
  var fill = document.getElementById('progress-fill');
  if (pct <= 0) {
    bar.className = 'progress-bar';
    fill.style.width = '0%';
  } else if (pct >= 1) {
    fill.style.width = '100%';
    setTimeout(function() { bar.className = 'progress-bar hidden'; }, 500);
  } else {
    bar.className = 'progress-bar';
    fill.style.width = Math.round(pct * 100) + '%';
  }
}

function debugLog(msg) {
  var el = document.getElementById('debug-info');
  if (el) el.innerHTML += msg + '<br>';
}

function displayResults(issues) {
  var container = document.getElementById('results');
  var list = document.getElementById('results-list');
  container.className = 'results';

  var counts = { red: 0, yellow: 0, green: 0, blue: 0, ok: 0 };
  for (var i = 0; i < issues.length; i++) {
    if (counts[issues[i].level] !== undefined) counts[issues[i].level]++;
  }

  document.getElementById('count-red').textContent = counts.red;
  document.getElementById('count-yellow').textContent = counts.yellow;
  document.getElementById('count-green').textContent = counts.green;
  document.getElementById('count-blue').textContent = counts.blue;
  document.getElementById('count-ok').textContent = counts.ok;

  var order = { red: 0, yellow: 1, green: 2, blue: 3, ok: 4, unknown: 5 };
  var sorted = issues.slice().sort(function(a, b) {
    return (order[a.level] || 4) - (order[b.level] || 4);
  });

  var html = '';
  for (var j = 0; j < sorted.length; j++) {
    var issue = sorted[j];
    html += '<div class="result-item level-' + issue.level + '">';
    html += '<div class="code">' + (issue.code || issue.name || issue.text) + '</div>';
    if (issue.name && issue.type === 'standard') {
      html += '<div class="name">' + issue.name + '</div>';
    }
    html += '<div class="message">' + issue.message + '</div>';
    if (issue.suggestion) {
      html += '<div class="suggestion">' + issue.suggestion + '</div>';
    }
    html += '</div>';
  }
  list.innerHTML = html;
}

// ============ Init ============
function initApp() {
  window._officeReady = true;
  debugLog('初始化完成');
  showStatus('就绪');
  document.getElementById('btn-scan').disabled = false;
}

if (typeof Office !== 'undefined' && Office.onReady) {
  Office.onReady(function() { initApp(); });
} else if (typeof Office !== 'undefined') {
  Office.initialize = function() { initApp(); };
}

// Fallback: force init after 5 seconds
setTimeout(function() {
  var btn = document.getElementById('btn-scan');
  if (btn && btn.disabled) {
    debugLog('Timeout fallback');
    initApp();
  }
}, 5000);

window.scanDocument = scanDocument;
window.clearHighlights = clearHighlights;
