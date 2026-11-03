/**
 * 生物育成シミュレーター／いのちの1万年すごろく 提出記録用 Apps Script
 * 使い方：
 *  1. 記録用のGoogleスプレッドシートの「拡張機能」→「Apps Script」に、このコードを丸ごと貼り付けて保存
 *  2. 「デプロイ」→「デプロイを管理」→ ✏️ →「バージョン：新バージョン」→「デプロイ」（URLは変わらない）
 *  3. クラスの班ランキング：ウェブアプリのURLの最後に ?page=ranking&kumi=2 を付けて開く（2は組の番号）
 * © Irie Kanji
 */
const SHEET_NAME = '記録';
const HEADERS = ['送信日時','種類','組','番号','氏名','題材','挑戦回数','ゲームオーバー回数',
  '設定（品種・方法など）','育成の目標','計画の理由','到達ターン','結果の概要',
  '評価1','評価2','評価3','評価4','合計（20点満点）',
  'ゲームオーバーの原因','ゲームオーバーの振り返り（30〜50字）',
  '振り返り1 工夫した管理と理由','振り返り2 立場ごとの評価','振り返り3 品種についての考え','振り返り4 次の改善'];
const SG_SHEET = '1万年すごろく';
const SG_HEADERS = ['送信日時','組','班','人数','コース','到達した時代','平均資産（万円）','班の得点（万円）','買われた施設数','個人の結果'];

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  let count = 0, isSg = false;
  try {
    const raw = (e.parameter && e.parameter.payload) ? e.parameter.payload : e.postData.contents;
    const data = JSON.parse(raw);
    const list = Array.isArray(data) ? data : [data];
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    list.forEach(d => {
      if (d.kind === 'すごろく') { writeSugoroku(ss, d); isSg = true; count++; return; }
      const sh = getSheet(ss, SHEET_NAME, HEADERS);
      const ax = d.axes || [];
      sh.appendRow([new Date(), d.kind, d.kumi, d.ban, d.name, d.sim, d.play, d.overCount,
        d.setting, d.goal, d.why, d.turn, d.summary,
        ax[0] || '', ax[1] || '', ax[2] || '', ax[3] || '', d.total || '',
        d.overCause, d.overReflect, d.q1 || '', d.q2 || '', d.q3 || '', d.q4 || '']);
      count++;
    });
  } catch (err) {
    return page('⚠️ 提出できませんでした', 'アプリにもどって、もう一度送ってください。<br><small>' + err + '</small>');
  } finally {
    lock.releaseLock();
  }
  if (isSg) return page('✅ 班の結果を提出できました', 'このタブを閉じて、アプリにもどってください。<br>クラスの班ランキングに反映されます。');
  return page('✅ 提出できました', (count > 1 ? count + '件の記録を受け付けました。<br>' : '') + 'このタブを閉じて、アプリにもどってください。');
}

function getSheet(ss, name, headers) {
  const sh = ss.getSheetByName(name) || ss.insertSheet(name);
  if (sh.getLastRow() === 0) {
    sh.appendRow(headers);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#e6efe9');
  }
  return sh;
}

function writeSugoroku(ss, d) {
  const sh = getSheet(ss, SG_SHEET, SG_HEADERS);
  sh.appendRow([new Date(), d.kumi, d.han, d.count, d.months, d.level, d.avgAssets, d.groupScore, d.props, d.members]);
}

/** ランキング画面から呼ばれる：組ごとに、各班の最新の結果を得点順に返す */
function getRanking(kumi) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SG_SHEET);
  if (!sh || sh.getLastRow() < 2) return [];
  const v = sh.getRange(2, 1, sh.getLastRow() - 1, SG_HEADERS.length).getValues();
  const latest = {};
  v.forEach(r => {
    if (kumi && String(r[1]) !== String(kumi)) return;
    const key = r[1] + '-' + r[2];
    latest[key] = { kumi: String(r[1]), han: String(r[2]), count: r[3], months: r[4], level: r[5], avg: r[6], score: r[7], props: r[8], members: String(r[9]),
      time: Utilities.formatDate(new Date(r[0]), 'Asia/Tokyo', 'M/d HH:mm') };
  });
  return Object.values(latest).sort((a, b) => b.score - a.score);
}

function doGet(e) {
  if (e && e.parameter && e.parameter.page === 'ranking') {
    const kumi = String(e.parameter.kumi || '').replace(/[^0-9]/g, '');
    return HtmlService.createHtmlOutput(rankingHtml(kumi)).setTitle('クラス班ランキング')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }
  return page('生物育成の提出先', '提出先は動作しています。');
}

function rankingHtml(kumi) {
  return `<!DOCTYPE html><html><head><meta charset="UTF-8"><style>
body{margin:0;font-family:"Hiragino Maru Gothic ProN","Noto Sans JP",Meiryo,sans-serif;background:linear-gradient(#bfe6ff,#e5f5dc);color:#22313a;min-height:100vh}
h1{text-align:center;margin:0;padding:14px;background:linear-gradient(90deg,#1f8a57,#2a9fc9);color:#fff;font-size:1.7rem}
.wrap{max-width:1100px;margin:0 auto;padding:16px}
.podium{display:flex;justify-content:center;align-items:flex-end;gap:16px;margin:10px 0 22px}
.pd{background:#fff;border-radius:18px 18px 0 0;width:240px;text-align:center;padding:12px;box-shadow:0 6px 0 rgba(0,0,0,.12);animation:up .8s both}
.pd .m{font-size:3rem}.pd .h{font-size:1.6rem;font-weight:800}.pd .s{font-size:1.4rem;font-weight:800;color:#1f7a4d}
.p1{height:230px;border:5px solid #f0c64a}.p2{height:190px;border:5px solid #b8c4cc;animation-delay:.2s}.p3{height:160px;border:5px solid #d9a37a;animation-delay:.4s}
table{width:100%;border-collapse:collapse;background:#fff;border-radius:14px;overflow:hidden;box-shadow:0 6px 0 rgba(0,0,0,.1)}
th,td{padding:9px 12px;border-bottom:1px solid #e2e8ee;text-align:left;font-size:1.05rem}th{background:#eef5f1}
td.sc{font-weight:800;color:#1f7a4d;font-size:1.2rem}.small{font-size:.85rem;color:#5a6b74}
.empty{text-align:center;font-size:1.3rem;padding:40px;background:#fff;border-radius:14px}
@keyframes up{from{transform:translateY(60px);opacity:0}to{transform:none;opacity:1}}
</style></head><body><h1>🏅 ${kumi ? kumi + '組 ' : ''}班ランキング 〜いのちの1万年すごろく〜</h1><div class="wrap"><div id="v" class="empty">読み込み中…</div>
<p class="small" style="text-align:center">10秒ごとに自動で更新されます。班の得点＝班員の平均資産＋時代ボーナス（令和まで発展すると🏆）</p></div>
<script>
function fmt(v){v=Math.round(v);var neg=v<0;v=Math.abs(v);var o=Math.floor(v/10000),m=v%10000;return (neg?'−':'')+(o?o+'億':'')+(m||!o?m.toLocaleString()+'万':'')+'円';}
function esc(s){return String(s).replace(/[&<>"]/g,function(c){return {'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c];});}
var last='';
function render(list){var key=JSON.stringify(list);if(key===last)return;last=key;var el=document.getElementById('v');
 if(!list.length){el.className='empty';el.innerHTML='まだ提出された班はありません。';return;}
 el.className='';var med=['🥇','🥈','🥉'],cls=['p1','p2','p3'],order=[1,0,2];
 var pod='<div class="podium">'+order.filter(function(i){return list[i];}).map(function(i){var r=list[i];return '<div class="pd '+cls[i]+'"><div class="m">'+med[i]+'</div><div class="h">'+esc(r.kumi)+'組 '+esc(r.han)+'班</div><div class="s">'+fmt(r.score)+'</div><div class="small">'+esc(r.level)+(String(r.level).indexOf('令和')>=0?' 🏆':'')+'</div></div>';}).join('')+'</div>';
 var tb='<table><tr><th>順位</th><th>班</th><th>班の得点</th><th>到達した時代</th><th>施設</th><th>メンバー（個人の順位）</th><th>提出</th></tr>'+list.map(function(r,i){return '<tr><td><b>'+(med[i]||'')+' '+(i+1)+'位</b></td><td>'+esc(r.kumi)+'組 '+esc(r.han)+'班</td><td class="sc">'+fmt(r.score)+'</td><td>'+esc(r.level)+(String(r.level).indexOf('令和')>=0?' 🏆':'')+'</td><td>'+r.props+'件</td><td class="small">'+esc(r.members)+'</td><td class="small">'+esc(r.time)+'</td></tr>';}).join('')+'</table>';
 el.innerHTML=pod+tb;}
function load(){google.script.run.withSuccessHandler(render).getRanking('${kumi}');}
load();setInterval(load,10000);
</script></body></html>`;
}

function page(title, msg) {
  return HtmlService.createHtmlOutput(
    '<div style="font-family:sans-serif;text-align:center;padding:48px 16px">' +
    '<h1 style="font-size:28px">' + title + '</h1><p style="font-size:18px;line-height:1.8">' + msg + '</p></div>'
  ).setTitle(title).addMetaTag('viewport', 'width=device-width, initial-scale=1');
}
