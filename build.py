# -*- coding: utf-8 -*-
"""年間単元指導計画表 Webアプリ ビルドスクリプト (横軸=月, 縦軸=教科)"""
import json, pathlib

SRC = pathlib.Path(r"C:\Users\nkmr2\gakkyu-keiei")
with open(SRC / "curriculum_data.json", encoding="utf-8") as f:
    data = json.load(f)
json_blob = json.dumps(data, ensure_ascii=False)

PART1 = r"""<!DOCTYPE html>
<html lang="ja">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>年間単元指導計画表</title>
<style>
/* ============================================================
   Reset & Base
   ============================================================ */
*{margin:0;padding:0;box-sizing:border-box;}
html{font-size:11px;}
body{
  font-family:"游ゴシック","Yu Gothic","Hiragino Sans",sans-serif;
  background:#f0f2f5;color:#333;
}

/* ============================================================
   Toolbar
   ============================================================ */
#toolbar{
  position:sticky;top:0;z-index:200;
  display:flex;align-items:center;gap:10px;
  padding:8px 16px;
  background:linear-gradient(135deg,#4a5568,#2d3748);
  color:#fff;box-shadow:0 2px 8px rgba(0,0,0,.2);
  flex-wrap:wrap;
}
#toolbar label{font-size:12px;font-weight:700;}
#toolbar select{
  padding:4px 8px;border-radius:6px;border:none;
  font-size:13px;font-weight:700;cursor:pointer;
}
.tb-btn{
  padding:6px 14px;border:none;border-radius:6px;
  font-size:12px;font-weight:700;cursor:pointer;
  color:#fff;transition:all .15s;
}
.tb-btn:hover{opacity:.85;}
.tb-btn.blue{background:#4299e1;}
.tb-btn.green{background:#48bb78;}
.tb-btn.orange{background:#ed8936;}
.tb-btn.red{background:#e53e3e;}
.tb-btn.purple{background:#7c3aed;}
.tb-btn:disabled{opacity:.4;cursor:default;}
.tb-btn.active{box-shadow:0 0 0 3px #fc8181;transform:scale(1.05);}
.tb-btn.active-ext{box-shadow:0 0 0 3px #b794f4;transform:scale(1.05);}
.tb-sep{width:1px;height:24px;background:rgba(255,255,255,.25);margin:0 4px;}
.tb-spacer{flex:1;}
.tb-note{font-size:10px;opacity:.6;}

/* ============================================================
   Title Bar
   ============================================================ */
#title-bar{
  text-align:center;padding:10px 0 5px;
  font-size:16px;font-weight:700;color:#2d3748;
  background:#fff;border-bottom:1px solid #e2e8f0;
}

/* ============================================================
   Table Wrapper (contains table + SVG overlay)
   ============================================================ */
#table-wrap{
  overflow:auto;padding:10px;background:#f0f2f5;
}
#table-inner{
  position:relative;display:inline-block;min-width:100%;
}
#plan-table{
  border-collapse:collapse;width:100%;
  table-layout:fixed;
  background:#fff;
  box-shadow:0 1px 4px rgba(0,0,0,.08);
}
#plan-table th,#plan-table td{
  border:1px solid #cbd5e0;
  padding:3px 4px;
  vertical-align:top;
  font-size:11px;line-height:1.45;
  word-break:break-all;
}
/* ヘッダー（月ヘッダ行） */
#plan-table thead th{
  position:sticky;top:0;z-index:10;
  background:#4a5568;color:#fff;
  text-align:center;vertical-align:middle;
  font-size:11px;font-weight:700;
  padding:6px 2px;white-space:nowrap;
}
/* 教科名セル（行ヘッダ） */
td.subj-cell{
  text-align:center;font-weight:700;
  vertical-align:middle;white-space:nowrap;
  width:80px;min-width:80px;font-size:12px;
  background:#f7fafc;
  cursor:pointer;
}
td.subj-cell:hover{background:#edf2f7;}
/* データセル */
td.data-cell{cursor:pointer;transition:background .12s;padding-bottom:14px;}
td.data-cell:hover{filter:brightness(.95);}
/* 合計列 */
td.total-cell,th.total-hdr{
  text-align:center;font-weight:700;
  background:#edf2f7 !important;font-size:12px;
  width:40px;min-width:40px;
}
/* 合計非表示 */
.hide-total td.total-cell,
.hide-total th.total-hdr,
.hide-total colgroup col.col-total{
  display:none;
}
/* 学期色 — 列に適用 */
.sem1{background:#E8F0FE;}
.sem2{background:#FFF8E1;}
.sem3{background:#E8F5E9;}
.sem-summer{background:#F5F5F5;}

/* ============================================================
   Unit Items (in table cells)
   ============================================================ */
.unit-item{
  padding:1px 2px;margin:1px 0;border-radius:3px;
  transition:background .1s,box-shadow .1s;
  position:relative;
}
.unit-hours{color:#718096;font-size:10px;}

/* Arrow mode: units become clickable targets */
body.arrow-mode td.data-cell{cursor:default;}
body.arrow-mode .unit-item{
  cursor:crosshair;border-radius:3px;
}
body.arrow-mode .unit-item:hover{
  background:rgba(229,62,62,.12);
}
.unit-item.arrow-selected{
  background:rgba(229,62,62,.2) !important;
  box-shadow:0 0 0 2px #e53e3e;
}

/* ============================================================
   SVG Arrow Layer
   ============================================================ */
#arrow-svg{
  position:absolute;top:0;left:0;
  pointer-events:none;overflow:visible;
}
.arrow-group{
  pointer-events:auto;
}
.arrow-line{
  stroke:#333;stroke-width:1.5;fill:none;
  marker-end:url(#arrowhead);
  pointer-events:none;
}
.arrow-hit{
  stroke:transparent;stroke-width:18;fill:none;
  pointer-events:stroke;cursor:pointer;
}
.arrow-group:hover .arrow-line{
  stroke:#d32f2f;stroke-width:2.8;
}
.arrow-group:hover .arrow-hit{
  stroke:rgba(211,47,47,.08);
}

/* ============================================================
   Modal (Cell Editor)
   ============================================================ */
#modal-overlay{
  display:none;position:fixed;inset:0;z-index:500;
  background:rgba(0,0,0,.45);
  justify-content:center;align-items:center;
}
#modal-overlay.open{display:flex;}
#modal{
  background:#fff;border-radius:14px;
  width:520px;max-width:92vw;max-height:80vh;
  display:flex;flex-direction:column;
  box-shadow:0 20px 60px rgba(0,0,0,.3);
  animation:modalIn .2s ease;
}
@keyframes modalIn{from{transform:scale(.92);opacity:0}to{transform:scale(1);opacity:1}}
#modal-header{
  padding:16px 20px 12px;border-bottom:1px solid #e2e8f0;
}
#modal-header h3{font-size:15px;color:#2d3748;}
#modal-header .sub{font-size:12px;color:#718096;margin-top:2px;}
#modal-body{
  padding:12px 20px;overflow-y:auto;flex:1;
}
#modal-footer{
  padding:12px 20px;border-top:1px solid #e2e8f0;
  display:flex;justify-content:space-between;align-items:center;
}

/* Modal unit rows */
.mu-row{
  display:flex;align-items:center;gap:6px;
  padding:6px 0;border-bottom:1px solid #f0f0f0;
}
.mu-row:last-child{border-bottom:none;}
.mu-drag{
  cursor:grab;font-size:16px;color:#a0aec0;
  padding:0 6px 0 2px;user-select:none;line-height:1;
}
.mu-drag:active{cursor:grabbing;}
.mu-row.dragging{opacity:.25;background:#e2e8f0;}
.mu-row.drag-over{box-shadow:inset 0 -3px 0 0 #4299e1;}
.mu-name{
  flex:1;padding:4px 6px;border:1px solid #cbd5e0;
  border-radius:5px;font-size:12px;font-family:inherit;
}
.mu-hours{
  width:50px;padding:4px 6px;border:1px solid #cbd5e0;
  border-radius:5px;font-size:12px;text-align:center;
}
.mu-del{
  background:none;border:none;cursor:pointer;
  font-size:16px;color:#e53e3e;padding:0 4px;
}
.mu-del:hover{color:#c53030;}

.modal-btn{
  padding:7px 18px;border:none;border-radius:7px;
  font-size:13px;font-weight:700;cursor:pointer;transition:opacity .15s;
}
.modal-btn:hover{opacity:.85;}
.modal-btn.add{background:#ebf8ff;color:#2b6cb0;}
.modal-btn.save{background:#4299e1;color:#fff;}
.modal-btn.cancel{background:#e2e8f0;color:#4a5568;}

/* ============================================================
   Arrow Mode Banner
   ============================================================ */
#arrow-banner{
  display:none;position:fixed;bottom:20px;left:50%;transform:translateX(-50%);
  z-index:300;
  background:#e53e3e;color:#fff;
  padding:10px 24px;border-radius:30px;
  font-size:13px;font-weight:700;
  box-shadow:0 4px 16px rgba(229,62,62,.4);
  animation:bannerIn .25s ease;
}
#arrow-banner.show{display:block;}
@keyframes bannerIn{from{transform:translateX(-50%) translateY(20px);opacity:0}to{transform:translateX(-50%) translateY(0);opacity:1}}

/* ============================================================
   Cell Drag-and-Drop (table level)
   ============================================================ */
.unit-item[draggable="true"]{cursor:grab;}
.unit-item.cell-dragging{opacity:.25;}
body.arrow-mode .unit-item[draggable="true"]{cursor:crosshair;}
td.data-cell.drop-target{
  outline:2px dashed #4299e1;outline-offset:-2px;
  background:rgba(66,153,225,.08) !important;
}

/* ============================================================
   Subject Editor Modal (教科別一覧編集)
   ============================================================ */
#subj-overlay{
  display:none;position:fixed;inset:0;z-index:500;
  background:rgba(0,0,0,.45);
  justify-content:center;align-items:center;
}
#subj-overlay.open{display:flex;}
#subj-modal{
  background:#fff;border-radius:14px;
  width:700px;max-width:95vw;max-height:88vh;
  display:flex;flex-direction:column;
  box-shadow:0 20px 60px rgba(0,0,0,.3);
  animation:modalIn .2s ease;
}
#subj-header{
  padding:14px 20px 10px;border-bottom:1px solid #e2e8f0;
  display:flex;justify-content:space-between;align-items:center;flex-wrap:wrap;gap:8px;
}
#subj-header h3{font-size:15px;color:#2d3748;}
.shift-controls{display:flex;gap:6px;align-items:center;}
.shift-btn{
  padding:4px 12px;border:1px solid #cbd5e0;border-radius:5px;
  background:#f7fafc;font-size:11px;font-weight:700;cursor:pointer;
}
.shift-btn:hover{background:#edf2f7;}
#subj-body{padding:10px 20px;overflow-y:auto;flex:1;}
#subj-footer{
  padding:12px 20px;border-top:1px solid #e2e8f0;
  display:flex;justify-content:flex-end;gap:8px;
}
.se-month{margin-bottom:10px;border:1px solid #e2e8f0;border-radius:8px;overflow:hidden;}
.se-month-hdr{
  padding:4px 10px;font-size:12px;font-weight:700;
  background:#f7fafc;border-bottom:1px solid #e2e8f0;
  display:flex;justify-content:space-between;align-items:center;
}
.se-month-body{padding:2px 6px;min-height:24px;transition:background .15s;}
.se-month-body.drag-over-month{background:rgba(66,153,225,.1);}
.se-add-btn{
  font-size:11px;color:#4299e1;cursor:pointer;background:none;
  border:none;padding:2px 4px;font-weight:700;
}
.se-add-btn:hover{text-decoration:underline;}

/* ============================================================
   Extract Mode (抽出版)
   ============================================================ */
td.data-cell.jiko-row{background:#FFF3E0 !important;}
td.subj-cell.jiko-subj{background:#e65100 !important;color:#fff !important;}

/* ============================================================
   Print
   ============================================================ */
@media print{
  @page{size:A3 landscape;margin:6mm;}
  body{background:#fff !important;}
  body.arrow-mode{/* reset */}
  #toolbar,#modal-overlay,#subj-overlay,#arrow-banner{display:none !important;}
  #title-bar{border:none;padding:3px 0 2px;font-size:13px;margin:0;}
  #table-wrap{padding:0;overflow:visible;margin:0;}
  #table-inner{display:block;min-width:0;width:100%;}
  #plan-table{box-shadow:none;width:100%;table-layout:fixed;}
  #plan-table th,#plan-table td{font-size:8.5px;padding:2px 3px;line-height:1.35;}
  #plan-table thead th{
    position:static;
    background:#4a5568 !important;color:#fff !important;
    -webkit-print-color-adjust:exact;print-color-adjust:exact;
    padding:3px 2px;
  }
  td.subj-cell{
    background:#f7fafc !important;
    -webkit-print-color-adjust:exact;print-color-adjust:exact;
    font-size:9px;
  }
  td.data-cell{padding-bottom:2px;}
  .unit-item{margin:0;padding:1px 1px;}
  .unit-hours{font-size:7.5px;}
  .sem1{background:#E8F0FE !important;-webkit-print-color-adjust:exact;print-color-adjust:exact;}
  .sem2{background:#FFF8E1 !important;-webkit-print-color-adjust:exact;print-color-adjust:exact;}
  .sem3{background:#E8F5E9 !important;-webkit-print-color-adjust:exact;print-color-adjust:exact;}
  .sem-summer{background:#F5F5F5 !important;-webkit-print-color-adjust:exact;print-color-adjust:exact;}
  td.total-cell,th.total-hdr{background:#edf2f7 !important;-webkit-print-color-adjust:exact;print-color-adjust:exact;}
  .hide-total td.total-cell,.hide-total th.total-hdr,.hide-total colgroup col.col-total{display:none;}
  #arrow-svg{pointer-events:none;}
  .arrow-hit{display:none;}
  td.data-cell.jiko-row{background:#FFF3E0 !important;-webkit-print-color-adjust:exact;print-color-adjust:exact;}
  td.subj-cell.jiko-subj{background:#e65100 !important;color:#fff !important;-webkit-print-color-adjust:exact;print-color-adjust:exact;}
}
</style>
</head>
<body>

<!-- ===== Toolbar ===== -->
<div id="toolbar">
  <label for="grade-select">学年：</label>
  <select id="grade-select">
    <option value="1">1年</option><option value="2">2年</option>
    <option value="3">3年</option><option value="4">4年</option>
    <option value="5" selected>5年</option><option value="6">6年</option>
  </select>
  <div class="tb-sep"></div>
  <button class="tb-btn" id="btn-reset" style="background:#718096;font-size:11px">リセット</button>
  <div class="tb-sep"></div>
  <button class="tb-btn red" id="btn-arrow">矢印モード</button>
  <div class="tb-sep"></div>
  <label for="jiko-pos">自己実現活動：</label>
  <select id="jiko-pos">
    <option value="first">先頭</option>
    <option value="3">3番目</option>
    <option value="5">5番目</option>
    <option value="default" selected>標準</option>
    <option value="last">最後尾</option>
  </select>
  <div class="tb-spacer"></div>
  <button class="tb-btn" id="btn-total" style="background:#718096;font-size:11px">合計を表示</button>
  <button class="tb-btn purple" id="btn-extract">抽出版</button>
  <div class="tb-sep"></div>
  <button class="tb-btn blue"   id="btn-save">保存</button>
  <button class="tb-btn green"  id="btn-load">読み込み</button>
  <button class="tb-btn orange" id="btn-pdf" onclick="window.print()">PDF出力</button>
  <span class="tb-note">※ 印刷ダイアログで「PDFとして保存」</span>
</div>

<!-- ===== Title ===== -->
<div id="title-bar"></div>

<!-- ===== Table + SVG ===== -->
<div id="table-wrap">
  <div id="table-inner">
    <table id="plan-table">
      <thead id="plan-thead"></thead>
      <tbody id="plan-tbody"></tbody>
    </table>
    <svg id="arrow-svg" xmlns="http://www.w3.org/2000/svg">
      <defs>
        <marker id="arrowhead" markerWidth="8" markerHeight="6"
                refX="7" refY="3" orient="auto" markerUnits="strokeWidth">
          <polygon points="0 0, 8 3, 0 6" fill="#333"/>
        </marker>
      </defs>
    </svg>
  </div>
</div>

<!-- ===== Modal (Cell Editor) ===== -->
<div id="modal-overlay">
  <div id="modal">
    <div id="modal-header">
      <h3 id="modal-title">単元の編集</h3>
      <div class="sub" id="modal-sub"></div>
    </div>
    <div id="modal-body"></div>
    <div id="modal-footer">
      <button class="modal-btn add" id="modal-add">＋ 単元を追加</button>
      <div style="display:flex;gap:8px">
        <button class="modal-btn cancel" id="modal-cancel">キャンセル</button>
        <button class="modal-btn save"   id="modal-save">保存</button>
      </div>
    </div>
  </div>
</div>

<!-- ===== Subject Editor Modal ===== -->
<div id="subj-overlay">
  <div id="subj-modal">
    <div id="subj-header">
      <h3 id="subj-title">教科 ─ 年間一覧編集</h3>
      <div class="shift-controls">
        <button class="shift-btn" id="shift-back">&larr; 1月 前にシフト</button>
        <button class="shift-btn" id="shift-fwd">1月 後にシフト &rarr;</button>
      </div>
    </div>
    <div id="subj-body"></div>
    <div id="subj-footer">
      <button class="modal-btn cancel" id="subj-cancel">キャンセル</button>
      <button class="modal-btn save"   id="subj-save">保存</button>
    </div>
  </div>
</div>

<!-- ===== Arrow Mode Banner ===== -->
<div id="arrow-banner">矢印モード：始点の単元をクリックしてください</div>

<script>
// ================================================================
//  Curriculum Data (embedded)
// ================================================================
const CURRICULUM_DATA = """

PART2 = r""";

// ================================================================
//  Constants
// ================================================================
const MONTHS      = [4,5,6,7,8,9,10,11,12,1,2,3];
const MONTH_LABELS= ['4月','5月','6月','7月','8月','9月','10月','11月','12月','1月','2月','3月'];
const ALL_SUBJECTS= ['国語','算数','社会','理科','音楽','図画工作','体育','家庭','外国語','道徳・特活','健康','食育','自己実現活動','安全'];
const JIKO = '自己実現活動';

function semClass(m){
  if(m>=4&&m<=7)return 'sem1';
  if(m===8)return 'sem-summer';
  if(m>=9&&m<=12)return 'sem2';
  return 'sem3';
}
function monthLabel(m){return MONTH_LABELS[MONTHS.indexOf(m)];}

// ================================================================
//  State
// ================================================================
let state = { grade:'5', units:{}, arrows:[] };
let uidCounter = 0;
function uid(){return 'u'+(++uidCounter);}
let aidCounter = 0;
function aid(){return 'a'+(++aidCounter);}
let extractMode = false;
let extractVisibleUids = null;
let extractArrowIds    = null;

function cellKey(subj,m){return subj+'\t'+m;}

function initGrade(grade){
  state.grade = grade;
  state.units = {};
  state.arrows = [];
  uidCounter = 0;
  aidCounter = 0;
  const gd = CURRICULUM_DATA[grade];
  if(!gd) return;
  for(const [subj, units] of Object.entries(gd)){
    for(const u of units){
      const k = cellKey(subj, u.month);
      if(!state.units[k]) state.units[k]=[];
      state.units[k].push({ id:uid(), unit:u.unit, hours:u.hours });
    }
  }
}

function getSubjects(){
  const gd = CURRICULUM_DATA[state.grade];
  if(!gd) return [];
  let subjects = ALL_SUBJECTS.filter(s=> gd[s]!==undefined);
  const pos = document.getElementById('jiko-pos')?.value || 'default';
  if(subjects.includes(JIKO) && pos!=='default'){
    subjects = subjects.filter(s=>s!==JIKO);
    if(pos==='first') subjects.unshift(JIKO);
    else if(pos==='last') subjects.push(JIKO);
    else { const n=parseInt(pos); if(n) subjects.splice(Math.min(n-1,subjects.length),0,JIKO); }
  }
  return subjects;
}

// ================================================================
//  Extract Mode Helpers (Phase 6)
// ================================================================
function computeExtractData(){
  const jikoUids = new Set();
  for(const [key, units] of Object.entries(state.units)){
    if(key.startsWith(JIKO+'\t')) units.forEach(u=>jikoUids.add(u.id));
  }
  extractVisibleUids = new Set(jikoUids);
  extractArrowIds = new Set();
  for(const a of state.arrows){
    if(jikoUids.has(a.fromId)||jikoUids.has(a.toId)){
      extractVisibleUids.add(a.fromId);
      extractVisibleUids.add(a.toId);
      extractArrowIds.add(a.id);
    }
  }
}

function getExtractSubjects(){
  computeExtractData();
  const connSubjs = new Set([JIKO]);
  for(const [key, units] of Object.entries(state.units)){
    const subj = key.split('\t')[0];
    for(const u of units){
      if(extractVisibleUids.has(u.id)) connSubjs.add(subj);
    }
  }
  return getSubjects().filter(s=>connSubjs.has(s));
}

// ================================================================
//  DOM refs
// ================================================================
const gradeSelect  = document.getElementById('grade-select');
const titleBar     = document.getElementById('title-bar');
const thead        = document.getElementById('plan-thead');
const tbody        = document.getElementById('plan-tbody');
const tableInner   = document.getElementById('table-inner');
const arrowSvg     = document.getElementById('arrow-svg');
const modalOverlay = document.getElementById('modal-overlay');
const modalBody    = document.getElementById('modal-body');
const modalTitle   = document.getElementById('modal-title');
const modalSub     = document.getElementById('modal-sub');
const arrowBanner  = document.getElementById('arrow-banner');

// ================================================================
//  Table Rendering — 横軸=月, 縦軸=教科
// ================================================================
function renderTable(){
  const subjects = extractMode ? getExtractSubjects() : getSubjects();

  // Title
  if(extractMode){
    titleBar.textContent = `令和7年度　第${state.grade}学年　年間単元指導計画表【抽出版：自己実現活動】`;
  } else {
    titleBar.textContent = `令和7年度　第${state.grade}学年　年間単元指導計画表`;
  }

  const table = document.getElementById('plan-table');

  // colgroup: 教科名列(80px) + 月列(12個均等) + 合計列(40px)
  let cg = table.querySelector('colgroup');
  if(cg) cg.remove();
  cg = document.createElement('colgroup');
  const cSubj = document.createElement('col');
  cSubj.style.width='80px'; cg.appendChild(cSubj);
  for(let i=0;i<MONTHS.length;i++){
    const c=document.createElement('col');
    cg.appendChild(c);
  }
  const cTotal = document.createElement('col');
  cTotal.style.width='40px'; cTotal.className='col-total'; cg.appendChild(cTotal);
  table.prepend(cg);

  // thead: 「教科」+ 月 + 「合計」
  let hhtml = '<tr><th style="width:80px">教科</th>';
  for(let mi=0;mi<MONTHS.length;mi++){
    const cls = semClass(MONTHS[mi]);
    hhtml += `<th class="${cls}">${MONTH_LABELS[mi]}</th>`;
  }
  hhtml += '<th class="total-hdr" style="width:40px">合計</th></tr>';
  thead.innerHTML = hhtml;

  // tbody: 各教科が1行
  let bhtml = '';
  for(const subj of subjects){
    const isJiko = extractMode && subj===JIKO;
    const jikoSubjCls = isJiko ? ' jiko-subj' : '';
    bhtml += `<tr>`;
    bhtml += `<td class="subj-cell${jikoSubjCls}" data-subj="${esc(subj)}">${esc(subj)}</td>`;
    let totalHrs = 0;
    for(let mi=0;mi<MONTHS.length;mi++){
      const m = MONTHS[mi];
      const cls = semClass(m);
      const k = cellKey(subj, m);
      let units = state.units[k]||[];
      // Extract mode: filter units for non-JIKO rows
      if(extractMode && subj!==JIKO){
        units = units.filter(u=>extractVisibleUids.has(u.id));
      }
      let ch = '';
      let hrs = 0;
      for(const u of units){
        ch += `<div class="unit-item" data-uid="${u.id}" draggable="true">${esc(u.unit)}<span class="unit-hours">（${u.hours}）</span></div>`;
        hrs += u.hours;
      }
      totalHrs += hrs;
      const jikoCls = isJiko ? ' jiko-row' : '';
      bhtml += `<td class="data-cell ${cls}${jikoCls}" data-subj="${esc(subj)}" data-month="${m}">${ch}</td>`;
    }
    bhtml += `<td class="total-cell">${totalHrs}</td>`;
    bhtml += `</tr>`;
  }
  tbody.innerHTML = bhtml;

  requestAnimationFrame(renderArrows);
}

function esc(s){const d=document.createElement('div');d.textContent=s;return d.innerHTML;}

// ================================================================
//  Cell Editing (Modal)
// ================================================================
let editKey = null;
let editUnits = [];

function openEditor(subj, month){
  editKey = cellKey(subj, month);
  editUnits = (state.units[editKey]||[]).map(u=>({...u}));
  modalTitle.textContent = '単元の編集';
  modalSub.textContent = `${subj} ─ ${monthLabel(month)}`;
  renderModalBody();
  modalOverlay.classList.add('open');
}
function closeEditor(){
  modalOverlay.classList.remove('open');
  editKey=null; editUnits=[];
}

function renderModalBody(){
  if(!editUnits.length){
    modalBody.innerHTML='<div style="text-align:center;padding:20px;color:#a0aec0">単元がありません</div>';
    return;
  }
  let h='';
  editUnits.forEach((u,i)=>{
    h+=`<div class="mu-row" data-idx="${i}" draggable="true">
      <span class="mu-drag" title="ドラッグで並べ替え">&#9776;</span>
      <input class="mu-name" value="${escAttr(u.unit)}" data-field="unit">
      <input class="mu-hours" type="number" min="0" value="${u.hours}" data-field="hours">
      <span style="font-size:11px;color:#718096">時間</span>
      <button class="mu-del" onclick="delUnit(${i})">&#10005;</button>
    </div>`;
  });
  modalBody.innerHTML=h;
  setupDragReorder();
}

function escAttr(s){return s.replace(/"/g,'&quot;').replace(/</g,'&lt;');}

function syncEditorInputs(){
  const rows=modalBody.querySelectorAll('.mu-row');
  rows.forEach((row,i)=>{
    if(!editUnits[i])return;
    const nameEl=row.querySelector('[data-field="unit"]');
    const hrsEl=row.querySelector('[data-field="hours"]');
    if(nameEl) editUnits[i].unit=nameEl.value;
    if(hrsEl)  editUnits[i].hours=parseInt(hrsEl.value)||0;
  });
}

// ── ドラッグ&ドロップ並べ替え ──
let dragFromIdx=null;
function setupDragReorder(){
  const rows=modalBody.querySelectorAll('.mu-row');
  rows.forEach((row,idx)=>{
    row.addEventListener('dragstart',(e)=>{
      const handle=row.querySelector('.mu-drag');
      if(!handle||!handle.contains(e.target)){e.preventDefault();return;}
      dragFromIdx=idx;
      row.classList.add('dragging');
      e.dataTransfer.effectAllowed='move';
      e.dataTransfer.setData('text/plain',String(idx));
    });
    row.addEventListener('dragend',()=>{
      row.classList.remove('dragging');
      rows.forEach(r=>r.classList.remove('drag-over'));
      dragFromIdx=null;
    });
    row.addEventListener('dragover',(e)=>{
      e.preventDefault();
      e.dataTransfer.dropEffect='move';
      rows.forEach(r=>r.classList.remove('drag-over'));
      if(dragFromIdx!==null&&idx!==dragFromIdx) row.classList.add('drag-over');
    });
    row.addEventListener('dragleave',()=>{
      row.classList.remove('drag-over');
    });
    row.addEventListener('drop',(e)=>{
      e.preventDefault();
      rows.forEach(r=>r.classList.remove('drag-over'));
      if(dragFromIdx===null||dragFromIdx===idx)return;
      syncEditorInputs();
      const item=editUnits.splice(dragFromIdx,1)[0];
      editUnits.splice(idx,0,item);
      dragFromIdx=null;
      renderModalBody();
    });
  });
}
window.delUnit=function(idx){
  syncEditorInputs();
  editUnits.splice(idx,1);
  renderModalBody();
};

document.getElementById('modal-add').addEventListener('click',()=>{
  syncEditorInputs();
  editUnits.push({id:uid(), unit:'', hours:1});
  renderModalBody();
  const rows=modalBody.querySelectorAll('.mu-row');
  if(rows.length) rows[rows.length-1].querySelector('.mu-name').focus();
});

document.getElementById('modal-save').addEventListener('click',()=>{
  syncEditorInputs();
  const newUnits = editUnits.filter(u=>u.unit.trim());
  const oldIds = new Set((state.units[editKey]||[]).map(u=>u.id));
  const newIds = new Set(newUnits.map(u=>u.id));
  const removed = [...oldIds].filter(id=>!newIds.has(id));
  if(removed.length){
    state.arrows = state.arrows.filter(a=>
      !removed.includes(a.fromId) && !removed.includes(a.toId)
    );
  }
  state.units[editKey] = newUnits;
  closeEditor();
  renderTable();
});

document.getElementById('modal-cancel').addEventListener('click', closeEditor);
modalOverlay.addEventListener('click',(e)=>{
  if(e.target===modalOverlay) closeEditor();
});

// ================================================================
//  Arrow Mode
// ================================================================
let arrowMode = false;
let arrowStartId = null;

const btnArrow = document.getElementById('btn-arrow');
btnArrow.addEventListener('click', toggleArrowMode);

function toggleArrowMode(){
  arrowMode = !arrowMode;
  document.body.classList.toggle('arrow-mode', arrowMode);
  btnArrow.classList.toggle('active', arrowMode);
  arrowStartId = null;
  clearArrowSelection();
  updateBanner();
}

function updateBanner(){
  if(!arrowMode){
    arrowBanner.classList.remove('show');
    return;
  }
  arrowBanner.classList.add('show');
  arrowBanner.textContent = arrowStartId
    ? '矢印モード：終点の単元をクリックしてください（Escでキャンセル）'
    : '矢印モード：始点の単元をクリックしてください（Escで解除）';
}

function clearArrowSelection(){
  document.querySelectorAll('.unit-item.arrow-selected').forEach(el=>{
    el.classList.remove('arrow-selected');
  });
}

function handleArrowClick(unitId){
  if(!arrowStartId){
    arrowStartId = unitId;
    const el = document.querySelector(`[data-uid="${unitId}"]`);
    if(el) el.classList.add('arrow-selected');
    updateBanner();
  } else {
    if(unitId === arrowStartId){
      arrowStartId = null;
      clearArrowSelection();
      updateBanner();
      return;
    }
    const dup = state.arrows.some(a=>a.fromId===arrowStartId && a.toId===unitId);
    if(!dup){
      state.arrows.push({ id:aid(), fromId:arrowStartId, toId:unitId });
    }
    arrowStartId = null;
    clearArrowSelection();
    updateBanner();
    if(extractMode) renderTable(); else renderArrows();
  }
}

// Esc key
document.addEventListener('keydown',(e)=>{
  if(e.key==='Escape'){
    if(arrowStartId){
      arrowStartId=null;
      clearArrowSelection();
      updateBanner();
    } else if(arrowMode){
      toggleArrowMode();
    }
    if(modalOverlay.classList.contains('open')){
      closeEditor();
    }
    if(document.getElementById('subj-overlay').classList.contains('open')){
      closeSubjectEditor();
    }
  }
});

// ================================================================
//  SVG Arrow Rendering
// ================================================================
function renderArrows(){
  arrowSvg.querySelectorAll('.arrow-group').forEach(g=>g.remove());
  const w=tableInner.scrollWidth, h=tableInner.scrollHeight;
  arrowSvg.setAttribute('width',w);
  arrowSvg.setAttribute('height',h);
  arrowSvg.style.width=w+'px';
  arrowSvg.style.height=h+'px';
  const pr=tableInner.getBoundingClientRect();
  const ns='http://www.w3.org/2000/svg';

  for(const arrow of state.arrows){
    if(extractMode && extractArrowIds && !extractArrowIds.has(arrow.id)) continue;
    const fromEl=document.querySelector(`[data-uid="${arrow.fromId}"]`);
    const toEl=document.querySelector(`[data-uid="${arrow.toId}"]`);
    if(!fromEl||!toEl) continue;
    const d=computeArrowPath(fromEl,toEl,pr);
    if(!d) continue;

    const g=document.createElementNS(ns,'g');
    g.setAttribute('class','arrow-group');
    g.setAttribute('data-aid',arrow.id);
    const hit=document.createElementNS(ns,'path');
    hit.setAttribute('d',d);hit.setAttribute('class','arrow-hit');
    g.appendChild(hit);
    const line=document.createElementNS(ns,'path');
    line.setAttribute('d',d);line.setAttribute('class','arrow-line');
    g.appendChild(line);

    const delHandler=(e)=>{
      e.preventDefault();e.stopPropagation();
      if(arrowMode) return;
      if(confirm('この矢印を削除しますか？')){
        state.arrows=state.arrows.filter(a=>a.id!==arrow.id);
        if(extractMode) renderTable(); else renderArrows();
      }
    };
    hit.addEventListener('click',delHandler);
    hit.addEventListener('contextmenu',delHandler);
    arrowSvg.appendChild(g);
  }
}

// 矢印パス計算 — 横軸=月なので同じ行=同教科が主要ルート
function computeArrowPath(fromEl,toEl,pr){
  const fR=fromEl.getBoundingClientRect();
  const tR=toEl.getBoundingClientRect();
  const fCell=fromEl.closest('td.data-cell');
  const tCell=toEl.closest('td.data-cell');
  if(!fCell||!tCell) return null;
  const fCR=fCell.getBoundingClientRect();
  const tCR=tCell.getBoundingClientRect();

  // 同じ行（同教科）かどうか判定
  const sameRow = Math.abs(fCR.top - tCR.top) < 5;

  if(sameRow){
    // ── 同一教科行：下端パディング領域を通るS字カーブ ──
    const routeY = fCR.bottom - pr.top - 4;
    const goRight = fR.left < tR.left;
    // 始点：ユニットの下端中央付近
    const x1 = fR.left - pr.left + fR.width/2;
    const y1 = fR.bottom - pr.top + 1;
    // 終点：ユニットの下端中央付近
    const x2 = tR.left - pr.left + tR.width/2;
    const y2 = tR.bottom - pr.top + 1;
    const dx = x2 - x1;
    if(Math.abs(dx) < 4) return `M${x1},${y1} L${x2},${y2}`;
    return `M${x1},${y1} C${x1+dx*0.25},${routeY} ${x1+dx*0.75},${routeY} ${x2},${y2}`;
  } else {
    // ── 異なる教科行：セル端から出入りするS字カーブ ──
    const goDown = fCR.top < tCR.top;
    // 始点：下向きなら下端、上向きなら上端
    const x1 = fR.left - pr.left + fR.width/2;
    const y1 = goDown ? (fR.bottom - pr.top + 2) : (fR.top - pr.top - 2);
    // 終点：下向きなら上端、上向きなら下端
    const x2 = tR.left - pr.left + tR.width/2;
    const y2 = goDown ? (tR.top - pr.top - 2) : (tR.bottom - pr.top + 2);
    const midY = (y1 + y2) / 2;
    const offsetX = Math.min(Math.abs(x2 - x1) * 0.3, 40);
    const curveDir = (x2 > x1) ? -1 : 1;
    return `M${x1},${y1} C${x1+curveDir*offsetX},${midY} ${x2-curveDir*offsetX},${midY} ${x2},${y2}`;
  }
}

// Re-render arrows on resize
let resizeTimer;
window.addEventListener('resize',()=>{
  clearTimeout(resizeTimer);
  resizeTimer=setTimeout(renderArrows,150);
});

// ================================================================
//  Table Click Delegation
// ================================================================
document.getElementById('plan-table').addEventListener('click',(e)=>{
  const unitEl = e.target.closest('.unit-item');
  const cellEl = e.target.closest('td.data-cell');
  const subjEl = e.target.closest('td.subj-cell');

  if(arrowMode){
    if(unitEl && unitEl.dataset.uid){
      handleArrowClick(unitEl.dataset.uid);
    }
    return;
  }

  // Subject cell click → open subject editor
  if(subjEl && subjEl.dataset.subj){
    openSubjectEditor(subjEl.dataset.subj);
    return;
  }

  // Data cell click → open cell editor
  if(cellEl){
    const subj  = cellEl.dataset.subj;
    const month = parseInt(cellEl.dataset.month);
    if(subj && !isNaN(month)){
      openEditor(subj, month);
    }
  }
});

// ================================================================
//  Grade Change
// ================================================================
gradeSelect.addEventListener('change',()=>{
  if(arrowMode) toggleArrowMode();
  initGrade(gradeSelect.value);
  renderTable();
});
// 自己実現活動の行位置変更
document.getElementById('jiko-pos').addEventListener('change',()=>{
  renderTable();
});

// ================================================================
//  Subject Editor (教科別一覧編集)
// ================================================================
let subjectEditSubj = null;
let subjectEditData = {};
const subjOverlay = document.getElementById('subj-overlay');
const subjBody    = document.getElementById('subj-body');

function openSubjectEditor(subj){
  subjectEditSubj=subj;
  subjectEditData={};
  for(const m of MONTHS){
    const k=cellKey(subj,m);
    subjectEditData[m]=(state.units[k]||[]).map(u=>({...u}));
  }
  document.getElementById('subj-title').textContent=`${subj} ─ 年間一覧編集`;
  renderSubjectBody();
  subjOverlay.classList.add('open');
}
function closeSubjectEditor(){
  subjOverlay.classList.remove('open');
  subjectEditSubj=null;subjectEditData={};
}

function renderSubjectBody(){
  let h='';
  MONTHS.forEach((m,mi)=>{
    const units=subjectEditData[m]||[];
    const hrs=units.reduce((s,u)=>s+u.hours,0);
    h+=`<div class="se-month" data-semonth="${m}">`;
    h+=`<div class="se-month-hdr"><span>${MONTH_LABELS[mi]}</span><span style="color:#a0aec0;font-size:10px">${hrs}時間</span></div>`;
    h+=`<div class="se-month-body" data-semonth="${m}">`;
    units.forEach((u,i)=>{
      h+=`<div class="mu-row" data-semonth="${m}" data-idx="${i}" draggable="true">
        <span class="mu-drag">&#9776;</span>
        <input class="mu-name" value="${escAttr(u.unit)}" data-field="unit">
        <input class="mu-hours" type="number" min="0" value="${u.hours}" data-field="hours" style="width:50px">
        <span style="font-size:11px;color:#718096">時間</span>
        <button class="mu-del" data-sm="${m}" data-si="${i}">&#10005;</button>
      </div>`;
    });
    h+=`<button class="se-add-btn" data-sm="${m}">＋ 追加</button>`;
    h+='</div></div>';
  });
  subjBody.innerHTML=h;
  setupSubjDrag();
  subjBody.querySelectorAll('.mu-del').forEach(btn=>{
    btn.addEventListener('click',()=>{
      syncSubjInputs();
      const sm=parseInt(btn.dataset.sm), si=parseInt(btn.dataset.si);
      (subjectEditData[sm]||[]).splice(si,1);
      renderSubjectBody();
    });
  });
  subjBody.querySelectorAll('.se-add-btn').forEach(btn=>{
    btn.addEventListener('click',()=>{
      syncSubjInputs();
      const sm=parseInt(btn.dataset.sm);
      if(!subjectEditData[sm])subjectEditData[sm]=[];
      subjectEditData[sm].push({id:uid(),unit:'',hours:1});
      renderSubjectBody();
      const sect=subjBody.querySelector(`.se-month-body[data-semonth="${sm}"]`);
      const rows=sect?.querySelectorAll('.mu-row');
      if(rows&&rows.length) rows[rows.length-1].querySelector('.mu-name').focus();
    });
  });
}

function syncSubjInputs(){
  subjBody.querySelectorAll('.se-month').forEach(sect=>{
    const m=parseInt(sect.dataset.semonth);
    const units=subjectEditData[m]||[];
    sect.querySelectorAll('.mu-row').forEach((row,i)=>{
      if(!units[i])return;
      const n=row.querySelector('[data-field="unit"]');
      const h=row.querySelector('[data-field="hours"]');
      if(n)units[i].unit=n.value;
      if(h)units[i].hours=parseInt(h.value)||0;
    });
  });
}

// Drag between months in subject editor
function setupSubjDrag(){
  let dMonth=null,dIdx=null;
  const rows=subjBody.querySelectorAll('.mu-row');
  rows.forEach(row=>{
    row.addEventListener('dragstart',(e)=>{
      const handle=row.querySelector('.mu-drag');
      if(!handle||!handle.contains(e.target)){e.preventDefault();return;}
      dMonth=parseInt(row.dataset.semonth);
      dIdx=parseInt(row.dataset.idx);
      row.classList.add('dragging');
      e.dataTransfer.effectAllowed='move';
      e.dataTransfer.setData('text/plain','subj');
    });
    row.addEventListener('dragend',()=>{
      row.classList.remove('dragging');
      subjBody.querySelectorAll('.drag-over-month').forEach(el=>el.classList.remove('drag-over-month'));
      dMonth=null;dIdx=null;
    });
  });
  const zones=subjBody.querySelectorAll('.se-month-body');
  zones.forEach(zone=>{
    zone.addEventListener('dragover',(e)=>{
      if(dMonth===null)return;
      e.preventDefault();e.dataTransfer.dropEffect='move';
      zones.forEach(z=>z.classList.remove('drag-over-month'));
      zone.classList.add('drag-over-month');
    });
    zone.addEventListener('dragleave',()=>zone.classList.remove('drag-over-month'));
    zone.addEventListener('drop',(e)=>{
      e.preventDefault();zone.classList.remove('drag-over-month');
      if(dMonth===null)return;
      const tm=parseInt(zone.dataset.semonth);
      syncSubjInputs();
      const src=subjectEditData[dMonth]||[];
      if(dIdx<0||dIdx>=src.length)return;
      const item=src.splice(dIdx,1)[0];
      if(!subjectEditData[tm])subjectEditData[tm]=[];
      subjectEditData[tm].push(item);
      dMonth=null;dIdx=null;
      renderSubjectBody();
    });
  });
}

// Month shift
document.getElementById('shift-back').addEventListener('click',()=>shiftSubjMonths(-1));
document.getElementById('shift-fwd').addEventListener('click',()=>shiftSubjMonths(1));
function shiftSubjMonths(dir){
  syncSubjInputs();
  const nd={};
  MONTHS.forEach(m=>nd[m]=[]);
  for(let i=0;i<MONTHS.length;i++){
    const ti=i+dir;
    if(ti>=0&&ti<MONTHS.length) nd[MONTHS[ti]]=subjectEditData[MONTHS[i]]||[];
  }
  subjectEditData=nd;
  renderSubjectBody();
}

// Save / Cancel
document.getElementById('subj-save').addEventListener('click',()=>{
  syncSubjInputs();
  const oldIds=new Set(), newIds=new Set();
  for(const m of MONTHS){
    (state.units[cellKey(subjectEditSubj,m)]||[]).forEach(u=>oldIds.add(u.id));
    (subjectEditData[m]||[]).forEach(u=>{if(u.unit.trim())newIds.add(u.id);});
  }
  const removed=[...oldIds].filter(id=>!newIds.has(id));
  if(removed.length){
    state.arrows=state.arrows.filter(a=>!removed.includes(a.fromId)&&!removed.includes(a.toId));
  }
  for(const m of MONTHS){
    state.units[cellKey(subjectEditSubj,m)]=(subjectEditData[m]||[]).filter(u=>u.unit.trim());
  }
  closeSubjectEditor();
  renderTable();
});
document.getElementById('subj-cancel').addEventListener('click',closeSubjectEditor);
subjOverlay.addEventListener('click',(e)=>{if(e.target===subjOverlay)closeSubjectEditor();});

// ================================================================
//  Cell-to-Cell Drag (表上のセル間ドラッグ移動) — 同教科(同行)内
// ================================================================
(function(){
  const tbl=document.getElementById('plan-table');
  let dragUid=null,dragSrcSubj=null,dragSrcMonth=null;

  tbl.addEventListener('dragstart',(e)=>{
    if(arrowMode){e.preventDefault();return;}
    const ui=e.target.closest('.unit-item');
    if(!ui||!ui.dataset.uid)return;
    const cell=ui.closest('td.data-cell');
    if(!cell)return;
    dragUid=ui.dataset.uid;
    dragSrcSubj=cell.dataset.subj;
    dragSrcMonth=parseInt(cell.dataset.month);
    ui.classList.add('cell-dragging');
    e.dataTransfer.effectAllowed='move';
    e.dataTransfer.setData('text/plain','cell');
  });

  tbl.addEventListener('dragend',(e)=>{
    tbl.querySelectorAll('.cell-dragging').forEach(el=>el.classList.remove('cell-dragging'));
    tbl.querySelectorAll('.drop-target').forEach(el=>el.classList.remove('drop-target'));
    dragUid=null;
  });

  tbl.addEventListener('dragover',(e)=>{
    if(!dragUid)return;
    const cell=e.target.closest('td.data-cell');
    if(!cell)return;
    // Same subject only
    if(cell.dataset.subj!==dragSrcSubj)return;
    e.preventDefault();e.dataTransfer.dropEffect='move';
    tbl.querySelectorAll('.drop-target').forEach(el=>el.classList.remove('drop-target'));
    cell.classList.add('drop-target');
  });

  tbl.addEventListener('dragleave',(e)=>{
    const cell=e.target.closest('td.data-cell');
    if(cell)cell.classList.remove('drop-target');
  });

  tbl.addEventListener('drop',(e)=>{
    if(!dragUid)return;
    e.preventDefault();
    tbl.querySelectorAll('.drop-target').forEach(el=>el.classList.remove('drop-target'));
    const cell=e.target.closest('td.data-cell');
    if(!cell)return;
    const tSubj=cell.dataset.subj;
    const tMonth=parseInt(cell.dataset.month);
    if(tSubj!==dragSrcSubj)return;
    if(tMonth===dragSrcMonth){dragUid=null;return;}
    const srcKey=cellKey(dragSrcSubj,dragSrcMonth);
    const tgtKey=cellKey(tSubj,tMonth);
    const srcArr=state.units[srcKey]||[];
    const idx=srcArr.findIndex(u=>u.id===dragUid);
    if(idx===-1){dragUid=null;return;}
    const unit=srcArr.splice(idx,1)[0];
    if(!state.units[tgtKey])state.units[tgtKey]=[];
    state.units[tgtKey].push(unit);
    dragUid=null;
    renderTable();
  });
})();

// ================================================================
//  Reset (初期データに戻す)
// ================================================================
document.getElementById('btn-reset').addEventListener('click',()=>{
  if(!confirm(`第${state.grade}学年のデータを初期状態に戻します。\n編集内容と矢印がすべて削除されます。\n\nよろしいですか？`))return;
  initGrade(state.grade);
  if(extractMode) toggleExtract();
  renderTable();
});

// ================================================================
//  Save / Load (Phase 4)
// ================================================================
document.getElementById('btn-save').addEventListener('click',()=>{
  const saveData = {
    version: 1,
    grade: state.grade,
    savedAt: new Date().toISOString(),
    jikoPos: document.getElementById('jiko-pos').value,
    units: state.units,
    arrows: state.arrows
  };
  const json = JSON.stringify(saveData, null, 2);
  const blob = new Blob([json], {type:'application/json'});
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  const d = new Date().toISOString().slice(0,10);
  a.href = url;
  a.download = `学級経営案_${state.grade}年_${d}.json`;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
});

document.getElementById('btn-load').addEventListener('click',()=>{
  if(!confirm('現在の編集内容が失われますが、よろしいですか？')) return;
  const input = document.createElement('input');
  input.type = 'file';
  input.accept = '.json';
  input.addEventListener('change',(e)=>{
    const file = e.target.files[0];
    if(!file) return;
    const reader = new FileReader();
    reader.onload = (ev)=>{
      try {
        const data = JSON.parse(ev.target.result);
        if(!data.grade||!data.units) throw new Error('Invalid');
        state.grade  = String(data.grade);
        state.units  = data.units;
        state.arrows = data.arrows || [];
        gradeSelect.value = state.grade;
        if(data.jikoPos) document.getElementById('jiko-pos').value = data.jikoPos;
        let maxU=0, maxA=0;
        for(const units of Object.values(state.units)){
          for(const u of units){ const n=parseInt(u.id.replace('u',''))||0; if(n>maxU)maxU=n; }
        }
        for(const a of state.arrows){ const n=parseInt(a.id.replace('a',''))||0; if(n>maxA)maxA=n; }
        uidCounter = maxU;
        aidCounter = maxA;
        if(extractMode) toggleExtract();
        renderTable();
      } catch(err){
        alert('ファイルの読み込みに失敗しました。\n正しいJSONファイルを選択してください。\n\n' + err.message);
      }
    };
    reader.readAsText(file);
  });
  input.click();
});

// ================================================================
//  Extract Mode Toggle (Phase 6)
// ================================================================
const btnExtract = document.getElementById('btn-extract');
btnExtract.addEventListener('click', toggleExtract);

function toggleExtract(){
  extractMode = !extractMode;
  btnExtract.classList.toggle('active-ext', extractMode);
  btnExtract.textContent = extractMode ? '全体表示に戻す' : '抽出版';
  extractVisibleUids = null;
  extractArrowIds = null;
  renderTable();
}

// ================================================================
//  Total Column Toggle (合計列の表示/非表示)
// ================================================================
const btnTotal = document.getElementById('btn-total');
const planTable = document.getElementById('plan-table');
// 初期状態は非表示
planTable.classList.add('hide-total');

btnTotal.addEventListener('click',()=>{
  const hidden = planTable.classList.toggle('hide-total');
  btnTotal.textContent = hidden ? '合計を表示' : '合計を非表示';
});

// ================================================================
//  Init
// ================================================================
initGrade(gradeSelect.value);
renderTable();
</script>
</body>
</html>
"""

html = PART1 + json_blob + PART2
out = SRC / "plan.html"
out.write_text(html, encoding="utf-8")
sz = out.stat().st_size
print(f"Generated: {out.name}  ({sz:,} bytes)")
