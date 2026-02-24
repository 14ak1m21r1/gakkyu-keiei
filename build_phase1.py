# -*- coding: utf-8 -*-
"""フェーズ1: 年間単元指導計画表 Webアプリ ビルドスクリプト"""
import json, pathlib

SRC = pathlib.Path(r"C:\Users\nkmr2\gakkyu-keiei")
with open(SRC / "curriculum_data.json", encoding="utf-8") as f:
    data = json.load(f)
json_blob = json.dumps(data, ensure_ascii=False)

HTML = r"""<!DOCTYPE html>
<html lang="ja">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>年間単元指導計画表</title>
<style>
/* ===== Reset & Base ===== */
*{margin:0;padding:0;box-sizing:border-box;}
html{font-size:11px;}
body{
  font-family:"游ゴシック","Yu Gothic","Hiragino Sans",sans-serif;
  background:#f0f2f5;color:#333;
}

/* ===== Toolbar ===== */
#toolbar{
  position:sticky;top:0;z-index:100;
  display:flex;align-items:center;gap:14px;
  padding:8px 20px;
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
  padding:6px 16px;border:none;border-radius:6px;
  font-size:12px;font-weight:700;cursor:pointer;
  color:#fff;transition:opacity .15s;
}
.tb-btn:hover{opacity:.85;}
.tb-btn.primary{background:#4299e1;}
.tb-btn.success{background:#48bb78;}
.tb-btn.warning{background:#ed8936;}
.tb-btn:disabled{opacity:.4;cursor:default;}
.tb-spacer{flex:1;}

/* ===== Title ===== */
#title-bar{
  text-align:center;padding:12px 0 6px;
  font-size:17px;font-weight:700;color:#2d3748;
  background:#fff;border-bottom:1px solid #e2e8f0;
}

/* ===== Table Container ===== */
#table-wrap{
  overflow:auto;padding:10px;background:#f0f2f5;
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
  font-size:11px;
  line-height:1.45;
  word-break:break-all;
}

/* ヘッダー */
#plan-table thead th{
  position:sticky;top:0;z-index:10;
  background:#4a5568;color:#fff;
  text-align:center;vertical-align:middle;
  font-size:11px;font-weight:700;
  padding:6px 2px;
  white-space:nowrap;
}
/* 月列 */
#plan-table td.month-cell{
  text-align:center;font-weight:700;
  vertical-align:middle;white-space:nowrap;
  width:38px;min-width:38px;
  font-size:12px;
}
/* 合計行 */
#plan-table tr.total-row td{
  text-align:center;font-weight:700;
  background:#edf2f7 !important;
  font-size:12px;
}

/* 学期色 */
.sem1{background:#E8F0FE;}
.sem2{background:#FFF8E1;}
.sem3{background:#E8F5E9;}
.sem-summer{background:#F5F5F5;}

/* セル内単元 */
.unit-item{
  padding:1px 0;
}
.unit-hours{
  color:#718096;font-size:10px;
}

/* ===== Print ===== */
@media print{
  @page{size:A3 landscape;margin:8mm;}
  body{background:#fff;}
  #toolbar{display:none !important;}
  #title-bar{border:none;padding:6px 0 4px;font-size:15px;}
  #table-wrap{padding:0;overflow:visible;}
  #plan-table{box-shadow:none;font-size:9px;}
  #plan-table th,#plan-table td{font-size:9px;padding:2px 3px;}
  #plan-table thead th{
    position:static;
    background:#4a5568 !important;color:#fff !important;
    -webkit-print-color-adjust:exact;print-color-adjust:exact;
  }
  .sem1{background:#E8F0FE !important;-webkit-print-color-adjust:exact;print-color-adjust:exact;}
  .sem2{background:#FFF8E1 !important;-webkit-print-color-adjust:exact;print-color-adjust:exact;}
  .sem3{background:#E8F5E9 !important;-webkit-print-color-adjust:exact;print-color-adjust:exact;}
  .sem-summer{background:#F5F5F5 !important;-webkit-print-color-adjust:exact;print-color-adjust:exact;}
  .total-row td{background:#edf2f7 !important;-webkit-print-color-adjust:exact;print-color-adjust:exact;}
}
</style>
</head>
<body>

<!-- ===== Toolbar ===== -->
<div id="toolbar">
  <label for="grade-select">学年：</label>
  <select id="grade-select">
    <option value="1">1年</option>
    <option value="2">2年</option>
    <option value="3">3年</option>
    <option value="4">4年</option>
    <option value="5" selected>5年</option>
    <option value="6">6年</option>
  </select>
  <div class="tb-spacer"></div>
  <button class="tb-btn primary" id="btn-save" disabled>保存</button>
  <button class="tb-btn success" id="btn-load" disabled>読み込み</button>
  <button class="tb-btn warning" id="btn-pdf" onclick="window.print()">
    PDF出力
  </button>
  <span style="font-size:10px;opacity:.7">※ 印刷ダイアログで「PDFとして保存」を選択</span>
</div>

<!-- ===== Title ===== -->
<div id="title-bar"></div>

<!-- ===== Table ===== -->
<div id="table-wrap">
  <table id="plan-table">
    <thead id="plan-thead"></thead>
    <tbody id="plan-tbody"></tbody>
  </table>
</div>

<script>
// ══════════════════════════════════════
//  埋め込みカリキュラムデータ
// ══════════════════════════════════════
const CURRICULUM_DATA = """ + json_blob + r""";

// ══════════════════════════════════════
//  定数
// ══════════════════════════════════════
const MONTHS = [4,5,6,7,8,9,10,11,12,1,2,3];
const MONTH_LABELS = ['4月','5月','6月','7月','8月','9月','10月','11月','12月','1月','2月','3月'];

// 学期判定: 月番号 → CSSクラス
function semClass(m) {
  if (m >= 4 && m <= 7) return 'sem1';
  if (m === 8) return 'sem-summer';
  if (m >= 9 && m <= 12) return 'sem2';
  return 'sem3'; // 1-3月
}

// 全教科の標準順序
const ALL_SUBJECTS = [
  '国語','算数','社会','理科','音楽','図画工作',
  '体育','家庭','外国語','道徳・特活','健康','食育',
  '自己実現活動','安全'
];

// ══════════════════════════════════════
//  表描画
// ══════════════════════════════════════
const gradeSelect = document.getElementById('grade-select');
const titleBar    = document.getElementById('title-bar');
const thead       = document.getElementById('plan-thead');
const tbody       = document.getElementById('plan-tbody');

function getSubjects(grade) {
  const d = CURRICULUM_DATA[grade];
  if (!d) return [];
  return ALL_SUBJECTS.filter(s => d[s] !== undefined);
}

function getUnits(grade, subject, month) {
  const arr = CURRICULUM_DATA[grade]?.[subject];
  if (!arr) return [];
  return arr.filter(u => u.month === month);
}

function renderTable(grade) {
  // タイトル
  titleBar.textContent = `令和7年度　第${grade}学年　年間単元指導計画表`;

  const subjects = getSubjects(grade);
  const numCols = subjects.length;

  // ── thead ──
  let hhtml = '<tr><th style="width:38px">月</th>';
  for (const s of subjects) {
    hhtml += `<th>${esc(s)}</th>`;
  }
  hhtml += '</tr>';
  thead.innerHTML = hhtml;

  // ── tbody ──
  // 年間時数集計用
  const totals = {};
  for (const s of subjects) totals[s] = 0;

  let bhtml = '';
  for (let mi = 0; mi < MONTHS.length; mi++) {
    const m = MONTHS[mi];
    const label = MONTH_LABELS[mi];
    const cls = semClass(m);

    bhtml += `<tr>`;
    bhtml += `<td class="month-cell ${cls}">${label}</td>`;

    for (const s of subjects) {
      const units = getUnits(grade, s, m);
      let cellHtml = '';
      let cellHours = 0;
      for (const u of units) {
        cellHtml += `<div class="unit-item">${esc(u.unit)}<span class="unit-hours">（${u.hours}）</span></div>`;
        cellHours += u.hours;
      }
      totals[s] += cellHours;
      bhtml += `<td class="${cls}">${cellHtml}</td>`;
    }
    bhtml += '</tr>';
  }

  // 合計行
  bhtml += '<tr class="total-row">';
  bhtml += '<td style="background:#edf2f7">合計</td>';
  for (const s of subjects) {
    bhtml += `<td>${totals[s]}</td>`;
  }
  bhtml += '</tr>';

  tbody.innerHTML = bhtml;

  // 列幅調整
  adjustColumnWidths(subjects);
}

function adjustColumnWidths(subjects) {
  const table = document.getElementById('plan-table');
  const n = subjects.length;
  // 月列は固定38px、残りを均等割
  const thCells = thead.querySelectorAll('th');
  if (thCells.length === 0) return;
  // tableのCSSでtable-layout:fixedなので、colgroup で制御
  let cg = table.querySelector('colgroup');
  if (cg) cg.remove();
  cg = document.createElement('colgroup');
  // 月列
  const col0 = document.createElement('col');
  col0.style.width = '38px';
  cg.appendChild(col0);
  // 教科列: 主要教科は広め
  const wide = {'国語':1.4,'算数':1.2,'社会':1.2,'理科':1.1,'道徳・特活':1.2};
  const totalWeight = subjects.reduce((sum, s) => sum + (wide[s] || 1), 0);
  for (const s of subjects) {
    const col = document.createElement('col');
    const pct = ((wide[s] || 1) / totalWeight * 100).toFixed(2);
    col.style.width = pct + '%';
    cg.appendChild(col);
  }
  table.prepend(cg);
}

// HTMLエスケープ
function esc(str) {
  const d = document.createElement('div');
  d.textContent = str;
  return d.innerHTML;
}

// ── 学年切替 ──
gradeSelect.addEventListener('change', () => {
  renderTable(gradeSelect.value);
});

// ── 初期描画 ──
renderTable(gradeSelect.value);
</script>
</body>
</html>
"""

out = SRC / "plan.html"
out.write_text(HTML, encoding="utf-8")
print(f"Generated: {out}  ({out.stat().st_size:,} bytes)")
