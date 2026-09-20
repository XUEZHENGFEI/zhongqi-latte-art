# -*- coding: utf-8 -*-
"""生成中秋拉花 3 阶数据看板 HTML（数据内嵌、单文件、本地方案）。"""
import openpyxl
import re
import json
import html as H
from collections import defaultdict

SRC = r"C:\Users\kxue\Desktop\中秋兔子拉花\作业明细数据_20260920170806.xlsx"
OUT = r"C:\Users\kxue\WorkBuddy\Kirin文件夹\中秋拉花看板.html"

# ---------- 1. 读取 Excel ----------
wb = openpyxl.load_workbook(SRC, data_only=True)
s = wb.worksheets[0]

COL_DEPT = 5
COL_NAME = 2
COL_POS = 7
COL_STATUS = 11
COL_APPROVER = [17, 20, 23, 26]

# ---------- 2. 区域映射 ----------
# 用户指定：江苏+合肥（含淮安 HA）= 1 卡；上海 = 1 卡；浙江 = 1 卡
JIANG_SU_HEFEI = {'NJ', 'SZ', 'WX', 'CZ', 'NT', 'YZ', 'KS', 'ZJG', 'HF', 'HA'}
SHANGHAI = {'SH', 'SHP', 'SHL'}
ZHEJIANG = {'HZ', 'NB', 'JH', 'SX', 'WZ', 'JX', 'TZ'}

PREFIX_TO_GROUP = {}
for p in JIANG_SU_HEFEI: PREFIX_TO_GROUP[p] = '江苏+合肥'
for p in SHANGHAI: PREFIX_TO_GROUP[p] = '上海'
for p in ZHEJIANG: PREFIX_TO_GROUP[p] = '浙江'

# ---------- 3. 中文名提取 ----------
def extract_cn_name(raw):
    """'吕静雯jane.lv@peets.cn' -> '吕静雯'"""
    if not raw:
        return None
    last_at = raw.rfind('@')
    if last_at == -1:
        return None
    head = raw[:last_at]
    first_cn = None
    for i, ch in enumerate(head):
        if '\u4e00' <= ch <= '\u9fff':
            first_cn = i
            break
    if first_cn is None:
        return None
    cn_part = head[first_cn:]
    for j, ch in enumerate(cn_part):
        if ch.isascii() and ch.isalpha():
            return cn_part[:j]
    return cn_part

def pass_status(st):
    """终态判定：已通过=passed；未通过/待重新提交=failed；其他=pending"""
    if st == '已通过':
        return 'passed'
    if st in ('未通过', '待重新提交'):
        return 'failed'
    return 'pending'

# ---------- 4. 提取所有记录 ----------
records = []
for r in range(3, s.max_row + 1):
    dept = s.cell(row=r, column=COL_DEPT).value
    if not dept:
        continue
    m = re.match(r'^([A-Z]+)(\d+)', dept)
    if not m:
        continue
    prefix = m.group(1)
    group = PREFIX_TO_GROUP.get(prefix)
    if not group:
        continue
    name = s.cell(row=r, column=COL_NAME).value
    position = s.cell(row=r, column=COL_POS).value
    # 用户指定：只保留 门店副经理 / 咖啡师 / 门店经理 / 值班经理 四个职级
    if position not in {'门店副经理', '咖啡师', '门店经理', '值班经理'}:
        continue
    status = s.cell(row=r, column=COL_STATUS).value
    approvers = []
    for c in COL_APPROVER:
        raw = s.cell(row=r, column=c).value
        cn = extract_cn_name(raw)
        if cn:
            approvers.append(cn)
    # 门店全名（去掉前缀编号前缀部分，保留 XX000-门店名 完整形式更友好）
    store = dept  # 完整保留 "NB005-宁波万象汇店"
    records.append({
        'group': group,
        'store': store,
        'name': name,
        'position': position,
        'status': status,
        'statusLabel': {'已通过':'通过','未通过':'未通过','待重新提交':'未通过','待批阅':'待批阅','未提交':'未提交'}.get(status, status or '—'),
        'pass': pass_status(status),
        'approvers': approvers
    })

# ---------- 5. 聚合到 3 阶 ----------
# group -> store -> {passed, failed, pending, persons[]}
agg = defaultdict(lambda: defaultdict(lambda: {'passed':0, 'failed':0, 'pending':0, 'persons':[]}))

for rec in records:
    sa = agg[rec['group']][rec['store']]
    if rec['pass'] == 'passed':
        sa['passed'] += 1
    elif rec['pass'] == 'failed':
        sa['failed'] += 1
    else:
        sa['pending'] += 1
    sa['persons'].append(rec)

# 排序：store 按编号稳定排序
def store_sort_key(store):
    m = re.match(r'^([A-Z]+)(\d+)-(.+)$', store)
    if m:
        return (m.group(1), int(m.group(2)), m.group(3))
    return (store, 0, '')

for g in agg:
    agg[g] = dict(sorted(agg[g].items(), key=lambda kv: store_sort_key(kv[0])))

# 区域元数据
group_meta = {
    '江苏+合肥': {
        'tagline': 'NJ · SZ · WX · CZ · NT · YZ · KS · ZJG · HA · HF',
        'icon': '江',
    },
    '上海': {
        'tagline': 'SH · SHP · SHL',
        'icon': '沪',
    },
    '浙江': {
        'tagline': 'HZ · NB · JH · SX · WZ · JX · TZ',
        'icon': '浙',
    },
}

# 输出 JSON（Python dict 转 JSON）
out_data = {
    'groups': ['江苏+合肥', '上海', '浙江'],
    'group_meta': group_meta,
    'agg': {
        g: {
            'passed': sum(v['passed'] for v in stores.values()),
            'failed': sum(v['failed'] for v in stores.values()),
            'pending': sum(v['pending'] for v in stores.values()),
            'stores': [
                {
                    'name': store,
                    'passed': v['passed'],
                    'failed': v['failed'],
                    'pending': v['pending'],
                    'persons': [
                        {
                            'name': p['name'],
                            'position': p['position'] or '—',
                            'statusLabel': p['statusLabel'],
                            'pass': p['pass'],
                            'approvers': p['approvers'],
                        } for p in v['persons']
                    ]
                } for store, v in stores.items()
            ]
        } for g, stores in agg.items()
    }
}

DATA_JSON = json.dumps(out_data, ensure_ascii=False, separators=(',', ':'))

# ---------- 6. HTML 模板 ----------
HTML = r"""<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no">
<title>中秋限定拉花 · 作业看板</title>
<style>
/* === 全局 === */
:root {
  --bg: #fff9f1;
  --card: #ffffff;
  --ink: #2a1f17;
  --ink-soft: #6b5b4f;
  --line: #f0e6d8;
  --primary: #c2410c;
  --primary-soft: #fed7aa;
  --amber: #d97706;
  --gold: #b45309;
  --green: #15803d;
  --green-bg: #dcfce7;
  --red: #b91c1c;
  --red-bg: #fee2e2;
  --gray: #6b7280;
  --gray-bg: #f3f4f6;
  --shadow: 0 4px 14px rgba(120, 60, 16, 0.08);
  --shadow-hover: 0 8px 22px rgba(120, 60, 16, 0.14);
}
* { box-sizing: border-box; -webkit-tap-highlight-color: transparent; }
html, body {
  margin: 0; padding: 0;
  font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", "PingFang SC", "Microsoft YaHei", sans-serif;
  background: var(--bg);
  color: var(--ink);
  font-size: 15px;
  line-height: 1.55;
  -webkit-font-smoothing: antialiased;
}
button { font-family: inherit; cursor: pointer; border: none; background: none; color: inherit; }

/* === 头部 === */
.header {
  background: linear-gradient(135deg, #fed7aa 0%, #fdba74 60%, #fb923c 100%);
  padding: 20px 24px 18px;
  color: #7c2d12;
  position: sticky; top: 0; z-index: 100;
  box-shadow: var(--shadow);
}
.header h1 {
  margin: 0; font-size: 22px; font-weight: 700;
  display: flex; align-items: center; gap: 10px;
}
.header .sub {
  margin-top: 4px; font-size: 13px; color: #9a3412;
  font-weight: 500;
}
.rabbit {
  width: 30px; height: 30px;
  display: inline-block; vertical-align: middle;
}

/* === 容器 === */
.container { max-width: 1180px; margin: 0 auto; padding: 18px 18px 40px; }

/* === 面包屑 === */
.breadcrumb {
  display: flex; align-items: center; gap: 6px;
  font-size: 13px; color: var(--ink-soft);
  margin-bottom: 14px;
  flex-wrap: wrap;
}
.breadcrumb button {
  padding: 4px 10px; border-radius: 8px;
  background: var(--gray-bg); color: var(--ink);
  transition: background .15s;
  font-size: 13px;
}
.breadcrumb button:hover { background: var(--primary-soft); }
.breadcrumb button.current { background: var(--primary); color: white; }
.breadcrumb .sep { color: var(--ink-soft); opacity: .5; }

/* === 摘要带 === */
.summary {
  background: var(--card);
  border-radius: 14px; padding: 14px 18px;
  box-shadow: var(--shadow);
  display: flex; gap: 24px;
  margin-bottom: 16px;
  flex-wrap: wrap;
  border-left: 4px solid var(--amber);
}
.summary .stat { display: flex; flex-direction: column; gap: 2px; }
.summary .stat .v { font-size: 22px; font-weight: 700; line-height: 1.1; }
.summary .stat .l { font-size: 12px; color: var(--ink-soft); }
.summary .stat.passed .v { color: var(--green); }
.summary .stat.failed .v { color: var(--red); }
.summary .stat.total .v { color: var(--primary); }
.summary .hint { font-size: 12px; color: var(--ink-soft); margin-left: auto; align-self: center; max-width: 360px; }

/* === 一阶：3 个区域卡 === */
.region-grid {
  display: grid;
  grid-template-columns: repeat(3, 1fr);
  gap: 16px;
}
.region-card {
  background: var(--card);
  border-radius: 16px;
  padding: 22px 20px 18px;
  cursor: pointer;
  transition: transform .18s, box-shadow .18s;
  box-shadow: var(--shadow);
  border-top: 4px solid var(--amber);
  position: relative;
  overflow: hidden;
}
.region-card:hover { transform: translateY(-3px); box-shadow: var(--shadow-hover); }
.region-card.active { border-top-color: var(--primary); background: linear-gradient(180deg, #fff7ed 0%, #ffffff 50%); }
.region-card .icon {
  width: 44px; height: 44px; border-radius: 12px;
  background: linear-gradient(135deg, #fb923c, #c2410c);
  color: white; font-weight: 700; font-size: 22px;
  display: flex; align-items: center; justify-content: center;
  margin-bottom: 12px;
  box-shadow: 0 4px 10px rgba(194, 65, 12, .25);
}
.region-card .name { font-size: 19px; font-weight: 700; color: var(--ink); }
.region-card .tag { font-size: 11px; color: var(--ink-soft); margin-top: 2px; letter-spacing: .5px; }
.region-card .stats {
  display: flex; gap: 10px; margin-top: 14px;
}
.region-card .pill {
  flex: 1; padding: 10px 8px;
  border-radius: 10px; text-align: center;
}
.region-card .pill .n { font-size: 22px; font-weight: 700; line-height: 1.1; }
.region-card .pill .l { font-size: 11px; opacity: .8; margin-top: 2px; }
.region-card .pill.passed { background: var(--green-bg); color: var(--green); }
.region-card .pill.failed { background: var(--red-bg); color: var(--red); }
.region-card .pill.pending { background: var(--gray-bg); color: var(--gray); }
.region-card .hint {
  font-size: 11px; color: var(--ink-soft);
  margin-top: 12px; padding-top: 10px;
  border-top: 1px dashed var(--line);
  text-align: right;
}

/* === 二阶：门店网格 === */
.store-grid {
  display: grid;
  grid-template-columns: repeat(auto-fill, minmax(220px, 1fr));
  gap: 12px;
}
.store-card {
  background: var(--card);
  border-radius: 12px;
  padding: 14px 14px 12px;
  cursor: pointer;
  box-shadow: var(--shadow);
  transition: transform .15s, box-shadow .15s;
  border-left: 3px solid var(--amber);
}
.store-card:hover { transform: translateY(-2px); box-shadow: var(--shadow-hover); border-left-color: var(--primary); }
.store-card .store-name { font-size: 14px; font-weight: 600; color: var(--ink); margin-bottom: 8px; line-height: 1.4; word-break: break-all; }
.store-card .store-stats {
  display: flex; gap: 6px; font-size: 12px;
}
.store-card .ss-cell {
  flex: 1; padding: 6px 4px;
  border-radius: 6px; text-align: center;
}
.store-card .ss-cell .n { font-size: 16px; font-weight: 700; }
.store-card .ss-cell .l { font-size: 10px; opacity: .75; }
.store-card .ss-cell.passed { background: var(--green-bg); color: var(--green); }
.store-card .ss-cell.failed { background: var(--red-bg); color: var(--red); }
.store-card .ss-cell.pending { background: var(--gray-bg); color: var(--gray); }

/* === 三阶：人员明细表格 === */
.detail-panel {
  background: var(--card);
  border-radius: 14px;
  padding: 18px;
  box-shadow: var(--shadow);
}
.detail-panel h2 {
  margin: 0 0 4px; font-size: 18px; color: var(--ink);
}
.detail-panel .desc { font-size: 12px; color: var(--ink-soft); margin-bottom: 14px; }

.table-wrap { overflow-x: auto; -webkit-overflow-scrolling: touch; }
table.people {
  width: 100%; border-collapse: collapse;
  font-size: 13px; min-width: 540px;
}
table.people th {
  background: #fff7ed; color: var(--gold);
  text-align: left; padding: 10px 12px;
  font-weight: 600; font-size: 12px;
  border-bottom: 2px solid var(--primary-soft);
  white-space: nowrap;
}
table.people td {
  padding: 10px 12px;
  border-bottom: 1px solid var(--line);
  vertical-align: middle;
}
table.people tr:hover td { background: #fffaf3; }
table.people .badge {
  display: inline-block; padding: 2px 8px;
  border-radius: 6px; font-size: 11px;
  font-weight: 600;
}
table.people .badge.passed { background: var(--green-bg); color: var(--green); }
table.people .badge.failed { background: var(--red-bg); color: var(--red); }
table.people .badge.pending { background: var(--gray-bg); color: var(--gray); }
table.people .approver {
  font-size: 12px; color: var(--ink-soft);
  background: #fff7ed; padding: 2px 6px;
  border-radius: 4px; margin-right: 4px; margin-bottom: 2px;
  display: inline-block;
}

/* 空态 */
.empty {
  text-align: center; padding: 60px 20px;
  color: var(--ink-soft); font-size: 14px;
}

/* 移动端 */
@media (max-width: 760px) {
  .header { padding: 16px 18px 14px; }
  .header h1 { font-size: 18px; }
  .container { padding: 14px 14px 30px; }
  .region-grid { grid-template-columns: 1fr; }
  .region-card { padding: 18px 16px 14px; }
  .summary { gap: 14px; padding: 12px 14px; }
  .summary .stat .v { font-size: 18px; }
  .summary .hint { display: none; }
  .store-grid { grid-template-columns: 1fr; }
  .detail-panel { padding: 14px; }
  .breadcrumb button { padding: 4px 8px; font-size: 12px; }
}

/* footer */
.footer {
  text-align: center; color: var(--ink-soft);
  font-size: 11px; padding: 20px 14px;
  border-top: 1px dashed var(--line); margin-top: 24px;
}
</style>
</head>
<body>

<div class="header">
  <h1>
    <svg class="rabbit" viewBox="0 0 64 64" fill="none" xmlns="http://www.w3.org/2000/svg">
      <ellipse cx="32" cy="48" rx="20" ry="13" fill="#fff5e6" stroke="#c2410c" stroke-width="2"/>
      <ellipse cx="20" cy="14" rx="6" ry="14" fill="#fff5e6" stroke="#c2410c" stroke-width="2"/>
      <ellipse cx="44" cy="14" rx="6" ry="14" fill="#fff5e6" stroke="#c2410c" stroke-width="2"/>
      <ellipse cx="20" cy="14" rx="2.5" ry="9" fill="#fb923c"/>
      <ellipse cx="44" cy="14" rx="2.5" ry="9" fill="#fb923c"/>
      <circle cx="25" cy="34" r="2.5" fill="#c2410c"/>
      <circle cx="39" cy="34" r="2.5" fill="#c2410c"/>
      <ellipse cx="32" cy="40" rx="2.5" ry="1.5" fill="#c2410c"/>
      <path d="M28 43 Q32 47 36 43" stroke="#c2410c" stroke-width="1.8" fill="none" stroke-linecap="round"/>
      <circle cx="25" cy="34" r="1" fill="#fff"/>
      <circle cx="39" cy="34" r="1" fill="#fff"/>
    </svg>
    中秋限定拉花 · 作业看板
  </h1>
  <div class="sub">江苏+合肥 · 上海 · 浙江 · 3 阶点击钻取</div>
</div>

<div class="container">
  <div class="breadcrumb" id="breadcrumb"></div>
  <div id="summary"></div>
  <div id="content"></div>
  <div class="footer">
    数据口径：仅统计 <strong>门店副经理 / 咖啡师 / 门店经理 / 值班经理</strong> 四个职级 · 通过率 = 已通过 ÷ 四类职级总数（含待批/未提交）<br>
    拖到手机浏览器里也能用 · 单文件离线可看
  </div>
</div>

<!-- 数据内嵌 -->
<script id="data" type="application/json">__DATA_JSON__</script>

<script>
(function () {
  const DATA = JSON.parse(document.getElementById('data').textContent);
  const STATE = { group: null, store: null };

  const $bc = document.getElementById('breadcrumb');
  const $sum = document.getElementById('summary');
  const $content = document.getElementById('content');

  // ===== 统一刷新入口（铁律 9：DAG 单向，绝不互调） =====
  function refreshAll() {
    renderBreadcrumb();
    renderSummary();
    renderContent();
  }

  // ===== 面包屑 =====
  function renderBreadcrumb() {
    const parts = [];
    parts.push({ label: '全部区域', active: !STATE.group, action: () => { STATE.group = null; STATE.store = null; refreshAll(); } });
    if (STATE.group) {
      const g = DATA.group_meta[STATE.group];
      parts.push({ label: STATE.group + ' · ' + g.tagline.split(' · ')[0], active: !STATE.store, action: () => { STATE.store = null; refreshAll(); } });
    }
    if (STATE.store) {
      parts.push({ label: STATE.store, active: true });
    }
    $bc.innerHTML = parts.map((p, i) =>
      `<button class="${p.active ? 'current' : ''}" data-i="${i}">${p.label}</button>${i < parts.length - 1 ? '<span class="sep">›</span>' : ''}`
    ).join('');
    $bc.querySelectorAll('button').forEach((btn) => {
      const i = parseInt(btn.dataset.i);
      btn.onclick = () => { if (!parts[i].active) parts[i].action(); };
    });
  }

  // ===== 摘要 =====
  function renderSummary() {
    const scope = STATE.group
      ? (STATE.store
          ? { passed: DATA.agg[STATE.group].stores.find(s => s.name === STATE.store).passed,
              failed: DATA.agg[STATE.group].stores.find(s => s.name === STATE.store).failed,
              pending: DATA.agg[STATE.group].stores.find(s => s.name === STATE.store).pending }
          : { passed: DATA.agg[STATE.group].passed, failed: DATA.agg[STATE.group].failed, pending: DATA.agg[STATE.group].pending })
      : {
          passed: DATA.groups.reduce((s, g) => s + DATA.agg[g].passed, 0),
          failed: DATA.groups.reduce((s, g) => s + DATA.agg[g].failed, 0),
          pending: DATA.groups.reduce((s, g) => s + DATA.agg[g].pending, 0),
        };
    const total = scope.passed + scope.failed + scope.pending;
    const rate = total > 0 ? ((scope.passed / total) * 100).toFixed(1) : '—';
    const scope_label = STATE.store ? STATE.store : (STATE.group ? STATE.group : '全部区域');
    $sum.innerHTML = `
      <div class="summary">
        <div class="stat total"><div class="v">${scope.passed + scope.failed + scope.pending}</div><div class="l">${scope_label} · 四类职级总人数</div></div>
        <div class="stat passed"><div class="v">${scope.passed}</div><div class="l">已通过</div></div>
        <div class="stat failed"><div class="v">${scope.failed}</div><div class="l">未通过/待重新提交</div></div>
        <div class="stat"><div class="v" style="color:var(--gold)">${rate}${rate !== '—' ? '%' : ''}</div><div class="l">通过率<br><span style="font-size:10px;color:var(--ink-soft);font-weight:400">已通过 / 四类职级总数</span></div></div>
        <div class="stat"><div class="v" style="color:var(--gray)">${scope.pending}</div><div class="l">待批阅/未提交</div></div>
        <div class="hint">点击区域卡片 → 展开门店；点击门店 → 展开人员明细（姓名、职位、通过状态、审批人中文名）。</div>
      </div>
    `;
  }

  // ===== 内容 =====
  function renderContent() {
    if (!STATE.group) {
      renderRegions();
    } else if (!STATE.store) {
      renderStores();
    } else {
      renderPeople();
    }
  }

  // 一阶：3 个区域
  function renderRegions() {
    const html = '<div class="region-grid">' + DATA.groups.map(g => {
      const a = DATA.agg[g];
      const m = DATA.group_meta[g];
      return `
        <div class="region-card" data-group="${g}">
          <div class="icon">${m.icon}</div>
          <div class="name">${g}</div>
          <div class="tag">${m.tagline}</div>
          <div class="stats">
            <div class="pill passed"><div class="n">${a.passed}</div><div class="l">已通过</div></div>
            <div class="pill failed"><div class="n">${a.failed}</div><div class="l">未通过</div></div>
            <div class="pill pending"><div class="n">${a.pending}</div><div class="l">待批/未交</div></div>
          </div>
          <div class="hint">${a.stores.length} 家门店 · 点击展开 ›</div>
        </div>
      `;
    }).join('') + '</div>';
    $content.innerHTML = html;
    $content.querySelectorAll('.region-card').forEach(card => {
      card.onclick = () => { STATE.group = card.dataset.group; STATE.store = null; refreshAll(); };
    });
  }

  // 二阶：门店列表
  function renderStores() {
    const stores = DATA.agg[STATE.group].stores;
    if (stores.length === 0) {
      $content.innerHTML = '<div class="empty">该区域暂无门店数据</div>';
      return;
    }
    const html = '<div class="store-grid">' + stores.map(s => `
      <div class="store-card" data-store="${s.name.replace(/"/g, '&quot;')}">
        <div class="store-name">${s.name}</div>
        <div class="store-stats">
          <div class="ss-cell passed"><div class="n">${s.passed}</div><div class="l">通过</div></div>
          <div class="ss-cell failed"><div class="n">${s.failed}</div><div class="l">未通过</div></div>
          <div class="ss-cell pending"><div class="n">${s.pending}</div><div class="l">待处理</div></div>
        </div>
      </div>
    `).join('') + '</div>';
    $content.innerHTML = html;
    $content.querySelectorAll('.store-card').forEach(card => {
      card.onclick = () => { STATE.store = card.dataset.store; refreshAll(); window.scrollTo({ top: 0, behavior: 'smooth' }); };
    });
  }

  // 三阶：人员明细
  function renderPeople() {
    const store = DATA.agg[STATE.group].stores.find(s => s.name === STATE.store);
    if (!store) {
      $content.innerHTML = '<div class="empty">未找到门店数据</div>';
      return;
    }
    // 排序：通过在前 → 未通过 → 待处理；同名按职位
    const order = { passed: 0, failed: 1, pending: 2 };
    const persons = store.persons.slice().sort((a, b) => {
      if (order[a.pass] !== order[b.pass]) return order[a.pass] - order[b.pass];
      return (a.name || '').localeCompare(b.name || '', 'zh-Hans-CN');
    });
    const html = `
      <div class="detail-panel">
        <h2>${store.name} · 人员明细</h2>
        <div class="desc">通过 <strong style="color:var(--green)">${store.passed}</strong> · 未通过 <strong style="color:var(--red)">${store.failed}</strong> · 待处理 ${store.pending} · 共 ${persons.length} 条作业</div>
        <div class="table-wrap">
          <table class="people">
            <thead>
              <tr>
                <th style="width:90px">姓名</th>
                <th style="width:130px">职位</th>
                <th style="width:90px">状态</th>
                <th>审批人（已去邮箱）</th>
              </tr>
            </thead>
            <tbody>
              ${persons.map(p => `
                <tr>
                  <td><strong>${p.name || '—'}</strong></td>
                  <td>${p.position || '—'}</td>
                  <td><span class="badge ${p.pass}">${p.statusLabel}</span></td>
                  <td>${p.approvers.length === 0 ? '<span style="color:var(--ink-soft);font-size:12px">无审批人</span>' : p.approvers.map(a => `<span class="approver">${a}</span>`).join('')}</td>
                </tr>
              `).join('')}
            </tbody>
          </table>
        </div>
      </div>
    `;
    $content.innerHTML = html;
  }

  // ===== 初始化 =====
  document.addEventListener('DOMContentLoaded', refreshAll);
})();
</script>

</body>
</html>
"""

final_html = HTML.replace('__DATA_JSON__', DATA_JSON)
with open(OUT, 'w', encoding='utf-8') as f:
    f.write(final_html)

print("✓ HTML 生成成功：", OUT)
print("  大小：", len(final_html), "bytes")
print("  区域数：", len(out_data['groups']))
total_persons = 0
for g, v in out_data['agg'].items():
    print(f"  {g}: 通过 {v['passed']} / 未通过 {v['failed']} / 待处理 {v['pending']} / {len(v['stores'])} 家门店")
    for s in v['stores']:
        total_persons += len(s['persons'])
print("  人员记录总数：", total_persons)