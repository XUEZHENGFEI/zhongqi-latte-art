# -*- coding: utf-8 -*-
"""生成中秋拉花 3 阶数据看板 HTML（数据内嵌、单文件、本地方案）。

v7 改造：北区/东区 顶部 tab 切换，每个 tab 内的城市为一阶。
两个源文件：
- 东区文件（28 列）：上海/江苏/合肥/浙江
- 北区文件（25 列）：北京/沈阳/大连/天津/呼和浩特/济南
"""
import openpyxl
import re
import json
import html as H
from collections import defaultdict

# ---------- 源文件配置 ----------
SRC_FILES = [
    {
        'path': r"C:\Users\kxue\Desktop\中秋兔子拉花\作业明细数据_20260921113052.xlsx",
        'tab': 'east',
        'tab_name': '东区',
    },
    {
        'path': r"C:\Users\kxue\Desktop\中秋兔子拉花\作业明细数据_20260921113208.xlsx",
        'tab': 'north',
        'tab_name': '北区',
    },
]

OUT = r"C:\Users\kxue\WorkBuddy\Kirin文件夹\中秋拉花看板.html"

# 列号（兼容两种文件结构，都统一到 col 5=所属部门）
COL_DEPT = 5
COL_NAME = 2
COL_POS = 7
COL_STATUS = 11
# 审批人姓名邮箱可能存在多个：col 17, 20, 23, 26（东区文件）
COL_APPROVER = [17, 20, 23, 26]

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

# ---------- 4. 提取所有记录（多文件循环） ----------
# 城市元数据：中文名 + 简称
CITY_META = {
    # 北区
    'BJ':  {'name': '北京',     'tagline': 'BJ'},
    'BJP': {'name': '北京Pop',  'tagline': 'BJP'},
    'TJ':  {'name': '天津',     'tagline': 'TJ'},
    'SY':  {'name': '沈阳',     'tagline': 'SY'},
    'DL':  {'name': '大连',     'tagline': 'DL'},
    'HT':  {'name': '呼和浩特', 'tagline': 'HT'},
    'JN':  {'name': '济南',     'tagline': 'JN'},
    # 东区
    'SH':  {'name': '上海',     'tagline': 'SH'},
    'SHP': {'name': '上海Pop',  'tagline': 'SHP'},
    'SHL': {'name': '上海岚',   'tagline': 'SHL'},
    'NJ':  {'name': '南京',     'tagline': 'NJ'},
    'SZ':  {'name': '苏州',     'tagline': 'SZ'},
    'WX':  {'name': '无锡',     'tagline': 'WX'},
    'CZ':  {'name': '常州',     'tagline': 'CZ'},
    'NT':  {'name': '南通',     'tagline': 'NT'},
    'YZ':  {'name': '扬州',     'tagline': 'YZ'},
    'KS':  {'name': '昆山',     'tagline': 'KS'},
    'ZJG': {'name': '张家港',   'tagline': 'ZJG'},
    'HF':  {'name': '合肥',     'tagline': 'HF'},
    'HA':  {'name': '淮安',     'tagline': 'HA'},
    'HZ':  {'name': '杭州',     'tagline': 'HZ'},
    'NB':  {'name': '宁波',     'tagline': 'NB'},
    'JH':  {'name': '金华',     'tagline': 'JH'},
    'SX':  {'name': '绍兴',     'tagline': 'SX'},
    'WZ':  {'name': '温州',     'tagline': 'WZ'},
    'JX':  {'name': '嘉兴',     'tagline': 'JX'},
    'TZ':  {'name': '台州',     'tagline': 'TZ'},
}

# tab_id → 城市前缀集合（用户确认：北区=北六省，东区=华东）
TAB_CITIES = {
    'north': {'BJ', 'BJP', 'TJ', 'SY', 'DL', 'HT', 'JN'},
    'east':  {'SH', 'SHP', 'SHL', 'NJ', 'SZ', 'WX', 'CZ', 'NT', 'YZ', 'KS', 'ZJG', 'HF', 'HA', 'HZ', 'NB', 'JH', 'SX', 'WZ', 'JX', 'TZ'},
}
TAB_ORDER = ['north', 'east']

# 东区子区域划分（v8 新增）
# 用户原话：江苏=江苏+合肥、上海=上海+桐乡+嘉兴、浙江
# 桐乡门店实际是 JX003-嘉兴桐乡万象汇，嘉兴/桐乡都归到 JX 前缀
EAST_SUB_REGIONS = [
    {
        'id': 'jiangsu',
        'name': '江苏',
        'tagline': '江苏 + 合肥',
        'prefixes': ['NJ', 'SZ', 'WX', 'CZ', 'NT', 'YZ', 'KS', 'ZJG', 'HA', 'HF'],
        'icon': '苏',
    },
    {
        'id': 'shanghai',
        'name': '上海',
        'tagline': '上海 + 桐乡 + 嘉兴',
        'prefixes': ['SH', 'SHP', 'SHL', 'JX'],
        'icon': '沪',
    },
    {
        'id': 'zhejiang',
        'name': '浙江',
        'tagline': '浙江',
        'prefixes': ['HZ', 'NB', 'JH', 'SX', 'WZ', 'TZ'],
        'icon': '浙',
    },
]

# 门店黑名单（剔除）
BLACKLIST_STORES = {'SHL01-上海始祖鸟会德丰咖啡店-联营', 'SH087-Roffee会议咖啡店'}

# 按门店前缀剔除（v9 新增：所有北京 pop-up 都不抓）
BLACKLIST_PREFIXES = {'BJP'}

records = []
for src in SRC_FILES:
    wb = openpyxl.load_workbook(src['path'], data_only=True)
    s = wb.worksheets[0]
    # 北区文件 25 列少 3 列（无审批人 2/3/4），过滤越界 col
    max_col = s.max_column
    approver_cols = [c for c in COL_APPROVER if c <= max_col]
    for r in range(3, s.max_row + 1):
        dept = s.cell(row=r, column=COL_DEPT).value
        if not dept:
            continue
        # 门店黑名单过滤（精确名称）
        if dept in BLACKLIST_STORES:
            continue
        # 前缀黑名单（如 BJP 北京 pop-up 全部剔除）
        m = re.match(r'^([A-Z]+)\d+', dept)
        if not m:
            continue
        if m.group(1) in BLACKLIST_PREFIXES:
            continue
        prefix = m.group(1)
        # 确定属于哪个 tab
        tab = None
        for tid, cities in TAB_CITIES.items():
            if prefix in cities:
                tab = tid
                break
        if not tab:
            continue
        # 东区下，确定属于哪个 sub_region
        sub_region = None
        if tab == 'east':
            for sr in EAST_SUB_REGIONS:
                if prefix in sr['prefixes']:
                    sub_region = sr['id']
                    break
        name = s.cell(row=r, column=COL_NAME).value
        position = s.cell(row=r, column=COL_POS).value
        # 用户指定：只保留 门店副经理 / 咖啡师 / 门店经理 / 值班经理 四个职级
        if position not in {'门店副经理', '咖啡师', '门店经理', '值班经理'}:
            continue
        status = s.cell(row=r, column=COL_STATUS).value
        approvers = []
        for c in approver_cols:
            raw = s.cell(row=r, column=c).value
            cn = extract_cn_name(raw)
            if cn:
                approvers.append(cn)
        store = dept
        records.append({
            'tab': tab,
            'sub_region': sub_region,
            'prefix': prefix,
            'city': CITY_META.get(prefix, {}).get('name', prefix),
            'store': store,
            'name': name,
            'position': position,
            'status': status,
            'statusLabel': {'已通过':'通过','未通过':'未通过','待重新提交':'未通过','待批阅':'待批阅','未提交':'未提交'}.get(status, status or '—'),
            'pass': pass_status(status),
            'approvers': approvers
        })

# ---------- 5. 聚合到 3 阶：tab → city → store → persons ----------
# city_key = tab + prefix
agg = defaultdict(lambda: defaultdict(lambda: {'passed':0, 'failed':0, 'pending':0, 'persons':[]}))

for rec in records:
    city_key = f"{rec['tab']}|{rec['prefix']}"
    sa = agg[city_key][rec['store']]
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

for ck in agg:
    agg[ck] = dict(sorted(agg[ck].items(), key=lambda kv: store_sort_key(kv[0])))

# tab_meta
tab_meta = {
    'north': {'name': '北区', 'tagline': '北京 · 天津 · 沈阳 · 大连 · 呼和浩特 · 济南', 'icon': '北'},
    'east':  {'name': '东区', 'tagline': '上海 · 江苏 · 合肥 · 浙江', 'icon': '东'},
}

# ---------- 6. 组装最终 JSON ----------

# 6.1 计算东区 sub_region 汇总（v8）
sub_region_agg = {sr['id']: {'passed':0, 'failed':0, 'pending':0, 'city_count':0, 'store_count':0, 'prefixes': []} for sr in EAST_SUB_REGIONS}
for rec in records:
    if rec['tab'] != 'east' or not rec['sub_region']:
        continue
    sr = sub_region_agg[rec['sub_region']]
    if rec['pass'] == 'passed': sr['passed'] += 1
    elif rec['pass'] == 'failed': sr['failed'] += 1
    else: sr['pending'] += 1

for sr in EAST_SUB_REGIONS:
    sr_id = sr['id']
    city_keys = {f"east|{p}" for p in sr['prefixes']}
    sub_region_agg[sr_id]['city_count'] = sum(1 for ck in city_keys if ck in agg)
    sub_region_agg[sr_id]['store_count'] = sum(len(agg[ck]) for ck in city_keys if ck in agg)
    sub_region_agg[sr_id]['prefixes'] = sorted(sr['prefixes'])

# 6.2 构造 cities 字典
def _build_city(tab_id, prefix, v):
    return {
        'tab': tab_id,
        'prefix': prefix,
        'name': CITY_META.get(prefix, {}).get('name', prefix),
        'tagline': CITY_META.get(prefix, {}).get('tagline', prefix),
        'passed': sum(s['passed'] for s in v.values()),
        'failed': sum(s['failed'] for s in v.values()),
        'pending': sum(s['pending'] for s in v.values()),
        'stores': [
            {
                'name': store,
                'passed': sv['passed'],
                'failed': sv['failed'],
                'pending': sv['pending'],
                'persons': [
                    {
                        'name': p['name'],
                        'position': p['position'] or '—',
                        'statusLabel': p['statusLabel'],
                        'pass': p['pass'],
                        'approvers': p['approvers'],
                    } for p in sv['persons']
                ]
            } for store, sv in v.items()
        ]
    }

out_data = {
    'tabs': [{'id': tid, 'name': tab_meta[tid]['name']} for tid in TAB_ORDER],
    'tab_meta': tab_meta,
    'tab_cities': {
        tid: sorted([p for p in TAB_CITIES[tid]], key=lambda x: (CITY_META.get(x, {}).get('name', x), x))
            for tid in TAB_ORDER
        },
    'east_sub_regions': [
        {**sr, 'passed': sub_region_agg[sr['id']]['passed'],
         'failed': sub_region_agg[sr['id']]['failed'],
         'pending': sub_region_agg[sr['id']]['pending'],
         'city_count': sub_region_agg[sr['id']]['city_count'],
         'store_count': sub_region_agg[sr['id']]['store_count']}
        for sr in EAST_SUB_REGIONS
    ],
    'city_meta': CITY_META,
    'cities': {ck: _build_city(*ck.split('|'), v) for ck, v in agg.items()},
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

/* === Tab 切换栏 === */
.tab-bar {
  display: flex; gap: 8px;
  margin-bottom: 14px;
  background: var(--card);
  padding: 6px;
  border-radius: 14px;
  box-shadow: var(--shadow);
}
.tab-btn {
  flex: 1;
  padding: 10px 16px;
  border-radius: 10px;
  display: flex; align-items: center; justify-content: space-between;
  gap: 10px;
  background: transparent;
  color: var(--ink-soft);
  font-size: 15px; font-weight: 600;
  transition: background .15s, color .15s, transform .12s;
}
.tab-btn:hover { background: var(--primary-soft); color: var(--ink); }
.tab-btn.active {
  background: linear-gradient(135deg, #fb923c, #c2410c);
  color: white;
  box-shadow: 0 4px 12px rgba(194, 65, 12, .25);
}
.tab-btn .tab-name { letter-spacing: 1px; }
.tab-btn .tab-stats { display: flex; gap: 4px; font-size: 12px; }
.tab-btn .tab-pill {
  padding: 2px 6px; border-radius: 5px;
  font-weight: 600; opacity: .92;
}
.tab-btn.active .tab-pill.green { background: rgba(255,255,255,.22); color: #d1fae5; }
.tab-btn.active .tab-pill.red   { background: rgba(255,255,255,.22); color: #fecaca; }
.tab-btn.active .tab-pill.gray  { background: rgba(255,255,255,.22); color: #e5e7eb; }
.tab-btn:not(.active) .tab-pill.green { background: var(--green-bg); color: var(--green); }
.tab-btn:not(.active) .tab-pill.red   { background: var(--red-bg); color: var(--red); }
.tab-btn:not(.active) .tab-pill.gray  { background: var(--gray-bg); color: var(--gray); }

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
.store-card .store-name.lvl-0 { color: #dc2626; font-weight: 700; }
.store-card .store-name.lvl-1 { color: #ea580c; font-weight: 700; }
.store-card .store-name.lvl-2 { color: var(--green); font-weight: 700; }
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
  <div class="sub">北区 · 东区 · 3 阶点击钻取</div>
</div>

<div class="container">
  <div class="tab-bar" id="tabBar"></div>
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
  const STATE = { tab: DATA.tabs[0].id, sub_region: null, city: null, store: null };

  const $tabBar = document.getElementById('tabBar');
  const $bc = document.getElementById('breadcrumb');
  const $sum = document.getElementById('summary');
  const $content = document.getElementById('content');

  // ===== 统一刷新入口（铁律 9：DAG 单向，绝不互调） =====
  function refreshAll() {
    renderTabs();
    renderBreadcrumb();
    renderSummary();
    renderContent();
  }

  // ===== Tab 切换栏 =====
  function renderTabs() {
    const counts = {};
    for (const t of DATA.tabs) {
      let p = 0, f = 0, n = 0;
      for (const ck of Object.keys(DATA.cities)) {
        const c = DATA.cities[ck];
        if (c.tab !== t.id) continue;
        p += c.passed; f += c.failed; n += c.pending;
      }
      counts[t.id] = { p, f, n };
    }
    $tabBar.innerHTML = DATA.tabs.map(t => {
      const active = STATE.tab === t.id;
      const c = counts[t.id];
      return `<button class="tab-btn ${active ? 'active' : ''}" data-tab="${t.id}">
        <span class="tab-name">${t.name}</span>
        <span class="tab-stats">
          <span class="tab-pill green">${c.p}</span>
          <span class="tab-pill red">${c.f}</span>
          <span class="tab-pill gray">${c.n}</span>
        </span>
      </button>`;
    }).join('');
    $tabBar.querySelectorAll('.tab-btn').forEach(btn => {
      btn.onclick = () => {
        if (STATE.tab === btn.dataset.tab) return;
        STATE.tab = btn.dataset.tab; STATE.sub_region = null; STATE.city = null; STATE.store = null;
        refreshAll();
      };
    });
  }

  // ===== 面包屑 =====
  function renderBreadcrumb() {
    const parts = [];
    parts.push({
      label: DATA.tab_meta[STATE.tab].name,
      active: !STATE.sub_region && !STATE.city,
      action: () => { STATE.sub_region = null; STATE.city = null; STATE.store = null; refreshAll(); }
    });
    // 东区有 sub_region 层级
    if (STATE.tab === 'east' && STATE.sub_region) {
      const sr = (DATA.east_sub_regions || []).find(s => s.id === STATE.sub_region);
      parts.push({
        label: sr ? sr.name : STATE.sub_region,
        active: !STATE.city,
        action: () => { STATE.city = null; STATE.store = null; refreshAll(); }
      });
    }
    if (STATE.city) {
      const cityKey = STATE.tab + '|' + STATE.city;
      const cityObj = DATA.cities[cityKey];
      const cityName = cityObj ? cityObj.name : STATE.city;
      parts.push({
        label: cityName,
        active: !STATE.store,
        action: () => { STATE.store = null; refreshAll(); }
      });
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
  function currentScope() {
    if (STATE.store) {
      const cityKey = STATE.tab + '|' + STATE.city;
      const store = DATA.cities[cityKey].stores.find(s => s.name === STATE.store);
      return { passed: store.passed, failed: store.failed, pending: store.pending };
    }
    if (STATE.city) {
      const cityKey = STATE.tab + '|' + STATE.city;
      const c = DATA.cities[cityKey];
      return { passed: c.passed, failed: c.failed, pending: c.pending };
    }
    if (STATE.sub_region && STATE.tab === 'east') {
      const sr = (DATA.east_sub_regions || []).find(s => s.id === STATE.sub_region);
      if (sr) return { passed: sr.passed, failed: sr.failed, pending: sr.pending };
    }
    // tab 总和
    let p=0,f=0,n=0;
    for (const ck of Object.keys(DATA.cities)) {
      const c = DATA.cities[ck];
      if (c.tab !== STATE.tab) continue;
      p += c.passed; f += c.failed; n += c.pending;
    }
    return { passed: p, failed: f, pending: n };
  }
  function currentLabel() {
    if (STATE.store) return STATE.store;
    if (STATE.city) {
      const cityKey = STATE.tab + '|' + STATE.city;
      return DATA.cities[cityKey].name;
    }
    if (STATE.sub_region && STATE.tab === 'east') {
      const sr = (DATA.east_sub_regions || []).find(s => s.id === STATE.sub_region);
      return sr ? sr.name : STATE.sub_region;
    }
    return DATA.tab_meta[STATE.tab].name;
  }
  function renderSummary() {
    const s = currentScope();
    const total = s.passed + s.failed + s.pending;
    const rate = total > 0 ? ((s.passed / total) * 100).toFixed(1) : '—';
    $sum.innerHTML = `
      <div class="summary">
        <div class="stat total"><div class="v">${total}</div><div class="l">${currentLabel()} · 四类职级总人数</div></div>
        <div class="stat passed"><div class="v">${s.passed}</div><div class="l">已通过</div></div>
        <div class="stat failed"><div class="v">${s.failed}</div><div class="l">未通过/待重新提交</div></div>
        <div class="stat"><div class="v" style="color:var(--gold)">${rate}${rate !== '—' ? '%' : ''}</div><div class="l">通过率<br><span style="font-size:10px;color:var(--ink-soft);font-weight:400">已通过 / 四类职级总数</span></div></div>
        <div class="stat"><div class="v" style="color:var(--gray)">${s.pending}</div><div class="l">待批阅/未提交</div></div>
        <div class="hint">点击城市卡片 → 展开门店；点击门店 → 展开人员明细（姓名、职位、通过状态、审批人中文名）。</div>
      </div>
    `;
  }

  // ===== 内容 =====
  function renderContent() {
    if (!STATE.store) {
      if (STATE.tab === 'east' && !STATE.sub_region) {
        renderSubRegions();
      } else if (!STATE.city) {
        renderCities();
      } else {
        renderStores();
      }
    } else {
      renderPeople();
    }
  }

  // 一阶-东区：sub_region 列表（江苏/上海/浙江）
  function renderSubRegions() {
    const list = DATA.east_sub_regions || [];
    const html = '<div class="region-grid">' + list.map(sr => `
        <div class="region-card" data-sr="${sr.id}">
          <div class="icon">${sr.icon}</div>
          <div class="name">${sr.name}</div>
          <div class="tag">${sr.tagline}</div>
          <div class="stats">
            <div class="pill passed"><div class="n">${sr.passed}</div><div class="l">已通过</div></div>
            <div class="pill failed"><div class="n">${sr.failed}</div><div class="l">未通过</div></div>
            <div class="pill pending"><div class="n">${sr.pending}</div><div class="l">待批/未交</div></div>
          </div>
          <div class="hint">${sr.city_count} 个城市 · ${sr.store_count} 家门店 ›</div>
        </div>
      `).join('') + '</div>';
    $content.innerHTML = html;
    $content.querySelectorAll('.region-card').forEach(card => {
      card.onclick = () => { STATE.sub_region = card.dataset.sr; refreshAll(); window.scrollTo({ top: 0, behavior: 'smooth' }); };
    });
  }

  // 二阶：当前 tab+sub_region 下的城市列表
  function renderCities() {
    let list = Object.values(DATA.cities).filter(c => c.tab === STATE.tab);
    // 东区下，按 sub_region 过滤
    if (STATE.tab === 'east' && STATE.sub_region) {
      const sr = (DATA.east_sub_regions || []).find(s => s.id === STATE.sub_region);
      if (sr) list = list.filter(c => sr.prefixes.includes(c.prefix));
    }
    if (list.length === 0) {
      $content.innerHTML = '<div class="empty">暂无城市数据</div>';
      return;
    }
    const html = '<div class="region-grid">' + list.map(c => `
        <div class="region-card" data-prefix="${c.prefix}">
          <div class="icon">${c.tagline}</div>
          <div class="name">${c.name}</div>
          <div class="tag">${c.prefix} · ${c.stores.length} 家门店</div>
          <div class="stats">
            <div class="pill passed"><div class="n">${c.passed}</div><div class="l">已通过</div></div>
            <div class="pill failed"><div class="n">${c.failed}</div><div class="l">未通过</div></div>
            <div class="pill pending"><div class="n">${c.pending}</div><div class="l">待批/未交</div></div>
          </div>
          <div class="hint">点击展开 ${c.stores.length} 家门店 ›</div>
        </div>
      `).join('') + '</div>';
    $content.innerHTML = html;
    $content.querySelectorAll('.region-card').forEach(card => {
      card.onclick = () => { STATE.city = card.dataset.prefix; refreshAll(); window.scrollTo({ top: 0, behavior: 'smooth' }); };
    });
  }

  // 二阶：当前 city 下的门店列表
  function renderStores() {
    const cityKey = STATE.tab + '|' + STATE.city;
    const cityObj = DATA.cities[cityKey];
    if (!cityObj || cityObj.stores.length === 0) {
      $content.innerHTML = '<div class="empty">该城市暂无门店数据</div>';
      return;
    }
    const html = '<div class="store-grid">' + cityObj.stores.map(s => {
      const lvl = s.passed === 0 ? 'lvl-0' : (s.passed === 1 ? 'lvl-1' : 'lvl-2');
      const lvlLabel = s.passed === 0 ? '全员未通过' : (s.passed === 1 ? '1 人通过' : `${s.passed} 人通过`);
      return `
      <div class="store-card" data-store="${s.name.replace(/"/g, '&quot;')}">
        <div class="store-name ${lvl}" title="${lvlLabel}">${s.name}</div>
        <div class="store-stats">
          <div class="ss-cell passed"><div class="n">${s.passed}</div><div class="l">通过</div></div>
          <div class="ss-cell failed"><div class="n">${s.failed}</div><div class="l">未通过</div></div>
          <div class="ss-cell pending"><div class="n">${s.pending}</div><div class="l">待处理</div></div>
        </div>
      </div>
    `;}).join('') + '</div>';
    $content.innerHTML = html;
    $content.querySelectorAll('.store-card').forEach(card => {
      card.onclick = () => { STATE.store = card.dataset.store; refreshAll(); window.scrollTo({ top: 0, behavior: 'smooth' }); };
    });
  }

  // 三阶：当前 store 的人员明细
  function renderPeople() {
    const cityKey = STATE.tab + '|' + STATE.city;
    const cityObj = DATA.cities[cityKey];
    const store = cityObj && cityObj.stores.find(s => s.name === STATE.store);
    if (!store) {
      $content.innerHTML = '<div class="empty">未找到门店数据</div>';
      return;
    }
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

# 同时复制到部署目录（Cloudflare Pages）
import shutil
DEPLOY_INDEX = r"C:\Users\kxue\WorkBuddy\Kirin文件夹\latte-art-zhongqi\index.html"
shutil.copy2(OUT, DEPLOY_INDEX)

print("✓ HTML 生成成功：", OUT)
print("✓ 同步到部署目录：", DEPLOY_INDEX)
print("  大小：", len(final_html), "bytes")
print("  Tab 数：", len(out_data['tabs']))
total_persons = 0
for tab in out_data['tabs']:
    city_list = [c for ck, c in out_data['cities'].items() if c['tab'] == tab['id']]
    p = sum(c['passed'] for c in city_list)
    f = sum(c['failed'] for c in city_list)
    n = sum(c['pending'] for c in city_list)
    print(f"  {tab['name']}: 通过 {p} / 未通过 {f} / 待处理 {n} / {len(city_list)} 个城市")
    for c in city_list:
        total_persons += sum(len(s['persons']) for s in c['stores'])
print("  人员记录总数：", total_persons)