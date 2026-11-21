# -*- coding: utf-8 -*-
"""从 index.html 内嵌的 RAW_ITEMS / RAW_REWARDS 生成打印版 Excel。"""
import re, os
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

HTML = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'index.html')
OUT = os.path.join(os.path.dirname(os.path.abspath(__file__)), '成长积分表-优化版.xlsx')

text = open(HTML, encoding='utf-8').read()
raw_items = re.search(r'const RAW_ITEMS = `(.*?)`;', text, re.S).group(1).strip().split('\n')
raw_rewards = re.search(r'const RAW_REWARDS = `(.*?)`;', text, re.S).group(1).strip().split('\n')

items = []
for line in raw_items:
    p = [x.strip() for x in line.split('|')]
    if len(p) < 2 or not p[1]:
        continue
    items.append({
        'cat': p[0], 'title': p[1], 'desc': p[2] if len(p) > 2 else '',
        'pts': p[3] if len(p) > 3 else '1', 'kind': p[4] if len(p) > 4 else 'plus',
        'peers': (p[5] if len(p) > 5 else '') == '1'
    })

rewards = []
for line in raw_rewards:
    p = [x.strip() for x in line.split('|')]
    if len(p) < 2:
        continue
    rewards.append({'title': p[0], 'cost': p[1], 'icon': p[2] if len(p) > 2 else '', 'daily': p[3] if len(p) > 3 else '0', 'note': p[4] if len(p) > 4 else ''})

# ---------- 样式 ----------
INK = '1F2937'
HEAD_FILL = PatternFill('solid', fgColor='3B82F6')
CAT_FILL = PatternFill('solid', fgColor='F1F5FB')
NAG_FILL = PatternFill('solid', fgColor='FEF2F2')
thin = Side(style='thin', color='DCE6F2')
BOX = Border(left=thin, right=thin, top=thin, bottom=thin)
F_BASE = Font(name='微软雅黑', size=10, color=INK)
F_HEAD = Font(name='微软雅黑', size=10, bold=True, color='FFFFFF')
F_BOLD = Font(name='微软雅黑', size=10, bold=True, color=INK)
F_TITLE = Font(name='微软雅黑', size=16, bold=True, color='1E3A8A')

wb = Workbook()

# ---------- Sheet1 每日打卡 ----------
ws = wb.active
ws.title = '每日打卡'
ws.sheet_view.showGridLines = False
ws['A1'] = '成长积分 · 每日打卡表（6 岁 · 深圳 · 一年级）'
ws['A1'].font = F_TITLE
ws.merge_cells('A1:J1')
ws['A2'] = '每天选 3–6 项重点打卡，不追求全打满；周一、周三有延时课，作业与阅读顺延 40 分钟。'
ws['A2'].font = Font(name='微软雅黑', size=9.5, color='6B7280')
ws.merge_cells('A2:J2')

heads = ['类别', '打卡项目', '怎么做', '分值', '周一', '周二', '周三', '周四', '周五', '周六', '周日']
# 11 列：类别/项目/说明/分值/7天 => 共 11
ws.insert_rows(3, 1)
for i, h in enumerate(heads):
    c = ws.cell(row=4, column=i + 1, value=h)
    c.font = F_HEAD
    c.fill = HEAD_FILL
    c.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    c.border = BOX
ws.row_dimensions[4].height = 24

row = 5
last_cat = None
cat_count = {}
for it in items:
    cat_count[it['cat']] = cat_count.get(it['cat'], 0) + 1
for it in items:
    if it['cat'] != last_cat:
        ws.cell(row=row, column=1, value=it['cat']).font = F_BOLD
        ws.merge_cells(start_row=row, start_column=1, end_row=row + cat_count[it['cat']] - 1, end_column=1)
        c = ws.cell(row=row, column=1)
        c.fill = CAT_FILL
        c.alignment = Alignment(horizontal='center', vertical='center')
        last_cat = it['cat']
    ws.cell(row=row, column=1).fill = CAT_FILL
    ws.cell(row=row, column=1).border = BOX

    t = ws.cell(row=row, column=2, value=it['title'])
    t.font = F_BASE
    t.alignment = Alignment(vertical='center', wrap_text=True)
    t.border = BOX
    if it['peers']:
        ws.cell(row=row, column=2).font = Font(name='微软雅黑', size=10, color='6D28D9')

    d = ws.cell(row=row, column=3, value=it['desc'])
    d.font = Font(name='微软雅黑', size=9, color='6B7280')
    d.alignment = Alignment(vertical='center', wrap_text=True)
    d.border = BOX

    p = ws.cell(row=row, column=4, value=it['pts'] if it['kind'] != 'minus' else it['pts'])
    p.font = Font(name='微软雅黑', size=10, bold=True, color='B91C1C' if it['kind'] == 'minus' else '047857')
    p.alignment = Alignment(horizontal='center', vertical='center')
    p.border = BOX

    for j in range(5, 12):
        c = ws.cell(row=row, column=j, value='')
        c.border = BOX
        c.alignment = Alignment(horizontal='center', vertical='center')
        if it['kind'] == 'minus':
            c.fill = NAG_FILL
    if it['kind'] == 'minus':
        ws.cell(row=row, column=2).fill = NAG_FILL
        ws.cell(row=row, column=3).fill = NAG_FILL
        ws.cell(row=row, column=4).fill = NAG_FILL
    ws.row_dimensions[row].height = 30
    row += 1

ws.cell(row=row, column=1, value='本周合计').font = F_BOLD
ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=4)
for j in range(1, 5):
    ws.cell(row=row, column=j).fill = PatternFill('solid', fgColor='FBFDFF')
    ws.cell(row=row, column=j).border = BOX
for j in range(5, 12):
    c = ws.cell(row=row, column=j, value=None)
    c.border = BOX
    c.fill = PatternFill('solid', fgColor='FBFDFF')

widths = {'A': 10, 'B': 30, 'C': 34, 'D': 6, 'E': 6, 'F': 6, 'G': 6, 'H': 6, 'I': 6, 'J': 6, 'K': 6, 'L': 6}
for col, w in widths.items():
    ws.column_dimensions[col].width = w
ws.freeze_panes = 'E5'

# ---------- Sheet2 作息 ----------
ws2 = wb.create_sheet('作息时间')
ws2.sheet_view.showGridLines = False
ws2['A1'] = '一日作息（普通日 / 延时课日）'
ws2['A1'].font = F_TITLE
ws2.merge_cells('A1:D1')
ws2['A3'] = '普通日（周二、周四、周五）'
ws2['C3'] = '延时课日（周一、周三 · 17:30 结束）'
for cell in ('A3', 'C3'):
    ws2[cell].font = F_HEAD
    ws2[cell].fill = HEAD_FILL
    ws2[cell].alignment = Alignment(horizontal='center')
ws2.merge_cells('A3:B3')
ws2.merge_cells('C3:D3')

normal = [
    ('06:50', '起床 + 自己穿衣', '闹钟响第一声就起，家长只提醒一次'),
    ('07:00', '洗漱 + 早餐', '20 分钟内吃完，不看屏幕'),
    ('07:35', '出门前检查书包', '作业 / 水壶 / 校卡，自己核对清单'),
    ('07:55', '到校', '进校门前和家长说再见'),
    ('15:55', '放学 → 到家', '先洗手喝水，15 分钟放空'),
    ('17:00', '户外 / 跳绳 / 球类', '先把身体电量放掉，再谈作业'),
    ('18:00', '晚餐', '30 分钟内吃完，桌上不摆玩具'),
    ('18:40', '亲子聊 10 分钟', '今天最开心 / 最难的一件事'),
    ('19:00', '作业时间', '番茄钟 10 分钟 + 休息 2 分钟，先做难的'),
    ('19:40', '阅读 / 朗读', '自主阅读 30 分钟或朗读 10 分钟'),
    ('20:10', '自由时间 / 兑换的游戏', '计时器响就结束'),
    ('20:40', '洗漱 + 整理明日书包', '自己完成，家长只检查'),
    ('21:00', '床上聊天 + 情绪复盘', '今天红/黄/绿哪个时候最多？为什么'),
    ('21:30', '熄灯', '手机 / iPad 已交给家长'),
]
late = [
    ('06:50', '起床 + 自己穿衣', '闹钟响第一声就起，家长只提醒一次'),
    ('07:00', '洗漱 + 早餐', '20 分钟内吃完，不看屏幕'),
    ('07:35', '出门前检查书包', '作业 / 水壶 / 校卡，自己核对清单'),
    ('07:55', '到校', '进校门前和家长说再见'),
    ('15:55', '放学 → 延时课', '校内课后服务，16:30 开始'),
    ('17:30', '延时课结束', '接到后给 10 分钟缓冲，不立刻问作业'),
    ('17:45', '到家 + 户外 15 分钟', '延时日缩短户外时间'),
    ('18:00', '晚餐', '30 分钟内吃完，桌上不摆玩具'),
    ('18:40', '亲子聊 10 分钟', '今天最开心 / 最难的一件事'),
    ('19:10', '作业时间', '番茄钟 10 分钟 + 休息 2 分钟'),
    ('19:50', '阅读 / 朗读', '自主阅读或朗读，二选一'),
    ('20:20', '自由时间 / 兑换的游戏', '计时器响就结束'),
    ('20:45', '洗漱 + 整理明日书包', '自己完成，家长只检查'),
    ('21:00', '床上聊天 + 情绪复盘', '今天红/黄/绿哪个时候最多？为什么'),
    ('21:30', '熄灯', '延时日不额外加作业'),
]
r2 = 4
for i in range(max(len(normal), len(late))):
    if i < len(normal):
        ws2.cell(row=r2 + i, column=1, value=normal[i][0]).font = Font(name='微软雅黑', size=10, bold=True, color='1D4ED8')
        ws2.cell(row=r2 + i, column=2, value=normal[i][1] + '｜' + normal[i][2]).font = F_BASE
    if i < len(late):
        ws2.cell(row=r2 + i, column=3, value=late[i][0]).font = Font(name='微软雅黑', size=10, bold=True, color='1D4ED8')
        ws2.cell(row=r2 + i, column=4, value=late[i][1] + '｜' + late[i][2]).font = F_BASE
    for j in range(1, 5):
        ws2.cell(row=r2 + i, column=j).border = BOX
        ws2.cell(row=r2 + i, column=j).alignment = Alignment(vertical='center', wrap_text=True)
    ws2.row_dimensions[r2 + i].height = 28
ws2.column_dimensions['A'].width = 9
ws2.column_dimensions['B'].width = 46
ws2.column_dimensions['C'].width = 9
ws2.column_dimensions['D'].width = 46

# ---------- Sheet3 奖励 ----------
ws3 = wb.create_sheet('奖励兑换')
ws3.sheet_view.showGridLines = False
ws3['A1'] = '积分奖励兑换表'
ws3['A1'].font = F_TITLE
ws3.merge_cells('A1:E1')
ws3['A2'] = '10 积分 = 1 元零花钱；屏幕时间固定在 20:10 之前，睡前 1 小时不给屏幕。'
ws3['A2'].font = Font(name='微软雅黑', size=9.5, color='6B7280')
ws3.merge_cells('A2:E2')
h3 = ['图标', '奖励', '所需积分', '每日上限', '说明']
for i, h in enumerate(h3):
    c = ws3.cell(row=4, column=i + 1, value=h)
    c.font = F_HEAD
    c.fill = HEAD_FILL
    c.alignment = Alignment(horizontal='center')
    c.border = BOX
for i, rw in enumerate(rewards):
    vals = [rw['icon'], rw['title'], rw['cost'], (rw['daily'] + ' 次') if rw['daily'] != '0' else '长期', rw['note']]
    for j, v in enumerate(vals):
        c = ws3.cell(row=5 + i, column=j + 1, value=v)
        c.font = F_BASE
        c.border = BOX
        c.alignment = Alignment(horizontal='center' if j in (0, 2, 3) else 'left', vertical='center', wrap_text=True)
    ws3.row_dimensions[5 + i].height = 30
ws3.column_dimensions['A'].width = 6
ws3.column_dimensions['B'].width = 30
ws3.column_dimensions['C'].width = 10
ws3.column_dimensions['D'].width = 10
ws3.column_dimensions['E'].width = 40

# ---------- Sheet4 说明 ----------
ws4 = wb.create_sheet('使用说明')
ws4.sheet_view.showGridLines = False
ws4['A1'] = '使用说明（家长页）'
ws4['A1'].font = F_TITLE
ws4.merge_cells('A1:B1')

notes = [
    ('一、这张表怎么用', '每天只挑 3–6 项重点打卡，不追求全部打满。全打满会让孩子把注意力放在「挣分」上，而不是行为本身。'),
    ('', '每天固定两个时间点：早上出门前 1 分钟对清单，晚上 21:00 床上复盘 2 分钟（红/黄/绿情绪 + 今天最棒的一件事）。'),
    ('', '周一、周三延时课，作业往后挪 40 分钟，当天不强求阅读与运动同时完成。'),
    ('二、为什么这样设计', '情绪不稳、社交理解偏弱的孩子，需要「看得见的规则 + 可预测的流程」，而不是临时的说教。表格的作用是把抽象的期待变成具体的动作。'),
    ('', '「情绪社交」类的 11 项参考 PEERS 训练思路：情绪命名 → 冷静策略 → 加入同伴 → 接受拒绝 → 冲突修复 → 求助表达。每一项都是可以当场练习的动作，不是性格评价。'),
    ('', '「需要改进」的 5 项是行为记录，不是人格标签。扣分只记录事实（发生了什么），不评价孩子是好是坏。'),
    ('三、几条硬规则', '积分当天有效，不跨天累积负债；扣分最低到 0，不倒扣成负数（页面里可以手动调整，但建议不用）。'),
    ('', '奖励要当场兑现，拖延兑换会让孩子不相信这套规则。零花钱按 10 分 = 1 元，当天最多 2 元。'),
    ('', '屏幕时间（电视 / 罗小黑 / 我的世界）只在 20:10 的自由时间使用，睡前 1 小时一律不给。'),
    ('', '家长不因为情绪失控临时加码或取消奖励，规则改了要提前说，不在冲突当下改。'),
    ('四、随年龄怎么调', '一年级（现在）：项目少一点、动作具体一点，重点是建立「做完就有反馈」的循环，分值以 1–2 分为主。'),
    ('', '二三年级：合并同类项（例如「作业独立完成」合并成一项），把「预习复习」换成「自己检查作业」，增加自我管理类项目。'),
    ('', '四年级以上：由孩子自己定每周 3 个目标，家长只做记录和复盘；积分改为周结算，加入「自己管理零花钱」的记账。'),
    ('', '青春期前后：重点从「打卡得分」转到「协商规则」，扣分项逐步取消，改成家庭会议讨论后果。'),
    ('五、需要注意', '这套表不能替代专业评估。若情绪爆发频率上升、影响上学或睡眠，应继续跟儿童精神科 / 发育行为科随访，必要时配合行为治疗。'),
    ('', '不要把积分表和「你表现好我才爱你」绑定。积分只兑换具体的东西，不兑换关注和情感。'),
]
ws4.cell(row=4, column=1, value='主题').font = F_HEAD
ws4.cell(row=4, column=2, value='内容').font = F_HEAD
for c in (ws4['A4'], ws4['B4']):
    c.fill = HEAD_FILL
    c.border = BOX
    c.alignment = Alignment(horizontal='center')
rr = 5
for a, b in notes:
    ca = ws4.cell(row=rr, column=1, value=a)
    cb = ws4.cell(row=rr, column=2, value=b)
    ca.font = F_BOLD
    ca.alignment = Alignment(vertical='top', wrap_text=True)
    cb.font = F_BASE
    cb.alignment = Alignment(vertical='top', wrap_text=True)
    for c in (ca, cb):
        c.border = BOX
    ws4.row_dimensions[rr].height = 46
    rr += 1
ws4.column_dimensions['A'].width = 18
ws4.column_dimensions['B'].width = 86

wb.save(OUT)
print('saved:', OUT)
print('items:', len(items), 'rewards:', len(rewards))
