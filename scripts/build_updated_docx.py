"""Build the Word companion from the same JSON/CSV evidence as main.tex."""
from pathlib import Path
import json, sys
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH

sys.path.insert(0, '/Users/xingyu/.codex/skills/math-modeling/tools/docx/scripts')
import paper_format as pf

ROOT = Path(__file__).resolve().parents[1]
summary = json.loads((ROOT / 'results/summary.json').read_text(encoding='utf-8'))
a, q1 = summary['aggregate'], summary['q1']
doc = pf.new_document(contest='cumcm')
pf.title(doc, '微网与外部电网电力调控策略研究（更新稿）')
pf.abstract_title(doc)
pf.body(doc, (
    f'本文研究固定/波动电价和逐步更新光伏信息下的微网购电与储能调度。'
    f'统一十分钟能量平衡模型显式区分常规购电、紧急购电、充放电和弃光，'
    f'并检查荷电状态、功率和跨日状态。问题一用互斥MILP，另以5 kWh有限网格库存价值DP交叉核验；'
    f'问题二至问题四使用日期因果预测。问题一购电量{q1["purchase_kwh"]:.3f} kWh，'
    f'费用{q1["objective"]:.3f}元；Q2、Q3、Q4-2、Q4-3全年本地回测费用分别为'
    f'{a["q2"]["cost"]:.2f}、{a["q3"]["cost"]:.2f}、{a["q4"]["cost"]:.2f}、{a["q4_3"]["cost"]:.2f}元。'
    '所有数字均由当前输入、代码和结果文件生成，不代表官方成绩；表1、表2、表3给出符号、全年结果和敏感性。'))
pf.keywords(doc, '微网；储能调度；混合整数线性规划；库存价值；滚动调整；动态电价')

def H(t): pf.heading1(doc, t)
def p(t): pf.body(doc, t)
def fig(name, cap):
    para = doc.add_paragraph(); para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    para.add_run().add_picture(str(ROOT / 'figures' / f'{name}.png'), width=Inches(5.8))
    c = doc.add_paragraph(cap); c.alignment = WD_ALIGN_PARAGRAPH.CENTER

H('一、问题与数据')
p('题面要求在十分钟粒度安排常规购电、紧急购电和储能充放电。附件1给出单日固定价格、负荷和光伏预测；附件2给出全年实际负荷与光伏；附件3给出每日0、6、12、18时的24小时光伏预报；附件4给出波动价格。全年回测只取2025-02-01至2025-12-31的334天，目标日预测只读取严格早于目标日的数据。题面与规则以官方材料[1]为准。')
p('如图1所示，原始数据统一到十分钟区间。')
fig('raw_q1_inputs', '图1  问题一输入序列')
p('图1用于核对十分钟采样、功率单位和日内光伏形状。')

H('二、统一物理模型')
p('令q_t、e_t、c_t、d_t、u_t分别表示常规购电、紧急购电、交流侧充电、交流侧放电和未利用供能，单位为kWh；L_t、G_t为功率，p_t为元/kWh，Δt=1/6 h。采用充电入库ηc_t、放电出库d_t/η，η=0.9。')
pf.equation(doc, r'q_t+e_t+d_t+G_t\Delta t=L_t\Delta t+c_t+u_t')
pf.equation(doc, r'S_{t+1}=S_t+\eta c_t-d_t/\eta')
pf.equation(doc, r'0\le c_t,d_t\le 5000\Delta t,\quad1200\le S_t\le10800,\quad S_0=6000')
p('问题一补充S_T=S_0和互斥二元变量；Q2至Q4跨日传递上一日末SOC。每次求解后执行非负性、功率、SOC和初始状态审计。模型结构参考微网优化文献[2]和Pyomo建模文献[3]。如图2所示，SOC轨迹接受边界审计。')
p('统一符号见表1。')
doc.add_paragraph('表1  统一符号与单位')
pf.three_line_table(doc, [['符号','含义','单位'],['q_t,e_t','常规/紧急购电量','kWh'],['c_t,d_t','交流侧充电/放电量','kWh'],['S_t','内部储能量','kWh'],['L_t,G_t','负荷/光伏功率','kW'],['p_t','交易价格','元/kWh']])
fig('process_q1_soc', '图2  问题一充放电与SOC')
p('图2显示互斥充放电与SOC边界。')

H('三、问题一与库存价值交叉核验')
p(f'问题一MILP得到全天购电量{q1["purchase_kwh"]:.3f} kWh、费用{q1["objective"]:.3f}元，SOC范围为{q1["soc_min"]:.0f}至{q1["soc_max"]:.0f} kWh，日末为{q1["terminal_soc"]:.0f} kWh。5 kWh状态网格DP的值为{q1["dp_grid_value"]:.3f}元，相对MILP差异{q1["dp_relative_gap_pct"]:.4f}%，该差异是离散网格近似误差。如图3所示，购电量随价格和净负荷变化。')
pf.equation(doc, r'F_t(S)=\min_{S\to S1}\{p_t(N_t+c_t-d_t)^++F_{t+1}(S1)\},\quad F_T(6000)=0')
fig('result_q1_dispatch', '图3  问题一价格与购电')
p('图3给出价格和购电量的时序对应。')

H('四、问题二：日期因果回测')
p(f'负荷预测将最近同类型日的日电量水平与归一化日内形状分别取中位数，再恢复为功率；光伏采用目标日前三日中位数并乘0.9。因果诊断显示，负荷MAPE为6.153%，逐时七日中位数为17.987%；光伏三日和七日中位数MAPE分别为4.027%和4.697%。Q2全年费用为{a["q2"]["cost"]:.2f}元，紧急购电量为{a["q2"]["emergency_kwh"]:.3f} kWh，见图4。')
pf.equation(doc, r'Lhat_{d,t}=Bhat_d s_hat_{d,t}/\Delta t,\quad\sum_t s_hat_{d,t}\Delta t=1')
fig('raw_q2_q3_q4_emergency', '图4  各场景全年紧急购电量')
p('图4用于比较预测误差在全年紧急购电上的传导。')

H('五、问题三：多时刻预报与B结算')
p('0、6、12、18时刻只重优化未执行的时间段，并锁定已执行前缀。对0时计划q^0和调整计划q^a，账单按计划基准费、0.5倍调减费、1.5倍调增费和5倍紧急费四项相加。Q3全年费用为{:.2f}元，紧急购电量为{:.3f} kWh；各分项和调整量保存在daily_metrics.csv，见图5。'.format(a['q3']['cost'], a['q3']['emergency_kwh']))
pf.equation(doc, r'K_3=\sum_t p_tq_t^0+\sum_t0.5p_t(q_t^0-q_t^a)^++\sum_t1.5p_t(q_t^a-q_t^0)^++\sum_t5p_te_t')
pf.equation(doc, r'e_t=\left[L_t\Delta t-G_t\Delta t-q_t-c_t+d_t\right]^+')
fig('process_q3_adjustment', '图5  计划—调整差额与紧急购电')
p('图5展示滚动调整量与紧急购电量的关系。')

H('六、问题四：波动电价的独立分支')
p('Q4-2将附件4价格替换固定价格并沿用Q2信息边界；Q4-3沿用Q3滚动预报和B结算。两个分支从每天相同的日初SOC独立启动，不把Q4-2的日末状态带入Q4-3。Q4-2费用为{:.2f}元、紧急购电{:.3f} kWh；Q4-3费用为{:.2f}元、紧急购电{:.3f} kWh，见图6和图7。'.format(a['q4']['cost'],a['q4']['emergency_kwh'],a['q4_3']['cost'],a['q4_3']['emergency_kwh']))
p('表2汇总四个场景的全年结果。')
doc.add_paragraph('表2  全年本地回测结果')
pf.three_line_table(doc, [['场景','费用（元）','紧急购电（kWh）'],['Q2',f'{a["q2"]["cost"]:.2f}',f'{a["q2"]["emergency_kwh"]:.3f}'],['Q3',f'{a["q3"]["cost"]:.2f}',f'{a["q3"]["emergency_kwh"]:.3f}'],['Q4-2',f'{a["q4"]["cost"]:.2f}',f'{a["q4"]["emergency_kwh"]:.3f}'],['Q4-3',f'{a["q4_3"]["cost"]:.2f}',f'{a["q4_3"]["emergency_kwh"]:.3f}']])
fig('result_q4_cost_box', '图6  各场景日费用分布')
p('图6给出不同信息和价格分支的日费用分布。')
fig('result_q4_3_series', '图7  问题4-3日费用序列')
p('图7给出波动价格和滚动调整下的逐日费用。')

H('七、稳健性、局限与复现')
p('本版本对经验材料进行了反向检查：经验材料的数字不直接使用；情景路径知道未来时只作为规划近似，不作为因果执行证据；LP互斥结论只在其假设成立时使用；Q3账单和Q4分支状态均逐日审计。有限网格DP不替代MILP，Q2至Q4的结果是本地附件回测而非官方成绩。图8给出本地回测的费用/紧急量对照。运行命令为 ../.venv/bin/python solve_c.py --seed 20260911；结果、输入哈希、预测诊断和DP对照分别见results目录。')
fig('process_q2_overview', '图8  本地回测费用与紧急量对照')
p('表3给出保守光伏系数的敏感性。')
doc.add_paragraph('表3  光伏保守系数敏感性（Q2，日电量—形状负荷预测）')
pf.three_line_table(doc, [['系数','全年费用（元）','紧急购电（kWh）'],['0.80','16622826.83','627200.633'],['0.90','16206550.45','789872.052'],['1.00','17440007.68','1333753.169']])
pf.equation(doc, r'\min\sum_t p_tq_t+5p_te_t')
H('AI工具使用声明')
p('Codex用于资料检索、代码组织、公式排版和结果复核；参赛队员负责题面解释、模型假设、参数选择、代码运行、结果解释和最终提交。')
H('参考文献')
p('[1] 2026年全国大学生数学建模竞赛题目与人工智能使用规定（以官方发布版本为准）。')
p('[2] Gulotta T 等. Microgrid energy management under uncertainty. International Journal of Electrical Power & Energy Systems, 2023.')
p('[3] Hart W E 等. Pyomo—optimization modeling in Python. Springer, 2017.')

out = pf.save_document(doc, ROOT, filename='完整论文.docx', contest='cumcm', overwrite=True)
print(out)
