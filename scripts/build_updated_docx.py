"""Build the competition-paper Word deliverable from the current run artifacts.

The manuscript is intentionally generated from summary/CSV/XLSX evidence so a
new solve run changes the reported numbers and tables together.
"""
from __future__ import annotations

import json
import sys
from datetime import datetime
from pathlib import Path

import openpyxl
import pandas as pd
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Inches, Pt

sys.path.insert(0, "/Users/xingyu/.codex/skills/math-modeling/tools/docx/scripts")
import paper_format as pf

ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / "results"
summary = json.loads((OUT / "summary.json").read_text(encoding="utf-8"))
agg = summary["aggregate"]
q1 = summary["q1"]
daily = pd.read_csv(OUT / "daily_metrics.csv")


def _date_key(value):
    if isinstance(value, datetime):
        return value.date().isoformat()
    return str(value)[:10].replace("/", "-")


def read_emergency_groups(filename: str, dates: list[str]):
    ws = openpyxl.load_workbook(OUT / filename, data_only=True, read_only=True)["紧急购电量"]
    result = {d: [] for d in dates}
    for row in ws.iter_rows(min_row=2, values_only=True):
        if row[0] is None or row[1] is None or row[2] is None:
            continue
        key = _date_key(row[0])
        if key in result:
            result[key].append((str(row[1]), float(row[2])))
    return result


def read_q1_selected():
    wb = openpyxl.load_workbook(OUT / "result1.xlsx", data_only=True, read_only=True)
    ws = wb["计划购电量"]
    wanted = {"10:00-10:10", "12:00-12:10", "14:00-14:10", "16:00-16:10", "18:00-18:10", "20:00-20:10"}
    purchase = {str(r[0]): float(r[1]) for r in ws.iter_rows(min_row=2, values_only=True) if r[0] in wanted}
    ws = wb["充放电量"]
    storage = []
    for r in ws.iter_rows(min_row=2, max_row=7, values_only=True):
        storage.append((str(r[0]), float(r[1] or 0), float(r[2] or 0)))
    return purchase, storage


q1_purchase, q1_storage = read_q1_selected()
specified_dates = ["2025-03-20", "2025-06-21", "2025-09-23", "2025-12-21"]
em2 = read_emergency_groups("result2.xlsx", specified_dates)
em3 = read_emergency_groups("result3.xlsx", specified_dates)

doc = pf.new_document(contest="cumcm")
pf.title(doc, "微网与外部电网电力调控策略研究")
pf.abstract_title(doc)
pf.body(
    doc,
    (
        "针对固定与波动电价、负荷和光伏信息逐步揭示条件下的微网购电与储能调度问题，"
        "本文在十分钟粒度建立统一的能量平衡模型，显式区分常规购电、紧急购电、充电、放电与弃光，"
        "并用储能状态递推、功率边界、容量边界和跨日状态约束保证物理可行。问题一采用带充放电互斥的混合整数线性规划，"
        "同时用5 kWh有限状态网格动态规划交叉核验；问题二把负荷预测拆为日电量水平和日内形状，并用目标日前三日光伏中位数乘0.9进行保守预测；"
        "问题三在0、6、12、18时锁定已执行前缀、滚动重优化，并按计划基准费、调减0.5倍、调增1.5倍和紧急费结算；"
        "问题四在附件4波动价格下独立重算问题二、问题三。问题一单日购电量为"
        f"{q1['purchase_kwh']:.3f} kWh、费用为{q1['objective']:.3f}元；"
        "2025年2月1日至12月31日的本地回测总费用分别为"
        f"Q2 {agg['q2']['cost']:.2f}元、Q3 {agg['q3']['cost']:.2f}元、"
        f"Q4-2 {agg['q4']['cost']:.2f}元和Q4-3 {agg['q4_3']['cost']:.2f}元。"
        "上述数值均由当前题目附件、代码和结果文件生成，不代表官方成绩。"
    ),
)
pf.keywords(doc, "微网；储能调度；混合整数线性规划；库存价值；滚动优化；光伏预测；动态电价")


def H(text):
    pf.heading1(doc, text)


def h(text):
    pf.heading2(doc, text)


def p(text):
    pf.body(doc, text)


def fig(stem, caption, width=5.8):
    para = doc.add_paragraph()
    para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    para.add_run().add_picture(str(ROOT / "figures" / f"{stem}.png"), width=Inches(width))
    cap = doc.add_paragraph(caption)
    cap.alignment = WD_ALIGN_PARAGRAPH.CENTER


def compact_table(rows, font_size=10):
    """Keep small result tables together without shrinking the whole paper."""
    table = pf.three_line_table(doc, rows)
    for row in table.rows:
        for cell in row.cells:
            for paragraph in cell.paragraphs:
                paragraph.paragraph_format.space_before = Pt(0)
                paragraph.paragraph_format.space_after = Pt(0)
                paragraph.paragraph_format.line_spacing = 1.0
                for run in paragraph.runs:
                    run.font.size = Pt(font_size)
    return table


H("一、问题重述与数据审计")
h("1.1 题目任务")
p(
    "题目要求在十分钟时间粒度制定微网的常规购电、紧急购电以及储能充放电策略。"
    "问题一给出单日价格、负荷和光伏预测，并要求日初、日末储能相同；问题二扩展到全年实际负荷和光伏，"
    "要求每天0:00制定计划，预测偏差由紧急购电补足；问题三在0、6、12、18时发布未来24小时整点光伏预报，"
    "允许调整尚未执行的计划；问题四把固定价格替换为附件4的波动价格，分别重算问题二和问题三。"
)
h("1.2 数据口径与信息边界")
p(
    "附件1包含144个十分钟记录，作为问题一的价格、负荷和光伏预测；附件2包含2025年365天的实际负荷和光伏；"
    "附件3包含每日四次、每次24个整点值的光伏预报；附件4包含每日144个十分钟价格。"
    "为避免未来信息泄漏，全年评估严格取2025-02-01至2025-12-31的334天，目标日预测仅读取严格早于目标日的行。"
    "功率乘以Δt=1/6 h后进入能量平衡，价格与电量相乘得到费用；时间标签0:00+1表示次日0:00。"
)
doc.add_paragraph("表1  数据文件与使用方式")
pf.three_line_table(
    doc,
    [
        ["文件", "内容", "模型使用"],
        ["附件1.xlsx", "单日价格、负荷、光伏预测", "问题一；固定价格基准"],
        ["附件2.xlsx", "全年实际负荷、光伏", "Q2—Q4实际结算与回测"],
        ["附件3.xlsx", "每日0/6/12/18时光伏预报", "Q3、Q4-3滚动更新"],
        ["附件4.xlsx", "全年十分钟波动价格", "Q4-2、Q4-3价格"],
    ],
)
h("1.3 信息集与预处理流程")
p(
    "四个子问题共享能量平衡和储能状态方程，但决策时可见的信息不同。问题一把附件1中的单日负荷、光伏预测和价格视为已知；"
    "问题二在每天0:00只使用目标日前的历史样本生成计划，实际负荷与实际光伏留到事后结算；问题三在0、6、12、18时接收新光伏预报，"
    "只重算尚未执行的区间；问题四只替换价格序列，分别沿用问题二、问题三的信息边界。这样处理可以把预测误差、信息更新时间和价格变化的影响分开。"
)
p(
    "数据读取先把Excel日期序列转换为真实日期，并检查全年日期连续、每天144个十分钟记录。附件3同一天的四条预报记录可能只有首行带日期，"
    "读取时沿用该日期；24个整点预测按小时中心线性插值到十分钟中心，端点使用最近值。输入审计显示，附件1有144条记录和432个数值字段，"
    "附件2和附件4均为365×144个功率或价格记录，附件3为365×4×24个预测值；检查未发现缺失或非数值单元格。"
)
p(
    "全年回测按结果模板取2025-02-01至2025-12-31的334天。所有功率在进入能量平衡前乘以Δt，所有费用由同一时段的价格乘购电量计算；"
    "原始附件不被改写，字段、范围和输入哈希由scripts/validate_input_xml.py记录，结果文件由当前代码写入results目录。"
)
fig("raw_q1_inputs", "图1  问题一输入序列（负荷、光伏预测与十分钟时间轴）", width=4.8)
p("图1只用于核对问题一输入的时间轴、负荷与光伏预测的相对变化，不把曲线形状直接当作全年实际运行结论。竞赛规则和论文格式以官方规范[1]为准。")
h("1.4 数据审计结论")
p(
    "审计结果表明，附件的主要风险不在缺失值，而在时间索引、功率与电量的换算以及预测信息的可得性。"
    "因此程序先按真实日期和十分钟序号建立索引，再执行功率乘Δt；结果模板中的‘0:00-0:10+1’只作为导出标签，"
    "不参与内部状态递推。附件3日期空白行只做同日标签的前向填充，不对预测数值作插值外的人工修正。"
)
p(
    "上述检查不能证明预测本身准确，只能证明输入结构完整、单位转换有迹可循。预测误差由问题二和问题三的因果回测衡量，"
    "实际负荷与实际光伏只在结算阶段进入误差和费用计算；若把实际值提前送入日内优化，会得到不可复现的完美信息结果。"
)
p(
    "由此，后文统一采用‘预测用于决策、实际值用于回放’的双轨记录：每个策略同时保存输入信息集、求解计划、实际能量平衡、"
    "紧急购电和费用分项。这样既能比较储能策略，也能把预测改进、预报更新时间和价格替换带来的收益分开归因。"
)

H("二、模型假设、符号与统一约束")
h("2.1 建模假设")
p(
    "作如下约定：①微网可从外网购电，光伏优先用于负荷或储能，富余光伏允许弃用；②储能只受题目给出的容量、"
    "充放电功率和效率约束，不考虑电池退化与自放电；③每个十分钟区间内功率保持不变；④紧急购电不参与计划优化，"
    "而是在实际数据到达后按剩余缺口结算；⑤Q3和Q4-3每次只重优化尚未执行区间，已执行前缀保持不变。"
)
p(
    "第一条假设把光伏定义为无边际成本的本地供能，并用弃光量u_t保证在光伏大于负荷且储能受限时仍有可行的能量平衡；"
    "题目没有给出上网电价，因此不把余电外送收益混入目标函数。第二、三条假设固定了状态量和控制量的时间口径，"
    "使每一步都能在同一组功率上限和容量边界下比较。"
)
p(
    "第四、五条假设规定了回测的先后顺序：先用当时可见的预测生成计划，再用实际负荷和实际光伏计算缺口与账单；"
    "预报更新只改变尚未执行的后缀，已经发生的功率和储能状态不能被事后改写。这样得到的费用才对应一个可以在线执行的策略。"
)
h("2.2 变量与能量平衡")
p(
    "令t=0,…,T−1表示十分钟时段，q_t、e_t、c_t、d_t、u_t分别表示常规购电、紧急购电、交流侧充电、交流侧放电和弃光量（kWh）；"
    "L_t、G_t表示负荷和光伏功率（kW），p_t表示价格（元/kWh），S_t表示时段起点储能量（kWh），η=0.9。"
)
pf.equation(doc, r"q_t+e_t+d_t+G_t\Delta t=L_t\Delta t+c_t+u_t")
pf.equation(doc, r"S_{t+1}=S_t+\eta c_t-d_t/\eta")
pf.equation(doc, r"0\le c_t,d_t\le5000\Delta t,\quad1200\le S_t\le10800,\quad S_0=6000")
doc.add_paragraph("表2  主要符号与单位")
pf.three_line_table(
    doc,
    [
        ["符号", "含义", "单位"],
        ["q_t,e_t", "常规/紧急购电量", "kWh"],
        ["c_t,d_t", "充电量/放电量", "kWh"],
        ["u_t", "未利用光伏量", "kWh"],
        ["S_t", "储能状态", "kWh"],
        ["L_t,G_t", "负荷/光伏功率", "kW"],
        ["p_t", "外网交易价格", "元/kWh"],
    ],
)
p(
    "能量平衡式左侧是可供给能量：常规购电、紧急购电、储能放电和光伏；右侧是负荷、储能充电以及弃光。"
    "因此光伏先抵消负荷，剩余部分只能进入储能或u_t；光伏不足时由储能放电、常规购电或紧急购电补齐。"
    "状态方程采用交流侧充放电口径：输入c_t中只有η比例存入电池，输出d_t需要从电池取出d_t/η。"
)
p(
    "由于Δt=1/6 h，5000 kW的功率上限对应每个十分钟最多833.333 kWh的充电或放电量。"
    "S_t是时段起点状态，因而共有T+1个状态点；问题一另外固定S_T=S_0，全年回测则把上一日末状态传给下一日初始状态。"
)
h("2.3 目标函数与审计")
p(
    "固定价格场景以常规购电费和紧急购电费最小为目标；问题一令e_t=0并补充S_T=S_0。"
    "问题三和问题4-3把0:00计划与后续调整分开计费。每次求解后检查能量平衡、变量非负性、充放电功率、SOC边界、初始状态和端点状态；"
    "Q3与Q4-3再检查四项账单之和是否等于日费用。"
)
h("2.4 互斥约束与求解接口")
p(
    "充放电互斥是储能模型中唯一的离散动作。问题一用二元变量z_t把它线性化："
    "c_t≤5000Δt z_t、d_t≤5000Δt(1−z_t)，其中大M取题目给出的单步功率上限，避免人为放大可行域。"
    "问题一因此采用MILP；问题二至问题四需要对334天反复滚动求解，采用连续LP并在目标中加入极小的充放电正则项，"
    "每次求解后记录充放电重叠量和SOC审计结果。若扩展模型引入退费、负价或其他耦合成本，LP不再自动保证互斥，必须重新启用二元约束。"
)
h("2.5 四问共享的输入—输出接口")
p(
    "四个问题都通过同一个调度接口接收负荷、光伏、价格和日初SOC，输出q_t、c_t、d_t、u_t、e_t及SOC轨迹。"
    "差别只体现在输入信息集和账单：问题一输入单日已知预测并强制日末回到初始SOC；问题二按日提前计划并在实际回放时补紧急电；"
    "问题三保留0:00计划基准、只对后缀滚动调整；问题四替换价格序列并保持前两问的信息边界。统一接口使不同策略的差异可以归因于预测、更新时间或价格，而不是代码分支的隐含口径。"
)
fig("process_q1_soc", "图2  问题一储能充放电与SOC边界审计", width=4.3)
p("图2同时核对充放电量、储能状态和上下界；尖峰对应单步功率上限，黑线的最低点和最高点分别检验1200 kWh与10800 kWh边界。")

H("三、问题一：单日确定性调度")
h("3.1 混合整数线性规划")
p(
    "在附件1预测值被视为已知的条件下，问题一只需为144个时段安排常规购电和储能动作。"
    "目标函数为"
)
pf.equation(doc, r"\min\sum_{t=0}^{143}p_tq_t")
p(
    "这个目标只为常规购电付费，因为问题一不设置紧急购电，也没有上网收益；日末SOC等于日初SOC则防止模型通过耗尽储能换取一次性低费用。"
    "在每个时段，储能只有在当前购电成本与未来购电成本比较后才值得充放电，因而q_t的零值本身也是储能机会成本作用的结果。"
)
p(
    "并以二元变量z_t施加c_t≤5000Δt z_t、d_t≤5000Δt(1−z_t)，从而消除同一时段同时充放电。"
    "大M直接取单步功率上限，既保留可行动作又不人为放大变量范围。求解器采用SciPy HiGHS的MILP接口，"
    "结果再通过能量平衡、SOC边界和端点条件复核；滚动优化与不确定性建模的写法参考文献[2-4]。"
)
h("3.2 库存价值动态规划交叉核验")
p(
    "为独立检查储能跨期机会成本，在1200—10800 kWh范围内按5 kWh离散状态，共得到1921个候选状态。"
    "令N_t=(L_t−G_t)Δt表示不考虑储能时该时段的净购电需求，则"
)
pf.equation(doc, r"F_t(S)=\min_{S\to S'}\{p_t\,[N_t+c_t-d_t]^++F_{t+1}(S')\},\quad F_T(6000)=0")
p(
    f"MILP费用为{q1['objective']:.3f}元，有限网格DP初值为{q1['dp_grid_value']:.3f}元，相对差异{q1['dp_relative_gap_pct']:.4f}%。"
    "递推从日末边界向前计算每个SOC的最低未来购电费，差异来自5 kWh状态离散和连续MILP之间的取整；提交轨迹采用MILP，DP只承担独立的数值交叉检查。"
)
doc.add_paragraph("表3  问题一主要结果")
pf.three_line_table(
    doc,
    [
        ["指标", "MILP", "5 kWh网格DP"],
        ["全天费用/元", f"{q1['objective']:.3f}", f"{q1['dp_grid_value']:.3f}"],
        ["全天常规购电/kWh", f"{q1['purchase_kwh']:.3f}", "—"],
        ["SOC最小/最大/kWh", f"{q1['soc_min']:.0f}/{q1['soc_max']:.0f}", "—"],
        ["日末SOC/kWh", f"{q1['terminal_soc']:.0f}", "6000（边界条件）"],
    ],
)
doc.add_paragraph("表4  问题一指定十分钟时段购电量")
compact_table(
    [["时间段", "购电量/kWh"]] + [[k, f"{q1_purchase[k]:.3f}"] for k in ["10:00-10:10", "12:00-12:10", "14:00-14:10", "16:00-16:10", "18:00-18:10", "20:00-20:10"]] + [["全天", f"{q1['purchase_kwh']:.3f}"]],
)
doc.add_page_break()
doc.add_paragraph("表5  问题一指定四小时储能充放电量")
compact_table([["时间段", "充电量/kWh", "放电量/kWh"]] + [[a, f"{b:.3f}", f"{c:.3f}"] for a, b, c in q1_storage] + [["0:00 / 24:00储电量", f"{q1['terminal_soc']:.0f}", f"{q1['terminal_soc']:.0f}"]])
p("表4给出题面要求的十分钟购电量，表5把连续轨迹压缩成六个四小时区间，便于检查充放电高峰是否与价格低谷和SOC边界一致。")
fig("result_q1_dispatch", "图3  问题一价格与常规购电量", width=4.8)
p("图3中橙线为常规购电量、蓝线为价格；橙线在低价时段并不必然为零，因为储能还受到容量、功率和日末SOC约束。")

H("四、问题二：全年因果预测与日计划")
h("4.1 负荷与光伏预测")
p(
    "问题二每天0:00只能使用目标日前已经发生的样本。负荷预测先把历史同类型日分成日电量水平和归一化日内形状："
    "对目标日d取最近至多8个同星期样本；若同类样本少于3天，则退回最近8天。分别对日电量和形状取中位数并重新归一化，最后恢复到十分钟功率："
)
pf.equation(doc, r"\hat L_{d,t}=\hat B_d\hat s_{d,t}/\Delta t,\quad\sum_t\hat s_{d,t}\Delta t=1")
p(
    "光伏预测取目标日前三日实际功率的逐时中位数，再乘0.9作为计划侧保守系数；这个系数只改变计划输入，实际光伏仍在回放时使用。"
    "表6比较日电量—形状负荷预测与逐时七日中位数基线，以及三日、七日光伏中位数。MAE和偏差单位为kW，MAPE按实际值设置小分母下限，避免夜间光伏接近0时被少数点主导。"
)
diag = summary["forecast_diagnostics"]
by_series = {r["series"]: r for r in diag}
doc.add_paragraph("表6  因果预测诊断（334天平均）")
pf.three_line_table(
    doc,
    [["序列", "MAE", "MAPE/%", "偏差"]]
    + [[name, f"{by_series[name]['mae']:.3f}", f"{by_series[name]['mape_pct']:.3f}", f"{by_series[name]['bias']:.3f}"] for name in ["load_level_shape", "load_median7", "pv_median3", "pv_median7"]],
)
p(
    f"表6显示，日电量—形状分解把负荷MAE从{by_series['load_median7']['mae']:.1f} kW降至{by_series['load_level_shape']['mae']:.1f} kW，MAPE从{by_series['load_median7']['mape_pct']:.2f}%降至{by_series['load_level_shape']['mape_pct']:.2f}%，且平均偏差接近0。"
    f"三日光伏中位数的MAPE和偏差分别为{by_series['pv_median3']['mape_pct']:.2f}%和{by_series['pv_median3']['bias']:.2f} kW，七日中位数的MAE略低但响应更慢，因此计划采用三日窗口。"
    "这里的比较是在同一334天回测窗口上的诊断，不能当作独立测试集上的泛化证明；正式应用还应采用嵌套滚动验证。"
)
h("4.2 计划执行与紧急结算")
p(
    "对每个评估日，先用预测负荷、保守光伏和固定电价求解计划，日末SOC不强制回到日初而是传给下一天。"
    "计划模型中的紧急购电只作为带5倍价格惩罚的可行性松弛；实际执行时固定q、c、d，重新代入实际负荷和实际光伏计算缺口。若实际负荷减去实际光伏、常规购电和储能放电后仍有缺口，则"
)
pf.equation(doc, r"e_t=[L_t\Delta t-G_t\Delta t-q_t-c_t+d_t]^+")
p(
    "相应的Q2账单为常规购电费与5倍价格紧急费之和。这样把预测阶段的决策和事后阶段的结算分开，避免把实际值提前泄漏到计划。"
)
p(
    f"Q2全年费用为{agg['q2']['cost']:.2f}元，紧急购电量为{agg['q2']['emergency_kwh']:.3f} kWh。"
    "其中316个评估日出现正的紧急购电量，占334个评估日的94.61%；紧急费为3028296.91元，占Q2总费用的18.69%。"
    "紧急购电量的中位数为772.785 kWh，P95为9469.708 kWh，峰值集中在夏季高负荷日；完整逐时计划与紧急时段写入results/result2.xlsx。"
)
fig("raw_q2_q3_q4_emergency", "图4  Q2、Q3与Q4-2全年紧急购电量", width=4.8)
p(
    "图4把三个场景放在同一坐标中作对照：Q2曲线反映日计划在预测误差下的缺口，Q3曲线包含滚动调整后的执行轨迹，Q4-2只改变价格序列。"
    "图中跨场景的数量比较用于解释预测与信息更新时间，费用差异仍需结合各自价格和结算规则判断。"
)
p(
    "紧急量是固定计划、实际负荷和实际光伏共同作用的残差，并不等同于单一预测指标。即使负荷MAE较低，若误差发生在储能已经接近下界或光伏快速下降的时段，仍可能形成较大的缺口；因此本题同时保留紧急购电量、紧急费用和日费用分布三个评价维度。"
)
p(
    "Q2中94.61%的评估日出现正紧急购电量，说明单日计划对预测误差仍较敏感；中位数772.785 kWh而P95达到9469.708 kWh，表明风险具有明显长尾。这个现象支持后文的保守系数敏感性分析，也提醒我们不能只报告全年平均费用。由于回测只覆盖一个年度且采用单一路径中位数预测，Q2结果应解释为当前附件下的可复现实验基线，而不是对其他年份的无偏保证。"
)

H("五、问题三：多时刻预报与滚动调整")
h("5.1 锁定前缀的滚动策略")
p(
    "每天0:00使用首份24小时预报生成计划q^0；在6:00、12:00、18:00收到新预报后，只对尚未执行的时段重新求解。"
    "设四次重优化起点为0、36、72、108个十分钟步，已执行前缀的q、c、d和SOC均保持不变，当前末状态作为下一次优化的初始状态。"
    "每次求解的剩余窗口分别为24、18、12和6小时；这一实现把信息到达时间写进算法，而不是用全天真实光伏回填历史决策。"
)
pf.equation(doc, r"x_t^{(k)}=x_t^{exec}\;(t<\tau_k),\quad \tau_k\in\{0,36,72,108\},\quad x=(q,c,d,S)")
p(
    "式中τ_k是第k次预报发布对应的十分钟索引。求解器只对τ_k之后的后缀优化，并把前缀末端SOC作为边界；因此每次更新既能利用新预报，又不会篡改已经发生的购电和储能动作。负荷仍沿用问题二的日期因果预测，Q3新增的信息只作用于光伏后缀。"
)
h("5.2 计划—调整结算")
p(
    "令q_t^a为最终调整计划，q_t^0为0:00基准计划。两者只在尚未执行的后缀上允许不同，题面结算规则写成"
)
pf.equation(doc, r"K_3=\sum_t p_tq_t^0+\sum_t0.5p_t[q_t^0-q_t^a]^++\sum_t1.5p_t[q_t^a-q_t^0]^++\sum_t5p_te_t")
p(
    "逐日记录基准费、调减费、调增费和紧急费，并检查四项之和。结算时始终以q^0计入基准费，调整部分按计划差额分为调减和调增两项：减少计划量按0.5倍价格计费，增加计划量按1.5倍价格计费，避免把最终调整计划误当作基准计划。"
)
row3 = daily[["q3_base_cost", "q3_reduction_fee", "q3_increase_fee", "q3_emergency_cost", "q3_cost"]].sum()
q2_positive_days = int((daily["q2_emergency_kwh"] > 1e-8).sum())
q3_positive_days = int((daily["q3_emergency_kwh"] > 1e-8).sum())
q3_adj_corr = float(daily[["q3_adjustment_abs_kwh", "q3_emergency_kwh"]].corr().iloc[0, 1])
doc.add_paragraph("表7  Q3全年结算分项")
pf.three_line_table(doc, [["分项", "全年金额/元", "占Q3总费用/%"], ["计划基准费", f"{row3['q3_base_cost']:.2f}", f"{100*row3['q3_base_cost']/row3['q3_cost']:.2f}"], ["调减费", f"{row3['q3_reduction_fee']:.2f}", f"{100*row3['q3_reduction_fee']/row3['q3_cost']:.2f}"], ["调增费", f"{row3['q3_increase_fee']:.2f}", f"{100*row3['q3_increase_fee']/row3['q3_cost']:.2f}"], ["紧急费", f"{row3['q3_emergency_cost']:.2f}", f"{100*row3['q3_emergency_cost']/row3['q3_cost']:.2f}"], ["合计", f"{row3['q3_cost']:.2f}", "100.00"]])
p(
    f"Q3全年费用为{agg['q3']['cost']:.2f}元，紧急购电量为{agg['q3']['emergency_kwh']:.3f} kWh，"
    f"计划—调整绝对差额总量为{daily['q3_adjustment_abs_kwh'].sum():.3f} kWh；其中调减量为{daily['q3_reduction_kwh'].sum():.3f} kWh，调增量为{daily['q3_increase_kwh'].sum():.3f} kWh。"
)
fig("process_q3_adjustment", "图5  问题三计划—调整差额与紧急购电量", width=4.8)
p(
    f"表7中计划基准费占Q3总费用的{100*row3['q3_base_cost']/row3['q3_cost']:.2f}%，调减和调增费用合计占{100*(row3['q3_reduction_fee']+row3['q3_increase_fee'])/row3['q3_cost']:.2f}%，紧急费占{100*row3['q3_emergency_cost']/row3['q3_cost']:.2f}%。"
    "基准计划仍然决定大部分账单，调整并不是把原计划完全替换掉。"
)
h("5.3 结果解读与边界")
p(
    f"与Q2相比，Q3全年费用增加{agg['q3']['cost']-agg['q2']['cost']:.2f}元（{100*(agg['q3']['cost']/agg['q2']['cost']-1):.2f}%），紧急购电量增加{agg['q3']['emergency_kwh']-agg['q2']['emergency_kwh']:.3f} kWh（{100*(agg['q3']['emergency_kwh']/agg['q2']['emergency_kwh']-1):.2f}%），出现正紧急量的日期由{q2_positive_days}天变为{q3_positive_days}天。"
    "在当前数据和结算规则下，滚动更新没有自动带来更低费用；调整费和新储能轨迹可能抵消信息更新带来的收益。"
)
p(
    f"图5的计划—调整绝对差与紧急购电量相关系数仅为{q3_adj_corr:.4f}，说明多调计划量不必然对应更大的实际缺口。散点图用于揭示这种解耦关系，而不是证明滚动策略具有因果改善；若要比较策略优劣，应在相同日初SOC、价格和信息集下做独立滚动回测，并报告费用、紧急量和调整量的联合分布。"
)
p(
    "从执行角度看，滚动策略的价值取决于新预报能否在储能尚有调节余量时改变后缀计划；如果更新到达时SOC已接近边界，或者调整主要发生在高价时段，新增信息会转化为调增费用或紧急费用。因而Q3应被看作带结算摩擦的在线策略实验，不能简单等同于“更新次数越多越好”。"
)

H("六、问题四：波动价格下的独立分支")
h("6.1 问题4-2")
p(
    "将附件4的144点价格替换问题二中的固定价格序列，保留相同的日期因果预测和跨日SOC传递。"
    f"Q4-2全年费用为{agg['q4']['cost']:.2f}元，紧急购电量为{agg['q4']['emergency_kwh']:.3f} kWh。"
)
h("6.2 问题4-3")
p(
    "在同一波动价格下，重新执行问题三的四次滚动预报和计划—调整结算。Q4-3从与Q4-2相同的日初SOC独立启动，"
    "不把Q4-2的日末状态带入Q4-3，从而避免两个实验分支互相污染。"
    f"Q4-3全年费用为{agg['q4_3']['cost']:.2f}元，紧急购电量为{agg['q4_3']['emergency_kwh']:.3f} kWh。"
)
row43 = daily[["q4_3_base_cost", "q4_3_reduction_fee", "q4_3_increase_fee", "q4_3_emergency_cost", "q4_3_cost"]].sum()
doc.add_paragraph("表8  四个场景的全年本地回测")
pf.three_line_table(doc, [["场景", "费用/元", "紧急购电/kWh", "日费用中位数/元", "日费用P95/元"]] + [[name, f"{agg[key]['cost']:.2f}", f"{agg[key]['emergency_kwh']:.3f}", f"{daily[cost].median():.2f}", f"{daily[cost].quantile(.95):.2f}"] for name, key, cost in [("Q2", "q2", "q2_cost"), ("Q3", "q3", "q3_cost"), ("Q4-2", "q4", "q4_cost"), ("Q4-3", "q4_3", "q4_3_cost")]])
doc.add_paragraph("表9  Q4-3全年结算分项")
pf.three_line_table(doc, [["分项", "全年金额/元", "占Q4-3总费用/%"], ["计划基准费", f"{row43['q4_3_base_cost']:.2f}", f"{100*row43['q4_3_base_cost']/row43['q4_3_cost']:.2f}"], ["调减费", f"{row43['q4_3_reduction_fee']:.2f}", f"{100*row43['q4_3_reduction_fee']/row43['q4_3_cost']:.2f}"], ["调增费", f"{row43['q4_3_increase_fee']:.2f}", f"{100*row43['q4_3_increase_fee']/row43['q4_3_cost']:.2f}"], ["紧急费", f"{row43['q4_3_emergency_cost']:.2f}", f"{100*row43['q4_3_emergency_cost']/row43['q4_3_cost']:.2f}"], ["合计", f"{row43['q4_3_cost']:.2f}", "100.00"]])
fig("result_q4_cost_box", "图6  Q2、Q3与Q4-2日费用分布")
fig("result_q4_3_series", "图7  Q4-3逐日费用序列")

H("七、题面指定日期结果")
p("下表给出题目要求的四个指定日期的日购电结果；逐十分钟完整计划保存在对应结果工作簿，表中紧急时段为连续正缺口的合并区间。")
for label, filename, groups, cost_col, em_col in [("问题二", "result2.xlsx", em2, "q2_cost", "q2_emergency_kwh"), ("问题三", "result3.xlsx", em3, "q3_cost", "q3_emergency_kwh")]:
    doc.add_paragraph(f"表{10 if label == '问题二' else 11}  {label}指定日期紧急购电")
    rows = [["日期", "日费用/元", "紧急购电量/kWh", "紧急购电时段（区间：电量/kWh）"]]
    for d in specified_dates:
        rec = daily.loc[daily.date == d].iloc[0]
        pieces = groups[d]
        desc = "；".join(f"{tm}：{val:.3f}" for tm, val in pieces) if pieces else "无"
        rows.append([d, f"{rec[cost_col]:.2f}", f"{rec[em_col]:.3f}", desc])
    pf.three_line_table(doc, rows)
    p(f"完整的144点计划、储能状态和所有紧急购电区间见{filename}；本表只压缩展示题面指定日期，避免正文被逐日明细淹没。")

H("八、敏感性、稳健性与局限")
h("8.1 光伏保守系数敏感性")
sens = pd.read_csv(OUT / "sensitivity_forecast.csv")
doc.add_paragraph("表12  Q2光伏保守系数敏感性")
pf.three_line_table(doc, [["系数", "全年费用/元", "紧急购电/kWh"]] + [[f"{r.pv_factor:.2f}", f"{r.q2_cost:.2f}", f"{r.q2_emergency_kwh:.3f}"] for r in sens.itertuples()])
p(
    "在当前附件和日电量—形状负荷预测下，0.9系数给出最低的回测费用；0.8虽然减少了紧急购电量，却增加了计划购电，"
    "1.0则在预测偏乐观时显著增加紧急购电。这个结论只对当前回测窗口和成本口径成立，不宣称0.9是普适最优。"
)
h("8.2 对抗性核验")
p(
    "外部经验贴仅用于方法讨论，本文不采用其中未经复现的数字或未来信息路径。"
    "重点检查包括：计划与调整结算是否分项相加、Q4-2与Q4-3是否从独立分支启动、负荷预测是否只使用目标日前数据、"
    "MILP与网格DP是否在同一端点条件下比较，以及结果工作簿是否保留题面要求的工作表。"
)
h("8.3 局限")
p(
    "当前模型把预测表示为单一路径，尚未建立完整场景树或概率预测区间，也未建模通信延迟、电池退化、负荷可移峰和CVaR风险度量。"
    "因此全年数值只能解释为本地附件回测基线；后续若引入鲁棒或随机模型，应重新估计场景、重算全部结果并重新审计。"
)
fig("process_q2_q4_compare", "图8  固定价格与波动价格下的日费用对比")

H("九、结论与可复现性")
p(
    "本文以一个统一的十分钟能量平衡模型贯穿四个问题：问题一用互斥MILP得到可执行的单日基准，"
    "网格DP提供库存价值的独立校验；问题二用日期因果预测完成全年计划和实际缺口结算；问题三把新预报转化为锁定前缀后的滚动再决策，"
    "并按题面规则拆解计划、调减、调增和紧急费用；问题四通过独立分支评估波动价格的影响。"
    f"在当前334天本地回测中，Q2至Q4-3费用依次为{agg['q2']['cost']:.2f}、{agg['q3']['cost']:.2f}、{agg['q4']['cost']:.2f}和{agg['q4_3']['cost']:.2f}元。"
)
p("复现命令为：../.venv/bin/python solve_c.py --seed 20260911。结果表位于results/result*.xlsx，逐日账单位于results/daily_metrics.csv，预测诊断位于results/forecast_diagnostics.csv，输入哈希和环境记录位于results/复现记录.json。")
p("正文中的图1、图2、图3、图4、图5、图6、图7、图8以及表1、表2、表3、表4、表5、表6、表7、表8、表9、表10、表11、表12、表13均在相应方法或结果段落中使用；完整逐十分钟结果保留在结果工作簿中。")
doc.add_paragraph("表13  交付物清单")
pf.three_line_table(doc, [["交付物", "位置", "用途"], ["求解脚本", "solve_c.py", "重算Q1—Q4"], ["结果工作簿", "results/result*.xlsx", "按题面模板复核和提交准备"], ["论文Word", "完整论文.docx", "可编辑正文和公式"], ["论文LaTeX", "完整论文-LaTeX/main.tex", "Overleaf编译"], ["图形", "figures/*.png,*.svg", "论文插图和编辑"], ["研究记录", "research/", "样本检索、经验贴审查与方法边界"]])

H("AI工具使用声明")
p("依据当届竞赛规定，Codex用于资料检索、代码组织、公式排版和结果复核；题面解释、模型假设、参数选择、代码运行、结果解释和最终提交由参赛队员审阅并负责。外部经验贴与公开论文只用于方法参考，不作为未经复现的结果来源。AI工具使用详情另见AI工具使用详情.pdf。")

H("参考文献")
p("[1] 全国大学生数学建模竞赛组委会. 全国大学生数学建模竞赛论文格式规范（2023年修订稿）. 中国大学生在线, 2023. https://dxs.moe.gov.cn/zx/a/hd_sxjm_gsyw/231205/1869343.shtml.")
p("[2] Gulotta F, Crespo del Granado P, Pisciella P, et al. Short-term uncertainty in the dispatch of energy resources for VPP: A novel rolling horizon model based on stochastic programming. International Journal of Electrical Power & Energy Systems, 2023, 153:109355. DOI:10.1016/j.ijepes.2023.109355.")
p("[3] Hönen J, Hurink J L, Zwart B. Dynamic Rolling Horizon-Based Robust Energy Management for Microgrids Under Uncertainty. arXiv:2307.05154, 2023.")
p("[4] Hart W E, Laird C, Watson J P, et al. Pyomo—Optimization Modeling in Python. Springer, 2017.")

out = pf.save_document(doc, ROOT, filename="完整论文.docx", contest="cumcm", overwrite=True)
print(out)
