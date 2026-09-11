# 2026 CUMCM C题研究工作区

本文件夹独立于项目根目录中的B题代码，保存C题题面、数据附件、求解代码、结果表、指标、图形、研究资料和论文草稿。

唯一复现命令（从本目录执行）：`../.venv/bin/python solve_c.py --seed 20260911`。

已完成内容：

- Q1：带充放电互斥和日末SOC约束的MILP；
- Q2：目标日前同类型日优先的日电量—日内形状负荷预测、三日光伏中位数（0.9保守系数）和2025-02-01至12-31因果回测；
- Q3：0/6/12/18四次发布、锁定已执行前缀的滚动调整，以及“计划基准费+调减0.5倍+调增1.5倍+紧急费”的分项结算；
- Q4：附件4动态价格下从同一日初SOC独立重算Q4-2和Q4-3；
- 状态价值函数交叉核验：Q1 5 kWh有限网格DP与MILP独立对照，差异写入 `results/value_dp_q1.json`；
- 对抗性审查：`research/经验贴对抗性审查.md` 记录经验贴可迁移假设、题面风险和本地复现边界，`research/对抗性更新说明.md` 记录本轮改动；
- 结果表、逐日CSV、复现清单、17组PNG/SVG图形、可编辑Word论文 `完整论文.docx` 与 Overleaf 工程 `完整论文-LaTeX/main.tex`；
- AI使用记录模板位于 [AI使用记录模板.md](AI使用记录模板.md)，已生成 [AI工具使用详情.pdf](AI工具使用详情.pdf)；提交前需由参赛队员审阅并按当届规定补充或调整。

本轮复现的本地附件回测总费用为：Q2 `16206550.45` 元，Q3 `18242015.85` 元，Q4-2 `16989713.48` 元，Q4-3 `19577410.54` 元。以上数字不是官方成绩。

主要验证命令：

`../.venv/bin/python validate_results.py`

`../.venv/bin/python scripts/validate_input_xml.py`

`../.venv/bin/python -m pytest -q`

工程化入口：`make solve` 运行复现，`make validate` 执行结果与输入校验，`make test` 运行针对性测试，`make sensitivity` 生成光伏系数敏感性，`make paper` 从最新结果生成 Word 稿；`make all` 串联全部本地步骤。依赖见 `requirements.txt`。

环境说明：求解器使用 `scipy.optimize.milp` 自带的 HiGHS 后端；角色环境检查若提示缺少 PuLP/OR-Tools，不影响本项目当前求解路径。XeLaTeX 未安装时，LaTeX 主稿可直接在 Overleaf 编译。
