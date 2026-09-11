# 2026 CUMCM C题研究工作区

本文件夹独立于项目根目录中的B题代码，保存C题题面、数据附件、求解代码、结果表、指标、图形、研究资料和论文草稿。

唯一复现命令（从本目录执行）：`../.venv/bin/python solve_c.py --seed 20260911`。

已完成内容：

- Q1：带充放电互斥和日末SOC约束的MILP；
- Q2：历史七日中位数预测、光伏保守校准和2025-02-01至12-31因果回测；
- Q3：0/6/12/18四次发布、锁定已执行前缀的滚动调整与差额结算；
- Q4：附件4动态价格下独立重算Q4-2和Q4-3；
- 结果表、逐日CSV、复现清单、15组PNG/SVG图形、可编辑Word论文 `完整论文.docx` 与 Overleaf 工程 `完整论文-LaTeX/main.tex`；
- AI使用记录模板位于 [AI使用记录模板.md](AI使用记录模板.md)，已生成 [AI工具使用详情.pdf](AI工具使用详情.pdf)；提交前需由参赛队员审阅并按当届规定补充或调整。

主要验证命令：

`../.venv/bin/python validate_results.py`

`../.venv/bin/python scripts/validate_input_xml.py`

工程化入口：`make solve` 运行复现，`make validate` 执行结果与输入校验；依赖见 `requirements.txt`。
