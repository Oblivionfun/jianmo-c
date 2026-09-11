# W1 evidence outline (internal)

- 摘要数值：Q1全天购电59482.699 kWh，费用35126.986元；SOC 1200–10800，端点6000；来源 results/summary.json q1。
- Q1模型：能量平衡/SOC/MILP互斥；图 raw_q1_inputs、process_q1_soc、result_q1_dispatch；代码 solve_c.py:q1/solve_dispatch。
- Q2结果：全年成本17197426.756元，紧急购电463147.305 kWh；来源 summary/ daily_metrics；图 raw_q2_overview、process_q2_overview、result_q2_overview。
- Q3结果：全年含调整费用成本20314332.606元，紧急购电758262.155 kWh，计划—调整差额为daily_metrics；采用0/6/12/18四次发布、锁定已执行段的滚动重优化。
- Q4结果：Q4-2全年成本17442616.322元、紧急购电463147.305 kWh；Q4-3滚动调整成本20547027.236元、紧急购电758389.163 kWh；价格替换附件4并独立重算调整策略。
- 数据/单位：数据审计.md、data_audit.json、题面 input/题面/C题.txt；每步Δt=1/6h，输出模板 result*.xlsx。
- 文献：research/references.bib、文献与开源矩阵.md；规则来源mcm.edu.cn AI规定。
- 局限：Q2因果预测为历史中位数+0.9PV校准；Q3/Q4-3为四次发布的确定性滚动近似，未建模预测分布和通信延迟；未将这些结果表述为官方最优或实测结论。
