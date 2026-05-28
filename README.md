# WOS 多校区 DID 小范围测试

本项目用于小范围测试“同城新建多校区与教师个人产出”的 WOS 清洗和基准 DID 回归流程。

## 数据口径

- WOS 原始数据目录：`D:\数据\WOS数据（数据不超过10w条的学校）`
- 同城新校区数据：`D:\数据\211_高校新校区_重新搜索.xlsx`
- 年份窗口：`2000-2022`
- 处理组：窗口期内同城新建多校区。
- 对照组：窗口期及之前不应有多校区记录；上海财经大学已按该口径剔除。

## 主要文件

- `scripts/clean_wos_sample.py`：读取 WOS 工作簿，拆分 `Author Full Names`，生成个人-论文和个人-年面板。
- `do/baseline_did_sample.do`：导入个人-年面板并运行基准 DID 回归。
- `outputs/sample_school_mapping.csv`：本轮 5:5 样本映射。
- `outputs/excluded_schools.csv`：剔除学校及原因。
- `outputs/validation_summary.txt`：清洗校验摘要。
- `outputs/baseline_did_sample.log`：Stata 回归日志。

## 运行方式

```powershell
& 'C:\Users\Dell\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe' scripts\clean_wos_sample.py
$p = Start-Process -FilePath 'D:\stata\StataMP-64.exe' -ArgumentList @('/e','do','do\baseline_did_sample.do') -WorkingDirectory (Get-Location) -WindowStyle Hidden -Wait -PassThru
```

大体量中间数据不会提交到 GitHub，包括作者-论文 CSV、个人-年 CSV 和 Stata DTA，可通过上述命令在本地复现。
