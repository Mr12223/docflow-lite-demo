# 评价集说明

这个目录用于存放毕业设计的“人工对照评价集”配置。

推荐流程：

1. 从真实样本中挑选 10~30 份有代表性的文档。
2. 在 `sample_manifest.json` 中为每个样本填写：
   - `path`：相对项目根目录的文件路径
   - `expected_success`：预期是否应该解析成功
   - `min_char_count`：最低提取字符数
   - `must_contain`：提取结果必须包含的关键词
   - `must_not_contain`：提取结果不应包含的关键词
3. 运行：

```powershell
.\.venv\Scripts\python.exe .\run_evaluation_set.py
```

生成报告后，可直接把 `reports/evaluation_xxx/evaluation.md` 的统计结果写入论文测试章节。
