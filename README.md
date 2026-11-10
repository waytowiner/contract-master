# 合同管理大师 Contract Master

> 让散落的合同扫描件，变成能核对、能追溯、能查询的合同资产。

合同的价值在条款里，档案的价值在**找得到、核得准、查得清**。

真正耗时的往往不是扫描，而是反复翻页、登记、核对金额、寻找原件，以及回答“这个数字从哪里来”。“合同管理大师”把本地 OCR、要素抽取、人工复核、Excel 台账和 SQLite 检索串成一条完整工作流，让合同整理从重复劳动变成有证据、有边界的数字化过程。

它不承诺“AI 永远正确”，而是让每一项结果都知道自己**来自哪份文件、哪一页、哪段原文**。

## 为什么选择它

- **本地离线，零 API 费用**：依赖安装完成后，OCR 与数据处理均在本机进行，合同不必上传第三方服务。
- **原件只读**：不修改、移动、重命名或删除原始合同；产物只能写入独立工作目录。
- **每个字段有出处**：来源文件、展示页码、物理页码、原文片段、坐标与置信度均可追溯。
- **不替人做最后判断**：存疑字段进入复核队列，每条复核项拥有唯一 ID，确认过程清晰可控。
- **保守归并，拒绝乱并**：可靠编号或精确金额与相同乙方同时匹配，核心字段不冲突才归并；全部页面保留。
- **交付即可使用**：生成 Excel/CSV 台账、复核清单和 SQLite 数据库，支持甲乙方、金额、编号及全文检索。

## 免费版功能

- JPG、PNG、TIFF、WebP 等图片与 PDF 的本地 OCR
- 合同编号、标题、甲乙方、金额、大写金额、签订地点、银行信息等规则抽取
- 大小写金额勾稽、低置信与空白页提醒、字段值卫生检查
- 文件盘点、哈希去重、缺页线索、保守副本归并与孤立页人工建议
- Excel 台账、CSV、复核队列、交付报告和运行日志
- SQLite 全文数据库与命令行查询
- 新增合同组的安全增量入库，以及变更场景的防误写保护

适合行政档案、财务、法务、审计复核，以及需要把历史扫描合同整理成可检索资产的个人和团队。

## 快速开始

### 1. 安装技能

从 [GitHub Releases](https://github.com/waytowiner/contract-master/releases/latest) 下载技能 ZIP，解压后安装。

将本目录完整复制到 WorkBuddy 用户技能目录，例如：

```text
~/.workbuddy/skills/contract-master/
```

重新打开 WorkBuddy 后，可以直接说：

```text
整理这批合同扫描件，生成台账和可查询数据库。
```

也可以手动执行命令行流程。

### 2. 安装依赖

```bash
python -m pip install rapidocr_onnxruntime pypdfium2 openpyxl
```

### 3. 运行标准流程

以下示例中，`D:/contracts` 是合同输入目录，`D:/contract-work` 是独立输出目录。输出目录不能位于技能目录或合同输入目录内。

```bash
python scripts/scan_files.py --input "D:/contracts" --out "D:/contract-work" --recursive
python scripts/ocr_extract.py --input "D:/contracts" --out "D:/contract-work" --recursive
python scripts/extract_fields.py --out "D:/contract-work"
python scripts/build_output.py --out "D:/contract-work"
```

人工编辑 `D:/contract-work/04_review_queue.xlsx`，逐项确认后回写：

```bash
python scripts/review_resolve.py --db "D:/contract-work/06_contracts.db" --queue "D:/contract-work/04_review_queue.xlsx"
```

查询示例：

```bash
python scripts/query_db.py --db "D:/contract-work/06_contracts.db" --party "某供应商"
python scripts/query_db.py --db "D:/contract-work/06_contracts.db" --keyword "设备"
```

## 新批次使用原则

不要直接拿一批新合同全量运行。请遵循：

1. **盘点**：检查文件数量、合同组数和异常文件；
2. **小样本校准**：抽取 10–20 页，人工核对 OCR 与台账字段；
3. **全量处理**：校准规则后再运行，并完成复核队列。

通用问题与处理规则见 [`references/rules-of-thumb.md`](references/rules-of-thumb.md)。

## 质量与安全

- 58 个单元、行为和临时目录集成测试；
- 16 项交付、金额勾稽、溯源、数据库、增量行为和安装一致性检查；
- 测试与发布文档使用合成或匿名化信息；
- 打包脚本只同步白名单文件，拒绝数据库、表格、合同图片、PDF 和 OCR 运行数据。

运行测试：

```bash
python -m unittest discover tests -v
```

仓库附带 Windows/Linux 双平台 [GitHub Actions 配置模板](https://github.com/waytowiner/contract-master/blob/main/ci/github-actions.yml)。当前发布凭据未获 `workflow` 权限，远端自动 CI 尚未启用；上述测试与验收已在本机运行通过。

**自动自检不等于图片级人工验收。OCR 与规则抽取仍可能出错，复核队列不能省略。**

## 免费版与专业版方向

当前发布的是完整可用的免费核心版，不通过削弱基础功能制造付费门槛。

后续专业版方向包括旧台账自动逐列比对、字段与规则定制、批量审计、更多合同模板、高级报表和团队协作。私有化部署、新批次校准、人工复核与技术支持也可作为独立服务提供。

这些能力仍在规划或演进中，不承诺固定上线日期。欢迎通过 GitHub Issues 留下真实使用场景和最希望解决的问题，这将直接决定后续优先级。

## 数据边界

本仓库只包含技能代码、文档和合成测试，不包含任何真实合同、OCR 全文、台账或数据库。

请勿在 Issues、Pull Request 或示例文件中上传合同原件、账号、个人隐私或其他敏感信息。

## 许可证

基于 [MIT License](LICENSE) 开源。

本技能不提供法律意见，不判断合同效力，不替代律师审查。
