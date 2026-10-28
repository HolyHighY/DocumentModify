# 03 内网操作规范：先验证，再确定本机流程

## 1. 给内网 AI 的执行指令

先阅读 [通用规范](01-common.md)、[报告类型](04-report-types.md) 和 [表单](05-templates.md)。**第一次在某台机器上使用本规范时，按本章 Q0—Q7 实际执行验证，再编写该机器的操作规程；不要跳过测试直接开始正式文件读写。**

系统、Python、WPS、字体、加密策略或读写工具变化时，重新验证受影响项目。可复用仍有效的证据，但要检查条件是否相同。

规则如下：

1. 验证使用合成样例和允许使用的文档副本，禁止覆盖原件、绕过加密或将材料送往外网。
2. 每一步记录命令、工作目录、退出码、预期与实际结果、产物位置、工具版本和时间。无证据不标为通过。
3. 状态使用 `PASS / FAIL / BLOCKED / NOT_RUN / N/A`。用户尚未执行的视觉检查是 `NOT_RUN`，不能凭脚本成功代签。
4. 失败后先区分环境问题、接口能力缺失、文档问题和未知问题；修正有依据的配置后重试。不要反复运行同一失败命令。
5. 某路线不可用时继续其他独立验证，记录降级路线。所有可用路线均无法完成关键读写时，停止正式写入，保留草稿和精确阻塞项。
6. 内网无执行工具或无法访问文件时，生成待执行命令和预期结果交给本机操作者；根据实际返回继续，不能自行补写测试结果。

## 2. Q0：建立工作区与确认边界

按本机实际路径新建独立项目工作区，例如：

```text
report-work/
  incoming/             # 原始交接包，验收前不改动
  originals/            # 原件，只读保留
  qualification/        # 合成样例、副本、日志与验证结果
  work/                 # 分章节草稿、中间文档和状态
  figures/              # 数据图、示意图及源文件
  delivery/             # 经复核的交付版本
```

记录允许读取的目录、输出位置、软件自动化权限、加密文档复制/另存规则和资料流向。默认不允许向外网传递内网文件、摘录、日志和截图。本机权限不足时请求具体操作权限，不要求关闭加密或安全系统。

本规范中的路径均为示例。命令执行前替换为实际路径；后续记录保留可直接复现的本机命令。

**通过条件**：工作区可用，读写边界已明确，测试不会覆盖原件。

## 3. Q1：环境与离线工具探测

先确认 Python 命令；Windows 示例：

```powershell
python --version
# 如果上条不存在，再尝试：
py -3 --version
```

确认一个可用入口后，以下统一以 `python` 表示。进入携带的规范目录执行：

```powershell
python tools/document_check.py probe --output "D:\report-work\qualification\environment.json"
```

该探测只记录平台、Python 和常见模块存在性，不读取环境变量或凭据，不导入 WPS、打开文档或确认 COM 可用性。另记录 WPS 的版本、位数、已安装字体、PDF 导出能力、是否有 GUI，以及终端加密软件是否要求特定应用打开文件。不要导出完整注册表或系统环境。

若准备使用 `python-docx`，执行：

```powershell
python -c "import docx; print(docx.__version__)"
```

若准备使用 Windows COM，执行：

```powershell
python -c "import win32com.client; print('pywin32 import OK')"
```

仅对选定路线安装必要依赖。已准备可信离线 wheel 时，可按实际包版本和目录执行，例如：

```powershell
python -m pip install --no-index --find-links "D:\offline-wheels" python-docx
# 仅 Windows COM 路线需要；不在其他平台执行：
python -m pip install --no-index --find-links "D:\offline-wheels" pywin32
```

安装前核对包来源、完整性、目标平台及 Python 版本。`--no-index` 确保不访问索引，但不替代完整性验证。没有兼容安装包时记录依赖阻塞，不能假设离线安装成功。

**通过条件**：至少一个候选路线的工具可启动；记录其余路线缺失情况。模块存在或导入成功不代表文档读写通过。

## 4. Q2：材料接入与可读性

1. 将外网包接入 `incoming`，从可信交接记录核对包来源。在未执行包内任意脚本前检查其用途和内容。
2. 若包含摘要清单，使用本规范随附工具验证：

```powershell
python tools/document_check.py verify --root "D:\report-work\incoming\writing-package" --manifest "D:\report-work\incoming\writing-package\MANIFEST.json"
```

退出码非零时停止使用该包，定位缺失、多余或不一致文件并重新接入。没有摘要时登记“未取得交接前摘要”，可建立本地基线，但不能声称已证明传输前后一致。

3. 用合成 CSV 验证中文、字段、单位、空值、小数和行数可正确读取；用合成照片验证方向、尺寸和插入能力。实际材料另做字段和数量核对，不把“能打开”当作数据含义正确。
4. 验证模板和代表性文档的可读路线。对允许分析的工作副本执行：

```powershell
python tools/document_check.py inspect-docx --input "D:\report-work\qualification\sample.docx" --output "D:\report-work\qualification\sample-structure.json"
```

退出码 `0` 只表示普通 DOCX 基本结构可解析。`2` 表示无法按普通 DOCX 确认，例如非 ZIP、加密容器、缺少主文档或 XML 无效；不能仅凭该结果断言具体加密机制。`.doc` 不适用此检查。文件操作错误使用退出码 `1`。

如果资源管理器可见文件、WPS 能打开，但 Python 无法直接解析，应继续 Q3 的授权 WPS 路线，不改扩展名伪装格式，不使用解密尝试。

**通过条件**：材料数量和完整性情况清楚；数据、图片、模板有已验证的接入方式。敏感材料提取结果仅存内网工作区。

## 5. Q3：验证 Python 与 WPS 读写路线

### 5.1 先制作覆盖格式的合成夹具

由操作者在目标 WPS 中创建测试文档，保存到 `qualification`。必须包含：

- 唯一正文标记 `DM_BODY_BASELINE_01`、中文、标题一级和二级、粗体/斜体；
- 2 行 2 列表格，单元格标记 `DM_TABLE_01`；
- 一张已知小图片、图题、正文引用；
- 分页符、页眉、页脚与页码；
- 若真实模板需要，加入公式、目录、交叉引用、合并单元格或浮动图片。

保存一份普通夹具（如本机允许），并通过本机批准流程保存一份处于真实加密策略下的夹具。不能拿未加密夹具的成功结果证明加密文档可用。保留初始文件、截图或允许生成的本地 PDF 基线；所有测试另用工作副本。

### 5.2 路线选择

| 路线 | 必须实际证明的能力 | 使用边界 |
| --- | --- | --- |
| Python 直接解析，如 python-docx | 能读到夹具正文、表格和所需结构；修改副本、保存、重开后保留格式 | 仅适用于实际可解析的文件；复杂域、公式等需单独验证 |
| Python 驱动授权 WPS 接口 | Python 可通过本机支持的接口打开受保护副本、提取、修改、保存并重开 | Windows 可调查 COM；其他平台按本机官方接口，不套用 Windows 命令 |
| WPS 内授权导入/导出配合 Python | WPS 可在允许位置导出受控中间内容，Python 处理后可回填并保留必要格式 | 中间文件仍按原资料范围管理；禁止默认导出不受保护的副本 |

先用最简单的可用路线验证，再决定正式方案。可能出现“正文用 WPS 提取、计算用 Python、合稿用 WPS”的组合。不能为了省事把包含表格、公式或图片的报告降为纯文本后宣称无损。

### 5.3 普通 DOCX 最小读写测试

仅在 Python 可直接解析普通夹具时执行。将以下代码保存为工作区的 `qualification/test_plain.py`，它拒绝覆盖已有输出：

```python
import sys
from pathlib import Path
from docx import Document

src, dst = map(Path, sys.argv[1:3])
assert src.resolve() != dst.resolve() and not dst.exists()
d = Document(src)
text = "\n".join(p.text for p in d.paragraphs)
assert "DM_BODY_BASELINE_01" in text
assert any("DM_TABLE_01" in c.text
           for t in d.tables for r in t.rows for c in r.cells)
assert len(d.inline_shapes) >= 1  # 本夹具要求一张行内图片
table_count, image_count = len(d.tables), len(d.inline_shapes)
d.add_paragraph("DM_WRITE_PLAIN_01")
d.save(dst)
r = Document(dst)
assert "DM_WRITE_PLAIN_01" in "\n".join(p.text for p in r.paragraphs)
assert len(r.tables) == table_count and len(r.inline_shapes) == image_count
print("PASS: basic text/table/inline-image round trip; visual checks pending")
```

```powershell
python "D:\report-work\qualification\test_plain.py" "D:\report-work\qualification\plain-fixture.docx" "D:\report-work\qualification\plain-result.docx"
```

该测试只覆盖基本对象，不证明复杂对象和版面正确。继续 Q6。在真实模板副本上做代表性替换、表格更新和插图后重复检查。

### 5.4 WPS 接口测试指令

Windows 候选 COM 标识可在本机注册信息中定向查询，例如：

```powershell
Get-ItemProperty -LiteralPath 'Registry::HKEY_CLASSES_ROOT\KWPS.Application\CLSID' -ErrorAction SilentlyContinue
```

这是候选标识，不保证所有 WPS 版本均注册。未发现时查本机 WPS 帮助或已批准的接口文档，不能通过禁用安全策略解决。接口对象、打开参数和保存方法必须以本机支持情况为准。

如果确认该版本支持 `DispatchEx` 和下列对象方法，可将代码保存为 `qualification/test_wps.py`。**传入的只能是事先准备好的测试工作副本，代码会修改该副本**；测试前关闭该副本，保留基线。先确认实例创建机制不会接管操作者正在编辑的文档。

```python
import sys
from pathlib import Path
import win32com.client

prog_id, filename = sys.argv[1:3]
p = Path(filename).resolve()
assert p.is_file()
app = win32com.client.DispatchEx(prog_id)
app.Visible = True  # 便于观察授权提示和保存结果，不隐藏失败对话框
doc = None
try:
    doc = app.Documents.Open(str(p))
    assert not doc.ReadOnly, "Test copy is read-only"
    assert "DM_BODY_BASELINE_01" in doc.Content.Text
    assert doc.Tables.Count >= 1
    assert "DM_TABLE_01" in doc.Tables.Item(1).Range.Text
    assert "DM_WRITE_WPS_01" not in doc.Content.Text, "Use a fresh test copy"
    doc.Content.InsertAfter("\rDM_WRITE_WPS_01\r")
    doc.Save()
    doc.Close(0)
    doc = None
    doc = app.Documents.Open(str(p))
    assert "DM_BODY_BASELINE_01" in doc.Content.Text
    assert "DM_WRITE_WPS_01" in doc.Content.Text
    assert "DM_TABLE_01" in doc.Tables.Item(1).Range.Text
    print("PASS: WPS basic read/write/reopen; image/layout checks pending")
finally:
    if doc is not None:
        doc.Close(0)
    # 不调用 Quit，避免关闭被实现复用的用户实例；人工确认后关闭测试窗口。
```

```powershell
python "D:\report-work\qualification\test_wps.py" "KWPS.Application" "D:\report-work\qualification\protected-working-copy.docx"
```

该片段是待本机验证的接口候选，不是已保证兼容的实现。接口异常、授权弹窗或超时均记录为失败或阻塞，不关掉提示假装成功。若创建了测试实例但脚本异常退出，人工仅关闭自己创建的测试窗口，不结束所有 WPS 进程。

完成基本测试后，AI 必须针对真实模板需要，在**新的测试副本**上补测：读出章节、按样式定位目标段落、定点替换、写入表格、插入图片及题注、处理所需公式与目录。记录对应 Python 方法、参数和验证结果；不能把追加一行成功当作所有能力通过。

**通过条件**：选定路线在真实加密策略下完成所需读写、保存和重新打开，且原件未改动。需要人工协助的环节明确登记为人工步骤。

## 6. Q4：验证分章写作与会话恢复

用合成材料建立两章小报告：背景章和结果章。结果章输入一份明确标记的合成数据和一张测试图片。

1. 按任务单、大纲、术语表、材料 ID 和当前章节材料起草一章，登记待补项、引用、图表和摘要。
2. 保存状态文件，启动独立新会话；若不能新建，则由操作者仅提供恢复包，标记为“恢复演练”。
3. 新会话只依据入口文件、任务单、提纲、术语、决定、章节摘要及本章材料完成下一章。
4. 检查是否正确延续术语、图表 ID 和限制，是否出现臆造数据或来源，是否能指出待补信息。
5. 合并两章，检查同一术语、结论和图表引用一致。对依赖缺失的内容保留草稿状态。

**通过条件**：不依赖完整聊天也能恢复任务，章节产物和状态一致。实际使用的上下文窗口、文件导入限制和章节拆分方法写入本机规程。

## 7. Q5：验证分析与正式写入之前的内容检查

先用已知答案的合成数据验证分析脚本：例如输入 `1, 2, 3`，均值应为 `2`；另验证单位换算、空值和明显非法值处理。此例只是最低演练，不替代真实报告所需的误差、拟合或统计验证。

确认图中的数据点、坐标和单位与输入对应，正文引用正确。正式材料接入后，保留原始数据和处理记录；没有实际数据时不能生成实测结果。

将已核对的两章内容通过 Q3 选定路线写入真实模板的工作副本，包含表格和图片。对照写入前后的预期变更范围，核对章节、对象及关键内容，没有要求的段落不得消失。真实报告写入前先形成修改清单和可恢复的本地版本。

**通过条件**：所需分析产生预期结果，内容能够定点写入且证据关系保留。

## 8. Q6：保存、重开与版面验收

在目标 WPS 中执行，不依赖 XML 计数代替视觉结果：

1. 保存、关闭测试文档，再重新打开；确认无需修复，正文标记、表格、图片、公式和样式保留。
2. 按模板需要更新目录、页码和交叉引用，检查更新后的显示结果。记录本机可用菜单或已验证接口。
3. 导出本地 PDF（如要求且允许）；核对所有页面的字体、分页、页眉页脚、图表截断、标题孤行、表格跨页、公式和目录页码。
4. 与基线及预期修改对照，确认原文未意外丢失或改变。若无 PDF 导出能力，使用 WPS 打印预览逐页检查并记录替代方式；交付明确要求 PDF 时，该替代不满足 PDF 交付要求。
5. 原件和工作副本在保存前后核对版本或摘要；加密系统可能使文件字节变化，语义和版面检查仍不可省略。

AI 若能在内网读取允许的页面渲染结果，可辅助逐页核对；否则交由操作者检查并记录结果。不上传页面或截图寻求外网检查。

**通过条件**：所需内容、复杂对象和版面均有检查证据。AI 无法视觉检查且操作者尚未确认时，排版验收保持 `NOT_RUN`。

## 9. Q7：固化本机操作规程并做最终演练

使用 [表单中的内网验证记录](05-templates.md#内网资格验证记录与本机规程)，在内网项目目录创建 `qualification/RESULTS.md` 和 `work/LOCAL-WORKFLOW.md`，不要把内网记录提交到外网仓库。

`LOCAL-WORKFLOW.md` 必须填写实测得到的内容：

- 适用机器、系统、Python/WPS 版本、加密条件、验证日期及失效条件；
- 材料接入、校验、提取、章节输入和恢复方法；
- Python 入口、离线依赖、读写路线和确切命令/接口；
- 保存、重开、目录更新、PDF 导出和逐页验收步骤；
- 失败恢复、版本保留、人工步骤、已知不支持对象和回退方法。

按该规程从一份新夹具副本重新演练：接入 → 分章恢复 → 分析 → 模板写入 → 保存重开 → 排版检查。仅重跑本次实际需要的路线。删除或注明未验证的候选方法，保留失败证据及限制。

最终按能力分别报告：材料接入、分章写作、Python/WPS 读写、分析、排版与交付。必要项目全通过才可声明“本机流程已验证”；阻塞或未运行项目存在时，只声明已证实的能力。

## 10. 正式任务的后续流程

在适用条件未变且验证记录有效时，使用本机规程接入真实材料、分章撰写、合稿和验收。每份报告仍执行事实、模板、保存重开和版面检查。资格验证通过不能代替每份报告自身的质量验收。
