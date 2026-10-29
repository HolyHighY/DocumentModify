# 03 内网操作规范：先验证，再确定本机流程

## 1. 给内网 AI 的执行指令

先阅读 [通用规范](01-common.md)、[报告类型](04-report-types.md) 和 [表单](05-templates.md)。先在任务单确定交付模式；未指定时沿用 `word`。**第一次在某台机器上使用本规范时，按下表执行所选模式需要的 Q0—Q7 验证，再编写该机器的操作规程；不要跳过所需验证直接开始正式材料处理。**

### 1.1 交付模式与责任

| 模式 | AI 的报告产物与工作 | 用户的工作 |
| --- | --- | --- |
| `word` | 按已验证路线写入模板工作副本，完成内容检查及可执行的文件、版面检查 | 完成 AI 无法执行的目标应用及视觉检查，并记录实际结果 |
| `md-only` | 仅交付 Markdown 正文及必要图片资源；读取真实模板，核对事实、来源、章节和资源引用 | 使用已有内网应用手动转换为 Word，应用模板样式并完成保存重开、目录/页码与逐页版面检查 |

选择 `md-only` 是任务范围约定，不以 Word 自动化失败为前提。AI 无须写入或转换 Word、生成 DOCX/PDF、安装转换器或完成版面验收；已有内网转换应用的信息未知也不阻塞 Markdown 内容交付。不能据此跳过必要模板读取、事实核对、来源引用和章节检查。

### 1.2 按模式确定资格验证范围

| 项目 | `word` | `md-only` |
| --- | --- | --- |
| Q0 工作区与边界 | 必需 | 必需，登记用户转换责任 |
| Q1 环境与工具 | 验证选定读写路线的工具 | 仅验证材料/模板读取、Markdown 与图片保存及所需分析工具；不要求 Word 转换工具或 COM |
| Q2 材料接入与可读性 | 必需 | 必需，包括真实模板的内容要求 |
| Q3 文档路线 | 验证所需读取、写入、保存及重开 | 仅验证需要的模板/材料读取；Word 写入、保存重开测试为 `N/A` |
| Q4 分章与恢复 | 必需 | 必需，使用 Markdown 章节及状态 |
| Q5 分析与合稿 | 验证所需分析与模板写入 | 验证所需分析与 Markdown 内容合并、引用和图片资源；Word 写入为 `N/A` |
| Q6 Word 版面 | 必需，可由用户完成视觉检查 | AI 的 Word 验收为 `N/A`；用户转换及排版另记 `NOT_RUN`，不阻塞 Markdown 交付 |
| Q7 固化与再演练 | 演练所需 Word 流程 | 演练接入、模板读取、Markdown 合稿与内容交付，不要求转换 Word |

系统、加密策略或选定模式所需的工具变化时，重新验证受影响项目；Python、WPS、字体等仅在所选路线实际依赖时纳入。可复用仍有效的证据，但要检查条件是否相同。由 `md-only` 改为 `word` 时补验新增的写入、保存重开和版面能力，不沿用 Markdown 成功记录证明 Word 可用。

**本机已经通过核验时**：先读取 `.local/environments/<环境ID>/LOCAL-WORKFLOW.md`、核验结果摘要及相关问题记录，确认适用条件和本次所需能力。满足条件即按本机规程写作，不重新跑全套核验；新增能力或条件变化仅补验相关项目。具体写作流程见 [START.md](../START.md)。

规则如下：

1. 验证使用合成样例和允许使用的文档副本，禁止覆盖原件、绕过加密或将材料送往外网。
2. 每一步记录命令、工作目录、退出码、预期与实际结果、产物位置、工具版本和时间。无证据不标为通过。
3. 状态使用 `PASS / FAIL / BLOCKED / NOT_RUN / N/A`。`N/A` 必须注明模式和范围；用户尚未执行的 Word 转换及视觉检查是 `NOT_RUN`，不能凭 Markdown 检查或脚本成功代签。
4. 失败后先区分环境问题、接口能力缺失、文档问题和未知问题；修正有依据的配置后重试。不要反复运行同一失败命令。
5. 某路线不可用时继续其他独立验证，记录降级路线。所有可用路线均无法完成所选模式必要的材料读取或内容输出时，保留草稿和精确阻塞项，不宣称相关检查通过。`word` 无法写入时停止正式 Word 写入；`md-only` 不因缺少 Word 写入、转换或排版能力而受阻。
6. 内网无执行工具或无法访问文件时，生成待执行命令和预期结果交给本机操作者；根据实际返回继续，不能自行补写测试结果。

## 2. Q0：建立工作区与确认边界

采用仓库内已忽略的 `.local/` 收拢本机工作区。环境记录跨报告复用，项目材料各自独立，例如：

```text
DocumentModify_main/
  docs/                         # 通用规范
  tools/                        # 通用工具
  .local/
    environments/office-pc/
      LOCAL-WORKFLOW.md         # 已验证操作方法
      ISSUES.md                 # 环境问题与解决记录
      scripts/                  # 已验证的本机脚本
      qualification/RESULTS.md  # 核验记录，夹具与日志放同目录
    projects/report-001/
      TASK.md                   # 引用本机环境及规程版本
      incoming/                 # 原始交接包，不改动
      originals/                # 原件，只读保留
      work/                     # 分章节草稿、模板约定与状态
      figures/                  # 图表与源文件
      delivery/                 # 经复核的交付版本
```

已有并列 `report-work` 时，按 [归并说明](../START.md#已有-report-work-的归并) 复制、核对和更新路径，不覆盖既有证据。本规范不保存本机测试日志；日志和问题进入环境目录，报告专属问题进入项目状态。

以下 Windows 命令从 `DocumentModify_main`（或实际仓库根目录）执行，先设置路径变量。替换环境和项目 ID 后创建所需目录，已有文件不会被覆盖：

```powershell
$Repo = (Get-Location).Path
$EnvRoot = Join-Path $Repo '.local\environments\office-pc'
$ProjectRoot = Join-Path $Repo '.local\projects\report-001'
New-Item -ItemType Directory -Force -Path "$EnvRoot\qualification", "$EnvRoot\scripts", "$ProjectRoot\incoming", "$ProjectRoot\originals", "$ProjectRoot\work", "$ProjectRoot\figures", "$ProjectRoot\delivery" | Out-Null
```

记录允许读取的目录、输出位置、软件自动化权限、加密文档复制/另存规则和资料流向。默认不允许向外网传递内网文件、摘录、日志和截图。本机权限不足时请求具体操作权限，不要求关闭加密或安全系统。

在 `TASK.md` 记录 `交付模式：word / md-only`，以及 AI 与用户各自负责的检查。`md-only` 可使用 `delivery/report.md` 和 `delivery/assets/`（仅有图片时创建），内容检查及交接结果记入 `work/STATUS.md` 或交付说明，不要求额外交付脚本、Word 或 PDF。

本规范中的路径均为示例。命令执行前替换为实际路径；后续记录保留可直接复现的本机命令。

**通过条件**：工作区可用，读写边界已明确，测试不会覆盖原件。

## 3. Q1：环境与离线工具探测

`md-only` 先验证能在允许目录保存、重读中文 Markdown，并按报告需要核对标题、表格、公式表达、本地图片引用及资源。只在读取或分析实际需要时使用 Python 或 WPS；不强制安装 `python-docx`、`pywin32`、Word 转换器或验证 PDF 导出。以下 Python 探测与依赖命令仅适用于选用这些工具的路线，不是 md-only 的统一前置条件。

选用 Python 时先确认命令；Windows 示例：

```powershell
python --version
# 如果上条不存在，再尝试：
py -3 --version
```

确认一个可用入口后，以下统一以 `python` 表示。进入携带的规范目录执行：

```powershell
python tools/document_check.py probe --output "$EnvRoot\qualification\environment.json"
```

该探测只记录平台、Python 和常见模块存在性，不读取环境变量或凭据，不导入 WPS、打开文档或确认 COM 可用性。按实际路线记录所用应用的版本和位数、是否有 GUI，以及终端加密软件是否要求特定应用打开文件；`word` 另记录需要的字体和 PDF 导出能力。不要导出完整注册表或系统环境。

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

**通过条件**：所选模式需要的工具可运行；`md-only` 能保存并重读 Markdown 和必要图片。未选用工具可注明 `N/A`；模块存在或导入成功不代表文档读取、写入或转换通过。

## 4. Q2：材料接入与可读性

1. 将外网包接入 `incoming`，从可信交接记录核对包来源。在未执行包内任意脚本前检查其用途和内容。
2. 若包含摘要清单，使用本规范随附工具验证：

```powershell
python tools/document_check.py verify --root "$ProjectRoot\incoming\writing-package" --manifest "$ProjectRoot\incoming\writing-package\MANIFEST.json"
```

退出码非零时停止使用该包，定位缺失、多余或不一致文件并重新接入。没有摘要时登记“未取得交接前摘要”，可建立本地基线，但不能声称已证明传输前后一致。

3. 按所需材料用合成 CSV 验证中文、字段、单位、空值、小数和行数可正确读取；用合成照片验证方向和尺寸。`word` 验证插入能力；`md-only` 验证图片保存及 Markdown 相对路径引用。实际材料另做字段和数量核对，不把“能打开”当作数据含义正确。
4. 两种模式都验证模板和代表性文档的可读路线；提取真实章名、顺序、层级、必填字段和图表要求，保留来源位置。使用 Python 检查普通 DOCX 时，对允许分析的工作副本执行：

```powershell
python tools/document_check.py inspect-docx --input "$EnvRoot\qualification\sample.docx" --output "$EnvRoot\qualification\sample-structure.json"
```

退出码 `0` 只表示普通 DOCX 基本结构可解析。`2` 表示无法按普通 DOCX 确认，例如非 ZIP、加密容器、缺少主文档或 XML 无效；不能仅凭该结果断言具体加密机制。`.doc` 不适用此检查。文件操作错误使用退出码 `1`。

如果资源管理器可见文件、WPS 能打开，但 Python 无法直接解析，应继续 Q3 的授权 WPS 路线，不改扩展名伪装格式，不使用解密尝试。

`md-only` 可使用可用的只读解析或授权应用读取；若 AI 无法直接读取，由用户在内网应用中核对并提供允许使用的章节、必填字段及要求摘录，记录原模板位置和人工核对范围。无需为读取而验证写入接口或转换工具。必要模板内容仍不可取得时，登记模板检查阻塞及受影响章节，不假定没有模板要求；可继续独立内容的草稿工作。

**通过条件**：材料数量和完整性情况清楚；数据、图片、模板有已验证的接入方式。敏感材料提取结果仅存内网工作区。

## 5. Q3：验证 Python 与 WPS 读写路线

`word` 按下述步骤验证所需读写能力。`md-only` 仅验证 Q2 选定的模板/材料读取路线：在允许副本上核对正文、表格、章节层级和所需信息与原文一致，记录工具或人工动作、提取范围及限制；已有 Q2 证据可复用。5.1 的写入夹具及 5.3、5.4 的写入/保存重开测试不作为 md-only 前置条件。若需要调查 WPS，只验证所需读取能力，不运行下面会修改副本的测试。不能省略公式、表格、图示中影响事实或模板的必要信息。

### 5.1 word：先制作覆盖格式的合成夹具

由操作者在目标 WPS 中创建测试文档，保存到环境的 `qualification/`。必须包含：

- 唯一正文标记 `DM_BODY_BASELINE_01`、中文、标题一级和二级、粗体/斜体；
- 2 行 2 列表格，单元格标记 `DM_TABLE_01`；
- 一张已知小图片、图题、正文引用；
- 分页符、页眉、页脚与页码；
- 若真实模板需要，加入公式、目录、交叉引用、合并单元格或浮动图片。

保存一份普通夹具（如本机允许），并通过本机批准流程保存一份处于真实加密策略下的夹具。不能拿未加密夹具的成功结果证明加密文档可用。保留初始文件、截图或允许生成的本地 PDF 基线；所有测试另用工作副本。

### 5.2 word：路线选择

| 路线 | 必须实际证明的能力 | 使用边界 |
| --- | --- | --- |
| Python 直接解析，如 python-docx | 能读到夹具正文、表格和所需结构；修改副本、保存、重开后保留格式 | 仅适用于实际可解析的文件；复杂域、公式等需单独验证 |
| Python 驱动授权 WPS 接口 | Python 可通过本机支持的接口打开受保护副本、提取、修改、保存并重开 | Windows 可调查 COM；其他平台按本机官方接口，不套用 Windows 命令 |
| WPS 内授权导入/导出配合 Python | WPS 可在允许位置导出受控中间内容，Python 处理后可回填并保留必要格式 | 中间文件仍按原资料范围管理；禁止默认导出不受保护的副本 |

先用最简单的可用路线验证，再决定正式方案。可能出现“正文用 WPS 提取、计算用 Python、合稿用 WPS”的组合。不能为了省事把包含表格、公式或图片的报告降为纯文本后宣称无损。

### 5.3 word：普通 DOCX 最小读写测试

仅在 Python 可直接解析普通夹具时执行。将以下代码保存为环境目录的 `qualification/test_plain.py`，它拒绝覆盖已有输出：

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
python "$EnvRoot\qualification\test_plain.py" "$EnvRoot\qualification\plain-fixture.docx" "$EnvRoot\qualification\plain-result.docx"
```

该测试只覆盖基本对象，不证明复杂对象和版面正确。继续 Q6。在真实模板副本上做代表性替换、表格更新和插图后重复检查。

### 5.4 word：WPS 接口测试指令

Windows 候选 COM 标识可在本机注册信息中定向查询，例如：

```powershell
Get-ItemProperty -LiteralPath 'Registry::HKEY_CLASSES_ROOT\KWPS.Application\CLSID' -ErrorAction SilentlyContinue
```

这是候选标识，不保证所有 WPS 版本均注册。未发现时查本机 WPS 帮助或已批准的接口文档，不能通过禁用安全策略解决。接口对象、打开参数和保存方法必须以本机支持情况为准。

如果确认该版本支持 `DispatchEx` 和下列对象方法，可将代码保存为环境目录的 `qualification/test_wps.py`。**传入的只能是事先准备好的测试工作副本，代码会修改该副本**；测试前关闭该副本，保留基线。先确认实例创建机制不会接管操作者正在编辑的文档。

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
python "$EnvRoot\qualification\test_wps.py" "KWPS.Application" "$EnvRoot\qualification\protected-working-copy.docx"
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

## 7. Q5：验证分析、内容检查与按模式合稿

先用已知答案的合成数据验证分析脚本：例如输入 `1, 2, 3`，均值应为 `2`；另验证单位换算、空值和明显非法值处理。此例只是最低演练，不替代真实报告所需的误差、拟合或统计验证。

确认图中的数据点、坐标和单位与输入对应，正文引用正确。正式材料接入后，保留原始数据和处理记录；没有实际数据时不能生成实测结果。

两种模式均核对事实、数据与单位、来源支持位置、必需章节、术语、图表和结论，缺少关键证据时保持草稿状态。没有分析任务时，分析子项可注明 `N/A` 及原因，内容检查仍须执行。

- `word`：将已核对的两章内容通过 Q3 选定路线写入真实模板的工作副本，包含表格和图片。对照写入前后的预期变更范围，核对章节、对象及关键内容，没有要求的段落不得消失。真实报告写入前先形成修改清单和可恢复的本地版本。
- `md-only`：将两章合并为 Markdown，按真实模板检查章名、顺序、层级和必填信息；重读文件，检查表格/公式表达、正文引用、参考文献及本地图片相对路径与文件对应，无须写入 Word。

**通过条件**：所需分析产生预期结果，合稿内容及证据关系保留；`word` 另须验证定点写入。`md-only` 的 Word 写入子项记为 `N/A`，不影响已完成的内容检查结论。

## 8. Q6：保存、重开与版面验收

本节的 Word 操作仅为 `word` 模式的 AI/操作者验收要求。`md-only` 将 AI 的本节操作记为 `N/A（模式范围外）`，将“Word 转换及排版待用户完成”独立记为 `NOT_RUN`；不执行本节作为 Markdown 交付门槛。用户后续在已有内网应用手动转换后，按下列适用项目检查并登记实际结果。

`word` 在目标 WPS/Word 中执行，不依赖 XML 计数代替视觉结果：

1. 保存、关闭测试文档，再重新打开；确认无需修复，正文标记、表格、图片、公式和样式保留。
2. 按模板需要更新目录、页码和交叉引用，检查更新后的显示结果。记录本机可用菜单或已验证接口。
3. 导出本地 PDF（如要求且允许）；核对所有页面的字体、分页、页眉页脚、图表截断、标题孤行、表格跨页、公式和目录页码。
4. 与基线及预期修改对照，确认原文未意外丢失或改变。若无 PDF 导出能力，使用 WPS 打印预览逐页检查并记录替代方式；交付明确要求 PDF 时，该替代不满足 PDF 交付要求。
5. 原件和工作副本在保存前后核对版本或摘要；加密系统可能使文件字节变化，语义和版面检查仍不可省略。

AI 若能在内网读取允许的页面渲染结果，可辅助逐页核对；否则交由操作者检查并记录结果。不上传页面或截图寻求外网检查。

**通过条件**：所需内容、复杂对象和版面均有检查证据。AI 无法视觉检查且操作者尚未确认时，排版验收保持 `NOT_RUN`。

## 9. Q7：固化本机操作规程并做最终演练

使用 [表单中的内网验证记录](05-templates.md#内网资格验证记录与本机规程)，在环境目录创建 `qualification/RESULTS.md`、`LOCAL-WORKFLOW.md` 和 `ISSUES.md`。它们分别保存证据、有效方法和问题经历，不写回本通用规范，不提交到外网仓库。项目 `TASK.md` 引用环境 ID、规程版本及证据 ID，项目专属模板要求和写作问题保留在项目内。

`LOCAL-WORKFLOW.md` 按交付模式填写实测得到的内容，模式范围外的能力明确注明 `N/A`：

- 交付模式、适用机器、系统、实际依赖的工具版本、加密条件、验证日期及失效条件；
- 材料接入、校验、提取、章节输入和恢复方法；
- 需要时的 Python 入口、离线依赖，所选材料/模板读取与内容输出路线及确切命令/人工动作；
- `word` 的保存、重开、目录更新、PDF 导出和逐页验收步骤；或 `md-only` 的 Markdown 合稿、必要图片资源和内容检查方法，以及用户转换和排版待办；
- 失败恢复、版本保留、人工步骤、已知不支持对象和回退方法。

规程标明版本和适用条件，并指向 `RESULTS.md` 的证据。将可复用脚本放在环境的 `scripts/`，不要只把有效代码留在聊天中。`ISSUES.md` 记录现象、原因或待查项、修正方法、复验结果和适用版本；仅复验通过的方法进入有效规程。

按该规程从一份新夹具副本重新演练所选路线：

- `word`：接入 → 模板读取 → 分章恢复 → 所需分析 → 模板写入 → 保存重开 → 排版检查。
- `md-only`：接入 → 模板读取 → 分章恢复 → 所需分析 → Markdown 合稿及图片引用 → 内容检查 → 交付并登记用户 Word 转换/排版待办。

仅重跑本次实际需要的路线。删除或注明未验证的候选方法，保留失败证据及限制。

最终按能力和交付模式分别报告：材料接入、模板读取、分章写作、分析、内容交付，以及 `word` 需要的写入与版面能力。所选模式的必要项目全通过才可声明该模式“本机流程已验证”；模式范围外的 `N/A` 不阻塞该结论，用户未做的 Word 工作仍单独列为 `NOT_RUN`，不得宣称 Word 已验证。

## 10. 正式任务的后续流程

在适用条件未变且验证记录有效时，使用本机规程接入真实材料、分章撰写、合稿和对应模式的检查。每份报告都执行事实、来源引用、模板内容和章节检查；`word` 另执行保存重开和版面检查，`md-only` 检查 Markdown 内容与必要图片资源后交接用户转换。资格验证通过不能代替每份报告自身的质量验收。

接入外网资料包后，先在项目内解析真实模板，建立 `work/template-contract.md`、`work/content-map.md` 和最终大纲。外网未见模板时提供的模块和建议结构需要重新适配，不直接当作最终章名。具体映射与启动指令见 [START.md](../START.md)。

`md-only` 合稿保存为 `delivery/report.md`，使用模板约定的章节和标题层级，正文内包含所需表格、公式表达、图题、引用及参考文献；必要图片存入 `delivery/assets/`，以相对路径引用，不依赖外网加载。分析源和事实核对证据留在本地项目中。检查所有必需章节、待补项、术语、数据与单位、来源定位和图片对应关系；已知转换语法限制如实记入交接，不把未经转换测试的效果写为已验证。

交付使用 [报告验收与交付记录](05-templates.md#报告验收与交付记录)，分别记录：

- **Markdown 内容检查结果**：逐项给出通过、失败或待验证及证据；关键事实、来源或章节缺口未解决时仍为草稿，全部必要内容检查通过可标“Markdown 内容已检查”。
- **Word 转换及排版待用户完成**：责任人为用户，状态 `NOT_RUN`；列出使用已有内网应用手动转换、应用模板样式、保存重开、检查目录/页码/表格/公式/图片及逐页版面的动作。AI 的 Word 写入、转换和验收标为 `N/A（md-only 模式范围外）`，不将用户待办标为通过或不适用。

用户转换及排版尚未完成不阻塞 Markdown 内容交付；在收到实际检查结果前，不宣称 Word 已定稿或已验收，也不要求 AI 先执行转换来证明 Markdown 可交付。
