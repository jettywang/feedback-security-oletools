# 加密附件分流实验（oletools 分支追踪 + 本地样本验证）

目标：解释“附件只带了看起来像加密 Office 文档的流名，宏扫描却标记为
`加密，已跳过`”这一结论的**实际判定依据**，并用 4 类本地小样本给出
**入库状态与触发条件**，让后续作业知道哪些附件应解密后重扫。

所有样本均在本地合成或复用仓库自带测试数据，无外部下载、无真实恶意代码。

## 1. 三条代码分支的真实走向

### 1.1 格式识别（ftguess）
入口：`oletools/ftguess.py` 的 `FileTypeGuesser`，按魔数（magic）分派。
- OLE/CFB：魔数 `D0 CF 11 E0 A1 B1 1A E1`，由
  `FType_Generic_OLE.recognize()` → `olefile.isOleFile(data)` 判定，
  成功后 `olefile.OleFileIO()` 真正打开并解析目录/FAT（`ftguess.py:318`）。
- ZIP/OpenXML：魔数 `PK`，`zipfile.is_zipfile()` 后再找
  `[Content_Types].xml`（`FType_Generic_OpenXML`，`ftguess.py:387`）。
- 关键点：**加密的 OOXML（.docx/.xlsx/.pptm 加密码）外层是 OLE 容器，不是
  ZIP**。所以加密文档在 ftguess 里只会得到 `container=OLE, filetype=OLE`
  （Generic OLE），看不到内部真实格式——真实格式要等解密后才知道。
- `olefile.isOleFile()` 只检查 8 字节魔数和少量 header 字段，**魔数正确不代表
  容器可解析**（见损坏样本 3a/3b）。

### 1.2 加密检查（oleid → crypto → msoffcrypto）
入口：`oletools/oleid.py:334 check_encrypted()`，仅当容器是 OLE 时调用
`oletools/crypto.py:154 is_encrypted(ole)`。判定分两层：

1. 首选 msoffcrypto（`crypto.py:183-196`）。
   `msoffcrypto.OfficeFile()`（`msoffcrypto/__init__.py:44-78`）先按魔数打开
   OLE，再按流名分派：
   - 存在 `EncryptionInfo` 流 → `OOXMLFile`（加密 OOXML 的标志）。
   - 存在 `wordDocument`/`Workbook`/`PowerPoint Document` → 对应 97 格式。
   - 都没有 → 抛 `FileFormatError: Unrecognized file format`。

   `OOXMLFile.__init__()` 会**真正解析 `EncryptionInfo`**
   （`msoffcrypto/format/ooxml.py:149-175`）：读版本号
   （agile=4.4 / standard=2|3|4.2），agile 还要 `parseString()` 解析 XML。
   内容非法即抛 `ExpatError`/`DecryptionError`/`struct.error`。
   - 97 格式不靠流名，而靠真实结构字段：Word 读 FIB 的 `fEncrypted`
     （`doc97.py:488`），Excel 在 Workbook BIFF 里找 `FilePass` 记录
     （`xls97.py:656`）。

2. msoffcrypto 抛异常时，oletools **吞掉异常**并回退到自带的
   `_is_encrypted_ole()`（`crypto.py:216-229, 247-288`）。该回退**只看
   名字/标志，不验证流内容是否有效**：
   - `\x05SummaryInformation` 的 PIDSI_DOC_SECURITY(0x13) 最低位=1；
   - 存在名为 `EncryptionInfo` 的流 → 直接 `return True`（`crypto.py:262`）；
   - 存在 `EncryptedSummary` 且无 `SummaryInformation`；
   - Word 的 `WordDocument` 头部 fEncrypted 位。

**结论（回答“需要什么证据”）**
- 标记“加密”并不要求解密成功，甚至不要求 `EncryptionInfo` 内容合法。
- 在装有 msoffcrypto 的当前环境，对 OOXML 型加密，结论实际由
  “**OLE 容器里存在名为 `EncryptionInfo` 的流**”这一结构事实 +
  msoffcrypto 成功解析共同支撑；一旦 msoffcrypto 解析失败，oletools 仍会
  仅凭流名回退为 `encrypted=True`。
- 因此“只有加密流名”的附件会得到与真加密**相同的 oleid 结论**，但二者在
  “能否解密”上截然不同（见第 3 节），必须用 msoffcrypto 的解析/解密结果
  做二次区分，不能只看 oleid 的布尔值。

### 1.3 宏提取（olevba）与“已跳过”
`oletools/olevba.py`：
- OLE 容器走 `open_ole()`，`type=OLE`（`olevba.py:2837`）。加密 OOXML 的
  OLE 外层没有 `VBA/_VBA_PROJECT`、`VBA/dir`、`PROJECT` 这些流，
  `find_vba_projects()`（`olevba.py:3183`）返回空 → 提取不到任何宏。
- 主流程 `process_file()` 先用“无加密”路径跑，随后再调用
  `crypto.is_encrypted()` 复核（`olevba.py:4513` 与异常路径 `4522`）。
- 一旦为加密：进入解密分支 `crypto.decrypt(...)`（`olevba.py:4570`），
  用 `options.password + crypto.DEFAULT_PASSWORDS`
  （含 `VelvetSweatshop`、`/01Hannes Ruescher/01`）尝试，成功则递归重扫
  解密产物（`olevba.py:4576`，`crypto_nesting+1`，受
  `MAX_NESTING_DEPTH` 限制）。
- 退出码（`olevba.py:527-536`）：
  - `0` 正常；`5` 打开失败；`6` 解析失败；
  - `8 RETURN_UNEXPECTED`：解密/解析抛未处理异常；
  - `9 RETURN_ENCRYPTED`：确认为加密但密码不对（`WrongEncryptionPassword`）。

## 2. 本地样本与实测证据

样本由 `scripts/build_samples.py` 生成，`scripts/triage.py` 采集证据，
原始结果保存在 `results/triage.json`。复现：

```bash
.venv/bin/python encryption-triage-lab/scripts/build_samples.py
.venv/bin/python encryption-triage-lab/scripts/triage.py
```

| 样本 | 容器/魔数 | ftguess | msoffcrypto | oleid `Encrypted` | olevba 退出码 | 含义 |
|---|---|---|---|---|---|---|
| `1_encrypted_valid.docx` | OLE，魔数正确 | OLE | `OOXMLFile`, encrypted=True，无异常 | **True** | **9** | 真加密，密码未知，宏在密文内不可见 |
| `2_fake_named_stream.bin` | OLE，结构合法 | OLE | 构造即抛 `ExpatError`（流内容非法） | **True**（回退仅凭流名） | **8** | 仅流名伪装；判定为加密但解密必崩 |
| `3a_corrupted_truncated.bin` | 魔数正确，正文截断 | Unknown | `NotOleFileError` | False（VBA 指标=Error） | **8** | 容器损坏，无法解析 |
| `3b_corrupted_fat.bin` | 魔数正确，目录/指针被破坏 | OLE | `FileFormatError: Unrecognized` | False | 0（流为空，未崩） | 损坏但 olefile 容错，无宏可提 |
| `4_plain_macro.docm` | ZIP/OpenXML | Word2007_DOCM | `OOXMLFile`, encrypted=False | False，**VBA=Yes** | 0，提取 4 个模块 | 正常未加密宏文档，直接扫描 |

补充实测：
- 真加密样本错误密码 → msoffcrypto `InvalidKeyError`；oletools `decrypt`
  返回 `(None, None)`，对应 olevba `9`。
- 伪流样本 `decrypt` 直接抛 `ExpatError`，对应 olevba `8`——**解密阶段就能
  把伪装者和真加密分开**。
- 标准“透明密码”样本
  `tests/test-data/encrypted/autostart-encrypt-standardpassword.xlsm`
  可用 `VelvetSweatshop` 自动解密，产物是含 `vbaProject.bin` 的 OpenXML，
  递归重扫检出 `Auto_Open`（AutoExec）。证明解密→重扫闭环有效。
- 97 格式加密（`encrypted.doc/.xls`）没有 `EncryptionInfo`，靠
  `WordDocument`/`Workbook` 内的 FIB `fEncrypted` / BIFF `FilePass` 判定，
  同样 `is_encrypted=True`。

## 3. 入库状态与触发条件（给后续作业）

建议给每个附件落库一个 `scan_state`，分流规则如下：

| scan_state | 触发条件（同时满足） | 是否解密重扫 | 处置 |
|---|---|---|---|
| `ENCRYPTED_VALID` | OLE 容器且 msoffcrypto **成功构造**并 `is_encrypted=True`（OOXML 有可解析的 `EncryptionInfo`+`EncryptedPackage`；或 97 格式 FIB/FilePass 置位） | **是** | 入解密队列：先试 `DEFAULT_PASSWORDS`，再用业务侧密码本；解密成功后以 `crypto_nesting+1` 递归重扫，最深 `MAX_NESTING_DEPTH` |
| `ENCRYPTED_NO_PASSWORD` | 同上但所有密码解密失败（`InvalidKeyError` / `decrypt→(None,None)`，olevba rc=9） | 是（待密码） | 挂起等待密码/外部投递；标记“密文未检”，不可判 benign |
| `ENCRYPTION_SPOOF` | OLE 中存在 `EncryptionInfo`/`EncryptedPackage` **流名**，但 msoffcrypto 构造或解析抛异常（`ExpatError/DecryptionError/struct.error`，olevba rc=8） | 否（解密无意义） | 判为结构异常/伪装，转人工或沙箱；oleid 的 True 仅来自流名回退，**不等于真加密** |
| `CORRUPT_CONTAINER` | 魔数正确但 `olefile.OleFileIO` 抛错（`NotOleFileError`，rc=8）或可打开却目录为空/`FileFormatError`（rc=0 但无流） | 否 | 判损坏，记录缺陷类型；无法支撑任何宏/加密结论，转修复或人工 |
| `PLAIN_MACRO` | OpenXML/ZIP（或未加密 OLE），msoffcrypto `is_encrypted=False`，且存在 VBA 工程（`PROJECT`+`VBA/_VBA_PROJECT`+`VBA/dir`，或 zip 内 `vbaProject.bin`） | 否（直接扫） | 直接跑 olevba/mraptor，按宏内容定级 |
| `PLAIN_CLEAN` | 未加密且无 VBA/XLM 工程 | 否 | 常规放行结论 |

判定要点（避免误分流）：
- **不要只信 oleid 的 `encrypted` 布尔值**：伪流（`ENCRYPTION_SPOOF`）与
  真加密（`ENCRYPTED_VALID`）该值都为 True。以 msoffcrypto 是否“无异常构造
  成功”作为第一道分界，再以 `decrypt` 结果区分“待密码”与“伪装/损坏”。
- **魔数正确 ≠ 可解析**：损坏容器（3a/3b）同样带 OLE 魔数，但不会被报加密；
  入库时把“无法解析”与“加密”分开计数。
- 解密后容器类型可能从 OLE 变为 OpenXML，必须对解密产物重新走完整
  ftguess→oleid→olevba 流程，而不是复用外层结论。
- 透明写保护（`VelvetSweatshop` 等）在结构上算加密，应自动解密重扫，
  不应停留在“已跳过”。
