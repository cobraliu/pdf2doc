# 扫描件转 Word

把**扫描版 / 拍照版 PDF**（没有文字层、复制不出字的那种）转换成**保留版式、可直接编辑**的 `.docx`。

全部处理在本机完成，**不联网、不上传、不调用任何云端 API**。

---

## 它做了什么

```
PDF ──渲染──> 页图 ──OCR──> 文字 + 坐标 ──版面重建──> Word
   PyMuPDF        RapidOCR              规则引擎      python-docx
```

OCR 只能给出「这一行字在哪个坐标」，把它还原成有结构的 Word 才是难点。本工具的版面重建做了这些：

| 处理 | 说明 |
|---|---|
| **两列还原** | 技术规格常见的「标签—值」两列（`Function` \| `to guide the cable`）还原成**无边框表格**，改右侧文字时左侧标签不会跟着移位 |
| **续行合并** | 被 OCR 拆断的长段落重新接回一段。判据是「上一行是否排到右边界」——没排满的行绝不吸收下一行，这样小标题不会被吞进正文 |
| **项目符号找回** | OCR 会直接丢弃 `•` 这类图形符号，靠「缩进超出正文基准」把列表项还原出来 |
| **动态行距** | 空行阈值按每页实际行距中位数计算，中英文混排、不同扫描分辨率都能自适应 |
| **噪声剔除** | 骑缝章、手写签名常被识别成乱码（如把印章读成「泊上场」），按「短字符串 + 低置信度」剔除 |
| **标题分级** | 识别 `一、` `1.1` `A.` `22.` 等编号形态，以及无编号的下划线小标题 |

---

## 使用

### 方式一：下载现成程序（推荐）

到 [Releases](../../releases) 下载对应平台的包：

| 平台 | 文件 | 说明 |
|---|---|---|
| Windows | `扫描件转Word.exe` | 双击运行；也可把 PDF 直接拖到 exe 图标上 |
| macOS | `扫描件转Word-macos.zip` | 解压后**右键 → 打开**（未做代码签名，直接双击会被拦） |
| Linux | `扫描件转Word` | `chmod +x` 后运行 |

界面：把 PDF（或整个文件夹）拖进左侧列表 → 点「开始转换」。

### 方式二：从源码运行

```bash
pip install -r requirements.txt

python gui.py                      # 图形界面
python convert.py 文件.pdf          # 命令行，可跟多个文件
python convert.py --rebuild 目录    # 跳过 OCR，用缓存重新出 docx（调参数时用）
```

> `--rebuild` 很有用：OCR 是最慢的一步，调完参数后用它秒出新版本，不必重跑识别。

---

## 示例

`examples/` 下有一份**合成的**扫描件示例（内容全部虚构，与任何真实合同无关）：

```bash
python convert.py examples/sample_scanned.pdf     # 生成 examples/sample_scanned.docx
python examples/make_sample.py                    # 需要时可重新合成这份示例 PDF
```

`sample_scanned.pdf` 是纯图像、无文字层的两页 A4，带轻微歪斜与噪点，覆盖了中英对照、
`1.` / `1.1` / `三、` 编号层级、两列「标签—值」、项目符号列表、页眉页脚与印章签名。
`sample_scanned.docx` 是它的转换结果，可以直接打开看效果。

---

## 参数说明

界面「参数」区对应引擎里的阈值，**改不出效果时优先动这两个**：

| 参数 | 默认 | 什么时候调 |
|---|---|---|
| **续行判定** | 0.78 | 段落被错误粘成一坨 → **调高**；本该连起来的句子被拆断 → **调低** |
| **两列分界** | 0.50 | 「标签—值」没被识别成表格 → 按右列实际起始位置调整 |
| 识别精度 | 2560 px | 小字/密排识别不准 → 调到高精度；追求速度 → 调到快速 |
| 项目符号缩进 | 0.030 | 列表项没被识别 → 调低；正文被误判成列表 → 调高 |
| 印章置信度 | 0.88 | 印章乱码混进正文 → 调高；正常短词被误删 → 调低 |

其余开关（页眉/页脚/印章剔除、页码标记、两列表格）按需勾选。

---

## 自己构建

PyInstaller **不支持交叉编译**——Windows 的 exe 必须在 Windows 上打，macOS 的 .app 必须在 macOS 上打。

```bash
pip install -r requirements.txt pyinstaller
python -c "from rapidocr import RapidOCR; RapidOCR()"   # 先下载模型, 否则打出空壳
pyinstaller pdfrec-gui.spec --noconfirm                  # 图形界面版
pyinstaller pdfrec.spec --noconfirm                      # 纯命令行版
```

仓库已配好 GitHub Actions（`.github/workflows/build.yml`）：打一个 `v*` 标签即可自动产出三平台安装包并发布 Release。

体积参考：命令行版约 137 MB，图形界面版约 165 MB（含 30 MB OCR 模型 + PySide6）。

---

## 已知限制

- **签名、印章、二维码、页眉页脚会被剔除**（默认行为，可在参数里关闭）。转出的 Word 若要作为正式文件，签章需另行处理。
- **原文的下划线、字符级格式不还原**，只保留标题加粗与段落结构。
- **图纸、示意图不会带过来**，只提取其上的文字。
- OCR 在型号、公差这类**密集数字**上仍可能出错，重要文件建议人工抽查一遍。

## 依赖与模型

| 组件 | 用途 | 授权 |
|---|---|---|
| PyMuPDF | PDF 渲染 | AGPL-3.0 |
| RapidOCR + PP-OCRv6 (ONNX) | 文字检测/识别/方向分类，共 30 MB | Apache-2.0 |
| python-docx | 生成 docx | MIT |
| PySide6 | 图形界面 | LGPL-3.0 |

> PyMuPDF 是 AGPL：若要闭源商用需购买商业授权，或替换为 pypdfium2（BSD）。
