# word-print-layout-pipeline

Word 文档排版一致性解决方案：**Linux 服务端 Word→PDF 转换** 与 **Windows 真实打印排版** 的双管线对比测试服务。

解决的核心问题：比对服务部署在 Linux 上，用 LibreOffice 等开源方案把 Word 转 PDF 后，**分页/折行结果与 Windows 上 Word/WPS 直接打印的排版不一致**（窜行、窜页），导致基于"文字 + 坐标"的文本比对出错。本项目提供两条可独立使用的管线和一个并排对比页面：

- **管线 A（linux-pipeline）**：纯 Linux，LibreOffice headless + UNO 折行修复器 → PDF → pdfplumber 解析文字和坐标。零授权成本。
- **管线 B（windows-printnode）**：Windows 节点，真实 Word/WPS 打印到 Microsoft XPS 虚拟打印机 → 解析 XPS 拿到**与打印预览逐页逐行一致**的文字、坐标和页面渲染图。需要一台装有 Word 或 WPS 的 Windows 机器（可复用办公电脑，零新增成本）。
- **对比页面（内置于管线 A）**：上传 Word 后两条管线同时跑，总体判定 + 逐页并排渲染 + 逐行文字/坐标对照，用于验收管线 A 的排版是否与管线 B（打印基准）一致。

## 为什么 LibreOffice 转 PDF 会和打印不一致

LibreOffice 的排版引擎与 Word 的折行规则不同，最典型的一类问题：**"数字串 + 全角标点"长串**。Word（及 WPS）把它当可断开的序列，而 LibreOffice 把它当成一个不可拆的西文单词整体推入下一行，造成整行下移、内容窜页。`uno_linewrap_fix.py` 通过 UNO API 读取文档真实字体、字号、行宽，按逆向出的 Word 断行规则模拟断行，在转换前修正这类文档，实测多个测试文档与 Windows 打印基准**逐页行数 100% 一致**。

## 目录结构

```
linux-pipeline/        管线 A：Linux word2pdf + 双管线对比 test 服务
  app.py                 Flask 服务（端口 8899）：上传 Word，A/B 双管线解析 + 对比页面
  uno_linewrap_fix.py    UNO 折行修复器（管线 A 核心，双断行模型）
  requirements.txt
  start_uno.sh           启动 LibreOffice UNO 监听（127.0.0.1:2002）
  lo-uno.service         systemd：UNO 监听常驻
  word2pdf-test.service  systemd：Flask 服务常驻
windows-printnode/     管线 B：Windows 打印解析节点（HTTP 服务，端口 8090）
  PrintParseService.cs   零外部依赖单文件服务：Word COM 打印 → XPS → 文字+坐标+页面 PNG
  setup.ps1              节点初始化：固定 XPS 打印机到文件端口、编译服务、注册自启
  setup.cmd              setup.ps1 的 cmd 包装（无人值守安装入口）
  autounattend.xml       Win10 VM 无人值守安装应答文件（密码已占位，需自行替换）
```

## 管线 A 部署（Linux，CentOS 7 验证通过）

```bash
# 1. 安装 LibreOffice 7.6 到 /opt/libreoffice7.6（版本固化，勿随意升级）
# 2. 部署代码
mkdir -p /opt/word2pdf-test && cp linux-pipeline/{app.py,uno_linewrap_fix.py,requirements.txt} /opt/word2pdf-test/
python3 -m venv /opt/word2pdf-test/venv && /opt/word2pdf-test/venv/bin/pip install -r /opt/word2pdf-test/requirements.txt
# 3. 注册并启动服务
cp linux-pipeline/{lo-uno.service,word2pdf-test.service} /etc/systemd/system/
systemctl daemon-reload && systemctl enable --now lo-uno word2pdf-test
# 4. 可选：接入管线 B 节点
systemctl edit word2pdf-test   # 加 Environment=XPS_API=http://<windows节点IP>:8090
```

打开 `http://<服务器>:8899`，上传 .doc/.docx 即可看到 A/B 两管线逐页并排对比。

## 管线 B 部署（任意装了 Word 或 WPS 的 Windows 机器）

不需要 VM——任何一台办公电脑即可：

```powershell
# 管理员 PowerShell，把 windows-printnode 目录拷到本机后：
powershell -NoProfile -ExecutionPolicy Bypass -File .\setup.ps1
```

setup.ps1 会：把 Microsoft XPS Document Writer 固定到文件端口（打印直接写文件、不弹窗）→ 用系统自带 csc.exe 编译 `PrintParseService.exe`（零外部依赖）→ 注册开机自启并启动。服务监听 8090：

```
GET  /health              → {"ok":true,"engine":"word"}
POST /api/print-parse     → 上传 doc/docx，返回每页文字+坐标+页面渲染图 URL
GET  /images/{job}/{n}.png → 页面渲染图
```

> 无人值守 VM 全新安装：`autounattend.xml` + `setup.cmd` + `setup.ps1` + `PrintParseService.cs` 打进 ISO 附载即可全自动完成（应答文件里的密码是占位符 `CHANGE-ME-Password`，务必替换）。

## 实测验证结果

| 测试文档 | 管线 A 每页行数 | 管线 B（打印基准）每页行数 | 结果 |
|---|---|---|---|
| 多页长数字串文档（7页） | 46,45,45,45,45,45,44 | 46,45,45,45,45,45,44 | ✅ 逐页一致 |
| 老版多页文档（7页） | 47,46,45,45,45,45,45 | 47,46,45,45,45,45,45 | ✅ 逐页一致 |
| 折行测试文档（1页） | 4 | 4 | ✅ |
| 单页授权书（1页） | 15 | 15 | ✅ |

## 授权说明

- 本项目代码：**MIT License**，可自由商用。
- 管线 A 依赖 LibreOffice（MPL 2.0），服务器端使用无授权问题。
- 管线 B 依赖 Windows + Microsoft Word 或 WPS：请确保所用机器本身的授权合法（复用已授权办公电脑最省事）。微软许可原则不允许在服务器上用 Office 自动化为无许可的客户端提供服务。
- 商业 SDK 替代方案（Aspose / Spire 等）实测折行与 Word 打印不一致，且授权费每年数千~数万美元，本项目不采用。

## License

MIT — 见 [LICENSE](LICENSE)。
