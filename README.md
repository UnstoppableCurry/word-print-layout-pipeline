<div align="center">

<img src="https://img.shields.io/badge/📄-Word%20Print%20Layout%20Pipeline-1f6feb?style=for-the-badge&labelColor=0d1117" alt="logo" height="46">

# Word Print Layout Pipeline

**Word 排版一致性双管线引擎 — Linux 转换 × Windows 真实打印，逐页逐行可验证**

[![Release](https://img.shields.io/github/v/release/UnstoppableCurry/word-print-layout-pipeline?style=flat-square&color=2da44e)](https://github.com/UnstoppableCurry/word-print-layout-pipeline/releases)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg?style=flat-square)](LICENSE)
[![Platform](https://img.shields.io/badge/Platform-Linux%20%7C%20Windows-1f6feb?style=flat-square)](#-架构总览)
[![LibreOffice](https://img.shields.io/badge/LibreOffice-7.6-18A303?style=flat-square&logo=libreoffice&logoColor=white)](#pipeline-a-部署)
[![.NET](https://img.shields.io/badge/.NET%20Framework-4.x-512BD4?style=flat-square)](#pipeline-b-部署)
[![Python](https://img.shields.io/badge/Python-3.x-3776AB?style=flat-square&logo=python&logoColor=white)](#pipeline-a-部署)

**[ English Documentation → README_EN.md ](README_EN.md)**

</div>

---

## 🎯 解决什么问题

文本比对服务部署在 Linux 上，用 LibreOffice 等开源方案把 Word 转成 PDF 再解析"文字 + 坐标"时，**分页与折行结果和 Windows 上 Word / WPS 直接打印的排版不一致**（窜行、窜页），导致比对误判。

本仓库提供一套**可独立部署、可交叉验证**的双管线方案：

| 管线 | 环境 | 技术路线 | 授权成本 |
|:---:|:---:|---|:---:|
| **A · `linux-pipeline`** | Linux | LibreOffice headless + **UNO 折行修复器** → PDF → pdfplumber 解析文字 + 坐标 | **零**（全开源） |
| **B · `windows-printnode`** | Windows | 真实 Word / WPS **打印**到 XPS 虚拟打印机 → 解析 XPS 文字 + 坐标 + 页面渲染图 | 复用已有 Office 授权 |
| **对比页** | 内置于 A | 上传 Word → 双管线同跑 → 总体判定 + 逐页并排渲染 + 逐行文字/坐标对照 | — |

## 🏗 架构总览

```mermaid
flowchart LR
    U["📄 上传 .doc / .docx"] --> A["<b>管线 A · Linux :8899</b><br/>LibreOffice + UNO 折行修复"]
    U --> B["<b>管线 B · Windows :8090</b><br/>Word COM 打印 → XPS 虚拟打印机"]
    A --> PA["PDF<br/>pdfplumber 文字 + 坐标"]
    B --> PB["XPS 解析<br/>文字 + 坐标 + 页面 PNG"]
    PA --> C["⚖️ 双管线对比页<br/>逐页并排 · 逐行对照 · 行数判定"]
    PB --> C
```

## ❓ 为什么 LibreOffice 转 PDF 会和打印不一致

LibreOffice 与 Word 的**折行规则不同**。最典型的一类：「数字串 + 全角标点」长串 —— Word / WPS 视其为可逐字断开的序列，而 LibreOffice 将其当作**不可拆分的西文单词**整体推入下一行，造成整行下移、内容窜页。

`uno_linewrap_fix.py` 通过 UNO API 读取文档真实字体、字号与行宽，按逆向还原的 Word 断行规则（双模型：token 模型 / 压缩逐字模型）模拟断行，并在转换前插入显式断行符强制对齐。实测多个文档与 Windows 打印基准**逐页行数 100% 一致**。

## 📁 目录结构

```
linux-pipeline/            管线 A · Linux word2pdf + 双管线对比 test 服务
├── app.py                   Flask 服务（:8899）上传 Word → A/B 双管线解析 + 对比页面
├── uno_linewrap_fix.py      UNO 折行修复器（管线 A 核心 · 双断行模型）
├── requirements.txt
├── start_uno.sh             启动 LibreOffice UNO 监听（127.0.0.1:2002）
├── lo-uno.service           systemd · UNO 监听常驻
└── word2pdf-test.service    systemd · Flask 服务常驻

windows-printnode/         管线 B · Windows 打印解析节点（:8090）
├── PrintParseService.cs     零外部依赖单文件服务：Word COM 打印 → XPS → 文字 + 坐标 + PNG
├── setup.ps1                节点初始化：固定 XPS 打印机文件端口 · 编译 · 注册自启
├── setup.cmd                setup.ps1 的 cmd 包装（无人值守安装入口）
└── autounattend.xml         Win10 VM 无人值守安装应答文件（密码为占位符，需替换）
```

## 🚀 Pipeline A 部署 · Linux（CentOS 7 验证通过）

```bash
# 1. 安装 LibreOffice 7.6 至 /opt/libreoffice7.6（版本固化，勿随意升级）

# 2. 部署代码
mkdir -p /opt/word2pdf-test
cp linux-pipeline/{app.py,uno_linewrap_fix.py,requirements.txt} /opt/word2pdf-test/
python3 -m venv /opt/word2pdf-test/venv
/opt/word2pdf-test/venv/bin/pip install -r /opt/word2pdf-test/requirements.txt

# 3. 注册并启动服务
cp linux-pipeline/{lo-uno.service,word2pdf-test.service} /etc/systemd/system/
systemctl daemon-reload
systemctl enable --now lo-uno word2pdf-test

# 4.（可选）接入管线 B 节点
systemctl edit word2pdf-test   # 添加 Environment=XPS_API=http://<windows节点IP>:8090
```

打开 `http://<服务器>:8899`，上传 `.doc / .docx` 即可查看 A/B 双管线逐页并排对比。

## 🖨 Pipeline B 部署 · Windows

> 无需虚拟机 —— **任何一台装有 Word 或 WPS 的办公电脑**即可：

```powershell
# 管理员 PowerShell，进入 windows-printnode 目录
powershell -NoProfile -ExecutionPolicy Bypass -File .\setup.ps1
```

`setup.ps1` 自动完成：XPS 打印机固定文件端口（打印直写文件、零弹窗）→ 系统自带 `csc.exe` 编译（零外部依赖）→ 注册开机自启并启动。

**HTTP API（:8090）**

| 方法 | 路径 | 说明 |
|:---:|---|---|
| `GET` | `/health` | 健康检查 → `{"status":"ok","engine":"word"}` |
| `POST` | `/api/print-parse` | 上传 doc/docx → 每页文字 + 坐标 + 页面渲染图 URL |
| `GET` | `/images/{job}/{n}.png` | 页面渲染图 |

<details>
<summary><b>🖥 无人值守 VM 全新安装（可选）</b></summary>

将 `autounattend.xml` + `setup.cmd` + `setup.ps1` + `PrintParseService.cs` 打包为 ISO 附载，Windows 安装全程自动完成并自愈部署服务。应答文件中的密码为占位符 `CHANGE-ME-Password`，**务必替换**。
</details>

## ✅ 实测验证

| 测试文档 | 管线 A 每页行数 | 管线 B（打印基准） | 结果 |
|---|:---:|:---:|:---:|
| 多页长数字串文档（7 页） | 46, 45, 45, 45, 45, 45, 44 | 46, 45, 45, 45, 45, 45, 44 | ✅ 逐页一致 |
| 老版多页文档（7 页） | 47, 46, 45, 45, 45, 45, 45 | 47, 46, 45, 45, 45, 45, 45 | ✅ 逐页一致 |
| 折行测试文档（1 页） | 4 | 4 | ✅ |
| 单页授权书（1 页） | 15 | 15 | ✅ |

## ⚖️ 授权说明

- **本项目代码**：[MIT License](LICENSE)，可自由商用
- **管线 A**：依赖 LibreOffice（MPL 2.0），服务器端使用无授权问题
- **管线 B**：依赖 Windows + Microsoft Word / WPS，请确保所用机器本身授权合法（复用已授权办公电脑最省事）；微软许可原则不允许在服务器端用 Office 自动化为无许可客户端提供服务
- **商业 SDK 替代方案**（Aspose / Spire 等）：实测折行与 Word 打印不一致，授权费每年数千至数万美元 —— 本项目不采用

---

<div align="center">
  <sub>Built with ❤️ for layout-faithful document processing · <a href="README_EN.md">English</a></sub>
</div>
