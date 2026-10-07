# 计量设备管理（Windows 桌面版）

本地离线运行的设备计量管理软件。主程序是原生 Windows EXE，界面按设备台账、送检、证书、结账对账、审计备份组织，业务数据保存在本机 SQLite。无需浏览器、服务端或网络。

## 直接使用

解压 `dist/计量设备管理-免安装版.zip`，保持 `计量设备管理.exe` 和 `_internal` 文件夹在同一目录，双击 EXE。整个文件夹可以复制到另一台 Windows 电脑；目标电脑无需安装 Python。首次运行自动创建数据库，数据存于 `%LOCALAPPDATA%\CNAS计量台账`，不会因移动或更新 EXE 消失。

当前交付的目录版已在 Windows 10 本机实际启动验证。`dist/计量设备管理.exe` 是另外构建的单文件包；在本开发环境中其自解压启动未完成，因此建议使用上述免安装目录版。

有 Python 3.11+ 的开发电脑也可运行 `start-local.cmd` 或 `python desktop.py`。安装依赖：`python -m pip install -r requirements.txt`。旧的本地浏览器界面仍可通过 `python server.py` 打开，用于开发排查。

## 工作流程

1. 在“设备台账”创建设备，或导入 Excel/CSV。资产编号按文本保存，支持搜索、编辑、归档和导出。
2. 在“送检记录”登记送检与返回日期；总览分别统计本月送检和校准。
3. 在“计量证书”批量导入可提取文本的 PDF。软件按哈希去重，提取设备编号、证书号、日期、结论等；字段完整、唯一精确匹配时自动确认并更新设备状态，有疑点则进入待确认。可以打开原件核对、更正。
4. 在“结账对账”导入 Excel/CSV 清单，核查重复项、证书/设备不匹配、金额差异与当月证书漏列。金额比较需要证书记录中有约定金额。
5. 在“审计与备份”导出日志和完整备份 ZIP，也可在软件内恢复。恢复前会自动保存现有数据的恢复包。

只处理可复制文本的 PDF，不做扫描/OCR。不同机构的证书模板可能需要人工确认。当前操作人是手动填写的记录字段，尚无多用户身份认证；适合单机本地使用。

## 构建与验证

源码测试：`python -m unittest discover -s tests -v`；桌面页面巡检：`python smoke_desktop.py`。构建依赖 PyInstaller 的 wheel：`python -m pip download pyinstaller -d wheelhouse`，再运行 `python build_exe.py --onedir`。构建程序把 wheel 解到本地临时目录，不向系统安装打包工具。构建时需要已安装 `requirements.txt` 的 Python，并在 Windows 上执行。运行 `python smoke_desktop.py --exe --onedir` 验证目录版 EXE 的启动与建库。

核心业务测试覆盖 Excel 导入、PDF 解析和自动回写、旧证书保护、异常、送检、对账及备份。未获得真实机构证书和结账清单样本，所以具体模板的字段提取准确率仍需用实际文件验证。

## 开源借鉴

直接使用 [pypdf](https://github.com/py-pdf/pypdf) 解析 PDF、[openpyxl](https://openpyxl.readthedocs.io/) 读写 Excel，并以 [PyInstaller](https://pyinstaller.org/) 构建 EXE。设备主档与校准历史分离的设计参考 [MSL Equipment](https://github.com/MSLNZ/msl-equipment)，台账模块与工作流布局参考 [InvenTree](https://github.com/inventree/InvenTree)。没有复制这两个项目的源文件。取舍见 [docs/OPEN_SOURCE.md](docs/OPEN_SOURCE.md)。
