# 离线计量台账

单机运行的设备计量管理系统。以原始证书和校准记录为依据，生成设备当前状态；无法可靠解析的证书进入待确认队列。

## 启动

在已安装 Python 3.11+、`pypdf` 和 `openpyxl` 的 Windows 电脑上运行：

```bat
python server.py
```

浏览器打开 `http://127.0.0.1:8765`。只监听本机，不访问云服务。数据保存在 `data/`，首次运行自动建库。用 `python -m unittest discover -s tests -v` 验证核心流程。

若需安装依赖，请在**有网络的同平台电脑**执行 `python -m pip download -r requirements.txt -d wheelhouse`，将项目和 `wheelhouse/` 拷贝到目标机，再执行 `python -m pip install --no-index --find-links wheelhouse -r requirements.txt`。将来可以制作同平台、同 Python 版本的离线安装包；当前仓库未宣称已验证目标云电脑环境。

## 工作流程

1. 创建设备，或上传 Excel 台账。资产编号按文本保存，保留前导零。设备可归档，历史证书与审计仍保留。
2. 记录送检日期；返回后登记返回日期。看板的送检统计取自实际送检记录。
3. 批量导入可提取文本的 PDF。系统计算 SHA-256 去重并提取字段；设备编号精确匹配且证书编号、日期、结论明确时自动确认并回写台账。字段冲突或缺失时进入待确认，原始 PDF 保留。
4. 对待确认的证书核对并更正字段。确认后的校准记录保留历史；最新有效记录计算台账日期，不允许旧证书覆盖较新记录。
5. 对异常设备手动标记及解除，并写明原因。异常优先于自动到期判定。最新证书不合格时也显示异常，新证书合格后按新结果重新判定。
6. 导入结账清单（Excel/CSV）。只有设备、证书编号可精确匹配时才自动核对；金额需有该项的约定金额才会比较。填写核对月份后，还会找出当月证书中未列入清单的记录。
7. 查看看板、导出审计记录、下载完整备份 ZIP（数据库与证书）。

恢复备份：先关闭服务，执行 `python restore.py calibration-backup.zip`。程序先检验 SQLite 和证书哈希，原有数据另存为 `data/pre-restore-*.zip`，再恢复数据库与 PDF。

## 边界

- 仅处理可提取文本的 PDF，不包含 OCR。
- 自动提取依赖证书字段标签及文本布局。自动确认只在字段完整、设备编号唯一且精确匹配时发生；不同机构的模板仍需在有样例后逐个校验。
- “本月校准”依据校准日期，“本月送检”依据实际登记的送检日期，二者分别统计。
- 默认只允许本机访问；操作人由用户在页面填写，单机环境下不构成身份认证。
- 当前为可运行的本地版本，尚无用户权限模块。上线到多人共用云电脑前需先补身份认证与权限控制。

## 开源复用

直接调用了 [pypdf](https://github.com/py-pdf/pypdf) 的 PDF 解析 API 和 [openpyxl](https://openpyxl.readthedocs.io/) 的 Excel 读写 API。设备历史与状态分离的设计参考 [MSL Equipment](https://github.com/MSLNZ/msl-equipment) 的设备登记与校准记录思路；可扩展台账界面参考 [InvenTree](https://github.com/inventree/InvenTree)。当前未复制这两个项目的源文件。详细选择见 [docs/OPEN_SOURCE.md](docs/OPEN_SOURCE.md)。
