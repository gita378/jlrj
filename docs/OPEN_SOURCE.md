# 开源复用与取舍

| 项目 | 复用位置 | 许可与取舍 |
| --- | --- | --- |
| [pypdf](https://github.com/py-pdf/pypdf) | `server.py` 直接调用 `PdfReader` 提取文本、尝试空密码解密 | BSD-3-Clause；本项目只处理可提取文本的 PDF。 |
| [openpyxl](https://openpyxl.readthedocs.io/) | `server.py` 直接调用 `load_workbook`、`Workbook` 读写台账与清单 | MIT；原始资产编号按文本处理。 |
| [MSL Equipment](https://github.com/MSLNZ/msl-equipment) | 借鉴设备登记与校准记录分离 | MIT；未复制源文件，也不引入仪器驱动。 |
| [InvenTree](https://github.com/inventree/InvenTree) | 借鉴台账、附件和可追溯记录的产品结构 | MIT；库存/采购子系统不适合本项目，未复制源文件。 |

原交接文档中的 LabOS 是付费软件，WCC_CMMS 附带 Commons Clause，Foxnic-EAM 部署依赖 Java/MySQL。均没有作为本地底座。PyMuPDF 是 AGPL/商业双授权且当前环境未安装，本版使用 pypdf。引用上游思想不等于复制其代码；本版实质代码复用集中在成熟、已安装的 PDF 和 Excel 库。
