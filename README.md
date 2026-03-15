由AI编写并推送

# 检索工具

一个基于 Python + Tkinter 的本地查询工具。

主要能力：
- 从 Excel 读取通用件数据
- 按动态字段组合查询
- 管理“类型 -> Excel/图标”映射
- 支持 GUI 与命令行两种模式

## 1. 开始

首次使用请按以下顺序，不要跳步：

1. 安装依赖

```bash
pip install -r requirements.txt
```

2. 启动 GUI

```bash
python retrieval_tool.py
```

3. 点击“设置”，输入口令 `2233`
4. 新增类型并填写：名称、Excel 路径、图标路径（图标可留空）
5. 回主界面后先选类型，再点“加载”
6. 状态栏出现“已加载类型”后再输入条件查询

高频误操作：
- 只切换类型但没点“加载”
- Excel 标题缺“编码/所属项目”或顺序错误
- 手工改 `types_config.json` 导致签名校验失败

## 2. 启动与参数

GUI（默认）：

```bash
python retrieval_tool.py
```

GUI 启动后自动预加载指定 Excel：

```bash
python retrieval_tool.py --excel "文件名.xlsx"
```

命令行交互：

```bash
python retrieval_tool.py --cli
```

命令行直接查询（旧参数）：

```bash
python retrieval_tool.py --excel "文件名.xlsx" --diameter 20 --length 1200
```

命令行直接查询（动态字段）：

```bash
python retrieval_tool.py --excel "文件名.xlsx" --field 直径 --value 20 --field 长度 --value 1200
```

运行自检：

```bash
python retrieval_tool.py --self-check
```

## 3. 打包与分发（Windows）

### 3.1 打包前检查

请确认以下文件都在项目根目录：
- `logo.png`（程序默认图标资源）
- `types_config.json`

### 3.2 单文件打包（推荐分发）

```bash
python -m PyInstaller --noconfirm --clean retrieval_tool_singlefile.spec
```

产物通常为：
- `dist/检索工具_单文件.exe`

### 3.3 目录版打包（适合调试）

```bat
build_exe.bat
```

产物通常为：
- `dist/检索工具/检索工具.exe`

### 3.4 分发前必做自检

```powershell
.\dist\检索工具_单文件.exe --self-check; echo $LASTEXITCODE
```

仅当退出码为 `0` 时再分发。

## 4. 关键规则

- Excel 标题必须包含“编码”和“所属项目”，且“编码”在前
- 可查询字段是“编码”和“所属项目”之间的列
- 建议只通过程序“设置”修改 `types_config.json`
- 图标建议使用 PNG，推荐 120x120 或 256x256

## 5. 常见问题

Q1：为什么切换类型后图标不立刻变化？

只有“加载成功”后，图标才切换为已加载类型，避免误判。

Q2：为什么提示配置被篡改？

通常是手工编辑配置文件或格式错误。请通过“设置”重新保存。

Q3：提示缺少 `openpyxl` 怎么办？

```bash
pip install -r requirements.txt
```

Q4：运行失败只看到一行错误怎么办？

程序会输出 `程序执行失败: ...` 并返回非 0 退出码。优先检查 Excel 路径、标题格式、依赖安装和 `--self-check` 结果。
