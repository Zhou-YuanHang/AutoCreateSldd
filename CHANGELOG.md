# 更新日志

> 所有重要的功能变更和修复记录在此文件。

---

## [1.3.0] — 2026-05-15

### 新增
- SLDD → Excel 反向同步功能（`sldd_to_excel.m` + `sldd_to_excel_helper.py`）
  - 读取 .sldd 数据字典，按类型分类写入同名 .xlsx
  - 扫描目录下 Enum .m 文件，解析类定义写入 Enum sheet
  - MATLAB 计算数据，Python + openpyxl 写入 Excel，保留全部格式
- 写入策略：
  - Signal / Parameter / Const / Bus：按第一列唯一键匹配，同名更新、无名追加
  - BusElement / Enum：全量替换（清空旧数据重写）

### 修复
- 自定义存储类（如 `ExportedGlobal`）的 HeaderFile/DefinitionFile 在 SLDD 中不存，反向同步时不再误判为"修改"
- Enum 多条同名的行不再挤在同一行（全量替换）
- BusElement 多条同 BusName 的行不再挤在同一行（全量替换）
- MATLAB 函数模式下 `evalin('base')` 确保 M 脚本在基础工作区执行
- `find(dataSect, 'IncludeReferences', false)` 兼容低版本 MATLAB
- `isequal` 对比改用 `struct` 转换绕开 handle 类陷阱

---

## [1.2.0] — 2026-05-15

### 新增
- `excel_to_sldd.m` 改为主入口函数，一键完成：读 Excel → 生成 M 脚本 → 执行 → 写入 SLDD
- `m_to_sldd.m` 独立函数：工作区对象 → 写入 .sldd（含对比逻辑）
- SLDD 外部引用保护（`open` 不重建，仅操作本字典条目）

### 变更
- `excel_to_sldd.m` 从脚本改为函数，运行后不污染基础工作区
- `listdlg` 默认选中改为按名称匹配（Signal / Parameter），不写死索引

### 修复
- `class(v)` 精确匹配改 `isa(v, ...)` 判断继承，支持 `myPackage.Signal` 等子类
- `CoderInfo.CustomAttributes` 用 `try-catch` 包裹，自定义包不支持的属性自动跳过

---

## [1.1.0] — 2026-05-14

### 新增
- BusElement 生成格式改为 MATLAB 标准 `saveVarsTmp{1}` + 列索引（参考 `Ref/Test/busTemp.m`）
- Const 支持：按 Parameter 处理，`CustomStorageClass` 固定为 `Define`
- Enum 支持：每个 EnumName 生成独立 `.m` 类定义文件

### 修复
- Bus 优先于 BusElement 处理，避免名称依赖问题
- `myPackage.Signal` 模式下 `CustomAttributes.HeaderFile` 报错
- `YS_SelectRAMSignal` 自定义存储类跟随 Package，不硬编码 Simulink
- Parameter 的 `Dimensions` 在 `Value` 前设置（Simulink 约束）

---

## [1.0.0] — 2026-05-14

### 新增
- 项目初始化
- `generate_template.py`：生成空白 Excel 模板，8 个 Sheet
- `Excel2Workspace.m`：读取 Excel → 生成 M 脚本
- `excel_to_sldd.m`：入口脚本
- 下拉列表引用 Config 表数据源（集中管理可选值）
- Signal / Parameter / Const / Enum / Bus / BusElement 完整字段映射

### 说明
- 参考 `Ref/f1/` 的分步模式
- 模板生成用 Python，转换器用 MATLAB
