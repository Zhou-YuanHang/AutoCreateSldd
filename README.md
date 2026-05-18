# AutoCreateSldd — Simulink 数据字典自动生成工具

从 Excel 模板自动创建 Simulink 数据字典 (.sldd)，支持双向同步。

## 功能概览

```
┌─────────────────────────────────────────────┐
│           AutoCreateSldd v1.0               │
├─────────────────────────────────────────────┤
│  Excel → SLDD  │  SLDD → Excel              │
│  ──────────────┼──────────────────────────  │
│  填写 Excel    │  编辑 SLDD                 │
│  ↓             │  ↓                         │
│  生成 M 脚本   │  反向导出到 Excel           │
│  ↓             │  ↓                         │
│  写入 .sldd    │  保留全部格式               │
└─────────────────────────────────────────────┘
```

## 快速开始

### 方式一：图形界面（推荐）

```matlab
>> sldd_tool
```

弹出窗口后：
1. 点击「浏览」选择或确认文件（自动识别当前文件夹第一个 .xlsx）
2. 检查校验提示是否通过
3. 点击「开始同步」
4. 查看日志输出

首次使用可点击「导出模板」生成空白 Excel 模板。

### 方式二：命令行

```matlab
>> excel_to_sldd      % Excel → M 脚本 → SLDD 一键完成
>> m_to_sldd          % 单独将工作区对象写入 SLDD
>> sldd_to_excel      % SLDD → Excel 反向同步
```

### 方式三：生成模板（Python）

```bash
python generate_template.py
```

---

## 文件说明

| 文件 | 用途 |
|------|------|
| `sldd_tool.m` | **图形界面**（App Designer 风格，推荐使用） |
| `excel_to_sldd.m` | **主入口函数**：Excel → M 脚本 → 执行 → SLDD 一键完成 |
| `Excel2Workspace.m` | 核心转换：读取 Excel 各 Sheet，生成 M 脚本 |
| `m_to_sldd.m` | 将基础工作区中的 Simulink 对象写入 .sldd |
| `sldd_to_excel.m` | **反向同步**：SLDD → Excel（MATLAB 端） |
| `sldd_to_excel_helper.py` | 反向同步 Python 端：保持 Excel 格式写数据 |
| `generate_template.py` | 生成空白 Excel 模板（Python，双击运行） |
| `CHANGELOG.md` | 更新日志 |
| `功能清单.md` | 已实现功能完整列表 |
| `讨论记录.md` | 开发过程中的问题讨论记录 |

---

## Excel 模板格式

### Signal / Parameter（14 列）

| 列 | 字段 | 必填 | 说明 |
|----|------|------|------|
| A | VariableName | ✅ | 变量名 |
| B | Package | ✅ | 包名（如 Simulink / myPackage） |
| C | Object | ✅ | Signal 或 Parameter |
| D | CustomStorageClass | ✅ | 存储类 |
| E | DataType | ✅ | 数据类型 |
| F | InitialValue | ✅* | 初始值（Signal 选填，Parameter 必填） |
| G | HeaderFile | ✅ | 头文件 |
| H | DefinitionFile | ✅ | 定义文件 |
| I | Description | 选填 | 描述 |
| J | Min | 选填 | 最小值 |
| K | Max | 选填 | 最大值 |
| L | Unit | 选填 | 单位 |
| M | Dimensions | 选填 | 维度 |
| N | Complexity | 选填 | real / complex |

### Const（4 列）

Name / Value / DataType / HeaderFile（全部必填，生成时固定 CustomStorageClass = "Define"）

### Enum（4 列）

EnumName / EnumNumbers / Value / DataType（全部必填）

### Bus（6 列）

BusName / Description / HeaderFile / Alignment / PreserveElementDimensions / DataScope（全部必填）

### BusElement（6 列）

BusName（必填）/ ElementName（必填）/ DataType（必填）/ Dimensions（必填）/ Description（选填）/ Unit（选填）

---

## 同步策略

| Sheet | 第一列 | 写入策略 |
|-------|--------|---------|
| Signal / Parameter / Const / Bus | 唯一键 | 按名匹配：同名更新、无名追加 |
| BusElement / Enum | 非唯一键 | 全量替换：清空旧数据重写 |

---

## 依赖

- **MATLAB R2020b+**（推荐）
- **Simulink**（数据对象和数据字典）
- **Python 3 + openpyxl**（仅 SLDD→Excel 反向同步需要）
- **无其他第三方依赖**
