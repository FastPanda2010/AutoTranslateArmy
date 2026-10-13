# Infinity 军表自动翻译工具

这个项目用于读取 *Infinity the Game* 官方 Army JSON，并按照自定义的词汇对照表生成格式精美的中文 Word 军书（DOCX）及 PDF。

---

## 核心特性

- **官方数据自动同步**：输入阵营编号即可自动调用 Corvus Belli 官方 Army API 更新本地数据。
- **完整的军书内容结构**：
  - **阵营封面页**：自动居中渲染阵营（及母阵营）Logo、阵营中英文名称、版本号与版权声明。
  - **火力组表（Fireteam Chart）**：位于文档前列，清晰展示小队编组选项。
  - **单位表与选项表**：包含单位主表、附属遥控兵种、可选装备包表格。
  - **特战升级表（Team-Ops / Spec-Ops / Tacball）**：自动提取特战单位的技能升级与 Tacball 支持表。
  - **命令图标与环境徽章**：特性与命令栏自动内嵌专属图标（Regular、Irregular、Lieutenant、Impetuous 等）与特殊环境徽章（如 Deepspace、Surface）。
- **单位模型立绘持久化**：支持从已排版的 Word 中一键反向提取手动插入的兵种立绘，后续重新生成军表时自动按原尺寸精准回填。
- **确定性术语翻译**：不依赖不可控的机器翻译，严格使用词汇对照表；内置单复数形态回退与等级技能推断。

---

## 快速开始

### 1. 生成中文军书

最常用的简写命令（以编号 `703` 为例）：

```powershell
py translate_army.py 703
```

该命令会自动完成以下流程：
1. 请求官方 API 更新 `samples/703.json`。
2. 读取词表 `translations.csv`，并自动加载 `Asset/unit_images/703/` 中的单位立绘。
3. 输出中文军书至 `output/703-cn.docx`。
4. 将未翻译词汇汇总写入 `output/missing.csv`。

#### 完整自定义命令与导出 PDF

```powershell
py translate_army.py samples\703.json output\703-cn.docx --glossary translations.csv --missing output\missing.csv --pdf-output output\703-cn.pdf
```

- `--pdf-output`（简写 `--pdf`）：生成 Word 后，调用本机 Microsoft Word 自动导出对应的 PDF 文件（仅限 Windows）。
- `--glossary`：词汇对照表路径，支持 `.csv` 或 `.json`（默认 `translations.csv`）。
- `--missing`：指定未翻译术语的导出路径。

---

## 单位立绘提取与自动回填

为了解决“重新生成军表会冲掉手动插入的单位立绘”的问题，项目提供了图片持久化工作流：

### 工作流程

```
[初次生成 DOCX] 
       ↓ 
[在 Word 中手动给各单位表头插入立绘并调整尺寸]
       ↓
[运行 extract_unit_images.py 提取归档到 Asset/unit_images/<faction>/]
       ↓
[后续无论何时用 translate_army.py 重新生成，图片与尺寸均自动精准回填]
```

### 提取命令

当你在生成的 Word 中手动插入并排版好单位立绘后，运行：

```powershell
py extract_unit_images.py output\703-cn.docx samples\703.json
```

程序会：
1. 自动解析 DOCX 内部的 OpenXML 结构，定位包含图片的表格表头。
2. 将图片导出到 `Asset/unit_images/703/` 目录下，并以清洗后的单位英文名命名。
3. 生成 `manifest.json`，精确记录每张图片的单位归属、文件哈希以及你在 Word 中调整的显示尺寸（EMU）。
4. 下次运行 `py translate_army.py 703` 时，程序会自动加载这些立绘并按原尺寸居中插入到对应表头中。跨阵营的同名单位也会自动复用已有的模型图。

---

## 词汇表（Glossary）配置

默认词表文件为 `translations.csv`（UTF-8 编码，包含表头）：

```csv
category,source,target
unit,MARUTS,摩卢陀
profile,MARUT,摩卢陀
skill,Courage,勇气
weapon,Combi Rifle,复合步枪
```

### 规则说明

1. **优先级机制**：`category` 可以填写具体类别，也可以使用 `*` 作为全局匹配。程序优先使用具体类别匹配，其次回退到全局 `*` 匹配。
2. **故意留空**：若希望保留某个特定英文原词不翻译，可保留空的 `target`。
3. **大小写与单复数回退**：匹配时不区分大小写；当遇到复数术语（如 `Drones`）且词表中只有单数（`Drone`）时，程序会自动进行词尾屈折推断并成功命中。
4. **规则推断**：针对带等级的特定技能（如 `Martial Arts L1` ~ `Martial Arts L5`），代码内嵌了模式推断规则（自动译为 `武术1级` 等），无需在词表中逐级穷举。

### 常用 Category 类别

| 类别 | 说明 | 示例 |
| :--- | :--- | :--- |
| `unit` | 单位总名 | `MARUTS` -> `摩卢陀` |
| `profile` | 单位具体兵种/配置名 | `MARUT` -> `摩卢陀` |
| `skill` | 特殊技能 | `Courage` -> `勇气` |
| `weapon` | 武器名称 | `Combi Rifle` -> `复合步枪` |
| `equip` | 装备名称 | `Hacker` -> `黑客装置` |
| `char` | 单位特性 | `Regular` -> `正规`, `Hackable` -> `可被黑入` |
| `type` | 兵种体型类型 | `LI`, `MI`, `HI`, `TAG`, `REM` |
| `category` | 部队类别 | `Line Troops`, `Veteran Troops` |
| `extra` | 属性修正/升级文本 | `-3`, `+1B`, `AP` |
| `order` | 命令类型 | `REGULAR`, `LIEUTENANT` |

---

## 补词与增量维护

1. 每次运行生成脚本后，检查 `output/missing.csv`。
2. 该文件中记录了当前军表中出现、但词表中尚未收录的所有专有名词。
3. 直接将需要翻译的行补充中文 `target` 后，复制粘贴进 `translations.csv`，再次运行生成命令即可完成汉化补充。
