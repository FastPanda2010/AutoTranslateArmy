# Infinity 军表自动翻译工具

这个项目用于读取 Infinity the Game 官方 Army JSON，并按你的词汇对照表生成中文 Word 军书。

## 快速开始

如果你有 WordReplacer II 的导出 txt，先把它转换成本工具自己的规则文件。这个步骤只需要在插件词库更新时做一次：

```powershell
& "C:\Users\zhaoj\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe" import_wordreplacer.py ".\英菲写表器中文插件_更新至黑风行动by天部，悠兔，9，蛋哥.txt" wordreplacer_rules.json
```

之后生成军书时，主工具只读取 `wordreplacer_rules.json` 和 `translations.csv`，不会直接读取原始 txt：

```powershell
& "C:\Users\zhaoj\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe" translate_army.py samples\703.json output\703-cn.docx --glossary translations.csv --rules wordreplacer_rules.json --missing output\missing.csv
```

生成结果：

- `output/703-cn.docx`：中文军书 Word 文件
- `output/missing.csv`：JSON 中出现但词汇表未覆盖的词，补进 `translations.csv` 后重新运行即可
- `wordreplacer_rules.json`：从 WordReplacer II txt 转出的本地规则文件

## 词汇表格式

`translations.csv` 使用 UTF-8 CSV，至少包含：

```csv
category,source,target
unit,MARUTS,摩卢陀
profile,MARUT,摩卢陀
skill,Courage,勇气
weapon,Combi Rifle,复合步枪
```

`category` 可以写具体类别，也可以写 `*` 作为全局翻译。程序会优先使用具体类别，其次使用全局翻译。

`translations.csv` 的优先级高于 `wordreplacer_rules.json`。如果某个插件译名你想覆盖，写进 `translations.csv` 即可；如果想故意省略某个词，保留空的 `target`。

常用类别：

- `unit`：单位总名
- `profile`：单位配置/兵种名
- `skill`：特殊技能
- `weapon`：武器
- `equip`：装备
- `char`：单位特性，如 Regular、Hackable
- `type`：兵种类型，如 LI、MI、TAG
- `category`：部队类别，如 Line Troops
- `extra`：括号中的修正/升级文本，如 `-3`、`+1B`
- `order`：命令类型，如 REGULAR、LIEUTENANT

## 设计说明

程序不会调用机器翻译。它只使用你提供的词汇表做确定性替换；没有翻译的内容会保留英文，并记录到 `missing.csv`。

主程序 `translate_army.py` 不读取 WordReplacer II 原始 txt。原始 txt 只由 `import_wordreplacer.py` 在转换阶段读取。
