# Infinity 军表自动翻译工具

这个项目用于读取 Infinity the Game 官方 Army JSON，并按你的词汇对照表生成中文 Word 军书。

程序会读取 `units` 生成单位表，并读取 `fireteamChart` 在文档最前方生成火力组表；`reinforcements` 增援数据不会读取。单位表只输出 `factions` 包含当前军表编号的单位，`factions=[]` 的佣兵池单位不会输出。

## 快速开始

```powershell
& py translate_army.py samples\703.json 你想输出路径/文件名（比如output/703-cn.docx） --glossary translations.csv --missing output\missing.csv
```

每次运行都会根据输入 JSON 的文件名自动更新官方数据：例如输入 `samples\205.json` 时，程序会以 `Origin: https://infinityuniverse.com` 请求 `https://api.corvusbelli.com/army/units/en/205`，确认返回有效 JSON 后再替换本地文件。因此输入文件名必须是纯数字形式，如 `205.json`。

生成结果：

- `output/703-cn.docx`：中文军书 Word 文件
- `output/missing.csv`：JSON 中出现但词汇表未覆盖的词，补进 `translations.csv` 后重新运行即可

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

如果想覆盖某个译名，直接修改 `translations.csv` 即可；如果想故意省略某个词，保留空的 `target`。

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
