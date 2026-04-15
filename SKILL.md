---
name: customs-declaration
description: |
  生成跨境电商出口报关全套资料（出口合同、Invoice、Packing List、报关单草稿）。
  当用户提到"报关单"、"报关资料"、"出口报关"、"customs declaration"、"报关草单"、
  "装箱单"、"packing list"、"invoice 发票"，或者用户提供了采购合同Excel和亚马逊FBA货件PDF
  并希望生成出口文件时，使用此skill。也适用于商检校验场景。
---

# 出口报关资料生成 Skill

## 概述

本skill将采购合同(Excel)和亚马逊FBA分仓货件(PDFs)作为输入，经过分仓归类、金额平摊、海运费计算等步骤，自动生成四份报关文件：出口合同、Invoice、Packing List、报关单草稿。

所有确定性的解析和生成工作由 `scripts/` 下的Python脚本完成，你负责流程调度、用户交互和异常处理。

## 依赖

运行脚本前确保安装了依赖：
```bash
pip install xlrd openpyxl --break-system-packages -q
```

## 完整流程

### Step 1: 收集输入文件

向用户确认以下输入是否齐全：

> **重要：采购合同和FBA货件文件必须由用户上传或明确指定文件路径/文件夹。如果用户未提供，应主动提醒用户提供，不要自行在用户系统中搜索或猜测文件位置。**

| # | 输入 | 形式 | 必需 |
|---|------|------|------|
| 1 | 采购合同 | Excel (.xls/.xlsx) | 是 |
| 2 | 亚马逊FBA分仓货件 | 多个PDF文件 | 是 |
| 3 | 分票报关指令 | 文字（哪些仓点合并为一票） | 是 |
| 4 | 报关汇率 | 数字（通常5~8，如5.5） | 是 |
| 5 | 海运费单价 | 数字（元/千克） | 是 |
| 6 | SKU知识库 | 腾讯文档在线表格 或 用户手动输入（海关编码、英文品名、申报要素、材质、法定第一单位、第二单位） | 是 |

**知识库获取方式：**
- **优先从腾讯文档获取**：使用腾讯文档skill读取在线表格。腾讯文档列名为：`产品 | 英文品名 | HS CODE | 单位一 | 单位二 | 材质`。其中"产品"列为产品中文名，需要用采购合同中解析出的产品中文名（name_cn）在表格中搜索匹配，匹配成功后获取该行的英文品名、HS CODE、单位一、单位二、材质等信息。
- **腾讯文档获取失败时的兜底方案**：如果腾讯文档连接失败、搜索不到匹配产品、或读取异常，可以让用户手动输入产品所需信息。**必须提醒用户提供以下所有字段**：英文品名、HS CODE、单位一、单位二、材质，确保信息完整，否则无法正确生成报关资料。

**知识库JSON格式：**
无论从腾讯文档获取还是用户手动输入，收集到信息后统一转为JSON文件（保存到 `/tmp/knowledge_base.json`），再传给脚本。格式如下：
```json
{
  "产品中文名或SKU": {
    "tariff_code": "3918909000",
    "english_name": "Artificial Grass Tiles",
    "declaration_elements": "0|0|塑料|人造草坪|无品牌|无型号",
    "material": "plastic",
    "unit_1": "千克",
    "unit_2": "平方米"
  }
}
```
其中键名为采购合同中的产品中文名（name_cn），需与采购合同解析结果中的 `name_cn` 字段一致。

**产品名称来源规则：**
- **中文名（name_cn）**：从采购合同A列解析。若单元格内有换行则取第一行，无换行则取整个名称。
- **英文名（english_name）**：仅从知识库获取。采购合同中不提取英文名。

如果用户没有提供知识库，可以从采购合同中提取基础信息，但要提醒用户补充海关编码、英文品名、申报要素、法定第一单位和第二单位。

### Step 2: 解析输入文件

#### 2a. 解析采购合同

```bash
python {baseDir}/scripts/parse_purchase_contract.py <采购合同路径> --output /tmp/purchase_contract.json
```

输出JSON结构：
```json
{
  "contract_no": "PO2603230466",
  "date": "2026-03-23",
  "supplier": {"name": "义乌市XXX有限公司", "city": "义乌"},
  "buyer": {"name": "深圳市艾进贸易有限公司"},
  "items": [
    {
      "name_cn": "仿真草坪",
      "spec": "绿色1米",
      "fba_sku": "PTR220001-P",
      "unit": "件",
      "quantity": 250,
      "packing_rate": 25,
      "unit_price_with_tax": 8.8,
      "package_size_cm": [43, 45, 54],
      "net_weight_kg": 15,
      "gross_weight_kg": 20,
      "total_amount": 2200
    }
  ],
  "grand_total": 22228.8
}
```

#### 2b. 解析FBA货件PDF

```bash
python {baseDir}/scripts/parse_fba_pdf.py <PDF文件夹路径> --output /tmp/fba_shipments.json
```

输出JSON结构：
```json
{
  "shipments": [
    {
      "file": "FBA199DN1ZLP-MDW2.pdf",
      "warehouse_code": "MDW2",
      "address": "250 EMERALD DR, Joliet, IL 60433-3280",
      "total_boxes": 14,
      "sku_breakdown": [
        {"sku": "PTR220001-P", "boxes": 5, "qty_per_box": 25, "total_qty": 125},
        {"sku": "PTR220002-P", "boxes": 3, "qty_per_box": 16, "total_qty": 48}
      ]
    }
  ],
  "matrix": {
    "PTR220001-P": {"MDW2": 125, "AVP1": 50, "PSP3": 25, "SCK4": 25, "RDU2": 25},
    "PTR220002-P": {"MDW2": 48, "AVP1": 80, "PSP3": 64, "SCK4": 96, "RDU2": 112}
  }
}
```

### Step 3: 呈现分仓矩阵，用户决策

将 `matrix` 以可读的表格形式呈现给用户，例如：

```
SKU              MDW2   AVP1   PSP3   SCK4   RDU2   合计
PTR220001-P       125     50     25     25     25    250
PTR220002-P        48     80     64     96    112    400
...
合计              XXX    XXX    XXX    XXX    XXX    910
```

然后询问用户：
1. **分票指令**：哪些仓点合并为一票？（如果用户还没提供）
2. **哪些票要报关**：用户勾选（可多选）
3. **报关汇率**和**海运费单价**（如果还没提供）

### Step 4: 生成报关资料

收集到所有参数后，调用生成脚本。该脚本一次性生成全部四份文件：

```bash
python {baseDir}/scripts/generate_all.py \
  --contract /tmp/purchase_contract.json \
  --shipments /tmp/fba_shipments.json \
  --knowledge-base /tmp/knowledge_base.json \
  --groups '<分票JSON>' \
  --selected-groups '0,1' \
  --exchange-rate 5.5 \
  --shipping-rate 6 \
  --output-dir <输出目录> \
  --template-dir {baseDir}/assets/templates
```

其中 `--groups` 是分票指令的JSON格式：
```json
[
  {"name": "美西", "warehouses": ["MDW2", "AVP1", "PSP3"]},
  {"name": "美东", "warehouses": ["SCK4"]},
  {"name": "美中", "warehouses": ["RDU2"]}
]
```

`--selected-groups` 是用户选择要报关的票的索引（从0开始）。

脚本会在输出目录下生成：
- `【{合同号}】出口合同.xlsx`
- `【{合同号}】IV&PL.xlsx`（含Invoice和Packing List两个sheet）
- `出口报关单草稿.xlsx`

每选中一票生成一套（多票时文件名含票名）。

### Step 5: 输出结果给用户

将生成的文件展示给用户，并附上关键数据摘要：
- 报关数量（件数）
- 申报总金额（美元）
- 换汇验证结果（应等于报关汇率）
- 总箱数、总毛重、总净重

### Step 6: 商检校验（后期可选）

当用户后续拿到商检单后，调用校验脚本：

```bash
python {baseDir}/scripts/validate_inspection.py \
  --declaration /tmp/purchase_contract.json \
  --inspection <商检单路径> \
  --output /tmp/validation_result.json
```

返回三种结果之一：
- `PASS`：商检和报关单据数据相符
- `ITEM_COUNT_MISMATCH`：项数不一致，需人工调整
- `VALUE_MISMATCH`：数值有误，需人工核对

---

## 核心计算公式参考

这些公式已固化在生成脚本中，此处列出供理解和排查问题。

### 箱数与重量
```
箱数 = 数量 ÷ 箱率
总净重 = 外箱净重 × 箱数
总毛重 = 外箱毛重 × 箱数
方数(CBM) = 长cm × 宽cm × 高cm ÷ 1,000,000 × 箱数
```

### 计费重（整个采购合同的全部SKU）
```
总毛重 = Σ(每SKU的 外箱毛重 × 该SKU总箱数)
总体积重 = Σ(每SKU的 长×宽×高÷6000 × 该SKU总箱数)
计费重 = max(总毛重, 总体积重)
```

### 海运费

海运费始终基于**全部合同**的计费重计算，不管选了几票、选了哪些仓库。分摊分两层：

**第一层：全局分摊到每个SKU**
```
总海运费(RMB) = 计费重 × 海运费单价
每SKU全局运费 = 总海运费 × (该SKU计费重贡献 ÷ 总计费重)
```
平摊依据取决于哪个更大——如果体积重>毛重，按体积重占比平摊；反之按毛重。

**第二层：按票分摊（多票时）**
```
该票该SKU运费 = 该SKU全局运费 × (该票该SKU数量 ÷ 所有选中票该SKU数量之和)
```
- 如果只选了一票，分母等于分子，比例=1，即**全部海运费归该票**
- 如果选了多票，海运费按各票数量占比分摊

### CNF单价（Invoice的核心）
```
不含税金额 = 该票该SKU采购金额 ÷ 1.13
CNF总价(RMB) = 不含税金额 + 该票该SKU分摊海运费
CNF单价(RMB) = CNF总价 ÷ 该票该SKU报关数量
CNF单价(USD) = CNF单价(RMB) ÷ 报关汇率
```

### 金额分摊（采购金额）
```
该票该SKU金额 = SKU采购总金额 × (该票该SKU数量 ÷ 所有选中票该SKU数量之和)
```
与海运费分摊逻辑一致：如果只选了一票，**全部采购金额摊入该票**。

### 换汇验证
```
(不含税总额 + 该票总海运费) ÷ Invoice美元总金额 ≈ 报关汇率
```

---

## 输出文件结构详情

### 出口合同

基于采购合同改造，主要变化：
- 数量：变为该票报关涉及仓点的SKU数量汇总
- 含税单价：= 原SKU采购金额 ÷ 该票报关数量（总金额不变，单价升高）
- 右侧附计算器区域：总毛重、总体积重、计费重、总海运费、每SKU的运费平摊和C&F价格

### Invoice

| 字段 | 来源 |
|------|------|
| Tariff Code | SKU知识库 |
| Description | SKU知识库（英文品名） |
| Qty | 该票该SKU报关数量 |
| Unit | 知识库/采购合同 |
| Unit Price | CNF单价(USD) |
| USD | Unit Price × Qty |
| material quality | SKU知识库 |

### Packing List

| 字段 | 来源 |
|------|------|
| Tariff Code | SKU知识库 |
| Description | SKU知识库 |
| Qty | 该票该SKU报关数量 |
| Box Qty | Qty ÷ 箱率 |
| N.W. | 外箱净重 × Box Qty |
| G.W. | 外箱毛重 × Box Qty |
| VOLUME | 长×宽×高÷1000000 × Box Qty |

### 报关单草稿

**核心原则：千克永远会出现在报关单上。**

每个SKU占三行，单位信息从知识库获取（按SKU维护法定第一单位、第二单位），第二单位若知识库未提供则取采购合同中的单位。

**数量填写规则（通用）：**
- 若该行单位是"千克" → 数量填净重（确定值，无需校验）
- 若该行单位不是"千克" → 校验采购合同中的单位是否与该行单位一致：
  - 一致 → 采用采购合同中的数量
  - 不一致 → 数量留空，并告警提醒用户手动填写（"报关单位"X"与采购合同单位"Y"不一致"）

**第一行 — 法定第一单位：**
- 单位：知识库中该SKU海关编码对应的法定第一计量单位（每类产品不同，需查海关编码确定）
- 数量：按上述规则填写
- 其他字段：品名、申报要素、海关编码、CNF单价(USD)、总价、原产国(中国)、目的国(美国)、货源地(供方城市)

**第二行 — 第二单位：**
- 单位来源优先级：知识库中的第二单位 > 采购合同中的单位
- 情况A（海关编码只有一个法定单位"千克"）：千克已作为第一单位出现在第一行，第二单位由企业自定，通常取采购合同中的单位（如"个"）
- 情况B（海关编码有两个法定单位，如"台"和"千克"）：第一单位"台"已在第一行，千克自动成为第二单位
- 数量：按上述规则填写
- 该SKU总价(USD)也在此行

**第三行 — 与第一行一致：**
- 单位和数量与第一行相同（同一SKU数据重复体现）
- 附加：币制(美元)

头部固定信息：境内发货人、境外收货人、合同号、件数(总箱数)、毛重、净重、成交方式(C&F)、运费(USD)等。

---

## 注意事项

- 采购合同可能是 `.xls` 旧格式，脚本已兼容处理
- FBA货件PDF的文本编码可能有问题，脚本用 `pdftotext -layout` 提取并做了容错处理
- 生成的输出文件统一为 `.xlsx` 格式
- 如果知识库不可用，用采购合同中的信息兜底，但提醒用户补充海关编码、申报要素、法定第一单位和第二单位
- 成交方式当前固定为 C&F (CNF)，脚本预留了 FOB 参数（`--price-term FOB` 时不平摊海运费）
