# DRAYEASY 付款发票识别

## 目标

付款组新增 `DRAYEASY` 发票类型。识别票面 `INVOICE` 右侧的发票号，以及 `Delivery Address`、`Container`、`Description`、`Rate`、`Qty`、`Amount` 标签下方的内容。

## 输出约定

Excel 固定为七列，且按明细行重复发票级字段：

1. INVOICE
2. Delivery Address
3. Container
4. Description
5. Rate
6. Qty
7. Amount

## 实现与验证

- 在共享 PDF 解析模块中增加基于文字坐标的 DRAYEASY 解析器；无文字层时沿用既有 OCR 回退。
- 在网页的付款组类型列表、接口校验和 Excel 导出规则中接入类型 `14`。
- 以构造的坐标化票面条目验证字段位置、续行合并及输出列顺序；再执行 Python 编译检查与前端构建。
