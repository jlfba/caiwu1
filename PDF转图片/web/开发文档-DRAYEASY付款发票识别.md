# DRAYEASY 付款发票识别

## 目标

付款组新增 `DRAYEASY` 发票类型。识别票面 `INVOICE` 右侧的发票号，以及 `Delivery Address`、`Container`、`Description`、`Rate`、`Qty`、`Amount` 标签下方的内容。

## 输出约定

Excel 固定为十列；付款组补充填写列置于前三列，后七列按明细行重复发票级字段：

1. 配仓单号
2. 单号
3. 费用名称
4. INVOICE
5. Delivery Address
6. Container
7. Description
8. Rate
9. Qty
10. Amount

## 实现与验证

- 在共享 PDF 解析模块中增加基于文字坐标的 DRAYEASY 解析器；无文字层时沿用既有 OCR 回退。
- Delivery Address 从明细表上方的独立字段读取完整地址并重复写入明细；Container 中的 ISO 柜号后续费用描述自动拆回 Description。
- PDF 文字层漏读 Amount 表头时，依据每行结尾的 Rate、Qty、Amount 三个数值回填正确列。
- Amount 只有空格时同样按空值处理，确保首条明细也能触发回填。
- 在网页的付款组类型列表、接口校验和 Excel 导出规则中接入类型 `14`。
- 以构造的坐标化票面条目验证字段位置、续行合并及输出列顺序；再执行 Python 编译检查与前端构建。
