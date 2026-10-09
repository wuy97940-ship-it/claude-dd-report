# MCP 数据源配置

尽调报告的数据底座。**没有商业数据源也能用**——见文末降级路径。

## 企查查全域 MCP（推荐）

覆盖：工商信息、股东/实控人/受益所有人穿透、35 项风险因子扫描、失信/被执行/限高、
股权冻结与出质、行政处罚、欠税、知产（专利/商标/软著）、招投标、涉诉与裁判文书、法规检索、文档解析。

### 接入

1. 在企查查开放平台开通 MCP 服务（ https://openapi.qcc.com ），获取鉴权
2. 在 Claude Code 的 MCP 配置（`.mcp.json` 或全局配置）中加入 qcc 系列 server：

```json
{
  "mcpServers": {
    "qcc-company":   { "command": "...", "args": ["..."] },
    "qcc-risk":      { "command": "...", "args": ["..."] },
    "qcc-ipr":       { "command": "...", "args": ["..."] },
    "qcc-executive": { "command": "...", "args": ["..."] },
    "qcc-operation": { "command": "...", "args": ["..."] },
    "qcc-history":   { "command": "...", "args": ["..."] },
    "qcc-legal-case": { "command": "...", "args": ["..."] },
    "qcc-legal-regulation": { "command": "...", "args": ["..."] },
    "qcc-document":  { "command": "...", "args": ["..."] }
  }
}
```

> 具体启动命令以企查查开放平台给你的接入文档为准（各 Server 支持 stdio/远程两种形态）。

### SKILL 内的数据路由（已内置）

| 尽调需求 | 首选工具 |
|---|---|
| 35项风险一次扫描 | `qcc-risk.get_company_risk_scan`（分诊）→ 按命中下钻对应原子工具 |
| 失信/被执行/限高/终本 | `qcc-risk.get_dishonest_info / judgment_debtor_info / high_consumption_restriction / terminated_cases` |
| 股东/实控人/受益所有人 | `qcc-company.get_shareholder_info / get_actual_controller / get_beneficial_owners` |
| 纳税信用/海关信用 | `qcc-operation.get_credit_evaluation`（注意：在 operation 不在 risk） |
| 董监高个人风险 | `qcc-executive.get_executive_risk_scan`（+18维原子工具） |
| 征信/报表扫描件解析 | `qcc-document.parse_document` |

### 使用纪律（SKILL 已内置）

- 实体锚定：searchKey 必须用企业**完整登记名或统一社会信用代码**，简称先过实体识别
- 拿到的数据要做**量级+勾稽校验**后才可入稿；跨源（企查查 vs 审计报表 vs 公开披露）交叉核对
- 上市公司股权分散 → `get_actual_controller` 返回"无记录"是正常，勿当"查不到"

## 财务数据接口

任选其一：同花顺 iFind、akshare（开源免费）、所属券商行情接口。
用途：可比公司财务数据与比率。纪律：**接口字段单位先验证再用**（毛利率字段有百分数与小数两种形态，
历史教训是直接引用导致口径错），必要时（营收-成本)/营收自算。

## 无商业数据源的降级路径

| 需求 | 免费渠道 |
|---|---|
| 工商登记/股东/年报 | 国家企业信用信息公示系统（www.gsxt.gov.cn） |
| 涉诉/被执行/失信 | 中国裁判文书网、中国执行信息公开网 |
| 上市/发债主体披露 | 交易所官网、巨潮资讯网 |
| 行政处罚/欠税 | 信用中国（www.creditchina.gov.cn） |
| 报表数据 | 委托方提供审计报告+征信报告（原件优先） |

SKILL 的数据纪律在降级路径下同样生效：查不到写"未获取到"+替代核查方案，绝不编造。
