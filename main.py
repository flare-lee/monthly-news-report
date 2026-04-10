import os
import csv
from datetime import datetime
from docx import Document
import smtplib
from email.message import EmailMessage

# =========================
# 基本設定
# =========================
month = datetime.utcnow().strftime("%Y_%m")
csv_file = f"news_{month}.csv"
word_file = f"report_{month}.docx"

# =========================
# 讀取新聞 CSV
# =========================
rows = []
with open(csv_file, newline="", encoding="utf-8") as f:
    reader = csv.DictReader(f)
    for row in reader:
        rows.append(row)

# =========================
# 分類新聞（市場 / 技術 / 財務）
# =========================
market_news = []
tech_news = []
finance_news = []

for r in rows:
    title = r["title"].lower()
    company = r["company"]

    if any(k in title for k in ["cloud", "ai", "market", "demand", "growth"]):
        market_news.append(f"- {company}：{r['title']}")

    if any(k in title for k in ["data center", "server", "rack", "infrastructure"]):
        tech_news.append(f"- {company}：{r['title']}")

    if any(k in title for k in ["capex", "investment", "revenue", "order"]):
        finance_news.append(f"- {company}：{r['title']}")


# =========================
# 組成「AI 專用指令＋新聞資料」（直接寫進 Word）
# =========================
ai_prompt = f"""
你是一位【華爾街資深產業分析師 / 顧問公司資深策略顧問】。

請根據以下已整理的新聞資料、公司資訊與產業背景，撰寫一份【Oracle & Wiwynn 產業分析月報】。

請嚴格遵守以下規範：
- 請直接輸出一份完整報告
- 需參考我提供的新聞與資料，將資訊整理為重點並進一步分析，不可只是摘要重述
- 請勿提及分析步驟、推理過程或你如何整理資料
- 請使用繁體中文
- 語氣需專業、冷靜、策略導向，風格需接近 McKinsey / BCG / Deloitte / sell-side 產業研究報告
- 每一段皆需為可直接貼入 Word 的完整段落
- 避免口語化表達，內容需具備管理層簡報與投資研究可讀性
- 請以「觀點 + 分析 + 商業含義」方式撰寫，而非單純新聞整理
- 報告最後請附上新聞或資料出處

＝＝＝＝ 指定報告結構（請嚴格遵守）＝＝＝＝

【一、產業與市場趨勢】
請從 Oracle 與 Wiwynn 在 AI、雲端、資料中心與企業 IT 領域的策略定位切入，
分析市場需求變化、客戶採購趨勢、產業週期與商業模式演進，
並連結供應鏈與競爭對手的發展方向，說明其產業意涵。

【二、技術演進與基礎設施設計重點】
請聚焦 AI Infrastructure、Data Center、Rack-level Server、高密度運算、
散熱與系統整合等關鍵技術主題，
分析技術演進如何影響資料中心架構、資本支出、產品規格與未來設計方向。

【三、財務與策略觀察（含競爭格局）】
請分析 Oracle 與 Wiwynn 的資本支出、營收結構、訂單動能與策略意圖，
並說明 Quanta、Wistron、Inventec、Dell、HPE、Supermicro 等競爭對手
在該市場中的角色變化、競爭格局，以及中長期風險與機會。

【四、結論與策略建議】
請總結本月最重要的產業訊號，
提出對管理層、業務團隊、產品團隊與投資人具備實質意義的策略建議，
內容需涵蓋優先順序、資源配置、競爭應對與未來觀察重點。

請清楚區分短期市場變化與中長期結構性趨勢，
並說明這些變化對 Oracle、ODM/OEM 與終端客戶各自的商業含義。

請不要只是描述新聞，
而要將新聞轉化為產業結論、競爭含義與策略推演，
產出可直接提交給主管、客戶或管理層閱讀的正式分析報告。

＝＝＝＝ 本月已整理新聞資料 ＝＝＝＝

【市場趨勢相關新聞】
{chr(10).join(market_news)}

【技術更新相關新聞】
{chr(10).join(tech_news)}

【財務與策略相關新聞】
{chr(10).join(finance_news)}
"""


# =========================
# 產出 Word（唯一交付檔案）
# =========================
doc = Document()
doc.add_heading("Oracle & Wiwynn 產業分析月報", level=1)

doc.add_paragraph(
    "【以下整段內容請直接全選，貼至 Gemini / ChatGPT，"
    "AI 會直接輸出完整專業報告】"
)

doc.add_paragraph("──────── AI 輸入內容（請勿修改）────────")
for line in ai_prompt.strip().split("\n"):
    doc.add_paragraph(line)
doc.add_paragraph("──────── AI 輸入內容結束 ────────")

doc.add_page_break()

doc.add_heading("【市場趨勢】", level=2)
doc.add_paragraph("（請貼上 AI 產生的完整段落內容）")

doc.add_heading("【技術更新】", level=2)
doc.add_paragraph("（請貼上 AI 產生的完整段落內容）")

doc.add_heading("【財務與策略訊號】", level=2)
doc.add_paragraph("（請貼上 AI 產生的完整段落內容）")

doc.save(word_file)
print("✅ Word report generated:", word_file)

# =========================
# 寄送 Email（只寄 Word）
# =========================
email_user = os.getenv("EMAIL_USER")
email_pass = os.getenv("EMAIL_PASS")
email_to = os.getenv("EMAIL_TO")

if email_user and email_pass and email_to:
    msg = EmailMessage()
    msg["Subject"] = f"Oracle & Wiwynn 產業分析月報 - {month}"
    msg["From"] = email_user
    msg["To"] = email_to
    msg.set_content(
        "附件為本月 Oracle & Wiwynn 產業分析月報。\n\n"
        "請在 Word 中『AI 輸入內容』區塊全選後貼到 AI，即可產出完整報告。"
    )

    with open(word_file, "rb") as f:
        msg.add_attachment(
            f.read(),
            maintype="application",
            subtype="vnd.openxmlformats-officedocument.wordprocessingml.document",
            filename=word_file
        )

    with smtplib.SMTP_SSL("smtp.gmail.com", 465) as smtp:
        smtp.login(email_user, email_pass)
        smtp.send_message(msg)
    print("✅ Email sent successfully")

