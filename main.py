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
你是一位【資料中心與雲端產業分析師】，
請以【提供企業管理層與策略決策使用】的專業語氣，
根據以下新聞資料，撰寫一份【產業分析月報】。

請嚴格依照下列結構輸出，
不要條列新聞、不要解釋分析方法，
只輸出可直接放進正式報告的「完整段落文字」，
語言請使用【繁體中文】。

＝＝＝＝ 指定輸出結構（不可更改）＝＝＝＝

【市場趨勢】
【技術更新】
【財務與策略訊號】

＝＝＝＝ 本月新聞資料（已整理）＝＝＝＝

【市場趨勢相關新聞】
{chr(10).join(market_news)}

【技術更新相關新聞】
{chr(10).join(tech_news)}

【財務與策略訊號相關新聞】
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

