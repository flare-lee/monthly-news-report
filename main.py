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
ai_input_file = f"ai_input_{month}.txt"
report_file = f"report_{month}.txt"
word_file = f"report_{month}.docx"

# =========================
# 讀取新聞 CSV（原始資料）
# =========================
rows = []
with open(csv_file, newline="", encoding="utf-8") as f:
    reader = csv.DictReader(f)
    for row in reader:
        rows.append(row)

# =========================
# 整理新聞成固定三類
# =========================
market_news = []
tech_news = []
finance_news = []

for r in rows:
    title = r["title"]

    # 市場趨勢
    if any(k in title for k in ["cloud", "AI", "market", "demand", "growth"]):
        market_news.append(f"- {r['company']}：{title}")

    # 技術更新
    if any(k in title.lower() for k in ["data center", "server", "rack", "infrastructure"]):
        tech_news.append(f"- {r['company']}：{title}")

    # 財務與策略
    if any(k in title.lower() for k in ["capex", "investment", "revenue", "order"]):
        finance_news.append(f"- {r['company']}：{title}")

# =========================
# 組「唯一要送 AI 的檔案」
# =========================
ai_input = f"""
你是一位【資料中心與雲端產業分析師】，
請以【提供企業管理層與策略決策使用】的專業語氣，
根據以下新聞資料，撰寫一份【產業分析月報】。

請嚴格依照指定結構輸出，
不要條列新聞、不要解釋分析方法，
只輸出【可直接放進正式報告】的完整段落文字，
語言請使用【繁體中文】。

＝＝＝＝ 指定輸出結構（不可更改）＝＝＝＝

【市場趨勢】
請說明 Oracle 在 AI 與雲端基礎建設的策略方向，
以及 Wiwynn 與資料中心需求所代表的市場趨勢。

【技術更新】
請聚焦 AI Infrastructure、Data Center、
Rack-level Server、高密度運算與系統架構演進。

【財務與策略訊號】
請分析 Oracle 的資本支出與雲端投資動向，
以及 ODM 競爭格局（Wiwynn、Quanta、Wistron、Inventec 等）。

＝＝＝＝ 本月新聞資料（已整理）＝＝＝＝

【市場趨勢相關新聞】
{chr(10).join(market_news)}

【技術更新相關新聞】
{chr(10).join(tech_news)}

【財務與策略訊號相關新聞】
{chr(10).join(finance_news)}
"""

with open(ai_input_file, "w", encoding="utf-8") as f:
    f.write(ai_input)

print("✅ ai_input file generated:", ai_input_file)

# =========================
# 產生 Word（交付用）
# =========================
doc = Document()
doc.add_heading("Oracle & Wiwynn 產業分析月報", level=1)

doc.add_heading("【市場趨勢】", level=2)
doc.add_paragraph("（請貼上 AI 產生的專業分析內容）")

doc.add_heading("【技術更新】", level=2)
doc.add_paragraph("（請貼上 AI 產生的專業分析內容）")

doc.add_heading("【財務與策略訊號】", level=2)
doc.add_paragraph("（請貼上 AI 產生的專業分析內容）")

doc.save(word_file)
print("✅ Word report generated:", word_file)

# =========================
# 寄 Email
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
        "附件包含本月產業分析報告與 AI 輸入檔。\n\n"
        "請開啟 ai_input 檔案，完整複製貼至 AI，即可產生報告正文。"
    )

    for file in [word_file, ai_input_file]:
        with open(file, "rb") as f:
            msg.add_attachment(
                f.read(),
                maintype="application",
                subtype="octet-stream",
                filename=file
            )

    with smtplib.SMTP_SSL("smtp.gmail.com", 465) as smtp:
        smtp.login(email_user, email_pass)
        smtp.send_message(msg)

    print("✅ Email sent successfully")

print("✅ Monthly workflow completed")
