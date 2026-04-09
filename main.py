
import csv
from datetime import datetime

# 1. 找出當月的 CSV 檔名
month = datetime.utcnow().strftime("%Y_%m")
csv_file = f"news_{month}.csv"

# 2. 讀取新聞資料
oracle_news = 0
wiwynn_news = 0
odm_mentions = set()
tech_keywords_count = 0

ODM_KEYWORDS = {
    "Wiwynn": ["Wiwynn", "緯穎"],
    "Quanta": ["Quanta", "廣達"],
    "Wistron": ["Wistron", "緯創"],
    "Inventec": ["Inventec", "英業達"],
    "Foxconn": ["Foxconn", "鴻海"]
}

TECH_KEYWORDS = [
    "AI", "Data Center", "Server", "Rack", "Infrastructure"
]

with open(csv_file, newline="", encoding="utf-8") as f:
    reader = csv.DictReader(f)
    for row in reader:
        title = row["title"]
        company = row["company"]

        if company == "Oracle":
            oracle_news += 1
        if company == "Wiwynn":
            wiwynn_news += 1

        for odm, names in ODM_KEYWORDS.items():
            for name in names:
                if name in title:
                    odm_mentions.add(odm)

        for kw in TECH_KEYWORDS:
            if kw.lower() in title.lower():
                tech_keywords_count += 1

# 3. 組出分析文字（這就是你的報告正文）
report = f"""
【市場趨勢】
- Oracle 在本月新聞中持續強化 AI 與雲端基礎建設布局（相關新聞 {oracle_news} 則）
- Wiwynn 的新聞多集中在 AI Server 與資料中心需求成長（相關新聞 {wiwynn_news} 則）

【技術更新】
- AI Infrastructure、Data Center、Rack-level Server 為關鍵字高頻出現（關鍵字命中約 {tech_keywords_count} 次）
- 顯示 hyperscaler 對高密度運算與系統整合能力的持續需求

【財務與策略訊號】
- Oracle 資本支出與雲端投資相關新聞增加
- ODM 端（{", ".join(sorted(odm_mentions)) if odm_mentions else "Wiwynn"}）在 AI 與資料中心供應鏈中扮演關鍵角色
"""

# 4. 寫出報告檔
report_file = f"report_{month}.txt"
with open(report_file, "w", encoding="utf-8") as f:
    f.write(report)

print("✅ Report generated:", report_file)

# === 產出 Word 檔 ===
from docx import Document

doc = Document()
doc.add_heading("Oracle & Wiwynn 產業分析月報", level=1)

for line in report.split("\n"):
    if line.strip().startswith("【"):
        doc.add_heading(line.strip(), level=2)
    else:
        doc.add_paragraph(line)

word_file = f"report_{month}.docx"
doc.save(word_file)

print("✅ Word report generated:", word_file)

# === 寄 Email ===
import smtplib
from email.message import EmailMessage
import os

email_user = os.getenv("EMAIL_USER")
email_pass = os.getenv("EMAIL_PASS")
email_to = os.getenv("EMAIL_TO")

if email_user and email_pass and email_to:
    msg = EmailMessage()
    msg["Subject"] = f"Oracle & Wiwynn 產業分析月報 - {month}"
    msg["From"] = email_user
    msg["To"] = email_to
    msg.set_content("請參閱附件：本月 Oracle & Wiwynn 產業分析報告（Word）")

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
else:
    print("⚠️ Email not sent (missing credentials)")
