
import os
import openai

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

rows = []

with open(csv_file, newline="", encoding="utf-8") as f:
    reader = csv.DictReader(f)
    for row in reader:
        rows.append(row)          # ✅ 關鍵就在這一行

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

news_text = ""
for r in rows:
    news_text += f"- {r['company']}: {r['title']}\n"


# 3. 組出分析文字（這就是你的報告正文）


prompt = f"""
你是一位資料中心與雲端產業分析師，
請根據以下新聞標題，撰寫一份產業分析月報，
必須包含以下三個段落，請用繁體中文撰寫：

【市場趨勢】
- Oracle AI 與雲端基礎建設布局
- Wiwynn 與 AI Server、市場需求

【技術更新】
- AI Infrastructure
- Data Center
- Rack-level Server
- hyperscaler 對高密度運算需求

【財務與策略訊號】
- Oracle 的資本支出與雲端投資方向
- ODM 競爭（Wiwynn、Quanta、Wistron、Inventec 等）

以下是本月新聞標題清單：
{news_text}
"""

response = openai.ChatCompletion.create(
    model="gpt-4o-mini",
    messages=[{"role": "user", "content": prompt}],
    temperature=0.3
)

report = response["choices"][0]["message"]["content"]


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

openai.api_key = os.getenv("OPENAI_API_KEY")
