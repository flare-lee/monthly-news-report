

ai_input_file = f"ai_input_{month}.txt"

with open(ai_input_file, "w", encoding="utf-8") as f:
    f.write(prompt)

print("✅ AI input file generated:", ai_input_file)


import os
import csv

from datetime import datetime
from docx import Document
import smtplib
from email.message import EmailMessage

# ========= 基本設定 =========
month = datetime.utcnow().strftime("%Y_%m")
csv_file = f"news_{month}.csv"
report_file = f"report_{month}.txt"
word_file = f"report_{month}.docx"

# ========= 讀取 CSV =========
rows = []
with open(csv_file, newline="", encoding="utf-8") as f:
    reader = csv.DictReader(f)
    for row in reader:
        rows.append(row)

# ========= 整理給 Gemini 的新聞文字 =========
news_text = ""
for r in rows:
    news_text += f"- {r['company']}: {r['title']}\n"

# ========= Gemini 設定 =========





prompt = f"""
你是一位資料中心與雲端產業分析師，
請以提供給管理層閱讀的專業語氣，
根據以下新聞標題撰寫一份產業分析月報，
請使用繁體中文，並務必包含三個段落：

【市場趨勢】
- Oracle 在 AI 與雲端基礎建設的布局
- Wiwynn 與 AI Server / 資料中心需求

【技術更新】
- AI Infrastructure
- Data Center
- Rack-level Server
- hyperscaler 對高密度運算需求

【財務與策略訊號】
- Oracle 資本支出與雲端投資動向
- ODM 競爭（Wiwynn、Quanta、Wistron、Inventec 等）

以下是本月新聞標題清單：
{news_text}
"""



# ========= 寫文字報告 =========
with open(report_file, "w", encoding="utf-8") as f:
    f.write(report)

# ========= 產出 Word =========
doc = Document()
doc.add_heading("Oracle & Wiwynn 產業分析月報", level=1)

for line in report.split("\n"):
    if line.strip().startswith("【"):
        doc.add_heading(line.strip(), level=2)
    else:
        doc.add_paragraph(line)

doc.save(word_file)

# ========= 寄 Email =========
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

print("✅ Monthly report completed")
