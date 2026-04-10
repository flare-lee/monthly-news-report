import os
import csv
from datetime import datetime
from docx import Document
import smtplib
from email.message import EmailMessage

# =========================
# 基本設定（一定要最前面）
# =========================
month = datetime.utcnow().strftime("%Y_%m")

csv_file = f"news_{month}.csv"
ai_input_file = f"ai_input_{month}.txt"
report_file = f"report_{month}.txt"
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
# 整理新聞成 AI 用文字
# =========================
news_text = ""
for r in rows:
    news_text += f"- {r['company']}: {r['title']}\n"

prompt = f"""
你是一位資料中心與雲端產業分析師，
請你以提供給管理層閱讀的專業語氣，
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

with open(ai_input_file, "w", encoding="utf-8") as f:
    f.write(prompt)

print("✅ AI input file generated:", ai_input_file)

# =========================
# 產生佔位分析內容（避免未定義錯誤）
# =========================
report = f"""
【本月產業分析尚未自動生成】

請開啟檔案：
{ai_input_file}

將檔案內容完整複製，貼至 Gemini 或 ChatGPT 網頁版，
取得 AI 產生的三段分析後，
再將內容貼回本 Word 報告中作為正式版本。
"""

# =========================
# 寫文字報告（txt）
# =========================
with open(report_file, "w", encoding="utf-8") as f:
    f.write(report)

# =========================
# 產出 Word 報告
# =========================
doc = Document()
doc.add_heading("Oracle & Wiwynn 產業分析月報", level=1)

for line in report.split("\n"):
    if line.strip().startswith("【"):
        doc.add_heading(line.strip(), level=2)
    else:
        doc.add_paragraph(line)

doc.save(word_file)
print("✅ Word report generated:", word_file)

# =========================
# 寄送 Email
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
        "附件為本月 Oracle & Wiwynn 產業分析報告。\n\n"
        "請參考 ai_input 檔案，將其內容貼入 AI 工具完成分析後更新 Word。"
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
else:
    print("⚠️ Email not sent (missing email credentials)")

print("✅ Monthly report flow completed")
