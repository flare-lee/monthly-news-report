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
# 組 AI 分析輸入內容
# =========================
news_text = ""
for r in rows:
    news_text += f"- {r['company']}: {r['title']}\n"

ai_prompt = f"""
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

# =========================
# 產出 Word 報告（使用者友善）
# =========================
doc = Document()
doc.add_heading("Oracle & Wiwynn 產業分析月報", level=1)

doc.add_paragraph(
    "以下內容請【直接全選】並貼至 Gemini 或 ChatGPT 進行分析："
)

doc.add_paragraph("─── AI 分析輸入區（請勿修改）───")
for line in ai_prompt.strip().split("\n"):
    doc.add_paragraph(line)
doc.add_paragraph("─── AI 分析輸入區 結束 ───")

doc.add_paragraph("")
doc.add_heading("【市場趨勢】", level=2)
doc.add_paragraph("（請貼上 AI 產生內容）")

doc.add_heading("【技術更新】", level=2)
doc.add_paragraph("（請貼上 AI 產生內容）")

doc.add_heading("【財務與策略訊號】", level=2)
doc.add_paragraph("（請貼上 AI 產生內容）")

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
        "附件為本月 Oracle & Wiwynn 產業分析報告。\n"
        "請於 Word 中直接全選 AI 輸入內容貼至 Gemini / ChatGPT。"
    )

    with open(word_file, "rb") as f:
        msg.add_attachment(
            f.read(),
            maintype="application",
            subtype="vnd.openxmlformats-officedocument.wordprocessingml.document",
            filename=word_file
        )

