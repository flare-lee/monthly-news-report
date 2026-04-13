import os
import csv
from datetime import datetime
from docx import Document
import smtplib
from email.message import EmailMessage
import google.generativeai as genai

# =========================
# 基本設定與環境變數
# =========================
month = datetime.now().strftime("%Y_%m")
csv_file = f"news_{month}.csv"
word_file = f"report_{month}.docx"

# 從 GitHub Secrets 或系統環境變數讀取
gemini_key = os.getenv("GEMINI_API_KEY")
email_user = os.getenv("EMAIL_USER")
email_pass = os.getenv("EMAIL_PASS")
email_to = os.getenv("EMAIL_TO")

# =========================
# 讀取新聞 CSV
# =========================
market_news = []
tech_news = []
finance_news = []

try:
    with open(csv_file, newline="", encoding="utf-8") as f:
        reader = csv.DictReader(f)
        for r in reader:
            title = r["title"].lower()
            company = r["company"]
            content = f"- {company}：{r['title']}"

            if any(k in title for k in ["cloud", "ai", "market", "demand", "growth"]):
                market_news.append(content)
            if any(k in title for k in ["data center", "server", "rack", "infrastructure"]):
                tech_news.append(content)
            if any(k in title for k in ["capex", "investment", "revenue", "order"]):
                finance_news.append(content)
except FileNotFoundError:
    print(f"❌ 找不到檔案: {csv_file}")
    exit(1)

# =========================
# Gemini AI 生成報告內容
# =========================
def call_gemini_api(market, tech, finance):
    if not gemini_key:
        return "錯誤：未設定 GEMINI_API_KEY"

    genai.configure(api_key=gemini_key)
    model = genai.GenerativeModel('gemini-1.5-flash')

    ai_prompt = f"""
    你是一位【華爾街資深產業分析師 / 顧問公司資深策略顧問】。
    請根據以下新聞，撰寫一份專業的【Oracle & Wiwynn 產業分析月報】。

    規範：
    - 使用繁體中文，語氣專業冷靜 (McKinsey/BCG 風格)。
    - 以「觀點 + 分析 + 商業含義」撰寫，不可只是摘要重述。
    - 每一段皆需為可直接放入 Word 的完整段落。

    報告結構：
    【一、產業與市場趨勢】分析 Oracle 與 緯穎 在 AI 雲端策略地位與需求變化。
    【二、技術演進與基礎設施設計重點】聚焦 AI Infrastructure、Rack-level Server、散熱與系統整合。
    【三、財務與策略觀察】分析資本支出、營收結構，並點評 Quanta, Foxconn, Dell, Supermicro 等競爭格局。
    【四、結論與策略建議】總結訊號，給予產品團隊與投資人的實質建議。

    資料來源：
    市場：{chr(10).join(market)}
    技術：{chr(10).join(tech)}
    財務：{chr(10).join(finance)}
    """
    
    try:
        response = model.generate_content(ai_prompt)
        return response.text
    except Exception as e:
        return f"AI 生成失敗: {str(e)}"

print("🚀 正在呼叫 Gemini API 生成專業報告...")
final_report_content = call_gemini_api(market_news, tech_news, finance_news)

# =========================
# 產出 Word（最終報告）
# =========================
doc = Document()
doc.add_heading(f"Oracle & Wiwynn 產業分析月報 - {month}", level=0)

# 將 AI 生成的內容寫入 Word
doc.add_paragraph(final_report_content)

# 附錄：原始資料
doc.add_page_break()
doc.add_heading("附錄：本月原始新聞參考", level=2)
for news in market_news + tech_news + finance_news:
    doc.add_paragraph(news, style='List Bullet')

doc.save(word_file)
print(f"✅ Word report generated: {word_file}")

# =========================
# 寄送 Email
# =========================
if email_user and email_pass and email_to:
    msg = EmailMessage()
    msg["Subject"] = f"【自動發送】Oracle & Wiwynn 產業分析報告 - {month}"
    msg["From"] = email_user
    msg["To"] = email_to
    msg.set_content(f"Flare 你好，附件為本月自動生成的 {month} 產業分析月報，請查收。")

    with open(word_file, "rb") as f:
        msg.add_attachment(
            f.read(),
            maintype="application",
            subtype="vnd.openxmlformats-officedocument.wordprocessingml.document",
            filename=word_file
        )

    try:
        with smtplib.SMTP_SSL("smtp.gmail.com", 465) as smtp:
            smtp.login(email_user, email_pass)
            smtp.send_message(msg)
        print("✅ Email sent successfully")
    except Exception as e:
        print(f"❌ Email 寄送失敗: {e}")
else:
    print("⚠️ 未偵測到完整 Email 設定，跳過寄送步驟。")
