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
# Gemini AI 生成報告內容 (整合你的最新 Prompt)
# =========================
def call_gemini_api(market, tech, finance):
    if not gemini_key:
        return "錯誤：未設定 GEMINI_API_KEY"

    try:
        genai.configure(api_key=gemini_key)
        # 使用最新穩定版名稱避免 404
        model = genai.GenerativeModel('gemini-1.5-flash-latest')

        # 這裡放入你改過後的完整 Prompt
        ai_prompt = f"""
你是一位【華爾街資深產業分析師 / 顧問公司資深策略顧問】。

請根據以下已整理的新聞資料、公司資訊與產業背景，撰寫一份【Oracle & Wiwynn 產業分析月報】。

請嚴格遵守以下規範：
- 請直接輸出一份完整的專業報告內容
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

＝＝＝＝ 本月已整理新聞資料 ＝＝＝＝

【市場趨勢相關新聞】
{chr(10).join(market)}

【技術更新相關新聞】
{chr(10).join(tech)}

【財務與策略相關新聞】
{chr(10).join(finance)}
"""
        
        response = model.generate_content(ai_prompt)
        return response.text if response.text else "AI 生成內容為空。"

    except Exception as e:
        return f"AI 生成失敗，錯誤訊息：{str(e)}"

print("🚀 正在呼叫 Gemini API 生成專業報告...")
final_report_content = call_gemini_api(market_news, tech_news, finance_news)

# =========================
# 產出 Word（最終報告）
# =========================
doc = Document()
# 移除原本叫你去貼給 AI 的段落，改為直接呈現分析內容
doc.add_heading(f"Oracle & Wiwynn 產業分析月報 - {month}", level=0)
doc.add_paragraph(final_report_content)

doc.save(word_file)
print(f"✅ Word report generated: {word_file}")

# =========================
# 寄送 Email
# =========================
if email_user and email_pass and email_to:
    msg = EmailMessage()
    msg["Subject"] = f"Oracle & Wiwynn 產業分析月報 - {month}"
    msg["From"] = email_user
    msg["To"] = email_to
    msg.set_content(f"Flare 你好，附件為本月由 AI 自動生成的產業分析月報。")

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
