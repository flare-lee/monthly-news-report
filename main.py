import os
import csv
from datetime import datetime
from docx import Document
import smtplib
from email.message import EmailMessage
import google.generativeai as genai

# =========================
# 1. 基本設定與環境變數
# =========================
month = datetime.now().strftime("%Y_%m")
csv_file = f"news_{month}.csv"
word_file = f"report_{month}.docx"

gemini_key = os.getenv("GEMINI_API_KEY")
email_user = os.getenv("EMAIL_USER")
email_pass = os.getenv("EMAIL_PASS")
email_to = os.getenv("EMAIL_TO")

# =========================
# 2. 讀取新聞 CSV 資料
# =========================
market_news, tech_news, finance_news = [], [], []

if not os.path.exists(csv_file):
    print(f"❌ 嚴重錯誤：找不到 {csv_file}")
    exit(1)

with open(csv_file, newline="", encoding="utf-8") as f:
    reader = csv.DictReader(f)
    for r in reader:
        title = r.get("title", "").lower()
        company = r.get("company", "未知公司")
        content = f"- {company}：{r.get('title', '')}"
        if any(k in title for k in ["cloud", "ai", "market", "demand", "growth"]):
            market_news.append(content)
        if any(k in title for k in ["data center", "server", "rack", "infrastructure"]):
            tech_news.append(content)
        if any(k in title for k in ["capex", "investment", "revenue", "order"]):
            finance_news.append(content)

# =========================
# 3. Gemini API 呼叫 (最強防錯版)
# =========================
def call_gemini_api(market, tech, finance):
    if not gemini_key:
        return "錯誤：未設定 GEMINI_API_KEY"

    # 完整保留你指定的專業 Prompt
    ai_prompt = f"""
你是一位【華爾街資深產業分析師 / 顧問公司資深策略顧問】。

請根據以下已整理的新聞資料、公司資訊與產業背景，撰寫一份【Oracle & Wiwynn 產業分析月報】。

請嚴格遵守以下規範：
- 請直接輸出一份word的完整報告
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
{chr(10).join(market)}

【技術更新相關新聞】
{chr(10).join(tech)}

【財務與策略相關新聞】
{chr(10).join(finance)}
"""

    try:
        genai.configure(api_key=gemini_key)
        
        # 嘗試清單：手動列出所有可能成功的模型路徑
        possible_models = ['models/gemini-1.5-flash', 'gemini-1.5-flash', 'models/gemini-pro']
        
        last_error = ""
        for m_name in possible_models:
            try:
                print(f"🔄 嘗試使用模型: {m_name}...")
                model = genai.GenerativeModel(model_name=m_name)
                response = model.generate_content(ai_prompt)
                if response.text:
                    print(f"✅ 成功使用 {m_name} 生成報告")
                    return response.text
            except Exception as e:
                last_error = str(e)
                continue
        
        # 如果走到這代表全掛了，我們印出所有可用的模型清單來診斷
        print("❌ 所有已知模型名稱皆失效。可用模型清單如下：")
        for m in genai.list_models():
            print(f"- {m.name} (支援方法: {m.supported_generation_methods})")
            
        return f"AI 生成最終失敗，最後一個錯誤：{last_error}"

    except Exception as e:
        return f"API 配置發生嚴重錯誤：{str(e)}"

# =========================
# 4. 執行與寄送
# =========================
print(f"🚀 開始分析 {month} 資料...")
report_text = call_gemini_api(market_news, tech_news, finance_news)

doc = Document()
doc.add_heading(f"Oracle & Wiwynn 產業分析月報 - {month}", 0)
doc.add_paragraph(report_text)
doc.save(word_file)

if email_user and email_pass:
    msg = EmailMessage()
    msg["Subject"] = f"Oracle & Wiwynn 產業分析月報 - {month}"
    msg["From"] = email_user
    msg["To"] = email_to
    msg.set_content(f"Flare 你好，本月分析報告已自動生成。")
    with open(word_file, "rb") as f:
        msg.add_attachment(f.read(), maintype="application", subtype="docx", filename=word_file)
    try:
        with smtplib.SMTP_SSL("smtp.gmail.com", 465) as smtp:
            smtp.login(email_user, email_pass)
            smtp.send_message(msg)
        print("✅ 郵件寄送成功！")
    except Exception as e:
        print(f"❌ 郵件寄送失敗: {e}")
