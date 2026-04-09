import feedparser
import csv
from datetime import datetime

# 全網新聞來源（Google News 搜尋）
rss_urls = {
    "Oracle": "https://news.google.com/rss/search?q=Oracle",
    "Wiwynn": "https://news.google.com/rss/search?q=Wiwynn+緯穎"
}

# 檔名：依月份產生
filename = f"news_{datetime.utcnow().strftime('%Y_%m')}.csv"

with open(filename, "w", newline="", encoding="utf-8") as f:
    writer = csv.writer(f)
    writer.writerow(["company", "title", "link"])

    for company, url in rss_urls.items():
        feed = feedparser.parse(url)
        for entry in feed.entries:
            writer.writerow([company, entry.title, entry.link])

print("✅ News saved to", filename)
