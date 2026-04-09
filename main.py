import feedparser

rss_urls = {
    "Oracle": "https://news.google.com/rss/search?q=Oracle",
    "Wiwynn": "https://news.google.com/rss/search?q=Wiwynn+緯穎"
}

for company, url in rss_urls.items():
    print("\n=== ", company, " ===")
    feed = feedparser.parse(url)
    for entry in feed.entries[:5]:
        print("-", entry.title)
        print(" ", entry.link)

