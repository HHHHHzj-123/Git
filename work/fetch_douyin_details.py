import asyncio
import json
from pathlib import Path

from crawlers.douyin.web.web_crawler import DouyinWebCrawler


IDS = [
    "7392785204122783028",
    "7397739121348726054",
    "7403557916621540660",
]
OUT = Path(r"C:\Users\HZJ\Desktop\Git\work\douyin_details")


async def main():
    OUT.mkdir(parents=True, exist_ok=True)
    crawler = DouyinWebCrawler()
    for aweme_id in IDS:
        result = await crawler.fetch_one_video(aweme_id)
        path = OUT / f"{aweme_id}.json"
        path.write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")
        print(aweme_id, type(result).__name__, path.stat().st_size)


if __name__ == "__main__":
    asyncio.run(main())
