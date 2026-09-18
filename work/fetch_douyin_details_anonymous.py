import importlib.util
import json
import secrets
import string
from pathlib import Path
from urllib.parse import quote, urlencode

import httpx


SIGNER_PATH = Path(
    r"C:\Users\HZJ\Desktop\Git\work\Douyin_TikTok_Download_API\crawlers\douyin\web\abogus.py"
)
OUT = Path(r"C:\Users\HZJ\Desktop\Git\work\douyin_details_anonymous")
IDS = ["7392785204122783028", "7397739121348726054", "7403557916621540660"]
UA = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
    "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/90.0.4430.212 Safari/537.36"
)


def load_signer():
    spec = importlib.util.spec_from_file_location("douyin_abogus", SIGNER_PATH)
    module = importlib.util.module_from_spec(spec)
    assert spec.loader is not None
    spec.loader.exec_module(module)
    return module.ABogus()


def params_for(aweme_id: str) -> dict[str, str]:
    alphabet = string.ascii_letters + string.digits + "-_"
    anonymous_ms_token = "".join(secrets.choice(alphabet) for _ in range(126)) + "=="
    return {
        "device_platform": "webapp",
        "aid": "6383",
        "channel": "channel_pc_web",
        "pc_client_type": "1",
        "version_code": "190500",
        "version_name": "19.5.0",
        "cookie_enabled": "true",
        "browser_language": "zh-CN",
        "browser_platform": "Win32",
        "browser_name": "Chrome",
        "browser_online": "true",
        "engine_name": "Blink",
        "os_name": "Windows",
        "os_version": "10",
        "platform": "PC",
        "screen_width": "1920",
        "screen_height": "1080",
        "browser_version": "90.0.4430.212",
        "engine_version": "90.0.4430.212",
        "cpu_core_num": "12",
        "device_memory": "8",
        "aweme_id": aweme_id,
        "msToken": anonymous_ms_token,
    }


def main():
    OUT.mkdir(parents=True, exist_ok=True)
    signer = load_signer()
    headers = {
        "User-Agent": UA,
        "Referer": "https://www.douyin.com/",
        "Accept-Language": "zh-CN,zh;q=0.9",
    }
    with httpx.Client(headers=headers, follow_redirects=True, timeout=30.0) as client:
        register_payload = {
            "region": "cn",
            "aid": 1768,
            "needFid": False,
            "service": "www.ixigua.com",
            "migrate_info": {"ticket": "", "source": "node"},
            "cbUrlProtocol": "https",
            "union": True,
        }
        register_response = client.post(
            "https://ttwid.bytedance.com/ttwid/union/register/",
            json=register_payload,
        )
        print("anonymous-session", register_response.status_code, sorted(client.cookies.keys()))
        for aweme_id in IDS:
            # Establish a fresh anonymous session. No browser or account cookies are read.
            client.get(f"https://www.douyin.com/video/{aweme_id}")
            params = params_for(aweme_id)
            a_bogus = quote(signer.get_value(params), safe="")
            url = (
                "https://www.douyin.com/aweme/v1/web/aweme/detail/?"
                + urlencode(params)
                + "&a_bogus="
                + a_bogus
            )
            response = client.get(url)
            path = OUT / f"{aweme_id}.json"
            path.write_bytes(response.content)
            print(aweme_id, response.status_code, response.headers.get("content-type"), len(response.content))


if __name__ == "__main__":
    main()
