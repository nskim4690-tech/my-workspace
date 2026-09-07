# -*- coding: utf-8 -*-
import sys
import json
import os
import urllib.request
import urllib.parse

CRED_PATH = os.path.join(os.path.dirname(__file__), "kakao_credentials.json")


def load_credentials():
    with open(CRED_PATH, "r", encoding="utf-8") as f:
        return json.load(f)


def save_credentials(cred):
    with open(CRED_PATH, "w", encoding="utf-8") as f:
        json.dump(cred, f, ensure_ascii=False, indent=2)


def refresh_access_token(cred):
    url = "https://kauth.kakao.com/oauth/token"
    data = urllib.parse.urlencode({
        "grant_type": "refresh_token",
        "client_id": cred["rest_api_key"],
        "client_secret": cred["client_secret"],
        "refresh_token": cred["refresh_token"],
    }).encode("utf-8")
    req = urllib.request.Request(url, data=data, method="POST")
    with urllib.request.urlopen(req) as res:
        result = json.loads(res.read().decode("utf-8"))
    access_token = result["access_token"]
    if "refresh_token" in result:
        cred["refresh_token"] = result["refresh_token"]
        save_credentials(cred)
    return access_token


def send_kakao_message(text):
    cred = load_credentials()
    access_token = refresh_access_token(cred)

    url = "https://kapi.kakao.com/v2/api/talk/memo/default/send"
    template_object = {
        "object_type": "text",
        "text": text,
        "link": {
            "web_url": "https://developers.kakao.com",
            "mobile_web_url": "https://developers.kakao.com"
        }
    }
    data = urllib.parse.urlencode({
        "template_object": json.dumps(template_object, ensure_ascii=False)
    }).encode("utf-8")

    req = urllib.request.Request(url, data=data, method="POST")
    req.add_header("Authorization", f"Bearer {access_token}")
    req.add_header("Content-Type", "application/x-www-form-urlencoded;charset=utf-8")

    with urllib.request.urlopen(req) as res:
        result = res.read().decode("utf-8")
        print(result)


if __name__ == "__main__":
    msg = sys.argv[1] if len(sys.argv) > 1 else "테스트 메시지"
    send_kakao_message(msg)
