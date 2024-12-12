import json

class CookiesCleaner:
    @staticmethod
    def clean_cookies(cookies):
        required_cookies = {"li_at", "li_rm", "JSESSIONID"}
        cleaned_cookies = []

        for cookie in cookies:
            if cookie["name"] in required_cookies:
                if cookie.get("sameSite") not in {"Strict", "Lax", "None"}:
                    cookie["sameSite"] = "Lax"
                cookie["secure"] = True
                if cookie["name"] in {"li_at", "li_rm"}:
                    cookie["httpOnly"] = True
                cleaned_cookies.append(cookie)

        return cleaned_cookies

json_file = r'Input Files\cookies.json'
with open(json_file, 'r', encoding='utf-8') as f:
    original_cookies = json.load(f)

cleaned_cookies = CookiesCleaner.clean_cookies(original_cookies)

with open(json_file, 'w', encoding='utf-8') as f:
    json.dump(cleaned_cookies, f, ensure_ascii=False, indent=4)
