import importlib.util
from collections import Counter
import re

spec = importlib.util.spec_from_file_location(
    "m", "c:/Users/andleeb/Desktop/DataScraping/scripts/uni_data_free_google.py"
)
m = importlib.util.module_from_spec(spec)
spec.loader.exec_module(m)

import requests

s = requests.Session()
s.headers.update(m.browser_headers())
r = s.get(
    m.BING_URL,
    params={"q": "Bahria University Asif Khaliq linkedin", "count": 10, "mkt": "en-US"},
    headers={**m.browser_headers(), "Referer": "https://www.bing.com/"},
    timeout=25,
)
t = r.text
print("len", len(t), "b_algo", t.count("b_algo"))
m2 = re.findall(r'class="(b_[a-z0-9_]+)"', t)
print("top classes", Counter(m2).most_common(20))
if "captcha" in t.lower() or "unusual traffic" in t.lower():
    print("possible block")
open("c:/Users/andleeb/Desktop/DataScraping/output/bing_debug.html", "w", encoding="utf-8").write(t)
