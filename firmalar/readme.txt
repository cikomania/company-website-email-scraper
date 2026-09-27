---------- MacOS ----------

1. Terminal:

"/Applications/Google Chrome.app/Contents/MacOS/Google Chrome" \
--remote-debugging-port=9222 \
--user-data-dir="$HOME/chrome_selenium"

2. Terminal:

cd ~/Desktop/firmalar
source venv/bin/activate
python3 sitemailbul.py


---------- Windows ----------

Tek Terminal:

& "C:\Program Files\Google\Chrome\Application\chrome.exe" 
--remote-debugging-port=9222 
--user-data-dir="$env:USERPROFILE\chrome_selenium"

cd C:\firmalar
venv\Scripts\activate
python sitemailbul.py
