---------- MacOS ----------

cd ~/Desktop/firmalar
source venv/bin/activate

"/Applications/Google Chrome.app/Contents/MacOS/Google Chrome" \
--remote-debugging-port=9222 \
--user-data-dir="$HOME/chrome_selenium"

python3 sitemailbul.py


---------- Windows ----------

cd C:\firmalar
source venv/bin/activate

& "C:\Program Files\Google\Chrome\Application\chrome.exe" 
--remote-debugging-port=9222 
--user-data-dir="$env:USERPROFILE\chrome_selenium"

python sitemailbul.py
