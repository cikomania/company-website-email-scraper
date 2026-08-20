---------- MacOS ----------

cd ~/Desktop/firmalar
python3 -m venv venv 
source venv/bin/activate
python3 -m pip install pandas openpyxl selenium webdriver-manager

"/Applications/Google Chrome.app/Contents/MacOS/Google Chrome" \
--remote-debugging-port=9222 \
--user-data-dir="$HOME/chrome_selenium"

python3 sitemailbul.py


---------- Windows ----------

cd C:\firmalar
python3 -m venv venv 
source venv/bin/activate
python -m pip install pandas openpyxl selenium webdriver-manager

& "C:\Program Files\Google\Chrome\Application\chrome.exe" 
--remote-debugging-port=9222 
--user-data-dir="$env:USERPROFILE\chrome_selenium"

python sitemailbul.py