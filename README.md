# AdvancedAI Project

## Installation Requirements

- Python 3.10
- ChromeDriver

### Installation Steps

1. Clone the repository
```bash
git clone <repository-url>
cd AdvancedAI_Project
```

2. Install dependencies
```bash
pip install -r requirements.txt
```

3. Modify the actual driver path
```python
# mail.py
def open_browser_with_cookies(cookie):
    driver_path = "actual driver path"
    # driver_path = "xxx/xxx/chromedriver.exe"
```

4. Adapt your base_url, api_key and model
```python
# backend.py
# Modify to your url and key
client = OpenAI(
    base_url="https://...",
    api_key="xxx",
)
# def ask(): modify "deepseek-r1:671b" to your model
model="deepseek-r1:671b",
```

## Usage
1. Run the python file
```bash
python backend.py
```

2. Enter the website address in the browser
```bash
http://127.0.0.1:5001/
```

3. Enter Cookie in the input box

4. Click the button
