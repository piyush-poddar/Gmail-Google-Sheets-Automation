# 📧 Gmail to Google Sheets Exporter

This project is a Python-based automation tool that:

- Connects to your Gmail inbox.
- Extracts emails data (date, sender, subject, body snippet).
- Creates a new Google Sheet and saves the extracted data in a tabular format.
- Supports token-based login (`token.json`) to avoid repeated authentication prompts.


## 🚀 Features

- 📬 Reads Emails from Gmail  
- 📄 Creates a New Google Sheet automatically  
- 📝 Saves:  
  - Date  
  - From (Sender)  
  - Subject  
  - A message snippet of the body  
- 🔐 Secure token storage for re-use  
- 🔄 Auto refreshes tokens when expired  


## 🗂️ Project Structure

```
📂 your_project/
 ┣ 📄 main.py            # Main script
 ┣ 📄 credentials.json   # OAuth credentials (download from Google Cloud Console)
 ┣ 📄 token.json         # Auto-generated after first login
```


## ⚙️ Setup Instructions

**1️⃣ Enable APIs on Google Cloud**  
- Gmail API  
- Google Sheets API  
- Download `credentials.json` and place it in your project directory.

**2️⃣ Install Dependencies**
```bash
pip install --upgrade google-api-python-client google-auth google-auth-oauthlib gspread
```

**3️⃣ Run the Script**
```bash
python main.py
```
First time → Browser opens → Login with your Google account → Grant permissions.

A new Google Sheet will be created. URL will be displayed in the terminal.


## ✅ Example Output

| Date      | From                  | Subject        | Snippet                |
|-----------|-----------------------|----------------|------------------------|
| Thu, ...  | someone@example.com   | Project Update | Here's the project...  |
