import os
import datetime
import sys

# อ่าน environment variables
token = os.getenv("SMARTSHEET_TOKEN")
created_at = os.getenv("TOKEN_CREATED_AT")

print("🔍 Checking Smartsheet Token Expiry...")

if not token:
    print("❌ ERROR: SMARTSHEET_TOKEN is missing!")
    sys.exit(1)

if not created_at:
    print("⚠️ TOKEN_CREATED_AT not set — cannot check expiry date.")
    sys.exit(0)

try:
    created_date = datetime.datetime.strptime(created_at, "%Y-%m-%d")
except ValueError:
    print("❌ TOKEN_CREATED_AT format invalid! Use YYYY-MM-DD")
    sys.exit(1)

# สมมุติว่า token มีอายุ 90 วัน
expiry_date = created_date + datetime.timedelta(days=90)
today = datetime.datetime.utcnow()

days_left = (expiry_date - today).days

if days_left <= 0:
    print("❌ Token has expired! Please generate a new one.")
    sys.exit(1)
elif days_left <= 7:
    print(f"⚠️ Token will expire in {days_left} days — please renew soon!")
else:
    print(f"✅ Token is valid. ({days_left} days left)")

sys.exit(0)
