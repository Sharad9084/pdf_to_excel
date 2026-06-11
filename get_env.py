import os

print("=== CHECKING DATABASE ENVIRONMENT VARIABLES ===")
db_url = os.environ.get("DATABASE_URL")
pg_url = os.environ.get("POSTGRES_URL")
print(f"DATABASE_URL configured: {bool(db_url)}")
if db_url:
    print(f"DATABASE_URL: {db_url[:20]}...")
print(f"POSTGRES_URL configured: {bool(pg_url)}")
if pg_url:
    print(f"POSTGRES_URL: {pg_url[:20]}...")
