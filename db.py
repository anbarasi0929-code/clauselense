
import os
from dotenv import load_dotenv
from pymongo import MongoClient

load_dotenv()

uri = os.getenv("MONGODB_URI")

if not uri:
    raise ValueError("MONGODB_URI is missing from .env")

client = MongoClient(uri, serverSelectionTimeoutMS=5000)
db = client["ClauseLens"]

if __name__ == "__main__":
    try:
        client.admin.command("ping")
        print("MongoDB Atlas connected successfully!")
    except Exception as e:
        print("MongoDB connection failed:", e)
