from dotenv import load_dotenv
import os

load_dotenv()
api_key = os.getenv("FIREBASE_API_KEY")

print("Happy coding!")
