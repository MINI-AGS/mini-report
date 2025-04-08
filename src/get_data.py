import os

import firebase_admin
from firebase_admin import credentials, firestore


def get_firebase_data():
    cwd = os.getcwd()
    cred = credentials.Certificate(cwd + "/src/firebase-admin-sdk.json")
    app = firebase_admin.initialize_app(cred)
    db = firestore.client()

    print("Firebase Admin SDK initialized!")

    docs = db.collection("data").stream()
    data = []

    for doc in docs:
        data.append(doc.to_dict())

    print("Data retrieved successfully!")
    return data


if __name__ == "__main__":
    get_firebase_data()
