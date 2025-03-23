import os

import firebase_admin
from firebase_admin import credentials, firestore

cwd = os.getcwd()
cred = credentials.Certificate(cwd + "/src/firebase-admin-sdk.json")
app = firebase_admin.initialize_app(cred)
db = firestore.client()

print("Firebase Admin SDK initialized!")
print(app)

print("Testing uploading a document to Firestore...")
doc_ref = db.collection("ejemplos").document("alovelace")
doc_ref.set({"first": "Ada", "last": "Lovelace", "born": 1815})

print("Document uploaded successfully!")

print("Testing reading a document from Firestore...")
doc = doc_ref.get()
print("Document data: {}".format(doc.to_dict()))

print("Document read successfully!")

print("Testing deleting a document from Firestore...")
doc_ref.delete()
print("Document deleted successfully!")


print("Happy coding!")
