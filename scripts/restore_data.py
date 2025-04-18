import os
import sys
import json

import firebase_admin
from firebase_admin import credentials, firestore


def initialize_db(path):
    """Initialize the SDK once and return a Firestore client."""
    if not firebase_admin._apps:
        cred = credentials.Certificate(path)
        firebase_admin.initialize_app(cred)
        print("Firebase Admin SDK initialized!")
    return firestore.client()


def load_data_from_json(filename="firestore_data.json"):
    """Load list-of-dicts from a JSON file."""
    if not os.path.isfile(filename):
        raise FileNotFoundError(f"No such file: {filename}")
    with open(filename, "r", encoding="utf-8") as f:
        data = json.load(f)
    print(f"Loaded {len(data)} records from {filename}")
    return data


def restore_data(db, data, batch_size=500):
    """
    Restore documents into 'data' collection.

    - If each item has an 'id' key, uses it as the document ID.
    - Otherwise, creates a new document with auto-generated ID.
    """
    batch = db.batch()
    for idx, record in enumerate(data, start=1):
        record_copy = record.copy()
        doc_id = record_copy.pop("id", None)
        coll = db.collection("data")

        if doc_id:
            # Restore with the original ID
            doc_ref = coll.document(doc_id)
            batch.set(doc_ref, record_copy)
        else:
            # Auto-generated ID
            doc_ref = coll.document()
            batch.set(doc_ref, record_copy)

        # Commit every batch_size operations
        if idx % batch_size == 0:
            batch.commit()
            print(f"Restored {idx} documents so far…")
            batch = db.batch()

    # Commit any remaining writes
    batch.commit()
    print(f"Restore complete: {idx} documents written.")


def resource_path(relative_path):
    """Get absolute path to resource (for PyInstaller)."""
    base_path = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base_path, relative_path)


if __name__ == "__main__":
    # Path to your service account JSON
    key_path = resource_path("firebase-admin-sdk.json")
    db = initialize_db(key_path)

    # Load your export file (adjust filename if needed)
    records = load_data_from_json("firestore_data.json")

    # Restore into Firestore
    restore_data(db, records)
