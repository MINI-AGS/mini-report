import csv
import json
import os
import sys
from datetime import date, datetime

import firebase_admin
from firebase_admin import credentials, firestore


def json_serial(obj):
    """
    JSON serializer for objects not serializable by default json code.
    Converts datetime/date (including Firestore’s DatetimeWithNanoseconds)
    into ISO‑formatted strings.
    """
    if isinstance(obj, (datetime, date)):
        return obj.isoformat()
    raise TypeError(f"Type {obj.__class__.__name__} not serializable")


def initialize_db(path):
    """Initialize the SDK once and return a Firestore client."""
    if not firebase_admin._apps:
        cred = credentials.Certificate(path)
        firebase_admin.initialize_app(cred)
        print("Firebase Admin SDK initialized!")
    return firestore.client()


def get_all_data(db):
    """Retrieve all documents from the 'data' collection."""
    docs = db.collection("data").stream()
    data = [doc.to_dict() for doc in docs]
    print(f"Retrieved {len(data)} documents.")
    return data


def delete_all_data(db, batch_size=500):
    """
    Delete all documents in 'data' using batched writes.
    Commits every `batch_size` deletes to stay within limits.
    """
    docs = db.collection("data").stream()
    batch = db.batch()
    count = 0

    for idx, doc in enumerate(docs, start=1):
        batch.delete(doc.reference)
        if idx % batch_size == 0:
            batch.commit()
            print(f"Deleted {idx} documents so far…")
            batch = db.batch()

    # Commit any remaining deletes
    if idx % batch_size != 0:
        batch.commit()
    print(f"Data deleted successfully: {idx} documents removed.")


def save_data_as_json(data, filename="data.json"):
    """
    Write list-of-dicts to a JSON file, converting any datetime
    fields into ISO strings via the json_serial() handler.
    """
    with open(filename, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=4, default=json_serial)
    print(f"Data saved to JSON file: {filename}")


def save_data_as_csv(data, filename="data.csv"):
    """Write list-of-dicts to a CSV file (keys become columns)."""
    if not data:
        print("No data to write to CSV.")
        return

    # Use the union of all keys to handle missing fields
    fieldnames = set().union(*(d.keys() for d in data))
    fieldnames = sorted(fieldnames)

    with open(filename, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(data)
    print(f"Data saved to CSV file: {filename}")


def resource_path(relative_path):
    """Get absolute path to resource (for PyInstaller)."""
    base_path = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base_path, relative_path)


if __name__ == "__main__":
    path = resource_path("firebase-admin-sdk.json")
    db = initialize_db(path)

    data = get_all_data(db)

    filename_json = f"firestore_data_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json"
    filename_csv = f"firestore_data_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv"

    save_data_as_json(data, filename=filename_json)
    save_data_as_csv(data, filename=filename_csv)

    delete_all_data(db)
