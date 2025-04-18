import csv
import json
import os
import sys
from datetime import date, datetime

import firebase_admin
from firebase_admin import credentials, firestore


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
    """Delete all documents in 'data' using batched writes."""
    docs = db.collection("data").stream()
    batch = db.batch()
    for idx, doc in enumerate(docs, start=1):
        batch.delete(doc.reference)
        if idx % batch_size == 0:
            batch.commit()
            print(f"Deleted {idx} so far…")
            batch = db.batch()
    batch.commit()
    print(f"Data deleted successfully: {idx} documents removed.")


def _json_serializer(obj):
    """Helper to make datetime/date serializable to ISO strings."""
    if isinstance(obj, (datetime, date)):
        return obj.isoformat()
    raise TypeError(f"Type {obj.__class__.__name__} not serializable")


def save_data_as_json(data, directory=".", prefix="data"):
    """
    Dump data to JSON at path:
      {directory}/{prefix}_{YYYYMMDD}.json
    """
    date_str = datetime.now().strftime("%Y%m%d")
    filename = os.path.join(directory, f"{prefix}_{date_str}.json")
    with open(filename, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=4, default=_json_serializer)
    print(f"Data saved to JSON: {filename}")
    return filename


def save_data_as_csv(data, directory=".", prefix="data"):
    """
    Dump data to CSV at path:
      {directory}/{prefix}_{YYYYMMDD}.csv
    """
    if not data:
        print("No data to write to CSV.")
        return None

    date_str = datetime.now().strftime("%Y%m%d")
    filename = os.path.join(directory, f"{prefix}_{date_str}.csv")

    # union of all keys in all docs
    fieldnames = sorted(set().union(*(d.keys() for d in data)))
    with open(filename, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(data)
    print(f"Data saved to CSV: {filename}")
    return filename
