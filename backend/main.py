from fastapi import FastAPI, HTTPException
from pydantic import BaseModel
from pymongo import MongoClient
import pandas as pd

app = FastAPI()

# MongoDB Connection
client = MongoClient("mongodb://localhost:27017/")
db = client["Employess"]
collection = db["Managers"]

# Excel Path
excel_path = "data.xlsx"

class RowData(BaseModel):
    unique_key: str
    data: dict

@app.get("/data")
async def get_data():
    data = list(collection.find({}, {'_id': 0}))
    return data

@app.put("/update")
async def update_data(row_data: RowData):
    try:
        # Update MongoDB
        collection.update_one({"unique_key": row_data.unique_key}, {"$set": row_data.data})

        # Update Excel
        df = pd.read_excel(excel_path)
        for index, row in df.iterrows():
            if str(row['unique_key']) == str(row_data.unique_key):
                for key, value in row_data.data.items():
                    df.at[index, key] = value
                df.to_excel(excel_path, index=False)
                break
        return {"message": "Data updated successfully"}
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))
