import os
import shutil
import tempfile
from pathlib import Path
from fastapi import FastAPI, UploadFile, File, HTTPException
from fastapi.responses import FileResponse
from tsk_final_engine import run

app = FastAPI(title="TSK Coil Inspection API")

@app.get("/")
async def root():
    return {"message": "TSK Backend is alive. Use POST /process"}

@app.post("/process")
async def process_excel(file: UploadFile = File(...)):
    if not file.filename.endswith(('.xlsx', '.xls')):
        raise HTTPException(400, "Only Excel files are accepted")

    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp_input:
        shutil.copyfileobj(file.file, tmp_input)
        input_path = tmp_input.name

    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp_output:
        output_path = tmp_output.name

    try:
        run(input_path, output_path)
    except Exception as e:
        os.unlink(input_path)
        os.unlink(output_path)
        raise HTTPException(500, f"Processing failed: {str(e)}")

    response = FileResponse(
        path=output_path,
        filename="TSK_Final_Results.xlsx",
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
    response.background = lambda: (os.unlink(input_path), os.unlink(output_path))
    return response
