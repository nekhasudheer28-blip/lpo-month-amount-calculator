from fastapi import FastAPI, File, Form, HTTPException, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import StreamingResponse

from services.lpo_mat_processor import process_lpo_workbook


app = FastAPI(title="LPO- Month & Amount Calculator API")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=False,
    allow_methods=["*"],
    allow_headers=["*"],
    expose_headers=["Content-Disposition", "X-Job-Numbers", "X-Summary"],
)


@app.get("/health")
def health():
    return {"status": "ok"}


@app.post("/process")
async def process_excel(
    file: UploadFile = File(...),
    selected_job_numbers: str = Form("[]"),
):
    if not file.filename.lower().endswith(".xlsx"):
        raise HTTPException(status_code=400, detail="Please upload an .xlsx file.")

    try:
        result = process_lpo_workbook(
            content=await file.read(),
            selected_job_numbers_json=selected_job_numbers,
            original_filename=file.filename,
        )
    except ValueError as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from exc
    except Exception as exc:
        raise HTTPException(status_code=500, detail=f"Failed to process workbook: {exc}") from exc

    headers = {
        "Content-Disposition": f'attachment; filename="{result.filename}"',
        "X-Job-Numbers": result.job_numbers_json,
        "X-Summary": result.summary_json,
    }

    return StreamingResponse(
        result.stream,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers=headers,
    )