from typing import Any

from fastapi import FastAPI, File, HTTPException, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import StreamingResponse

from .schemas import (
    DuplicateRemovalRequest,
    ExportRequest,
    ReportRequest,
    SheetRequest,
    UploadResponse,
    ValidateRequest,
    ValidateResponse,
)
from .services.excel_service import (
    generate_error_report,
    generate_error_report_from_rows,
    export_rows_to_excel,
    parse_sheet,
    process_uploaded_file,
    remove_rows,
    revalidate,
    session_store,
)

app = FastAPI(title="Excel Checker API", version="0.1.0")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)


@app.post("/api/upload", response_model=UploadResponse)
async def upload_excel(file: UploadFile = File(...)) -> Any:
    file_bytes = await file.read()
    try:
        session, payload = process_uploaded_file(file_bytes, file.filename)
    except ValueError as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from exc
    session_id = session_store.create_session(session)
    payload["sessionId"] = session_id
    return payload


@app.post("/api/validate", response_model=ValidateResponse)
async def validate_excel(payload: ValidateRequest) -> Any:
    try:
        session = session_store.get(payload.sessionId)
    except KeyError as exc:
        raise HTTPException(status_code=404, detail="Session not found") from exc
    rows = [row.dict() for row in payload.rows]
    columns_payload = [column.dict() for column in payload.columns]
    result = revalidate(session, rows, payload.columnTypes, columns_payload)
    session_store.update(payload.sessionId, session)
    result["sessionId"] = payload.sessionId
    return result


@app.post("/api/duplicates/remove", response_model=ValidateResponse)
async def remove_duplicates(request: DuplicateRemovalRequest) -> Any:
    try:
        session = session_store.get(request.sessionId)
    except KeyError as exc:
        raise HTTPException(status_code=404, detail="Session not found") from exc
    updated_rows = remove_rows(session, request.rowIds)
    result = revalidate(session, updated_rows, session.overrides)
    session_store.update(request.sessionId, session)
    result["sessionId"] = request.sessionId
    return result


@app.get("/api/report/{session_id}")
async def download_report(session_id: str) -> StreamingResponse:
    try:
        session = session_store.get(session_id)
    except KeyError as exc:
        raise HTTPException(status_code=404, detail="Session not found") from exc
    report_bytes = generate_error_report(session)
    filename = session.original_filename.rsplit(".", 1)[0] + "_report.xlsx"
    return StreamingResponse(
        iter([report_bytes]),
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f"attachment; filename={filename}"},
    )


@app.post("/api/export")
async def export_sheet(request: ExportRequest) -> StreamingResponse:
    if not request.rows:
        raise HTTPException(status_code=400, detail="No data provided for export.")
    export_bytes = export_rows_to_excel(
        [row.dict() for row in request.rows],
        [column.dict() for column in request.columns],
    )
    filename = (request.sessionId or "edited").split(".")[0] + "_edited.xlsx"
    return StreamingResponse(
        iter([export_bytes]),
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f"attachment; filename={filename}"},
    )


@app.post("/api/report")
async def download_report_from_payload(request: ReportRequest) -> StreamingResponse:
    if not request.rows or not request.errors:
        raise HTTPException(status_code=400, detail="Rows and errors are required.")
    report_bytes = generate_error_report_from_rows(
        [row.dict() for row in request.rows],
        [error.dict() for error in request.errors],
    )
    filename = "validation_report.xlsx"
    return StreamingResponse(
        iter([report_bytes]),
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f"attachment; filename={filename}"},
    )


@app.post("/api/sheet", response_model=UploadResponse)
async def switch_sheet(request: SheetRequest) -> Any:
    try:
        session = session_store.get(request.sessionId)
    except KeyError as exc:
        raise HTTPException(status_code=404, detail="Session not found") from exc
    if request.sheetName not in session.sheet_names:
        raise HTTPException(status_code=400, detail="Sheet not found in workbook.")
    try:
        (
            rows,
            columns,
            detected_types,
            column_info,
            errors,
            duplicate_groups,
        ) = parse_sheet(session.workbook_bytes, request.sheetName)
    except ValueError as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from exc
    session.rows = rows
    session.columns = columns
    session.detected_types = detected_types
    session.column_info = column_info
    session.errors = errors
    session.duplicate_groups = duplicate_groups
    session.sheet_name = request.sheetName
    session.overrides = {}
    session_store.update(request.sessionId, session)
    payload = {
        "columns": column_info,
        "rows": rows,
        "errors": errors,
        "duplicateGroups": duplicate_groups,
        "sheetName": request.sheetName,
        "sheetNames": session.sheet_names,
        "sessionId": request.sessionId,
    }
    return payload

