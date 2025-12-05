from typing import Any, Dict, List, Optional

from pydantic import BaseModel, Field


class ColumnInfo(BaseModel):
    name: str
    detectedType: str
    overrideType: Optional[str] = None
    nullable: bool = True


class CellError(BaseModel):
    rowId: int
    column: str
    expectedType: str
    actualValue: Any
    message: str


class RowPayload(BaseModel):
    rowId: int
    values: Dict[str, Any]


class UploadResponse(BaseModel):
    sessionId: str
    columns: List[ColumnInfo]
    rows: List[RowPayload]
    errors: List[CellError]
    duplicateGroups: List[List[int]]
    sheetName: Optional[str] = None
    sheetNames: List[str] = Field(default_factory=list)


class ValidateRequest(BaseModel):
    sessionId: str
    rows: List[RowPayload]
    columnTypes: Dict[str, str] = Field(default_factory=dict)
    columns: List[ColumnInfo] = Field(default_factory=list)


class ExportRequest(BaseModel):
    sessionId: Optional[str] = None
    rows: List[RowPayload] = Field(default_factory=list)
    columns: List[ColumnInfo] = Field(default_factory=list)


class ReportRequest(BaseModel):
    rows: List[RowPayload] = Field(default_factory=list)
    columns: List[ColumnInfo] = Field(default_factory=list)
    errors: List[CellError] = Field(default_factory=list)


class SheetRequest(BaseModel):
    sessionId: str
    sheetName: str


class DuplicateRemovalRequest(BaseModel):
    sessionId: str
    rowIds: List[int] = Field(default_factory=list)


class ValidateResponse(UploadResponse):
    pass

