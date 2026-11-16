from pydantic import BaseModel
from typing import Any, Optional

# Model API

class SubmitResponse(BaseModel):
    error_code: int
    error_message: str
    request_id: str


class StatusResponse(BaseModel):
    request_id: str
    status: str
    result: Optional[Any] = None
