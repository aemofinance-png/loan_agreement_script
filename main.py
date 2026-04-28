from fastapi import FastAPI
from fastapi.responses import StreamingResponse
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel
import io
from generator import generate_agreement  # we'll create this next

app = FastAPI()

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)

class LoanRequest(BaseModel):
    clientName: str
    loanAmount: float
    annualRatePct: float
    loanTermMonths: int
    monthlyPayment: float
    firstPaymentDate: str  # YYYY-MM-DD
    agreementDate: str     # e.g. "March 08, 2026"
    referenceNo: str
    currencySymbol: str = "$"

@app.post("/generate-agreement")
def generate(data: LoanRequest):
    pdf_bytes = generate_agreement(data)
    return StreamingResponse(
        io.BytesIO(pdf_bytes),
        media_type="application/pdf",
        headers={"Content-Disposition": f"inline; filename=agreement_{data.referenceNo}.pdf"}


    )


@app.get("/")
def root():
    return {"message": "Loan Agreement API is running"}