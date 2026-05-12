from reportlab.lib.pagesizes import letter
from reportlab.platypus import (SimpleDocTemplate, Paragraph, Spacer, Table,
                                TableStyle, HRFlowable, Image)
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib import colors
from reportlab.lib.units import inch
from reportlab.lib.enums import TA_CENTER, TA_JUSTIFY, TA_RIGHT, TA_LEFT
from datetime import date
from dateutil.relativedelta import relativedelta
import io
import base64
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
import os

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
# Register Arial font to support international currency symbols
pdfmetrics.registerFont(TTFont('Arial', os.path.join(BASE_DIR, 'Arial.ttf')))
pdfmetrics.registerFont(TTFont('Arial-Bold', os.path.join(BASE_DIR, 'Arial-Bold.ttf')))

LOGO_PATH = os.path.join(BASE_DIR, 'aemo_logo.png')

# Branding Colors
NAVY = colors.HexColor('#1A3560')
RED = colors.HexColor('#C0272D')
LIGHT_BLUE = colors.HexColor('#F4F7FB')
BG_GREY = colors.HexColor('#F9FAFB')
TEXT_MAIN = colors.HexColor('#111827')
TEXT_DIM = colors.HexColor('#6B7280')
BORDER_COLOR = colors.HexColor('#E5E7EB')
WHITE = colors.white

def generate_agreement(data) -> bytes:
    buffer = io.BytesIO()

    monthly_rate = data.annualRatePct / 100 / 12
    first_date = date.fromisoformat(data.firstPaymentDate)

    # Amortization logic
    def build_schedule():
        balance = data.loanAmount
        rows = []
        for i in range(1, data.loanTermMonths + 1):
            interest = round(balance * monthly_rate, 2)
            if i == data.loanTermMonths:
                princ = balance
                pmt = round(princ + interest, 2)
            else:
                princ = round(data.monthlyPayment - interest, 2)
                pmt = data.monthlyPayment
            balance = round(balance - princ, 2)
            due = first_date + relativedelta(months=i - 1)
            rows.append((i, due.strftime('%b %d, %Y'),
                         data.currencySymbol + f'{pmt:,.2f}',
                         data.currencySymbol + f'{princ:,.2f}',
                         data.currencySymbol + f'{interest:,.2f}',
                         data.currencySymbol + f'{max(balance, 0):,.2f}'))
        return rows

    schedule = build_schedule()
    total_repay = round(sum(float(r[2].replace(data.currencySymbol, '').replace(',', '')) for r in schedule), 2)
    total_interest = round(total_repay - data.loanAmount, 2)
    last_date = (first_date + relativedelta(months=data.loanTermMonths - 1)).strftime('%B %d, %Y')

    # Styles
    def S(name, **kw): return ParagraphStyle(name, **kw)
    sTitle = S('sTitle', fontSize=18, fontName='Arial-Bold', textColor=NAVY, alignment=TA_LEFT, spaceAfter=8)
    sSubTitle = S('sSubTitle', fontSize=10, fontName='Arial', textColor=TEXT_DIM, alignment=TA_LEFT, spaceAfter=12)
    sSectionHead = S('sSectionHead', fontSize=9, fontName='Arial-Bold', textColor=NAVY, spaceBefore=10, spaceAfter=6, letterSpacing=0.5)
    sBody = S('sBody', fontSize=9, fontName='Arial', textColor=TEXT_MAIN, leading=13, alignment=TA_JUSTIFY)
    sSmall = S('sSmall', fontSize=8, fontName='Arial', textColor=TEXT_DIM, leading=11)
    sLabel = S('sLabel', fontSize=7, fontName='Arial-Bold', textColor=TEXT_DIM, alignment=TA_LEFT, spaceAfter=2)
    sValue = S('sValue', fontSize=10, fontName='Arial-Bold', textColor=TEXT_MAIN, alignment=TA_LEFT)
    sMetricLabel = S('sMetricLabel', fontSize=8, fontName='Arial', textColor=WHITE, alignment=TA_CENTER)
    sMetricValue = S('sMetricValue', fontSize=14, fontName='Arial-Bold', textColor=WHITE, alignment=TA_CENTER)
    sTCHeader = S('sTCHeader', fontSize=9, fontName='Arial-Bold', textColor=TEXT_MAIN, spaceBefore=8, spaceAfter=4)
    sTCBody = S('sTCBody', fontSize=8.5, fontName='Arial', textColor=TEXT_MAIN, leading=11, alignment=TA_JUSTIFY, leftIndent=10)

    W = letter[0] - 1.3 * inch

    doc = SimpleDocTemplate(buffer, pagesize=letter,
                            rightMargin=0.65 * inch, leftMargin=0.65 * inch,
                            topMargin=0.5 * inch, bottomMargin=0.65 * inch)
    story = []

    # --- Header ---
    aemo = Image(LOGO_PATH, width=1.4 * inch, height=0.65 * inch)
    ref_data = [
        [aemo, [Paragraph('DATE OF AGREEMENT', sLabel), Paragraph(data.agreementDate, sValue)],
               [Paragraph('REFERENCE NUMBER', sLabel), Paragraph(data.referenceNo, sValue)]]
    ]
    ht = Table(ref_data, colWidths=[2 * inch, W / 2 - 1 * inch, W / 2 - 1 * inch])
    ht.setStyle(TableStyle([
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
        ('LEFTPADDING', (0, 0), (-1, -1), 0),
    ]))
    story += [ht, Spacer(1, 15)]

    story += [
        Paragraph('Personal Loan Agreement', sTitle),
        Paragraph('This document outlines the binding financial commitment between AEMO Finance and the Borrower.', sSubTitle),
    ]

    # --- Key Facts Statement (Premium Disclosure Box) ---
    kfs_data = [
        [Paragraph('LOAN AMOUNT', sMetricLabel), Paragraph('ANNUAL RATE', sMetricLabel), 
         Paragraph('LOAN TERM', sMetricLabel), Paragraph('TOTAL INTEREST', sMetricLabel)],
        [Paragraph(data.currencySymbol + f'{data.loanAmount:,.2f}', sMetricValue), 
         Paragraph(f'{data.annualRatePct:.2f}%', sMetricValue), 
         Paragraph(f'{data.loanTermMonths} MO', sMetricValue), 
         Paragraph(data.currencySymbol + f'{total_interest:,.2f}', sMetricValue)]
    ]
    kfs_table = Table(kfs_data, colWidths=[W / 4] * 4)
    kfs_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, -1), NAVY),
        ('ROUNDEDCORNERS', [8, 8, 8, 8]),
        ('TOPPADDING', (0, 0), (-1, 0), 12),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 0),
        ('TOPPADDING', (0, 1), (-1, 1), 2),
        ('BOTTOMPADDING', (0, 1), (-1, 1), 12),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
    ]))
    story += [kfs_table, Spacer(1, 20)]

    # --- Section: Borrower & Management ---
    story += [Paragraph('PARTIES & MANAGEMENT', sSectionHead)]
    pm_data = [
        [[Paragraph('BORROWER FULL NAME', sLabel), Paragraph(data.clientName, sValue)],
         [Paragraph('LOAN PRODUCT', sLabel), Paragraph('Standard Personal Loan', sValue)]],
        [[Paragraph('LENDER / MANAGER', sLabel), Paragraph('AEMO Finance', sValue)],
         [Paragraph('REPAYMENT FREQUENCY', sLabel), Paragraph('Monthly Installments', sValue)]]
    ]
    pm_table = Table(pm_data, colWidths=[W / 2, W / 2])
    pm_table.setStyle(TableStyle([
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 10),
        ('GRID', (0, 0), (-1, -1), 0.5, BORDER_COLOR),
        ('BACKGROUND', (0, 0), (-1, -1), BG_GREY),
        ('LEFTPADDING', (0, 0), (-1, -1), 10),
        ('TOPPADDING', (0, 0), (-1, -1), 8),
    ]))
    story += [pm_table, Spacer(1, 15)]

    # --- Section: Repayment Framework ---
    story += [Paragraph('REPAYMENT FRAMEWORK', sSectionHead)]
    story += [
        Paragraph(
            f'The Borrower is committed to a monthly installment of <b>{data.currencySymbol}{data.monthlyPayment:,.2f}</b>. '
            f'Deductions will be processed between the <b>28th and 30th</b> of every month. '
            f'The first payment is due on <b>{first_date.strftime("%B %d, %Y")}</b>, concluding on <b>{last_date}</b>.',
            sBody
        ),
        Spacer(1, 12)
    ]

    # --- Section: Amortization Schedule ---
    story += [Paragraph('AMORTIZATION SCHEDULE', sSectionHead)]
    amort_data = [['#', 'DUE DATE', 'PAYMENT', 'PRINCIPAL', 'INTEREST', 'BALANCE']] + \
                 [[str(r[0]), r[1], r[2], r[3], r[4], r[5]] for r in schedule]
    at = Table(amort_data, colWidths=[0.45 * inch, 1.1 * inch, 1.1 * inch, 1 * inch, 0.9 * inch, 1.15 * inch], repeatRows=1)
    at.setStyle(TableStyle([
        ('FONTNAME', (0, 0), (-1, 0), 'Arial-Bold'),
        ('FONTSIZE', (0, 0), (-1, 0), 7),
        ('TEXTCOLOR', (0, 0), (-1, 0), TEXT_DIM),
        ('LINEBELOW', (0, 0), (-1, 0), 1, BORDER_COLOR),
        ('FONTNAME', (0, 1), (-1, -1), 'Arial'),
        ('FONTSIZE', (0, 1), (-1, -1), 8),
        ('ALIGN', (2, 1), (-1, -1), 'RIGHT'),
        ('TOPPADDING', (0, 0), (-1, -1), 6),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 6),
        ('ROWBACKGROUNDS', (0, 1), (-1, -1), [WHITE, BG_GREY]),
    ]))
    story += [at, Spacer(1, 20)]

    # --- Section: Expanded Terms & Conditions ---
    story += [Paragraph('TERMS & CONDITIONS', sSectionHead)]
    
    tcs = [
        ('1. LOAN DISBURSEMENT', 'Funds will be disbursed to the bank account provided by the Borrower upon final verification and execution of this agreement.'),
        ('2. REPAYMENT OBLIGATION', f'The Borrower agrees to pay {data.loanTermMonths} consecutive monthly installments. Failure to maintain sufficient funds in the designated account may result in additional charges.'),
        ('3. LATE FEES & PENALTIES', 'A late payment fee of 0.5% of the installment amount will be applied for each day the payment remains outstanding, commencing after a 5-day grace period from the due date.'),
        ('4. PREPAYMENT', 'The Borrower reserves the right to repay the loan in full or in part at any time without incurring prepayment penalties.'),
        ('5. EVENTS OF DEFAULT', 'Default occurs if the Borrower fails to make payments, provides false information, or enters insolvency. Upon default, AEMO Finance may demand immediate repayment of the full outstanding balance.'),
        ('6. DATA PRIVACY & COMMUNICATION', 'The Borrower consents to AEMO Finance processing personal data for loan management and receiving notifications via SMS, Email, or Phone.'),
        ('7. GOVERNING LAW', 'This agreement is governed by the laws of the jurisdiction where AEMO Finance is registered.'),
    ]

    for title, content in tcs:
        story += [Paragraph(title, sTCHeader), Paragraph(content, sTCBody)]
    
    story += [Spacer(1, 25)]

    # --- Section: Approval & Execution ---
    sig_data = [
        [[Paragraph('FOR AND ON BEHALF OF AEMO FINANCE', sLabel), Spacer(1, 25), 
          HRFlowable(width=2*inch, thickness=1, color=TEXT_MAIN, hAlign='LEFT'), Paragraph('Authorized Signatory', sSmall)],
         [Paragraph('CLIENT / BORROWER ACKNOWLEDGMENT', sLabel), Spacer(1, 25), 
          HRFlowable(width=2*inch, thickness=1, color=TEXT_MAIN, hAlign='LEFT'), Paragraph(data.clientName, sSmall)]]
    ]
    sig_table = Table(sig_data, colWidths=[W/2, W/2])
    story += [sig_table, Spacer(1, 20)]

    story += [
        HRFlowable(width=W, thickness=0.5, color=BORDER_COLOR),
        Spacer(1, 5),
        Paragraph('AEMO Finance  •  Support: finance@aemostack.com  •  Confidential Financial Document', S('Footer', fontSize=7, textColor=TEXT_DIM, alignment=TA_CENTER))
    ]

    doc.build(story)
    buffer.seek(0)
    return buffer.read()
