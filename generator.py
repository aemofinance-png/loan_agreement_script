from reportlab.lib.pagesizes import letter
from reportlab.platypus import (SimpleDocTemplate, Paragraph, Spacer, Table,
                                TableStyle, HRFlowable)
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib import colors
from reportlab.lib.units import inch
from reportlab.lib.enums import TA_CENTER, TA_JUSTIFY, TA_RIGHT
from datetime import date
from dateutil.relativedelta import relativedelta
import io
import base64
from reportlab.lib.utils import ImageReader
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, HRFlowable, Image

import os

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
LOGO_PATH = os.path.join(BASE_DIR, 'aemo_logo.png')
LOGO_B64 = base64.b64encode(open(LOGO_PATH, 'rb').read()).decode('utf-8')
# colours
NAVY  = colors.HexColor('#1A3560')
RED   = colors.HexColor('#C0272D')
LIGHT = colors.HexColor('#F2F5FA')
WHITE = colors.white
GREY  = colors.HexColor('#555555')
LGREY = colors.HexColor('#DDDDDD')

def generate_agreement(data) -> bytes:
    buffer = io.BytesIO()

    monthly_rate = data.annualRatePct / 100 / 12
    first_date   = date.fromisoformat(data.firstPaymentDate)

    # amortization
    def build_schedule():
        balance = data.loanAmount
        rows = []
        for i in range(1, data.loanTermMonths + 1):
            interest = round(balance * monthly_rate, 2)
            if i == data.loanTermMonths:
                princ = balance
                pmt   = round(princ + interest, 2)
            else:
                princ = round(data.monthlyPayment - interest, 2)
                pmt   = data.monthlyPayment
            balance = round(balance - princ, 2)
            due = first_date + relativedelta(months=i - 1)
            rows.append((i, due.strftime('%b %d, %Y'),
                         '$' + f'{pmt:,.2f}',
                         '$' + f'{princ:,.2f}',
                         '$' + f'{interest:,.2f}',
                         '$' + f'{max(balance, 0):,.2f}'))
        return rows

    schedule       = build_schedule()
    total_repay    = round(sum(float(r[2].replace('$','').replace(',','')) for r in schedule), 2)
    total_interest = round(total_repay - data.loanAmount, 2)
    last_date      = (first_date + relativedelta(months=data.loanTermMonths - 1)).strftime('%B %d, %Y')

    # styles
    def S(name, **kw): return ParagraphStyle(name, **kw)
    sTitle  = S('sTitle',  fontSize=16, fontName='Helvetica-Bold', textColor=NAVY, alignment=TA_CENTER, spaceAfter=4)
    sSub    = S('sSub',    fontSize=9,  fontName='Helvetica', textColor=GREY, alignment=TA_CENTER)
    sHead   = S('sHead',   fontSize=10, fontName='Helvetica-Bold', textColor=WHITE)
    sNorm   = S('sNorm',   fontSize=9,  fontName='Helvetica', textColor=colors.black, leading=13, alignment=TA_JUSTIFY)
    sSmall  = S('sSmall',  fontSize=8,  fontName='Helvetica', textColor=GREY, leading=12, alignment=TA_JUSTIFY)
    sRef    = S('sRef',    fontSize=8,  fontName='Helvetica-Bold', textColor=NAVY, alignment=TA_RIGHT)
    sFooter = S('sFooter', fontSize=7.5, fontName='Helvetica', textColor=GREY, alignment=TA_CENTER)
    sTC     = S('sTC',     fontSize=8.5, fontName='Helvetica', textColor=colors.black, leading=12, alignment=TA_JUSTIFY)

    W = letter[0] - 1.3*inch

    def section_bar(title):
        t = Table([[Paragraph('  ' + title, sHead)]], colWidths=[W])
        t.setStyle(TableStyle([
            ('BACKGROUND', (0,0), (-1,-1), NAVY),
            ('TOPPADDING', (0,0), (-1,-1), 6),
            ('BOTTOMPADDING', (0,0), (-1,-1), 6),
            ('LEFTPADDING', (0,0), (-1,-1), 8),
        ]))
        return t

    def kv_table(rows, col1=1.8*inch):
        t = Table(rows, colWidths=[col1, W - col1])
        t.setStyle(TableStyle([
            ('FONTNAME', (0,0), (0,-1), 'Helvetica-Bold'),
            ('FONTNAME', (1,0), (1,-1), 'Helvetica'),
            ('FONTSIZE', (0,0), (-1,-1), 9),
            ('TEXTCOLOR', (0,0), (0,-1), NAVY),
            ('ROWBACKGROUNDS', (0,0), (-1,-1), [LIGHT, WHITE]),
            ('TOPPADDING', (0,0), (-1,-1), 5),
            ('BOTTOMPADDING', (0,0), (-1,-1), 5),
            ('LEFTPADDING', (0,0), (-1,-1), 8),
            ('GRID', (0,0), (-1,-1), 0.3, LGREY),
        ]))
        return t

    doc = SimpleDocTemplate(buffer, pagesize=letter,
                            rightMargin=0.65*inch, leftMargin=0.65*inch,
                            topMargin=0.5*inch, bottomMargin=0.65*inch)
    story = []

     # header
    logo_bytes = base64.b64decode(LOGO_B64)
    logo_stream = io.BytesIO(logo_bytes)
    aemo = Image(logo_stream, width=1.6*inch, height=0.75*inch)

    ht = Table([[aemo]], colWidths=[1.7*inch])
    ht.setStyle(TableStyle([
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('LEFTPADDING', (0,0), (-1,-1), 0),
        ('RIGHTPADDING', (0,0), (-1,-1), 0),
    ]))
    story += [ht, Spacer(1, 6),
              HRFlowable(width=W, thickness=2, color=NAVY),
              Spacer(1, 4),
              HRFlowable(width=W, thickness=1.5, color=RED),
              Spacer(1, 10)]

    

    # title
    story += [
        Paragraph('PERSONAL LOAN AGREEMENT', sTitle),
        Paragraph('Loan Terms &amp; Repayment Framework', sSub),
        Spacer(1, 6),
    ]

    ref_t = Table([
        [Paragraph('<b>Date:</b> ' + data.agreementDate, sSmall),
         Paragraph('<b>Reference No.:</b> ' + data.referenceNo, sRef)]
    ], colWidths=[W/2, W/2])
    ref_t.setStyle(TableStyle([
        ('LEFTPADDING', (0,0), (-1,-1), 0),
        ('RIGHTPADDING', (0,0), (-1,-1), 0),
    ]))
    story += [ref_t, Spacer(1, 10)]

    # borrower
    story += [section_bar('BORROWER DETAILS'), Spacer(1, 4),
              kv_table([
                  ['Full Name:', data.clientName],
                  ['Loan Type:', 'Personal Loan'],
                  ['Account Managed By:', 'AEMO Finance'],
              ]), Spacer(1, 10)]

    # loan summary
    ls_data = [
        ['Principal Loan Amount:',        '$' + f'{data.loanAmount:,.2f}'],
        ['Annual Percentage Rate (APR):', f'{data.annualRatePct:.2f}%'],
        ['Loan Term:',                    f'{data.loanTermMonths} Months'],
        ['Repayment Frequency:',          'Monthly'],
        ['Monthly Installment:',          '$' + f'{data.monthlyPayment:,.2f}'],
        ['Total Repayment Amount:',       '$' + f'{total_repay:,.2f}'],
        ['Total Interest Payable:',       '$' + f'{total_interest:,.2f}'],
    ]
    lt = Table(ls_data, colWidths=[2.8*inch, W - 2.8*inch])
    lt.setStyle(TableStyle([
        ('FONTNAME', (0,0), (0,-1), 'Helvetica-Bold'),
        ('FONTNAME', (1,0), (1,-1), 'Helvetica'),
        ('FONTSIZE', (0,0), (-1,-1), 9),
        ('TEXTCOLOR', (0,0), (0,-1), NAVY),
        ('FONTNAME', (1,4), (1,6), 'Helvetica-Bold'),
        ('TEXTCOLOR', (1,4), (1,6), RED),
        ('ROWBACKGROUNDS', (0,0), (-1,-1), [LIGHT, WHITE]),
        ('TOPPADDING', (0,0), (-1,-1), 5),
        ('BOTTOMPADDING', (0,0), (-1,-1), 5),
        ('LEFTPADDING', (0,0), (-1,-1), 8),
        ('GRID', (0,0), (-1,-1), 0.3, LGREY),
    ]))
    story += [section_bar('LOAN SUMMARY'), Spacer(1, 4), lt, Spacer(1, 10)]

    # repayment framework
    story += [section_bar('REPAYMENT FRAMEWORK'), Spacer(1, 5),
              Paragraph(
                  'Monthly deductions of <b>$' + f'{data.monthlyPayment:,.2f}' +
                  '</b> will be debited between the <b>28th - 30th of each month</b>, commencing '
                  '<b>' + first_date.strftime('%B %d, %Y') + '</b>, through to <b>' + last_date + '</b>. '
                  'All payments are processed through AEMO Finance with the full knowledge and approval of the client.',
                  sNorm),
              Spacer(1, 10)]

    # amortization
    col_w = [0.45*inch, 1.1*inch, 1.0*inch, 1.0*inch, 0.95*inch, 1.15*inch]
    amort_data = [['#', 'Due Date', 'Payment', 'Principal', 'Interest', 'Balance']] + \
                 [[str(r[0]), r[1], r[2], r[3], r[4], r[5]] for r in schedule]
    at = Table(amort_data, colWidths=col_w, repeatRows=1)
    row_bg = [('BACKGROUND', (0,i), (-1,i), LIGHT if i % 2 == 0 else WHITE)
              for i in range(1, len(amort_data))]
    at.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), NAVY),
        ('TEXTCOLOR', (0,0), (-1,0), WHITE),
        ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
        ('FONTSIZE', (0,0), (-1,0), 8),
        ('ALIGN', (0,0), (-1,0), 'CENTER'),
        ('FONTNAME', (0,1), (-1,-1), 'Helvetica'),
        ('FONTSIZE', (0,1), (-1,-1), 7.5),
        ('ALIGN', (0,1), (1,-1), 'CENTER'),
        ('ALIGN', (2,1), (-1,-1), 'RIGHT'),
        ('BACKGROUND', (0,-1), (-1,-1), colors.HexColor('#EAF0FF')),
        ('FONTNAME', (0,-1), (-1,-1), 'Helvetica-Bold'),
        ('TEXTCOLOR', (5,1), (5,-1), NAVY),
        ('TOPPADDING', (0,0), (-1,-1), 4),
        ('BOTTOMPADDING', (0,0), (-1,-1), 4),
        ('LEFTPADDING', (0,0), (-1,-1), 5),
        ('RIGHTPADDING', (0,0), (-1,-1), 5),
        ('GRID', (0,0), (-1,-1), 0.3, LGREY),
        ('LINEBELOW', (0,0), (-1,0), 1.5, RED),
    ] + row_bg))
    story += [section_bar('FULL AMORTIZATION SCHEDULE'), Spacer(1, 4), at, Spacer(1, 12)]

    # T&Cs
    tcs = [
        'The borrower agrees to repay the loan in ' + str(data.loanTermMonths) +
        ' equal monthly installments of $' + f'{data.monthlyPayment:,.2f}.',
        'Payments are due between the 28th - 30th of each month.',
        'Late payments may attract a penalty fee as determined by AEMO Finance.',
        'The borrower may request early repayment subject to applicable terms.',
        'This agreement is governed by applicable local laws.',
        'AEMO Finance reserves the right to amend terms with prior written notice to the borrower.',
    ]
    story += [section_bar('TERMS & CONDITIONS'), Spacer(1, 5)]
    for i, tc in enumerate(tcs, 1):
        story += [Paragraph(str(i) + '.  ' + tc, sTC), Spacer(1, 3)]
    story.append(Spacer(1, 8))

    # approval
    story += [section_bar('APPROVAL & AUTHORIZATION'), Spacer(1, 6),
              Paragraph('This loan has been reviewed and approved with the full knowledge and consent of the client.', sNorm),
              Spacer(1, 14)]

    sig_t = Table([
        [Paragraph('<b>Approved by:</b> AEMO Finance<br/>'
                   '<b>Date:</b> ' + data.agreementDate, sSmall),
         Paragraph('<b>Client:</b> ' + data.clientName + '<br/>'
                   '<b>Date:</b> ' + data.agreementDate, sSmall)],
        [Paragraph('____________________________<br/><i>Authorized Signature</i>', sSmall),
         Paragraph('____________________________<br/><i>Client Signature</i>', sSmall)],
    ], colWidths=[W/2 - 0.1*inch, W/2 - 0.1*inch])
    sig_t.setStyle(TableStyle([
        ('VALIGN', (0,0), (-1,-1), 'BOTTOM'),
        ('TOPPADDING', (0,0), (-1,-1), 6),
        ('LEFTPADDING', (0,0), (-1,-1), 0),
    ]))
    story += [sig_t, Spacer(1, 16),
              HRFlowable(width=W, thickness=1, color=LGREY),
              Spacer(1, 4),
              Paragraph('AEMO Finance  |  Tel: +297-588-0101', sFooter)]

    doc.build(story)
    buffer.seek(0)
    return buffer.read()