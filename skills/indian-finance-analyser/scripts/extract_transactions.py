#!/usr/bin/env python3
"""
Universal Indian bank/credit card PDF transaction extractor.
Auto-detects format from PDF content and applies the right parser.
Supports: HDFC CC, HDFC Bank, Kotak CC, Kotak Bank, ICICI CC, ICICI Bank,
          SBI, Axis Bank, Yes Bank, and a generic table fallback.
"""
import argparse, re, sys
from datetime import datetime
from pathlib import Path
import pdfplumber
import pandas as pd

# ── Format detection ────────────────────────────────────────────────────────

def detect_format(first_page_text):
    """Identify bank and statement type from first-page text."""
    t = first_page_text.upper()
    if 'HDFC BANK' in t or 'HDFCBANK' in t:
        if any(k in t for k in ['CREDIT CARD', 'CARD STATEMENT', 'REGALIA', 'MILLENNIA',
                                  'RUPAY', 'MONEYBACK', 'DINERS', 'INFINIA']):
            return 'hdfc_cc'
        return 'hdfc_bank'
    if 'KOTAK' in t or 'AAACK4409' in t or 'KKBK' in t:
        if any(k in t for k in ['CREDIT CARD', 'CARD STATEMENT', 'BILLED',
                                 'TOTALAMOUNTDUE', 'MINIMUMAMOUNT', 'MINIMUMAMOUNTDUE',
                                 'TOTALAMOUNT', 'CARDNUMBER', 'PRIMARY CARD']):
            return 'kotak_cc'
        return 'kotak_bank'
    if 'ICICI BANK' in t or 'ICICIDIRECT' in t:
        if any(k in t for k in ['CREDIT CARD', 'CARD STATEMENT']):
            return 'icici_cc'
        return 'icici_bank'
    if 'STATE BANK OF INDIA' in t or 'SBI' in t:
        return 'sbi_bank'
    if 'AXIS BANK' in t:
        if 'CREDIT CARD' in t:
            return 'axis_cc'
        return 'axis_bank'
    if 'YES BANK' in t:
        return 'yes_bank'
    if 'INDUSIND' in t:
        return 'indusind_cc'
    if 'FEDERAL BANK' in t:
        return 'federal_bank'
    if 'AU SMALL FINANCE' in t or 'AU BANK' in t:
        return 'au_bank'
    if 'RBL BANK' in t:
        return 'rbl_cc'
    if 'PAYTM' in t:
        return 'paytm_cc'
    return 'generic'

# ── Individual parsers ──────────────────────────────────────────────────────

def parse_hdfc_cc(pdf, source_name):
    """HDFC Credit Card: DD/MM/YYYY| HH:MM  DESCRIPTION  AMOUNT"""
    txns = []
    for page in pdf.pages:
        text = page.extract_text() or ''
        for line in text.split('\n'):
            m = re.match(
                r'(\d{2}/\d{2}/\d{4})\|\s*\d{2}:\d{2}\s+(.+?)\s+[+\-]?\s*[Cc]?\s*([\d,]+\.\d{2})\s*[l|]?\s*$',
                line.strip())
            if m:
                date_str, desc, amt = m.group(1), m.group(2).strip(), m.group(3)
                if any(k in desc.upper() for k in ['AUTOPAY','THANK YOU','PAYMENT RECEIVED']): continue
                try:
                    txns.append({'Date': datetime.strptime(date_str, '%d/%m/%Y'),
                                 'Description': desc.strip(),
                                 'Amount': float(amt.replace(',', '')),
                                 'Type': 'Debit', 'Source': source_name})
                except: pass
    return txns

def parse_kotak_cc(pdf, source_name):
    """Kotak Credit Card: DD/MM/YYYY  MERCHANT  CATEGORY  AMOUNT [Cr]"""
    txns = []
    for page in pdf.pages:
        text = page.extract_text() or ''
        for line in text.split('\n'):
            m = re.match(r'(\d{2}/\d{2}/\d{4})\s+(.+?)\s+(?:[A-Za-z\s&/]+)\s+([\d,]+\.\d{2})\s*(Cr)?\s*$', line.strip())
            if m:
                date_str, desc, amt, cr = m.group(1), m.group(2).strip(), m.group(3), m.group(4)
                if any(k in desc.upper() for k in ['PAYMENT RECEIVED','AUTODEBIT','REFUND']): continue
                if cr: continue  # credit line
                try:
                    txns.append({'Date': datetime.strptime(date_str, '%d/%m/%Y'),
                                 'Description': desc,
                                 'Amount': float(amt.replace(',', '')),
                                 'Type': 'Debit', 'Source': source_name})
                except: pass
    return txns

def parse_kotak_bank(pdf, source_name):
    """Kotak Bank: numbered rows with withdrawal/deposit/balance columns."""
    full_text = "".join(p.extract_text() or '' for p in pdf.pages)
    txns = []
    prev_balance = None
    for line in full_text.split('\n'):
        line = line.strip()
        # Opening balance
        ob = re.search(r'Opening\s*Balance.*?([\d,]+\.\d{2})', line, re.IGNORECASE)
        if ob and prev_balance is None:
            prev_balance = float(ob.group(1).replace(',', ''))
            continue
        m = re.match(r'^\d+\s+(\d{2}\s+\w{3}\s+\d{4})\s+(.+?)\s+([\w\-\/]+)\s+([\d,]+\.\d{2})\s+([\d,]+\.\d{2})\s*$', line)
        if not m or prev_balance is None: continue
        date_str, desc, ref = m.group(1), m.group(2).strip(), m.group(3)
        amount, balance = float(m.group(4).replace(',','')), float(m.group(5).replace(',',''))
        try:
            date = datetime.strptime(date_str, '%d %b %Y')
        except: continue
        is_debit = abs(prev_balance - amount - balance) < 2
        is_credit = abs(prev_balance + amount - balance) < 2
        prev_balance = balance
        if is_debit:
            # Skip obvious non-spending debits generically
            skip_patterns = [
                'CARD DUES', 'AUTOPAY', 'NEFT SCBLH',  # CC payments, salary-related
                'NACH-MUT', 'NACH MUT', 'MUTUAL FUND',  # Investments
                'NACH-ECS-DR',                           # ECS debits (loans etc)
            ]
            if any(k in (desc+ref).upper() for k in skip_patterns): continue
            txns.append({'Date': date, 'Description': desc, 'Amount': amount,
                         'Type': 'Debit', 'Source': source_name})
    return txns

def parse_icici_cc(pdf, source_name):
    """ICICI Credit Card: tabular format DD/MM/YYYY  DESCRIPTION  AMOUNT"""
    txns = []
    for page in pdf.pages:
        text = page.extract_text() or ''
        for line in text.split('\n'):
            # ICICI: date, description, amount (Dr/Cr)
            m = re.match(r'(\d{2}/\d{2}/\d{4})\s+(.+?)\s+([\d,]+\.\d{2})\s*(Dr|Cr)?\s*$', line.strip())
            if m:
                date_str, desc, amt, dr_cr = m.group(1), m.group(2).strip(), m.group(3), m.group(4)
                if dr_cr == 'Cr': continue
                if any(k in desc.upper() for k in ['PAYMENT','REFUND','REVERSAL']): continue
                try:
                    txns.append({'Date': datetime.strptime(date_str, '%d/%m/%Y'),
                                 'Description': desc, 'Amount': float(amt.replace(',','')),
                                 'Type': 'Debit', 'Source': source_name})
                except: pass
    return txns

def parse_bank_generic_debit_credit(pdf, source_name, date_fmt='%d/%m/%Y'):
    """
    Generic parser for tabular bank statements.
    Looks for rows with: DATE  DESCRIPTION  DEBIT  CREDIT  BALANCE
    Extracts only debits.
    """
    txns = []
    for page in pdf.pages:
        # Try table extraction first
        tables = page.extract_tables()
        for table in tables:
            if not table or len(table) < 2: continue
            # Detect columns
            header = [str(h).upper().strip() if h else '' for h in table[0]]
            date_col = next((i for i,h in enumerate(header) if 'DATE' in h), None)
            desc_col = next((i for i,h in enumerate(header) if any(k in h for k in ['DESC','NARR','PARTI','REMARK','DETAIL'])), None)
            debit_col = next((i for i,h in enumerate(header) if any(k in h for k in ['DEBIT','DR','WITHDRAWAL','PAID'])), None)
            credit_col = next((i for i,h in enumerate(header) if any(k in h for k in ['CREDIT','CR','DEPOSIT','RECEIVED'])), None)
            if date_col is None or desc_col is None: continue
            for row in table[1:]:
                try:
                    date_raw = str(row[date_col]).strip() if row[date_col] else ''
                    desc = str(row[desc_col]).strip() if row[desc_col] else ''
                    debit_raw = str(row[debit_col]).strip() if debit_col is not None and row[debit_col] else ''
                    credit_raw = str(row[credit_col]).strip() if credit_col is not None and row[credit_col] else ''
                    if not date_raw or not desc or date_raw in ('', 'None'): continue
                    # Parse amount — debit column takes priority
                    amt_str = debit_raw or ''
                    if not amt_str or amt_str in ('', '-', '0', 'None'): continue
                    amt = float(re.sub(r'[^\d.]', '', amt_str))
                    if amt <= 0: continue
                    # Parse date — try multiple formats
                    date = None
                    for fmt in ['%d/%m/%Y','%d-%m-%Y','%d %b %Y','%d-%b-%Y','%d/%m/%y','%d-%m-%y']:
                        try: date = datetime.strptime(date_raw.split()[0], fmt); break
                        except: pass
                    if not date: continue
                    if any(k in desc.upper() for k in ['OPENING BAL','CLOSING BAL','TRANSFER TO OWN']): continue
                    txns.append({'Date': date, 'Description': desc, 'Amount': amt,
                                 'Type': 'Debit', 'Source': source_name})
                except: pass
    return txns

def parse_sbi(pdf, source_name):
    """SBI: Dr/Cr column or separate debit column."""
    txns = []
    for page in pdf.pages:
        text = page.extract_text() or ''
        for line in text.split('\n'):
            # SBI format: DD/MM/YYYY  DESCRIPTION  REF  DEBIT  CREDIT  BALANCE
            m = re.match(r'(\d{2}/\d{2}/\d{4})\s+(.+?)\s+([\d,]+\.\d{2})\s+([\d,]+\.\d{2})?\s+([\d,]+\.\d{2})\s*$', line.strip())
            if m:
                date_str, desc = m.group(1), m.group(2).strip()
                debit = m.group(3)
                credit = m.group(4)
                if credit: continue  # If credit column has value, it's incoming
                if any(k in desc.upper() for k in ['BY TRANSFER','INTEREST','OPENING']): continue
                try:
                    txns.append({'Date': datetime.strptime(date_str,'%d/%m/%Y'),
                                 'Description': desc, 'Amount': float(debit.replace(',','')),
                                 'Type': 'Debit', 'Source': source_name})
                except: pass
    # If text parsing yielded nothing, try generic table parser
    if not txns:
        txns = parse_bank_generic_debit_credit(pdf, source_name)
    return txns

# ── Dispatcher ──────────────────────────────────────────────────────────────

PARSERS = {
    'hdfc_cc':    parse_hdfc_cc,
    'kotak_cc':   parse_kotak_cc,
    'kotak_bank': parse_kotak_bank,
    'icici_cc':   parse_icici_cc,
    'sbi_bank':   parse_sbi,
}

def extract_from_pdf(filepath, password=None):
    name = filepath.stem
    try:
        with pdfplumber.open(filepath, password=password or '') as pdf:
            first_text = pdf.pages[0].extract_text() or ''
            fmt = detect_format(first_text)
            print(f"    Format detected: {fmt}")
            source_name = f"{name} ({fmt.replace('_', ' ').title()})"
            parser = PARSERS.get(fmt)
            if parser:
                txns = parser(pdf, source_name)
            else:
                # Generic fallback
                print(f"    Using generic table parser")
                txns = parse_bank_generic_debit_credit(pdf, source_name)
            return txns, fmt
    except Exception as e:
        err = str(e).lower()
        if 'password' in err or 'incorrect' in err or 'encrypt' in err:
            return None, 'password_protected'
        print(f"    ⚠️  Error: {e}", file=sys.stderr)
        return [], 'error'

# ── Main ────────────────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(description='Universal Indian bank PDF extractor')
    parser.add_argument('--input-dir', required=True, help='Folder containing PDF statements')
    parser.add_argument('--output', required=True, help='Output CSV path')
    parser.add_argument('--password', default=None, help='PDF password (if all PDFs share one)')
    parser.add_argument('--start-date', default=None, help='Filter from date (YYYY-MM-DD)')
    parser.add_argument('--end-date', default=None, help='Filter to date (YYYY-MM-DD)')
    args = parser.parse_args()

    input_dir = Path(args.input_dir)
    pdfs = sorted(input_dir.rglob('*.pdf'))
    if not pdfs:
        print(f"❌ No PDFs found in {input_dir}", file=sys.stderr)
        sys.exit(1)

    print(f"📂 Found {len(pdfs)} PDF(s) in {input_dir}\n")
    all_txns = []
    skipped = []

    for pdf_path in pdfs:
        print(f"  📄 {pdf_path.name}")
        txns, fmt = extract_from_pdf(pdf_path, args.password)
        if fmt == 'password_protected':
            print(f"    🔒 Password-protected — skipped. Re-run with --password")
            skipped.append(pdf_path.name)
            continue
        if txns is None or fmt == 'error':
            skipped.append(pdf_path.name)
            continue
        print(f"    ✓ {len(txns)} transactions extracted")
        all_txns.extend(txns)

    if not all_txns:
        print("❌ No transactions extracted.", file=sys.stderr)
        sys.exit(1)

    df = pd.DataFrame(all_txns)
    df['Date'] = pd.to_datetime(df['Date'])
    df = df.drop_duplicates(subset=['Date','Amount','Description','Source'])
    if args.start_date:
        df = df[df['Date'] >= args.start_date]
    if args.end_date:
        df = df[df['Date'] <= args.end_date]
    df = df.sort_values('Date').reset_index(drop=True)
    df.to_csv(args.output, index=False)

    print(f"\n{'='*50}")
    print(f"✅ Extracted {len(df)} transactions → {args.output}")
    if skipped:
        print(f"⚠️  Skipped {len(skipped)} file(s): {', '.join(skipped)}")

if __name__ == '__main__':
    main()
