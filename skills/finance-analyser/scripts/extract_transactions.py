#!/usr/bin/env python3
"""
Universal bank/credit card PDF transaction extractor.
Auto-detects format from PDF content and applies the right parser.
Works with any bank worldwide — falls back to generic table extraction
if the specific bank format is not recognised.
"""
import argparse, re, sys
from datetime import datetime
from pathlib import Path
import pdfplumber
import pandas as pd

# ── Format detection ─────────────────────────────────────────────────────────

def detect_format(text):
    t = text.upper()
    # India
    if 'HDFC BANK' in t or 'HDFCBANK' in t:
        return 'hdfc_cc' if any(k in t for k in ['CREDIT CARD','CARD STATEMENT','REGALIA','RUPAY','MILLENNIA','MONEYBACK','INFINIA']) else 'hdfc_bank'
    if 'KOTAK' in t or 'AAACK4409' in t or 'KKBK' in t:
        return 'kotak_cc' if any(k in t for k in ['CREDIT CARD','CARD STATEMENT','TOTALAMOUNTDUE','MINIMUMAMOUNTDUE','PRIMARY CARD']) else 'kotak_bank'
    if 'ICICI BANK' in t:
        return 'icici_cc' if 'CREDIT CARD' in t else 'icici_bank'
    if 'STATE BANK OF INDIA' in t or ('SBI' in t and 'STATEMENT' in t):
        return 'sbi_bank'
    if 'AXIS BANK' in t:
        return 'axis_cc' if 'CREDIT CARD' in t else 'axis_bank'
    # UK
    if 'BARCLAYS' in t:
        return 'barclays'
    if 'HSBC' in t:
        return 'hsbc'
    if 'MONZO' in t:
        return 'monzo'
    if 'REVOLUT' in t:
        return 'revolut'
    if 'LLOYDS' in t:
        return 'lloyds'
    if 'NATWEST' in t or 'NATIONAL WESTMINSTER' in t:
        return 'natwest'
    # US
    if 'CHASE' in t and ('BANK' in t or 'CARD' in t):
        return 'chase'
    if 'BANK OF AMERICA' in t:
        return 'bofa'
    if 'WELLS FARGO' in t:
        return 'wells_fargo'
    if 'CITIBANK' in t or 'CITI BANK' in t:
        return 'citi'
    if 'AMERICAN EXPRESS' in t or 'AMEX' in t:
        return 'amex'
    # Australia
    if 'COMMONWEALTH BANK' in t or 'COMMBANK' in t:
        return 'commbank'
    if 'ANZ BANK' in t or 'AUSTRALIA AND NEW ZEALAND' in t:
        return 'anz'
    if 'WESTPAC' in t:
        return 'westpac'
    # Canada
    if 'ROYAL BANK' in t or 'RBC' in t:
        return 'rbc'
    if 'TD BANK' in t or 'TORONTO-DOMINION' in t:
        return 'td_canada'
    # UAE / Middle East
    if 'EMIRATES NBD' in t:
        return 'emirates_nbd'
    if 'MASHREQ' in t:
        return 'mashreq'
    return 'generic'

def detect_currency(text):
    """Detect predominant currency from statement text."""
    patterns = [
        (r'₹|INR|Rs\.', '₹'),
        (r'£|GBP', '£'),
        (r'€|EUR', '€'),
        (r'A\$|AUD', 'A$'),
        (r'S\$|SGD', 'S$'),
        (r'AED|Dhs', 'AED'),
        (r'CAD|C\$', 'C$'),
        (r'CHF', 'CHF'),
        (r'\$|USD', '$'),
    ]
    for pattern, symbol in patterns:
        if re.search(pattern, text):
            return symbol
    return None   # caller will fall back to format-based default

FORMAT_CURRENCY = {
    'hdfc_cc': '₹', 'hdfc_bank': '₹',
    'kotak_cc': '₹', 'kotak_bank': '₹',
    'icici_cc': '₹', 'icici_bank': '₹',
    'sbi_bank': '₹', 'axis_cc': '₹', 'axis_bank': '₹',
    'barclays': '£', 'hsbc': '£', 'monzo': '£',
    'revolut': '£', 'lloyds': '£', 'natwest': '£',
    'commbank': 'A$', 'anz': 'A$', 'westpac': 'A$',
    'rbc': 'C$', 'td_canada': 'C$',
    'emirates_nbd': 'AED', 'mashreq': 'AED',
    'chase': '$', 'bofa': '$', 'wells_fargo': '$',
    'citi': '$', 'amex': '$',
    'generic': '$',
}

# ── Parsers ──────────────────────────────────────────────────────────────────

def skip_line(desc):
    """Return True if line is a payment/credit/internal entry to skip."""
    return any(k in desc.upper() for k in [
        'AUTOPAY','THANK YOU','PAYMENT RECEIVED','AUTODEBIT',
        'CARD DUES','OPENING BALANCE','CLOSING BALANCE',
        'NACH-MUT','NACH MUT','MUTUAL FUND','NEFT SCBLH',
        'TRANSFER TO OWN','REFUND',
    ])

def parse_hdfc_cc(pdf, source):
    txns = []
    for page in pdf.pages:
        for line in (page.extract_text() or '').split('\n'):
            m = re.match(r'(\d{2}/\d{2}/\d{4})\|\s*\d{2}:\d{2}\s+(.+?)\s+[+\-]?\s*[Cc]?\s*([\d,]+\.\d{2})\s*[l|]?\s*$', line.strip())
            if m:
                date_str, desc, amt = m.group(1), m.group(2).strip(), m.group(3)
                if skip_line(desc): continue
                try:
                    txns.append({'Date': datetime.strptime(date_str,'%d/%m/%Y'), 'Description': desc,
                                 'Amount': float(amt.replace(',','')), 'Type': 'Debit', 'Source': source})
                except: pass
    return txns

def parse_kotak_cc(pdf, source):
    txns = []
    for page in pdf.pages:
        for line in (page.extract_text() or '').split('\n'):
            m = re.match(r'(\d{2}/\d{2}/\d{4})\s+(.+?)\s+(?:[A-Za-z\s&/]+)\s+([\d,]+\.\d{2})\s*(Cr)?\s*$', line.strip())
            if m and not m.group(4) and not skip_line(m.group(2)):
                try:
                    txns.append({'Date': datetime.strptime(m.group(1),'%d/%m/%Y'), 'Description': m.group(2).strip(),
                                 'Amount': float(m.group(3).replace(',','')), 'Type': 'Debit', 'Source': source})
                except: pass
    return txns

def parse_kotak_bank(pdf, source):
    full_text = "".join(p.extract_text() or '' for p in pdf.pages)
    txns = []; prev_bal = None
    for line in full_text.split('\n'):
        ob = re.search(r'Opening\s*Balance.*?([\d,]+\.\d{2})', line, re.IGNORECASE)
        if ob and prev_bal is None:
            prev_bal = float(ob.group(1).replace(',','')); continue
        m = re.match(r'^\d+\s+(\d{2}\s+\w{3}\s+\d{4})\s+(.+?)\s+([\w\-\/]+)\s+([\d,]+\.\d{2})\s+([\d,]+\.\d{2})\s*$', line.strip())
        if not m or prev_bal is None: continue
        desc, amount, balance = m.group(2).strip(), float(m.group(4).replace(',','')), float(m.group(5).replace(',',''))
        is_debit = abs(prev_bal - amount - balance) < 2
        prev_bal = balance
        if is_debit and not skip_line(desc):
            try:
                txns.append({'Date': datetime.strptime(m.group(1),'%d %b %Y'), 'Description': desc,
                             'Amount': amount, 'Type': 'Debit', 'Source': source})
            except: pass
    return txns

def parse_generic_table(pdf, source):
    """
    Fallback: use pdfplumber table extraction.
    Looks for columns containing Date, Description/Narration, and Debit/Amount.
    Works for most structured bank PDFs worldwide.
    """
    txns = []
    date_fmts = ['%d/%m/%Y','%d-%m-%Y','%d %b %Y','%d-%b-%Y','%d/%m/%y',
                 '%m/%d/%Y','%Y-%m-%d','%b %d, %Y','%d %B %Y']
    for page in pdf.pages:
        for table in (page.extract_tables() or []):
            if not table or len(table) < 2: continue
            header = [str(h).upper().strip().replace('\n',' ') if h else '' for h in table[0]]
            date_col  = next((i for i,h in enumerate(header) if re.search(r'DATE|DT\b', h)), None)
            desc_col  = next((i for i,h in enumerate(header) if re.search(r'DESC|NARR|PARTI|REMARK|DETAIL|PARTICULARS|TRANSACTION', h)), None)
            debit_col = next((i for i,h in enumerate(header) if re.search(r'DEBIT|DR\b|WITHDRAWAL|AMOUNT|PAID OUT|SPEND', h)), None)
            credit_col= next((i for i,h in enumerate(header) if re.search(r'CREDIT|CR\b|DEPOSIT|PAID IN', h)), None)
            if date_col is None or desc_col is None or debit_col is None: continue
            for row in table[1:]:
                try:
                    date_raw  = str(row[date_col] or '').strip()
                    desc      = str(row[desc_col] or '').strip()
                    debit_raw = str(row[debit_col] or '').strip()
                    credit_raw= str(row[credit_col] or '').strip() if credit_col is not None else ''
                    if not date_raw or not desc or not debit_raw: continue
                    if debit_raw in ('-','','None','0'): continue
                    if credit_raw and credit_raw not in ('-','','None','0'): continue  # skip credits
                    amt = float(re.sub(r'[^\d.]', '', debit_raw))
                    if amt <= 0: continue
                    date = None
                    for fmt in date_fmts:
                        try: date = datetime.strptime(date_raw.split()[0], fmt); break
                        except: pass
                    if not date: continue
                    if skip_line(desc): continue
                    txns.append({'Date': date, 'Description': desc,
                                 'Amount': amt, 'Type': 'Debit', 'Source': source})
                except: pass
    return txns

# Specific parsers for international banks (all use generic table as backbone)
def parse_with_generic(pdf, source): return parse_generic_table(pdf, source)

PARSERS = {
    'hdfc_cc':    parse_hdfc_cc,
    'kotak_cc':   parse_kotak_cc,
    'kotak_bank': parse_kotak_bank,
    # All others use generic table extraction — works well for most structured PDFs
}

# ── Main ─────────────────────────────────────────────────────────────────────

def extract_from_pdf(path, password=None):
    try:
        with pdfplumber.open(path, password=password or '') as pdf:
            first_text = pdf.pages[0].extract_text() or ''
            fmt = detect_format(first_text)
            currency = detect_currency(first_text) or FORMAT_CURRENCY.get(fmt, '$')
            source = f"{path.stem} ({fmt.replace('_',' ').title()})"
            parser = PARSERS.get(fmt, parse_generic_table)
            if fmt not in PARSERS:
                print(f"    ℹ️  Using generic table parser for {fmt}")
            txns = parser(pdf, source)
            return txns, fmt, currency
    except Exception as e:
        if any(k in str(e).lower() for k in ['password','incorrect','encrypt']):
            return None, 'password_protected', None
        print(f"    ⚠️  Error: {e}", file=sys.stderr)
        return [], 'error', None

def main():
    parser = argparse.ArgumentParser(description='Universal bank PDF transaction extractor')
    parser.add_argument('--input-dir', required=True, help='Folder containing PDF statements')
    parser.add_argument('--output', required=True, help='Output CSV path')
    parser.add_argument('--password', default=None, help='PDF password if protected')
    parser.add_argument('--start-date', default=None)
    parser.add_argument('--end-date', default=None)
    args = parser.parse_args()

    pdfs = sorted(Path(args.input_dir).rglob('*.pdf'))
    if not pdfs:
        print(f"❌ No PDFs found in {args.input_dir}", file=sys.stderr); sys.exit(1)

    print(f"📂 Found {len(pdfs)} PDF(s)\n")
    all_txns, skipped, currencies = [], [], []

    for p in pdfs:
        print(f"  📄 {p.name}")
        txns, fmt, currency = extract_from_pdf(p, args.password)
        if fmt == 'password_protected':
            print(f"    🔒 Password-protected — skipped. Re-run with --password"); skipped.append(p.name); continue
        if fmt == 'error' or txns is None:
            skipped.append(p.name); continue
        print(f"    ✓ {len(txns)} transactions | format: {fmt} | currency: {currency}")
        for t in txns:
            t['Currency'] = currency
        all_txns.extend(txns)
        if currency: currencies.append(currency)

    if not all_txns:
        print("❌ No transactions extracted.", file=sys.stderr); sys.exit(1)

    df = pd.DataFrame(all_txns)
    df['Date'] = pd.to_datetime(df['Date'])
    df = df.drop_duplicates(subset=['Date','Amount','Description','Source'])
    if args.start_date: df = df[df['Date'] >= args.start_date]
    if args.end_date:   df = df[df['Date'] <= args.end_date]
    df = df.sort_values('Date').reset_index(drop=True)
    df.to_csv(args.output, index=False)

    dominant_currency = max(set(currencies), key=currencies.count) if currencies else '$'
    print(f"\n{'='*55}")
    print(f"✅ {len(df)} transactions saved → {args.output}")
    print(f"   Dominant currency: {dominant_currency}")
    if skipped: print(f"   ⚠️  Skipped: {', '.join(skipped)}")

if __name__ == '__main__':
    main()
