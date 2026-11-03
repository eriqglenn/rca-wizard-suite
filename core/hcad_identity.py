"""Shared HCAD property identity for workbook exports and downloads."""
import io
import math
import re
from decimal import Decimal, InvalidOperation


def clean_account(value):
    """Normalize valid 13-digit accounts, including one lost leading zero."""
    if isinstance(value, float):
        if not math.isfinite(value) or not value.is_integer():return ''
        value=int(value)
    text=str(value or '').strip()
    if len(text)>100 or text.startswith('-'):return ''
    if re.fullmatch(r'\d+(?:\.\d+)?[eE][+]?\d+|\d+\.0+',text):
        try:
            number=Decimal(text)
            if not number.is_finite() or number.adjusted()>12 or number.adjusted()<0:return ''
            if number!=number.to_integral_value():return ''
            text=str(int(number))
        except (InvalidOperation,ValueError,OverflowError):return ''
    digits=re.sub(r'\D+','',text)
    if len(digits)==12:digits='0'+digits
    return digits if len(digits)==13 else ''


def collect_hcad_records(path):
    """One record per normalized account, with address fallback when unavailable."""
    from openpyxl import load_workbook
    from .workbooks import read_shared_bytes
    book=load_workbook(io.BytesIO(read_shared_bytes(path)),data_only=True,read_only=True)
    try:
        sheet='Step 1 - Input Properties Info'
        if sheet not in book.sheetnames:raise ValueError(f"Workbook must contain the '{sheet}' tab.")
        ws=book[sheet]
        heads={str(v or '').strip():i for i,v in enumerate(next(ws.iter_rows(max_row=1,values_only=True)))}
        records=[];seen=set()
        for row in ws.iter_rows(min_row=2,values_only=True):
            def value(name):
                index=heads.get(name)
                return row[index] if index is not None and index<len(row) else None
            def text(name):
                raw=value(name)
                if raw is None:return ''
                result=str(raw).strip()
                return '' if result.casefold() in ('none','null','nan','<null>') else result
            account=next((a for field in ('TAX_ID','ALT_TAX_ID','HCAD_NUM','HCAD_NUM_1') if (a:=clean_account(value(field)))),'')
            address=text('Site_addr_1') or text('Property Description')
            unit=text('str_unit')
            if address and unit and unit.upper() not in address.upper().split():address+=' '+unit
            address=re.sub(r'\s+',' ',address).strip(' ,').upper()
            if not account and not address:continue
            key=('account',account) if account else ('address',address)
            if key in seen:continue
            seen.add(key);records.append({'Address':address,'TAX_ID':account})
        return records
    finally:book.close()
