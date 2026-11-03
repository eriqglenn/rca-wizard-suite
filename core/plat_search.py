"""Recorded subdivision-plat search using the County Clerk's public form."""
from __future__ import annotations

from dataclasses import asdict, dataclass
from datetime import datetime
from http.cookiejar import CookieJar
import io
import json
from pathlib import Path
import re
import ssl
import threading
from urllib.parse import urlencode, urljoin, urlsplit
from urllib.request import build_opener, HTTPSHandler, HTTPCookieProcessor, Request

from lxml import html
from openpyxl import load_workbook

SEARCH_URL = 'https://www.cclerk.hctx.net/Applications/WebSearch/MC.aspx'
LOGIN_URL = 'https://www.cclerk.hctx.net/Applications/WebSearch/Registration/Login.aspx'
CLERK_HOSTS = {'www.cclerk.hctx.net', 'cclerk.hctx.net'}


class PlatError(ValueError):
    pass


def clean(value):
    return re.sub(r'\s+', ' ', str(value or '')).strip()


def clerk_url(value, base=SEARCH_URL):
    url = urljoin(base, value)
    parts = urlsplit(url)
    if parts.scheme != 'https' or parts.hostname not in CLERK_HOSTS or parts.username or parts.password:
        raise PlatError('The County returned an unexpected document link. Open the County search website to check it.')
    return url


@dataclass(frozen=True)
class PlatRecord:
    file_number: str
    name: str
    file_date: str
    record_type: str
    map_pages: int
    volume_page: str
    viewer_url: str


def parse_results(content, url):
    tree = html.fromstring(content)
    page_text = clean(tree.text_content())
    records = []
    for number in tree.xpath('//span[substring(@id,string-length(@id)-9)="_lblFileNo"]'):
        row = number.xpath('ancestor::tr[1]')[0]
        def value(suffix):
            found = row.xpath('.//*[@id]')
            return next((clean(e.text_content()) for e in found if e.get('id','').endswith(suffix)), '')
        links = row.xpath('.//a[@href]')
        link = next((e for e in links if e.get('id','').endswith('_HyperLinkVolNoEC') and clean(e.text_content())), None)
        # Film-code links are separate instruments; do not substitute them for the plat.
        if link is None:
            viewer = ''
        else:
            viewer = clerk_url(link.get('href'), url)
        pages = value('_lblMapPgs')
        records.append(PlatRecord(clean(number.text_content()), value('_lblDesc'), value('_lblFileDate'),
                                  value('_lblIType'), int(pages) if pages.isdigit() else 0,
                                  clean(link.text_content()) if link is not None else '', viewer))
    count = re.search(r'([\d,]+)\s+Record\(s\)\s+Found', page_text, re.I)
    total = int(count[1].replace(',','')) if count else len(records)
    if not records and not count:
        if re.search(r'no (?:records|results)|0 record', page_text, re.I):
            total = 0
        else:
            raise PlatError('The County did not return recognizable search results. Try again or open the County website.')
    return {'records': [asdict(r) for r in records], 'total': total, 'truncated': total > len(records), 'url': url}


class PlatSearchClient:
    def __init__(self, cancel=None):
        self.cancel = cancel or threading.Event()
        self.opener = None

    def request(self, url, fields=None):
        if self.cancel.is_set():
            raise PlatError('Plat search cancelled.')
        url = clerk_url(url)
        if self.opener is None:
            try:
                from .map_connection import map_ssl_context
            except ModuleNotFoundError:
                context = ssl.create_default_context()
            else:
                context = map_ssl_context()
            self.opener = build_opener(HTTPSHandler(context=context), HTTPCookieProcessor(CookieJar()))
        body = urlencode(fields).encode() if fields is not None else None
        try:
            with self.opener.open(Request(url, data=body, headers={'User-Agent': 'Houston-RCA-Plat-Finder/1.0'}), timeout=45) as response:
                final = clerk_url(response.url)
                content = response.read(8*1024*1024+1)
        except OSError as exc:
            raise PlatError(f'Could not reach the County plat search. Check the connection and try again.\n{exc}') from exc
        if self.cancel.is_set():
            raise PlatError('Plat search cancelled.')
        if len(content) > 8*1024*1024:
            raise PlatError('This search returned too much data. Use a more specific subdivision name, file number, or volume/page.')
        return content, final

    def search(self, name='', file_number='', volume='', page=''):
        name, file_number, volume, page = map(clean, (name, file_number, volume, page))
        if not any((name, file_number, volume, page)):
            raise PlatError('Enter a subdivision name, file number, or volume and page.')
        if bool(volume) != bool(page):
            raise PlatError('Enter both the volume and page, or leave both blank.')
        if volume and (not volume.isdigit() or len(volume)>4 or not page.isdigit() or len(page)>3):
            raise PlatError('Use numeric volume (up to 4 digits) and page (up to 3 digits).')
        if len(name)>180 or len(file_number)>30:
            raise PlatError('Shorten the subdivision name or file number and try again.')
        content, url = self.request(SEARCH_URL)
        tree = html.fromstring(content)
        fields = {e.get('name'): e.get('value','') for e in tree.xpath('//input[@type="hidden" and @name]')}
        if not tree.xpath('//input[@name="ctl00$ContentPlaceHolder1$btnSearch"]'):
            raise PlatError('The County search form has changed. Open the County website to search.')
        prefix = 'ctl00$ContentPlaceHolder1$'
        fields.update({prefix+'txtDesc': name, prefix+'txtFileNo': file_number,
                       prefix+'txtVolNo': volume, prefix+'txtPageNo': page, prefix+'btnSearch': 'Search'})
        # Match the page-number input actually advertised by the public form.
        page_fields = tree.xpath('//input[@aria-label="Page"]/@name')
        if page_fields:
            fields.pop(prefix+'txtPageNo', None)
            fields[page_fields[0]] = page
        content, url = self.request(url, fields)
        result = parse_results(content, url)
        result['query'] = {'name': name, 'file_number': file_number, 'volume': volume, 'page': page}
        result['retrieved_at'] = datetime.now().astimezone().isoformat()
        return result


def search_name(name):
    """Suggest a broader name; keep the full worksheet wording visible in the UI."""
    name = clean(name).upper()
    name = re.split(r'\s+(?:SEC(?:TION)?\.?\s*\d|\d+(?:ST|ND|RD|TH)\s+PAR|PAR(?:TIAL)?\s+R/P|R/P\b|REPLAT\b)', name, maxsplit=1)[0]
    return name


def read_plat_context(path):
    """Read subdivision names without requiring a completed MLS/MBL calculation."""
    result = {'names': [], 'reference': '', 'warnings': []}
    if not path or not Path(path).is_file():
        result['warnings'].append('Choose a calculation sheet to suggest subdivision names, or type a name below.')
        return result
    from .workbooks import WorkbookReadCache
    resolved = WorkbookReadCache().resolve(path)
    wb = load_workbook(resolved, read_only=True, data_only=True)
    try:
        names = {}
        step1 = next((s for s in wb if s.title.lower().startswith('step 1')), None)
        if step1 is not None:
            rows = step1.iter_rows(values_only=True)
            headers = [re.sub(r'[^a-z0-9]', '', clean(v).lower()) for v in next(rows, [])]
            columns = [i for i,k in enumerate(headers) if k in {'legaldscr2','subdivision','subdivisionname','subdivname'}]
            for row in rows:
                for i in columns:
                    value = clean(row[i]) if i<len(row) else ''
                    if value and not value.startswith(('=', '#')) and value not in {'0','None'}:
                        names[value] = names.get(value,0)+1
        ci = next((s for s in wb if re.sub(r'[^a-z]','',s.title.lower())=='caseinfo'), None)
        if ci is not None:
            for row in ci.iter_rows(values_only=True):
                if len(row)<2: continue
                key = re.sub(r'[^a-z0-9]','',clean(row[0]).lower())
                value = clean(row[1])
                if key in {'subdivision','subdivisionname'} and value:
                    names.setdefault(value,0)
                if key in {'casenumber','referencenumber','filenumber'} and value:
                    result['reference'] = value
        result['names'] = [{'name':name, 'properties':count, 'search_name':search_name(name)} for name,count in names.items()]
        if not names:
            result['warnings'].append('No subdivision name was found in the worksheet. Enter its name or recording volume/page below.')
        if len(names)>1:
            result['warnings'].append('The worksheet lists multiple subdivision names or replats. Review each applicable recording; more than one plat may be needed.')
    finally:
        wb.close()
    return result


def validated_pdf(content):
    import fitz
    try:
        if not content.lstrip().startswith(b'%PDF-'):
            raise PlatError('The download is not a PDF.')
        with fitz.open(stream=content, filetype='pdf') as doc:
            if not doc.is_pdf or doc.needs_pass or len(doc)<1:
                raise PlatError('The downloaded plat is empty or password protected.')
            return len(doc)
    except Exception as exc:
        raise PlatError('The download is not a readable PDF plat. Use the County viewer to download a PDF and choose Add Downloaded Plat.') from exc


def save_plats(items, output_dir):
    """Combine only planner-selected recorded PDFs; retain each page unchanged."""
    import fitz
    if not items:
        raise PlatError('Download and review at least one plat first.')
    folder = Path(output_dir)
    folder.mkdir(parents=True, exist_ok=True)
    stem = 'Subdivision_Plats_'+datetime.now().strftime('%Y%m%d_%H%M%S_%f')
    target, meta = folder/(stem+'.pdf'), folder/(stem+'.plats.json')
    for item in items:
        validated_pdf(item['pdf'])
    if len(items)==1:
        content = items[0]['pdf']
    else:
        with fitz.open() as combined:
            for item in items:
                with fitz.open(stream=item['pdf'],filetype='pdf') as doc:
                    combined.insert_pdf(doc)
            content = combined.tobytes(garbage=3,deflate=True)
    record = {'created_at':datetime.now().astimezone().isoformat(), 'source':SEARCH_URL,
              'plats':[{k:v for k,v in item.items() if k!='pdf'} for item in items]}
    written = []
    try:
        for path, data in ((meta,json.dumps(record,indent=2).encode()),(target,content)):
            with path.open('xb') as stream:
                written.append(path)
                stream.write(data)
    except Exception:
        for path in written: path.unlink(missing_ok=True)
        raise
    return target
