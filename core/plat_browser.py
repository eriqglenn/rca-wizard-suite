"""A visible, planner-operated County browser with PDF download capture."""
from __future__ import annotations
import os
import gc
from pathlib import Path
import threading
import time
from urllib.parse import urlsplit

from .plat_search import CLERK_HOSTS, PlatError, clerk_url, validated_pdf

MAX_PDF_BYTES = 100*1024*1024


class PlatBrowser:
    def __init__(self, events, profile=None, headless=False):
        self.events = events
        self.profile = Path(profile) if profile else Path(os.environ.get('LOCALAPPDATA',Path.home()))/'RCAWizardSuite'/'CountyClerkBrowser'
        self.headless = headless
        self.cancel = threading.Event()
        self.thread = None

    def start(self, record):
        if not record.get('viewer_url'):
            raise PlatError('This County record has no online plat image link. Try its volume/page on the County website or add a downloaded copy.')
        clerk_url(record['viewer_url'])
        if self.thread and self.thread.is_alive():
            raise PlatError('Finish or cancel the current County download first.')
        gc.collect()  # Dispose of old Tk cycles on the calling UI thread.
        self.cancel.clear()
        self.thread = threading.Thread(target=self._run,args=(dict(record),),daemon=True)
        self.thread.start()

    def close(self):
        self.cancel.set()

    def _run(self, record):
        context = None
        captured = []
        try:
            from playwright.sync_api import sync_playwright
            with sync_playwright() as p:
                errors = []
                for channel in ('chrome','msedge'):
                    try:
                        context = p.chromium.launch_persistent_context(str(self.profile),channel=channel,
                            headless=self.headless,accept_downloads=True,viewport={'width':1280,'height':850})
                        break
                    except Exception as exc:
                        errors.append(str(exc).splitlines()[0])
                if context is None:
                    raise PlatError('Could not open Chrome or Edge for the County download. Close any other wizard County browser and retry, or use Open County Website and Add Downloaded Plat.\n'+'\n'.join(errors))
                page = context.pages[0] if context.pages else context.new_page()
                for old in list(context.pages):
                    if old != page: old.close()
                page.goto('about:blank')
                def accept(content, source):
                    if captured or self.cancel.is_set(): return
                    if len(content)>MAX_PDF_BYTES:
                        raise PlatError('The plat is larger than 100 MB. Download it in the County viewer and choose Add Downloaded Plat.')
                    pages = validated_pdf(content)
                    captured.append({'record':record,'pdf':content,'pages':pages,'source_url':source})
                def response(r):
                    if captured or self.cancel.is_set(): return
                    try:
                        if urlsplit(r.url).hostname not in CLERK_HOSTS: return
                        if 'application/pdf' not in r.headers.get('content-type','').lower() or r.status!=200: return
                        size = int(r.headers.get('content-length','0'))
                        if size>MAX_PDF_BYTES: return
                        try:
                            content = r.body()
                        except Exception:
                            content = b''
                        if not content.lstrip().startswith(b'%PDF-'):
                            # Chromium's built-in PDF viewer can own the response.
                            # Fetch only that observed PDF URL using this browser's
                            # authenticated session; never guess image endpoints.
                            fetched = context.request.get(r.url,timeout=30000)
                            if fetched.status != 200: return
                            content = fetched.body()
                        accept(content,r.url)
                    except Exception:
                        # A PDF viewer may prevent response-body access; its Download
                        # control still raises the standard browser download event.
                        pass
                def download(d):
                    if captured or self.cancel.is_set(): return
                    try:
                        path = d.path()
                        if path and Path(path).stat().st_size<=MAX_PDF_BYTES:
                            accept(Path(path).read_bytes(),d.url)
                    except Exception as exc:
                        self.events.put(('browser_status','The browser download could not be read. Download a PDF, then use Add Downloaded Plat.'))
                def watch(tab):
                    tab.on('response',response)
                    tab.on('download',download)
                watch(page)
                context.on('page',watch)
                self.events.put(('browser_status','In the County browser, sign in if requested and open/download the complete plat PDF. The wizard will capture it for review.'))
                try:
                    page.goto(record['viewer_url'],wait_until='domcontentloaded',timeout=60000)
                except Exception:
                    if not captured and not context.pages:
                        raise PlatError('The County browser was closed before a plat was downloaded.')
                deadline = time.monotonic()+1200
                while not captured and not self.cancel.is_set() and time.monotonic()<deadline:
                    open_pages = [tab for tab in context.pages if not tab.is_closed()]
                    if not open_pages: break
                    open_pages[0].wait_for_timeout(250)
                if captured and not self.cancel.is_set():
                    self.events.put(('downloaded',captured[0]))
                elif not self.cancel.is_set():
                    self.events.put(('browser_error','No PDF was downloaded. You can retry after signing in, or choose Add Downloaded Plat for a PDF saved from the County viewer.'))
                context.close()
        except Exception as exc:
            if not self.cancel.is_set(): self.events.put(('browser_error',str(exc)))
        finally:
            self.events.put(('browser_done',None))
