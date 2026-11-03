"""HCAD Parcel Viewer prints with a yellow application boundary."""
from __future__ import annotations

from datetime import datetime
import hashlib
import json
import math
from pathlib import Path
import re
import ssl
import threading
from urllib.parse import urlencode, urlsplit
from urllib.request import Request, urlopen
import uuid

from shapely.geometry import Polygon
from shapely.geometry.polygon import orient

from .boundary_map import MapError, application_boundary

PRINT_SERVICE = 'https://arcweb.hcad.org/server/rest/services/MapTemplate/HCADMaps/GPServer/Export%20Web%20Map'
PARCEL_VIEWER = 'https://arcweb.hcad.org/parcel-viewer-v2.0/'
PARCEL_LAYER = 'https://arcweb.hcad.org/server/rest/services/public/public_query/MapServer/0'
FACET_LAYER = 'https://arcweb.hcad.org/server/rest/services/public/FACET_Page_Grid/MapServer/0'
BASEMAP_STYLE = 'https://www.arcgis.com/sharing/rest/content/items/020ac8d7e59943888ba5308d8769b167/resources/styles/root.json'
LAYOUT = 'Letter Landscape Side'


def default_settings():
    return {'show_platted_lot_lines':True,'lot_divisions':{},'zoom': 1.0, 'pan_x': 0.0, 'pan_y': 0.0}


def tax_map_boundary(data):
    """Enclose every selected parcel with one continuous, hole-free outline.

    Close streets and gaps while retaining stepped parcel edges where possible.
    Keep this rule local to tax maps; City boundary maps have separate controls.
    Coordinates are State Plane feet. Never subtract context parcels: doing so
    can split the outline again or introduce interior yellow rings.
    """
    selected = application_boundary(data, include_street=False)
    if selected.geom_type == 'Polygon':
        return Polygon(selected.exterior)
    hull = selected.convex_hull
    span = max(hull.bounds[2] - hull.bounds[0], hull.bounds[3] - hull.bounds[1])
    distance = 45.0
    while distance <= max(45.0, span * 2):
        joined = selected.buffer(distance, join_style=2).buffer(-distance, join_style=2)
        joined = joined.intersection(hull).union(selected)
        if joined.geom_type == 'Polygon' and joined.is_valid:
            return Polygon(joined.exterior)
        distance *= 2
    # Unusual separated geometries still get a single enclosing outer ring.
    return hull


def map_extent(data, settings):
    boundary = tax_map_boundary(data)
    x0, y0, x1, y1 = boundary.bounds
    zoom = float(settings.get('zoom', 1))
    px, py = float(settings.get('pan_x', 0)), float(settings.get('pan_y', 0))
    if not all(math.isfinite(v) for v in (x0, y0, x1, y1, zoom, px, py)) or not .5 <= zoom <= 2:
        raise MapError('Choose a framing value between 0.5 and 2.')
    span = max(650, max(x1 - x0, y1 - y0) * 1.65) / zoom
    cx, cy = (x0 + x1) / 2 + px, (y0 + y1) / 2 + py
    return boundary, [cx - span / 2, cy - span / 2, cx + span / 2, cy + span / 2]


def boundary_geometry(boundary):
    rings = []
    for poly in ([boundary] if boundary.geom_type == 'Polygon' else boundary.geoms):
        # Esri polygons require clockwise shells and counterclockwise holes.
        poly = orient(poly, sign=-1)
        rings.append([list(p) for p in poly.exterior.coords])
        rings.extend([list(p) for p in hole.coords] for hole in poly.interiors)
    return {'rings': rings, 'spatialReference': {'wkid': 2278}}


def build_web_map(data, settings):
    boundary, extent = map_extent(data, settings)
    symbol = {'type': 'esriSFS', 'style': 'esriSFSNull', 'color': [0, 0, 0, 0],
              'outline': {'type': 'esriSLS', 'style': 'esriSLSSolid', 'color': [255, 255, 0, 255], 'width': 3}}
    graphic = {'id': 'application_boundary', 'title': 'Application boundary', 'opacity': 1,
               'visibility': True, 'featureCollection': {'layers': [{
                   'layerDefinition': {'name': 'Application boundary', 'geometryType': 'esriGeometryPolygon',
                                       'drawingInfo': {'renderer': {'type': 'simple', 'symbol': symbol}}},
                   'featureSet': {'geometryType': 'esriGeometryPolygon', 'spatialReference': {'wkid': 2278},
                                  'features': [{'geometry': boundary_geometry(boundary), 'symbol': symbol}]}}]}}
    return {
        'mapOptions': {'extent': dict(zip(('xmin', 'ymin', 'xmax', 'ymax'), extent), spatialReference={'wkid': 2278}),
                       'spatialReference': {'wkid': 2278}, 'rotation': 0, 'showAttribution': True},
        'baseMap': {'title': 'HCAD Parcel Basemap', 'baseMapLayers': [{
            'id': 'hcad_parcel_basemap', 'type': 'VectorTileLayer', 'layerType': 'VectorTileLayer',
            'title': 'HCAD Parcel Basemap', 'styleUrl': BASEMAP_STYLE, 'visibility': True, 'opacity': 1}]},
        'operationalLayers': [
            {'id': 'parcels', 'url': PARCEL_LAYER, 'title': 'Parcels', 'opacity': .2, 'visibility': True},
            {'id': 'facet_page', 'url': FACET_LAYER, 'title': 'Facet Page', 'opacity': 1, 'visibility': True},
            graphic],
        'exportOptions': {'dpi': 200, 'outputSize': [1500, 1500]},
        'layoutOptions': {'titleText': 'HCAD_Map', 'authorText': '', 'copyrightText': 'HCAD-GIS',
                          'legendOptions': {'operationalLayers': [{'id': 'parcels'}, {'id': 'facet_page'}]}}
    }


class HCADPrintClient:
    def __init__(self, cache_dir, progress=None, cancel=None, use_cache=True):
        self.cache_dir = Path(cache_dir)
        self.cache_dir.mkdir(parents=True, exist_ok=True)
        self.progress = progress or (lambda text: None)
        self.cancel = cancel or threading.Event()
        self.use_cache = use_cache
        self.cached = False
        self._context = None

    def check_cancelled(self):
        if self.cancel.is_set():
            raise MapError('Tax map creation cancelled.')

    def request(self, url, params=None):
        self.check_cancelled()
        if self._context is None:
            try:
                from .map_connection import map_ssl_context
            except ModuleNotFoundError:
                self._context = ssl.create_default_context()
            else:
                self._context = map_ssl_context()
        body = urlencode(params).encode('utf-8') if params is not None else None
        req = Request(url, data=body, headers={'User-Agent': 'Houston-RCA-Tax-Map/1.0'})
        try:
            with urlopen(req, context=self._context, timeout=90) as response:
                content = response.read(35 * 1024 * 1024 + 1)
        except OSError as exc:
            raise MapError(f'Could not reach HCAD to print the tax map. Check the connection and retry.\n{exc}') from exc
        self.check_cancelled()
        if len(content) > 35 * 1024 * 1024:
            raise MapError('HCAD returned a map larger than 35 MB. Zoom closer and retry.')
        return content

    def print_map(self, web_map):
        key = hashlib.sha256(json.dumps([LAYOUT, web_map, datetime.now().date().isoformat()], sort_keys=True).encode()).hexdigest()
        cache = self.cache_dir / (key + '.pdf')
        if self.use_cache and cache.is_file():
            content = cache.read_bytes()
            if content.startswith(b'%PDF-'):
                self.cached = True
                self.progress('Using the saved HCAD tax map for this view.')
                return content
        self.cached = False
        self.progress('HCAD is printing the tax map...')
        raw = self.request(PRINT_SERVICE + '/execute', {'f': 'json', 'Web_Map_as_JSON': json.dumps(web_map),
                          'Format': 'PDF', 'Layout_Template': LAYOUT})
        try:
            result = json.loads(raw)
        except (ValueError, UnicodeError) as exc:
            raise MapError('HCAD returned an unreadable print response. Retry the tax map.') from exc
        if result.get('error'):
            err = result['error']
            raise MapError('HCAD could not print this view: ' + str(err.get('message', err)) + '\n' + '\n'.join(err.get('details', [])))
        url = next((r.get('value', {}).get('url') for r in result.get('results', []) if r.get('paramName') == 'Output_File'), None)
        parts = urlsplit(url or '')
        if parts.hostname != 'arcweb.hcad.org' or parts.scheme not in ('http', 'https'):
            raise MapError('HCAD did not return a downloadable map. Retry the tax map.')
        content = self.request(parts._replace(scheme='https').geturl())
        if not content.startswith(b'%PDF-'):
            raise MapError('HCAD returned an invalid PDF. Retry the tax map.')
        temp = cache.with_suffix('.' + uuid.uuid4().hex + '.tmp')
        try:
            temp.write_bytes(content)
            temp.replace(cache)
        finally:
            temp.unlink(missing_ok=True)
        return content


def render_tax_map(data, settings, client):
    import fitz
    web_map = build_web_map(data, settings)
    pdf = client.print_map(web_map)
    with fitz.open(stream=pdf, filetype='pdf') as doc:
        if len(doc) != 1 or doc[0].rect.width <= doc[0].rect.height:
            raise MapError('HCAD did not return the expected single-page landscape tax map. Please retry.')
    boundary, extent = map_extent(data, settings)
    from .lot_divisions import overlay_tax_lots
    pdf,lot_notes=overlay_tax_lots(pdf,data,settings,extent)
    notes = list(data.get('warnings', []))+lot_notes
    x0, y0, x1, y1 = boundary.bounds
    if x0 < extent[0] or y0 < extent[1] or x1 > extent[2] or y1 > extent[3]:
        notes.append('Part of the yellow boundary is outside this view. Zoom out or reset the framing to show it all.')
    notes.append('HCAD Parcel Viewer basemap; one continuous yellow outline encloses the properties and intervening streets or gaps. Review the outer boundary before use.')
    if client.cached:
        notes.append('Saved HCAD print reused for this view.')
    return {'pdf': pdf, 'issues': notes, 'extent': extent, 'web_map': web_map}


def save_tax_map(data, settings, rendered, output_dir):
    if not rendered or not rendered.get('pdf'):
        raise MapError('Create a tax map preview before saving.')
    folder = Path(output_dir)
    folder.mkdir(parents=True, exist_ok=True)
    stem = re.sub(r'[^A-Za-z0-9_-]+', '_', data['case']['reference']).strip('_') + '_Tax_Map'
    if any((folder / (stem + ext)).exists() for ext in ('.pdf', '.taxmap.json')):
        stem += '_' + datetime.now().strftime('%Y%m%d_%H%M%S_%f')
    target = folder / (stem + '.pdf')
    meta = folder / (stem + '.taxmap.json')
    record = {'schema': 1, 'kind': 'hcad_tax_map', 'case': data['case'], 'settings': settings,
              'sources': data.get('sources', []), 'print_service': PRINT_SERVICE, 'basemap_style': BASEMAP_STYLE,
              'extent': rendered['extent'], 'created_at': datetime.now().astimezone().isoformat(),
              'warnings': rendered['issues'],
              'parcel_accounts': [m['feature']['attributes'].get('HCAD_NUM', '') for m in data['matched']]}
    written = []
    try:
        for path, content in ((meta, json.dumps(record, indent=2).encode('utf-8')), (target, rendered['pdf'])):
            with path.open('xb') as stream:
                written.append(path)
                stream.write(content)
    except Exception:
        for path in written:
            path.unlink(missing_ok=True)
        raise
    return target
