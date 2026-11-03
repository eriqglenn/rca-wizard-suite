# Measured City layout and deterministic vector boundary-map rendering.
from __future__ import annotations
from datetime import datetime
from functools import lru_cache
from pathlib import Path
import io
import json
import math
import os
import re
import struct
import threading
import tempfile
from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.colors import black, white, HexColor
from reportlab.lib.utils import ImageReader
from shapely.geometry import LineString, box
from shapely.affinity import rotate
from shapely.ops import unary_union, linemerge, nearest_points
from .boundary_map import MapError, MAP_SIGNATURE, map_case, parcel_polygon, application_boundary, road_key, road_name
from .paths import WIZARD_FILES_DIR
from .map_text import display_text

STREET_FONT_SIZE = 18
ADDRESS_FONT_SIZE = 12
MIN_ADDRESS_FONT_SIZE = 5
FONT_LOCK = threading.RLock()

CITY_LAYOUT = {
    'outer_frame': [17.75, 19.25, 576, 756],
    'map_frame': [22.25, 162.25, 567, 607.5],
    'footer_frame': [22.25, 22.75, 567, 135],
    'heading_x': 27, 'heading_baselines': [141, 121.5, 102, 82.5],
    'heading_font': 'Calibri-Bold', 'heading_size': 16,
    'source_x': 457.5, 'source_baselines': [147.5, 138, 128.5],
    'source_font': 'Calibri-Italic', 'source_size': 8,
    'disclaimer_x': 457.5, 'disclaimer_y': 112, 'disclaimer_width': 123.5,
    'disclaimer_font': 'Calibri', 'disclaimer_size': 6, 'disclaimer_leading': 7.2,
    'boundary_key_box': [30, 29.5, 23, 14],
    'boundary_key_x': 57, 'boundary_key_baseline': 33.5,
    'seal_box': [453, 27, 132, 52.8],
}

@lru_cache(maxsize=1)
def city_seal():
    """Embed the supplied City seal unchanged, including its department wordmark."""
    asset = WIZARD_FILES_DIR / 'assets' / 'maps' / 'houston_planning_seal.tif'
    if not asset.is_file():
        raise MapError('The City seal asset is missing. Restore the complete Wizard_Files/assets folder from the wizard package, then retry.')
    return ImageReader(str(asset))


def text(c,x,y,s,size=10,font='Calibri',align='left',color=black):
    c.setFillColor(color)
    c.setFont(font,size)
    if align=='center': c.drawCentredString(x,y,s)
    elif align=='right': c.drawRightString(x,y,s)
    else: c.drawString(x,y,s)

@lru_cache(maxsize=1024)
def label_metrics(content):
    """Measure Arial glyph ink, including rounded digit overshoot, in em units."""
    face = pdfmetrics.getFont('Arial').face
    glyf = face.get_table('glyf')
    pen, bounds = 0.0, []
    for char in content:
        glyph = face.charToGlyph.get(ord(char), 0)
        start, end = face.glyphPos[glyph:glyph+2]
        if end > start:
            _, left, bottom, right, top = struct.unpack_from('>hhhhh',glyf,start)
            bounds.append((pen+left/face.unitsPerEm,bottom/face.unitsPerEm,
                           pen+right/face.unitsPerEm,top/face.unitsPerEm))
        pen += face.charWidths.get(ord(char),face.defaultWidth)/1000
    if not bounds:
        return (0,0,pen,0,pen)
    return (min(b[0] for b in bounds),min(b[1] for b in bounds),
            max(b[2] for b in bounds),max(b[3] for b in bounds),pen)


def label(c,x,y,s,angle=0,size=12,color=black):
    c.saveState(); c.translate(x,y); c.rotate(angle)
    _,bottom,_,top,_ = label_metrics(s)
    text(c,0,-(top+bottom)*size/2,s,size,font='Arial',align='center',color=color)
    c.restoreState()

def city_footer(c,case):
    s=CITY_LAYOUT
    c.setStrokeColor(black);c.setLineWidth(1)
    c.rect(*s['footer_frame'],stroke=1,fill=0)
    for line,y in zip(case['footer_lines'],case.get('heading_baselines',s['heading_baselines'])):
        text(c,s['heading_x'],y,line,s['heading_size'],s['heading_font'])
    date=datetime.now().astimezone()
    date_label=f'{date:%B} {date.day}, {date:%Y}'
    for line,y in zip(['Source: Harris Central Appraisal District',
                       'Date: '+date_label,'Reference: '+case['ref']],s['source_baselines']):
        text(c,s['source_x'],y,line,s['source_size'],s['source_font'])
    # The source uses a justified four-line disclaimer.
    disclaimer=['This map is made available for reference purposes',
                'only and should not be substituted for a survey',
                'product. The City of Houston will not accept',
                'liability of any kind in conjunction with its use.']
    for i,line in enumerate(disclaimer):
        c.saveState()
        tx=c.beginText(s['disclaimer_x'],s['disclaimer_y']-i*s['disclaimer_leading'])
        tx.setFont(s['disclaimer_font'],s['disclaimer_size'])
        width=pdfmetrics.stringWidth(line,s['disclaimer_font'],s['disclaimer_size'])
        assert width<=s['disclaimer_width']+.1
        tx.setWordSpace((s['disclaimer_width']-width)/line.count(' '))
        tx.textOut(line);c.drawText(tx)
        c.restoreState()
    c.setLineWidth(1.5);c.setStrokeColor(black)
    c.rect(*s['boundary_key_box'],stroke=1,fill=0)
    text(c,s['boundary_key_x'],s['boundary_key_baseline'],case['boundary_legend'],9,'Calibri')
    c.drawImage(city_seal(), *s['seal_box'], preserveAspectRatio=True, anchor='c')

def city_north_and_scale(c,angle,scale,stops=None):
    # Match the compact circular north symbol and alternating scale at the map's foot.
    nx,ny=480,171
    sx,sy=506,168
    c.setFillColor(white);c.rect(469,163.5,109,20,stroke=0,fill=1)
    c.saveState();c.translate(nx,ny);c.rotate(angle)
    c.setFillColor(black);c.circle(0,0,6.5,stroke=0,fill=1)
    c.setFillColor(white)
    arrow=c.beginPath();arrow.moveTo(0,5.4);arrow.lineTo(-4.5,.1)
    arrow.lineTo(-1.7,.1);arrow.lineTo(-1.7,-.8);arrow.lineTo(1.7,-.8)
    arrow.lineTo(1.7,.1);arrow.lineTo(4.5,.1);arrow.close()
    c.drawPath(arrow,stroke=0,fill=1)
    text(c,0,-5.1,'N',6.6,'Arial-Bold',align='center',color=white)
    c.restoreState()
    stops=stops or [0,5,10,20,30]
    for i in range(len(stops)-1):
        c.setFillColor(black if i%2==0 else white);c.setStrokeColor(black);c.setLineWidth(.45)
        c.rect(sx+stops[i]*scale,sy,(stops[i+1]-stops[i])*scale,3.2,stroke=1,fill=1)
    label_end = float('-inf')
    for n in stops:
        width = pdfmetrics.stringWidth(str(n), 'Arial', 6.5)
        start = sx + n*scale - width/2
        if start < label_end + 1.5 and n != stops[-1]:
            continue
        text(c,sx+n*scale,sy+5.8,str(n),6.5,'Arial',align='center')
        label_end = sx + n*scale + width/2
    text(c,sx+stops[-1]*scale+2.5,sy+.3,'Feet',6.5,'Arial')


def register_fonts():
    with FONT_LOCK:
        if 'Calibri' in pdfmetrics.getRegisteredFontNames():
            return
        font_dir = Path(os.environ.get('WINDIR', 'C:/Windows')) / 'Fonts'
        for name, file in [('Calibri', 'calibri.ttf'), ('Calibri-Bold', 'calibrib.ttf'),
                           ('Calibri-Italic', 'calibrii.ttf'), ('Arial', 'arial.ttf'), ('Arial-Bold', 'arialbd.ttf')]:
            path = font_dir / file
            if not path.exists():
                raise MapError(f'The City map requires {name}. Install the Windows/Office font, then retry.')
            pdfmetrics.registerFont(TTFont(name, str(path)))


def footer_lines(case):
    register_fonts()
    font, size = 'Calibri-Bold', 16
    if pdfmetrics.stringWidth(case.title, font, size) > 425:
        raise MapError('The map title is too long for the City footer.')
    lines = []
    for paragraph in [f'{case.block}, {case.sides},',
                      f'between {case.cross_streets[0]} and {case.cross_streets[1]}']:
        # Legend typography only; keep map street labels and source data intact.
        paragraph = re.sub(r'\b(\d+)(st|nd|rd|th)\b',
                           lambda m: m[1] + m[2].lower(), paragraph, flags=re.I)
        current = ''
        for word in paragraph.split():
            candidate = (current + ' ' + word).strip()
            width = 420 if not lines else 318
            if pdfmetrics.stringWidth(candidate, font, size) > width:
                if not current:
                    raise MapError('A description word is too long for the City footer.')
                lines.append(current)
                current = word
            else:
                current = candidate
        if current:
            lines.append(current)
    if len(lines) > 3:
        raise MapError('The block description exceeds the City footer at its standard font size. Shorten the description in Excel (for example, use a block range).')
    return [case.title, *lines, case.minimum]


def _road_lines(data, name):
    parts = []
    for f in data['roads']:
        if road_key(f['attributes'].get('FULL_NAME', '')) != road_key(name):
            continue
        parts.extend(LineString(p) for p in f['geometry']['paths'] if len(p) >= 2)
    if not parts:
        return []
    merged = unary_union(parts)
    if merged.geom_type != 'LineString':
        merged = linemerge(merged)
    return [merged] if merged.geom_type == 'LineString' else list(merged.geoms)


def default_settings(data):
    case = map_case(data)
    selected = unary_union([parcel_polygon(m['feature']) for m in data['matched']])
    roads = _road_lines(data, case.street)
    if not roads:
        return {'rotation': 0.0, 'zoom': 1.0, 'pan_x': 0.0, 'pan_y': 0.0,
                'include_street': case.sides.endswith('sides'), 'address_rotation': 0,
                'show_interior_lot_lines': True, 'show_platted_lot_lines': True, 'lot_divisions': {}, 'text_overrides': {}, 'street_styles': {}, 'label_offsets': {}}
    line = min(roads, key=lambda p: p.distance(selected.centroid))
    # Use the nearest segment's direction, not a curved street's distant end.
    segments = [LineString([a, b]) for a, b in zip(line.coords, list(line.coords)[1:])]
    segment = min(segments, key=lambda p: p.distance(selected.centroid))
    a, b = segment.coords
    angle = math.degrees(math.atan2(b[1] - a[1], b[0] - a[0]))
    while angle > 90:
        angle -= 180
    while angle <= -90:
        angle += 180
    return {'rotation': -90 - angle, 'zoom': 1.0, 'pan_x': 0.0, 'pan_y': 0.0,
            'include_street': case.sides.endswith('sides'), 'address_rotation': 0,
            'show_interior_lot_lines': True, 'show_platted_lot_lines': True, 'lot_divisions': {}, 'text_overrides': {}, 'street_styles': {}, 'label_offsets': {}}


def street_style(settings, key):
    style = settings.get('street_styles', {}).get(key, {})
    try:
        size = float(style.get('size', STREET_FONT_SIZE))
        angle = style.get('angle')
        angle = None if angle is None else float(angle)
        if not math.isfinite(size) or not 6 <= size <= 72:
            raise ValueError()
        if angle is not None and (not math.isfinite(angle) or not -180 <= angle <= 180):
            raise ValueError()
    except (TypeError, ValueError, AttributeError):
        raise MapError('Use a street-name size from 6 to 72 points and a rotation from -180 to 180 degrees.') from None
    return size, angle


def _angle(degrees):
    while degrees > 90:
        degrees -= 180
    while degrees < -90:
        degrees += 180
    return degrees


def label_rectangle(content, x, y, degrees, size):
    left,bottom,right,top,advance = label_metrics(content)
    height = (top-bottom)*size
    return rotate(box(x+(left-advance/2)*size-.1, y-height/2-.1,
                      x+(right-advance/2)*size+.1, y+height/2+.1), degrees, origin=(x,y))


def fit_address(content, parcel, preferred, occupied, min_size=MIN_ADDRESS_FONT_SIZE, inset=.35):
    """Find a readable label wholly inside one component of its own parcel."""
    anchor = parcel.representative_point()
    parts = [parcel] if parcel.geom_type == 'Polygon' else list(parcel.geoms)
    orientations = [preferred, _angle(preferred+90)]
    for part in parts:
        coords = list(part.minimum_rotated_rectangle.exterior.coords)
        for a, b in zip(coords, coords[1:]):
            angle = _angle(math.degrees(math.atan2(b[1]-a[1], b[0]-a[0])))
            if not any(abs(_angle(angle-old)) < .1 for old in orientations):
                orientations.append(angle)
    # Work in text-aligned coordinates so long, thin and angled parcels get
    # useful candidate positions instead of a grid across their bounding box.
    candidates = []
    for degrees in orientations:
        for part in parts:
            aligned = rotate(part, -degrees, origin=(0,0))
            x0,y0,x1,y1 = aligned.bounds
            center = aligned.representative_point()
            points = [(center.x,center.y), (aligned.centroid.x,aligned.centroid.y)]
            points += [(x0+(x1-x0)*u,y0+(y1-y0)*v)
                       for u in (.15,.3,.5,.7,.85) for v in (.15,.3,.5,.7,.85)]
            radians = math.radians(degrees)
            for x,y in points:
                px = x*math.cos(radians)-y*math.sin(radians)
                py = x*math.sin(radians)+y*math.cos(radians)
                candidates.append((px,py,degrees))
    interior = parcel.buffer(-inset)
    for half_size in range(int(ADDRESS_FONT_SIZE*2), int(min_size*2)-1, -1):
        size = half_size/2
        choices = []
        for x,y,degrees in candidates:
            rect = label_rectangle(content,x,y,degrees,size)
            if interior.covers(rect) and not any(rect.intersection(r).area>.01 for _,r in occupied):
                score = math.hypot(x-anchor.x,y-anchor.y) + abs(_angle(degrees-preferred))*.1
                choices.append((score,x,y,degrees,size))
        if choices:
            return min(choices)[1:]
    return None


def render_map(data, settings):
    """Return PDF bytes and page-coordinate hit boxes for the preview editor."""
    with FONT_LOCK:
        return _render_map(data, settings)


def _render_map(data, settings):
    from .lot_divisions import expand_map_lots
    original_data=data
    data=expand_map_lots(data,settings)
    register_fonts()
    case = map_case(data)
    footer = footer_lines(case)
    legend_labels = []
    for i, original in enumerate(footer):
        key = f'legend:{i}'
        content = display_text(settings, key, original)
        content = re.sub(r'\b(\d+)(st|nd|rd|th)\b', lambda m: m[1]+m[2].lower(), content, flags=re.I)
        limit = 425 if i == 0 else 420 if i == 1 else 318
        width = pdfmetrics.stringWidth(content, 'Calibri-Bold', CITY_LAYOUT['heading_size'])
        if width > limit:
            raise MapError('This legend line is too long for the City format. Shorten the wording; the font and spacing will stay unchanged.')
        footer[i] = content
        baseline = 141 - 19.5*i
        legend_labels.append({'key': key, 'text': content, 'size': CITY_LAYOUT['heading_size'],
                              'bounds': [27, baseline-4, 27+max(width,40), baseline+15]})
    raw_polys = [parcel_polygon(m['feature']) for m in data['matched']]
    selected = unary_union(raw_polys)
    origin = (selected.centroid.x, selected.centroid.y)
    rotation = float(settings['rotation'])
    def orient(p):
        return rotate(p, rotation, origin=origin)
    polys = [orient(p) for p in raw_polys]
    boundary = orient(application_boundary(data, settings.get('include_street', False)))
    # Include the actual end intersections. Parcel-only framing can crop an
    # angled end street and leave its label with no position on the page.
    framing = list(polys)
    main_lines = _road_lines(data, case.street)
    if main_lines:
        main = unary_union(main_lines)
        for name in case.cross_streets:
            cross_lines = _road_lines(data, name)
            if cross_lines:
                cross = unary_union(cross_lines)
                intersections = main.intersection(cross)
                points = ([intersections] if intersections.geom_type == 'Point' else
                          [g for g in getattr(intersections, 'geoms', []) if g.geom_type == 'Point'])
                if points:
                    framing.append(orient(min(points, key=lambda p: p.distance(selected.centroid))))
                elif main.distance(cross) < 40:
                    framing.append(orient(nearest_points(main, cross)[0]))
    bx0, by0, bx1, by1 = unary_union(framing).bounds
    mx, my, mw, mh = CITY_LAYOUT['map_frame']
    scale = min((mw - 80) / max(bx1 - bx0, 1), (mh - 96) / max(by1 - by0, 1)) * float(settings['zoom'])
    cx = (bx0 + bx1) / 2 + float(settings.get('pan_x', 0)) / scale
    cy = (by0 + by1) / 2 + float(settings.get('pan_y', 0)) / scale
    def xy(x, y):
        return mx + mw / 2 + (x - cx) * scale, my + mh / 2 + (y - cy) * scale
    def page_geometry(p):
        from shapely.affinity import affine_transform
        return affine_transform(p, [scale, 0, 0, scale, mx + mw / 2 - cx * scale, my + mh / 2 - cy * scale])
    view_rect = box(mx + 3, my + 24, mx + mw - 3, my + mh - 3)
    issues = list(data.get('warnings', []))
    if not view_rect.covers(page_geometry(unary_union(polys))):
        issues.append('An application parcel is outside the map frame. Zoom out or reset the framing.')
    pdf = io.BytesIO()
    c = canvas.Canvas(pdf, pagesize=(612, 792), pageCompression=1)
    c.setTitle(f'{case.reference} - {case.block} - Boundary Map')
    c.setAuthor('City Planning RCA Wizard Suite')
    c.setSubject('Boundary map generated from calculation-sheet properties and public HCAD geometry.')
    c.setKeywords(f'{MAP_SIGNATURE}|{case.wizard}|{case.workbook_hash}')
    c.setStrokeColor(black)
    c.setLineWidth(2)
    c.rect(*CITY_LAYOUT['outer_frame'], stroke=1, fill=0)
    c.saveState()
    clip = c.beginPath()
    clip.rect(mx, my, mw, mh)
    c.clipPath(clip, stroke=0, fill=0)
    def shape(p, fill=0, exterior_only=False):
        if p.geom_type in ('MultiPolygon', 'GeometryCollection'):
            for part in p.geoms:
                if part.geom_type in ('Polygon', 'MultiPolygon'):
                    shape(part, fill, exterior_only)
            return
        path = c.beginPath()
        for ring in ([p.exterior] if exterior_only else [p.exterior, *p.interiors]):
            points = list(ring.coords)
            path.moveTo(*xy(*points[0]))
            for point in points[1:]:
                path.lineTo(*xy(*point))
            path.close()
        c.drawPath(path, stroke=1, fill=fill, fillMode=0)
    c.setStrokeColor(HexColor('#B1B1B1'))
    c.setLineWidth(.5)
    for f in data['context']:
        shape(orient(parcel_polygon(f)))
    c.setFillColor(white)
    # Clear context edges beneath the selected parcels with the white map
    # background. Draw each interior division only when requested; the outer
    # application boundary and parcel/address geometry are independent of this.
    c.setStrokeColor(HexColor('#B1B1B1'))
    c.setLineWidth(.5)
    shape(unary_union(polys), fill=1)
    if settings.get('show_interior_lot_lines', True):
        for match in original_data['matched']:
            shape(orient(parcel_polygon(match['feature'])))
    c.setStrokeColor(black)
    c.setLineWidth(1.65)
    shape(boundary, exterior_only=True)
    from .lot_divisions import interior_lines,line_parts
    lot_lines,lot_notes=interior_lines(original_data,settings);issues.extend(lot_notes)
    c.saveState();c.setStrokeColor(HexColor('#666666'));c.setLineWidth(.65);c.setDash(3,2)
    for geometry in lot_lines:
        for line in line_parts(orient(geometry)):
            path=c.beginPath();points=list(line.coords);path.moveTo(*xy(*points[0]))
            for point in points[1:]:path.lineTo(*xy(*point))
            c.drawPath(path,stroke=1,fill=0)
    c.restoreState()
    labels, occupied, rejected_offsets = [], [], []
    def draw_label(key, x, y, content, degrees, size, parcel=None, advisory=False):
        label_issues = issues
        offset = settings.get('label_offsets', {}).get(key, [0, 0])
        moved = label_rectangle(content,x+offset[0],y+offset[1],degrees,size)
        if parcel is not None and not parcel.covers(moved):
            rejected_offsets.append(key)
            issues.append(f'{content}: the requested move leaves its parcel. The address was kept inside its own lot.')
        else:
            x, y = x + offset[0], y + offset[1]
        rect = label_rectangle(content,x,y,degrees,size)
        if not view_rect.covers(rect):
            label_issues.append(f'{content}: label is outside the map frame. Move it or adjust framing.')
        if parcel is not None and not parcel.covers(rect):
            issues.append(f'{content}: address label does not fit inside its parcel. Rotate address labels or adjust the framing.')
        for other_text, other in occupied:
            if rect.intersection(other).area > 1:
                label_issues.append(f'Labels overlap: {content} / {other_text}. Drag a label in the preview.')
        occupied.append((content, rect))
        label(c, x, y, content, angle=degrees, size=size)
        labels.append({'key': key, 'text': content, 'x': x, 'y': y, 'angle': degrees,
                       'size': size, 'bounds': list(rect.bounds), 'corners': list(rect.exterior.coords)})
    for m, p in zip(data['matched'], polys):
        prop = m['property']
        # Zero is the City's display convention for an unaddressed parcel;
        # leave its account, source address and matching data unchanged.
        content = str(prop.get('label') or prop.get('number') or '0').strip() or '0'
        key = 'parcel:' + str(prop['row'])
        content = display_text(settings, key, content)
        preferred = float(settings.get('address_rotation', 0))
        page_p = page_geometry(p)
        placement = fit_address(content,page_p,preferred,occupied)
        if placement is None:
            placement = fit_address(content,page_p,preferred,occupied,min_size=2,inset=.05)
            if placement:
                issues.append(f'{content}: small parcel; address text reduced to {placement[3]:g} pt to keep it inside its own lot. Zoom in for larger text.')
        if placement:
            px,py,degrees,size = placement
            draw_label(key,px,py,content,degrees,size,page_p)
        else:
            if key in settings.get('text_overrides', {}):
                raise MapError('The edited address cannot fit inside its parcel. Shorten the text or adjust the framing before editing it.')
            issues.append(f'{content}: its lot is too small for a readable address in this layout. Zoom in or change the orientation to keep the address inside its own lot.')
    selected_page = page_geometry(unary_union(polys))
    layout_occupied = list(occupied)
    for street_index, name in enumerate([case.street, *case.cross_streets]):
        street_key = 'street:' + road_key(name)
        size, manual_angle = street_style(settings, street_key)
        public_name = next((f['attributes']['FULL_NAME'] for f in data['roads']
                            if road_key(f['attributes'].get('FULL_NAME','')) == road_key(name)),name)
        content = display_text(settings, 'street:' + road_key(name), road_name(public_name))
        width = pdfmetrics.stringWidth(content, 'Arial', STREET_FONT_SIZE)
        if width > mw-20 and 'street:' + road_key(name) in settings.get('text_overrides', {}):
            raise MapError('The edited street name is too long for the map at the City font size. Shorten the wording.')
        candidates = []
        for raw_line in _road_lines(data, name):
            line = page_geometry(orient(raw_line))
            visible = line.intersection(view_rect)
            parts = [visible] if visible.geom_type == 'LineString' else list(getattr(visible, 'geoms', []))
            for part in parts:
                if part.geom_type != 'LineString' or part.length < width + 8:
                    continue
                for i in range(21):
                    distance = width / 2 + 4 + i / 20 * (part.length - width - 8)
                    q = part.interpolate(distance)
                    a = part.interpolate(max(0, distance - width / 2))
                    b = part.interpolate(min(part.length, distance + width / 2))
                    degrees = _angle(math.degrees(math.atan2(b.y - a.y, b.x - a.x)))
                    rect = rotate(box(q.x - width / 2, q.y - 9, q.x + width / 2, q.y + 9), degrees, origin=(q.x, q.y))
                    # Prefer a straight, uncluttered street corridor near the case.
                    conflict = rect.intersection(selected_page).area + sum(rect.intersection(r).area for _, r in layout_occupied)
                    curve = part.hausdorff_distance(LineString([a, b])) if part.length < width + 20 else q.distance(LineString([a, b]))
                    score = conflict * 100 + curve * 4 + q.distance(selected_page.centroid)
                    if view_rect.covers(rect):
                        candidates.append((score, q.x, q.y, degrees))
        if candidates:
            _, x, y, degrees = min(candidates)
        else:
            # Keep unresolved labels visible and movable without preventing save.
            x, y, degrees = mx + mw / 2, my + mh - 30 - street_index * 26, 0
            positions = []
            for py in range(int(my + mh - 25), int(my + 35), -26):
                for px in (mx + mw/2, mx + width/2 + 10, mx + mw - width/2 - 10):
                    rect = label_rectangle(content, px, py, 0, STREET_FONT_SIZE)
                    if view_rect.covers(rect) and not rect.intersects(selected_page) and not any(rect.intersects(r) for _,r in layout_occupied):
                        positions.append((abs(py-y)+abs(px-x),px,py))
            if positions:
                _,x,y = min(positions)
            if 'street:' + road_key(name) not in settings.get('label_offsets', {}):
                issues.append(f'Warning: {content} has no verified label position in this framing. The map can still be saved; drag the label into its street or adjust framing.')
        # Keep automatic placement independent of appearance overrides. This
        # rotates/resizes about the existing center and leaves other names put.
        offset = settings.get('label_offsets', {}).get(street_key, [0, 0])
        layout_occupied.append((content, label_rectangle(content,x+offset[0],y+offset[1],degrees,STREET_FONT_SIZE)))
        draw_label(street_key, x, y, content, degrees if manual_angle is None else manual_angle, size, advisory=True)
    c.restoreState()
    # Compact scale in the same footer position, using nice survey-foot units.
    choices = [n * 10 ** power for power in range(-1, 6) for n in (1, 1.5, 2, 3, 5, 7.5)]
    total = max(n for n in choices if n * scale <= 67)
    stops = [0, total / 6, total / 3, 2 * total / 3, total]
    # Integer/half-foot values prevent long decimal tick labels.
    if any(abs(n * 2 - round(n * 2)) > .001 for n in stops):
        stops = [0, total / 2, total]
    stops = [int(n) if int(n) == n else n for n in stops]
    city_north_and_scale(c, rotation, scale, stops)
    c.setStrokeColor(black)
    c.setLineWidth(1)
    c.rect(*CITY_LAYOUT['map_frame'], stroke=1, fill=0)
    city_footer(c, {'footer_lines': footer, 'heading_baselines': [141 - 19.5 * i for i in range(len(footer))],
                    'ref': case.reference, 'boundary_legend': case.boundary_legend})
    c.save()
    return {'pdf': pdf.getvalue(), 'labels': labels, 'editable_labels': labels + legend_labels, 'issues': list(dict.fromkeys(issues)),
            'rejected_offsets': rejected_offsets,
            'blocking': [], 'scale': scale, 'scale_feet': total,
            'page_transform': [scale,0,0,scale,mx+mw/2-cx*scale,my+mh/2-cy*scale],
            'rotation_origin': origin,
            'footer_lines': footer, 'rotation': rotation, 'selected_count': len(polys)}


def pdf_png(pdf_bytes, dpi=110):
    import fitz
    with fitz.open(stream=pdf_bytes, filetype='pdf') as doc:
        return doc[0].get_pixmap(dpi=dpi, alpha=False).tobytes('png')


def save_map(data, settings, rendered, output_dir):
    """Save exactly the reviewed PDF and its provenance, without replacing prior maps."""
    if not rendered.get('pdf'):
        raise MapError('Create a map preview before saving.')
    folder = Path(output_dir)
    folder.mkdir(parents=True, exist_ok=True)
    case = map_case(data)
    stem = case.reference.replace(' ', '_') + '_Boundary_Map'
    if any((folder / (stem + ext)).exists() for ext in ('.pdf', '.png', '.boundary.json')):
        stem += '_' + datetime.now().strftime('%Y%m%d_%H%M%S_%f')
    pdf = folder / (stem + '.pdf')
    png = folder / (stem + '.png')
    manifest = folder / (stem + '.boundary.json')
    metadata = {'schema': 1, 'case': data['case'], 'sources': data['sources'], 'settings': settings,
                'scale_points_per_foot': rendered['scale'], 'scale_bar_feet': rendered['scale_feet'],
                'footer_lines': rendered['footer_lines'], 'created_at': datetime.now().astimezone().isoformat(),
                'warnings': list(dict.fromkeys(rendered.get('issues', []) + rendered.get('blocking', []))),
                'parcel_matches': [{'row':m['property']['row'], 'workbook_account':m['property']['account'],
                                    'matched_account':m['feature']['attributes']['HCAD_NUM'],
                                    'method':m.get('match_method','account'),
                                    'service':m['feature'].get('_parcel_service',data['sources'][0]['service'] if data['sources'] else '')}
                                   for m in data['matched']]}
    written = []
    try:
        for path, content in [(png, pdf_png(rendered['pdf'], 200)),
                              (manifest, json.dumps(metadata, indent=2).encode('utf-8')),
                              (pdf, rendered['pdf'])]:
            with path.open('xb') as stream:
                written.append(path)
                stream.write(content)
    except Exception:
        for path in written:
            path.unlink(missing_ok=True)
        raise
    return pdf


def export_map(rendered, destination):
    """Download the reviewed map as PDF or PNG, independent of attachment selection."""
    if not rendered or not rendered.get('pdf'):
        raise MapError('Create a map preview before downloading.')
    target = Path(destination)
    suffix = target.suffix.lower()
    if suffix not in ('.pdf', '.png'):
        raise MapError('Save the map as a PDF or PNG file.')
    content = rendered['pdf'] if suffix == '.pdf' else pdf_png(rendered['pdf'], 200)
    # A failed write must not truncate a map the user already has.
    temporary = None
    try:
        with tempfile.NamedTemporaryFile(dir=target.parent, prefix='.boundary_map_', suffix='.tmp', delete=False) as stream:
            temporary = Path(stream.name)
            stream.write(content)
        os.replace(temporary, target)
    finally:
        if temporary is not None:
            temporary.unlink(missing_ok=True)
    return target
