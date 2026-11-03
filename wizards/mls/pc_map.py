"""Separate MLS PC land-use/support map; ordinary maps retain their appearance."""
from __future__ import annotations
from datetime import datetime
from pathlib import Path
import io
import json
import math
from reportlab.pdfgen import canvas
from reportlab.lib.colors import HexColor, black, white
from reportlab.pdfbase import pdfmetrics
from shapely.geometry import box, LineString, Point, Polygon
from shapely.affinity import rotate, affine_transform
from shapely.ops import unary_union, nearest_points
from core.boundary_map import parcel_polygon, application_boundary, map_case, road_key, road_name, MapError
from core.boundary_map_render import (register_fonts, text, label, label_rectangle, fit_address,
    street_style, city_seal, pdf_png, FONT_LOCK, _road_lines, _angle, default_settings)
from core.map_text import display_text
from .pc_landuse import (COLORS, NAMES, LEGEND_USES, context_use, selected_use,
    TRANSPORT_HATCH, BELOW_HATCH, SUPPORT_COLOR, PROTEST_COLOR)

PAGE = (792, 612)
FRAME = (188, 18, 586, 576)

def pc_defaults(data):
    result = default_settings(data)
    result['rotation'] = 0
    result['zoom'] = 1
    return result

def polygon_path(c, polygon):
    path=c.beginPath()
    for ring in [polygon.exterior,*polygon.interiors]:
        points=list(ring.coords);path.moveTo(*points[0])
        for q in points[1:]:path.lineTo(*q)
        path.close()
    return path

def hatch(c, polygon, color, spacing=4, descending=False):
    """Clip pattern to the same polygon used for its fill, including holes."""
    if polygon.geom_type != 'Polygon':
        for part in getattr(polygon,'geoms',[]):hatch(c,part,color,spacing,descending)
        return
    c.saveState();c.clipPath(polygon_path(c,polygon),stroke=0,fillMode=0)
    c.setStrokeColor(HexColor(color));c.setLineWidth(.55)
    a,b,d,e=polygon.bounds;height=e-b
    for x in range(math.floor(a-height-spacing),math.ceil(d+height+spacing),spacing):
        c.line(x,e if descending else b,x+height,b if descending else e)
    c.restoreState()

def support_arrow(x, y, size, degrees):
    arrow=Polygon([(x-size,y-size),(x+size,y),(x-size,y+size),(x-size*.35,y)])
    return rotate(arrow,degrees,origin=(x,y))

def support_direction(street, x, y, fallback):
    if street.is_empty:return fallback
    target=nearest_points(Point(x,y),street)[1]
    return math.degrees(math.atan2(target.y-y,target.x-x))

def status_symbol(c, kind, x, y, size, degrees=-90):
    c.saveState()
    if kind=='support':
        c.setFillColor(HexColor(SUPPORT_COLOR))
        c.drawPath(polygon_path(c,support_arrow(x,y,size,degrees)),fill=1,stroke=0)
    else:
        c.setStrokeColor(HexColor(PROTEST_COLOR));c.setLineWidth(max(1,size*.4))
        c.lines([(x-size,y-size,x+size,y+size),(x-size,y+size,x+size,y-size)])
    c.restoreState()

def render_pc_map(data, settings):
    with FONT_LOCK:
        return _render(data, settings)

def expand_pc_lots(data, settings):
    try:
        from core.lot_divisions import expand_map_lots
    except ModuleNotFoundError:  # V26work has no legal-lot map editor.
        return data
    from .pc_data import attach_pc_rows
    if not data.get('pc_facts'):return data
    expanded=expand_map_lots(data,settings)
    if len(expanded['matched'])==len(data['matched']):return data
    result=attach_pc_rows(expanded,data['pc_facts'],data.get('pc_protested_rows',[]))
    # A reviewed division changes only that property's join to the Step 2 lots.
    # Preserve reviewed categories and symbols on all unchanged properties.
    import copy
    for match,previous in zip(result['matched'],expanded['matched']):
        if not previous.get('platted_lot') and 'pc' in previous:
            match['pc']=copy.deepcopy(previous['pc'])
    return result

def _render(data, settings):
    original_data=data
    data=expand_pc_lots(data,settings)
    register_fonts()
    case = map_case(data)
    raw = [parcel_polygon(m['feature']) for m in data['matched']]
    union = unary_union(raw)
    origin = (union.centroid.x, union.centroid.y)
    angle = float(settings['rotation'])
    polys = [rotate(p, angle, origin=origin) for p in raw]
    frame = box(*[FRAME[0], FRAME[1], FRAME[0]+FRAME[2], FRAME[1]+FRAME[3]])
    x0,y0,x1,y1 = unary_union(polys).bounds
    mx,my,mw,mh = FRAME
    scale = min((mw-156)/max(x1-x0,1),(mh-116)/max(y1-y0,1))*float(settings['zoom'])
    cx=(x0+x1)/2+float(settings.get('pan_x',0))/scale
    cy=(y0+y1)/2+float(settings.get('pan_y',0))/scale
    matrix=[scale,0,0,scale,mx+mw/2-cx*scale,my+mh/2-cy*scale]
    def pg(g): return affine_transform(rotate(g,angle,origin=origin),matrix)
    page_polys=[pg(p) for p in raw]
    boundary=pg(application_boundary(data,settings.get('include_street',False)))
    view=box(mx+3,my+28,mx+mw-3,my+mh-3)
    issues=list(data.get('warnings',[])); labels=[]; occupied=[]; rejected=[]
    output=io.BytesIO(); c=canvas.Canvas(output,pagesize=PAGE,pageCompression=1,invariant=1)
    c.setTitle(f'{case.reference} - PC Land Use and Support Map')
    c.setAuthor('City Planning RCA Wizard Suite')
    c.setKeywords('RCA_PC_MAP_V1|mls|'+case.workbook_hash)
    c.setLineWidth(1);c.setStrokeColor(black);c.rect(18,18,756,576,fill=0)
    c.saveState();clip=c.beginPath();clip.rect(*FRAME);c.clipPath(clip,stroke=0)
    def shape(p,fill=0,stroke=1):
        if p.geom_type!='Polygon':
            for g in getattr(p,'geoms',[]): shape(g,fill,stroke)
            return
        c.drawPath(polygon_path(c,p),fill=fill,stroke=stroke,fillMode=0)
    c.setLineWidth(.4);c.setStrokeColor(HexColor('#AAA5A5'))
    used=set();land_use_records=[]
    for f in data['context']:
        p=pg(parcel_polygon(f))
        if not p.intersects(frame):continue
        lu=context_use(f);used.add(lu)
        c.setFillColor(HexColor(COLORS[lu]));shape(p,fill=1)
        if lu=='TRN':hatch(c,p,TRANSPORT_HATCH,descending=True)
    for m,p in zip(data['matched'],page_polys):
        info=m.get('pc',{});lu=selected_use(m)
        land_use_records.append({'row':m['property']['row'],'category':lu,'color':COLORS[lu],
                                'worksheet':info.get('map_land_use',''),'public':m['feature'].get('pc_land_use',{})})
        used.add(lu);c.setFillColor(HexColor(COLORS[lu]));c.setStrokeColor(HexColor('#777777'));shape(p,fill=1,stroke=int(settings.get('show_platted_lot_lines' if m.get('platted_lot') else 'show_interior_lot_lines',True)))
        if lu=='TRN':hatch(c,p,TRANSPORT_HATCH,descending=True)
        if info.get('below'):
            hatch(c,p,BELOW_HATCH)
    c.setStrokeColor(black);c.setLineWidth(1.8);shape(boundary)
    def draw(key,content,x,y,degrees,size,parcel=None):
        dx,dy=settings.get('label_offsets',{}).get(key,[0,0])
        rect=label_rectangle(content,x+dx,y+dy,degrees,size)
        if parcel is not None and not parcel.covers(rect):
            rejected.append(key);issues.append(f'{content}: address kept inside its parcel.')
        else:x+=dx;y+=dy
        rect=label_rectangle(content,x,y,degrees,size)
        if not view.covers(rect):issues.append(f'{content}: label near or outside the map frame; adjust its placement.')
        if any(rect.intersection(g).area>1 for _,g in occupied):issues.append(f'{content}: overlaps another label or symbol; adjust its placement.')
        label(c,x,y,content,degrees,size)
        occupied.append((content,rect))
        labels.append({'key':key,'text':content,'x':x,'y':y,'angle':degrees,'size':size,'bounds':list(rect.bounds),'corners':list(rect.exterior.coords)})
    try:
        from core.lot_divisions import interior_lines,line_parts
        lot_lines,lot_notes=interior_lines(original_data,settings);issues.extend(lot_notes)
    except ModuleNotFoundError:
        lot_lines=[]  # V26work shows its existing parcel boundaries.

    c.saveState();c.setStrokeColor(HexColor('#555555'));c.setLineWidth(.65);c.setDash(3,2)
    for geometry in lot_lines:
        for line in line_parts(pg(geometry)):
            points=list(line.coords);path=c.beginPath();path.moveTo(*points[0])
            for point in points[1:]:path.lineTo(*point)
            c.drawPath(path,stroke=1,fill=0)
    c.restoreState()
    symbol_records=[]
    support_street=unary_union([pg(line) for line in _road_lines(data,case.street)])
    if support_street.is_empty and any(m.get('pc',{}).get('signed') for m in data['matched']):
        issues.append('Support arrow direction could not be checked against the application street; review the downward-facing arrows.')
    for m,p in zip(data['matched'],page_polys):
        prop=m['property'];info=m.get('pc',{});key='parcel:'+str(prop['row'])
        content=display_text(settings,key,str(prop.get('number') or '0'))
        # Reserve symbol space first so narrow lots can fit both their address
        # and status instead of letting the address consume the entire lot.
        for kind,shown in [('support',info.get('signed')),('protest',info.get('protested'))]:
            if not shown:continue
            candidates=[]
            a,b,d,e=p.bounds
            for size in (4,3,2):
                for fx in (.2,.8,.5,.35,.65,.1,.9):
                    for fy in (.2,.8,.5,.35,.65,.1,.9):
                        x=a+(d-a)*fx;y=b+(e-b)*fy
                        degrees=support_direction(support_street,x,y,-90+angle) if kind=='support' else 0
                        r=box(*support_arrow(x,y,size,degrees).bounds) if kind=='support' else box(x-size,y-size,x+size,y+size)
                        if p.covers(r) and not any(r.intersects(g) for _,g in occupied):candidates.append((size,x,y,r,degrees))
                if candidates:break
            if not candidates:
                issues.append(f'{content}: {kind} symbol cannot fit at this scale; zoom in.');continue
            size,x,y,r,degrees=candidates[0]
            symbol_records.append({'row':prop['row'],'kind':kind,'bounds':list(r.bounds),
                                   'center':[x,y],'size':size,'angle':degrees})
            occupied.append((kind,r));status_symbol(c,kind,x,y,size,degrees)
        place=fit_address(content,p,float(settings.get('address_rotation',0)),occupied)
        if not place:place=fit_address(content,p,float(settings.get('address_rotation',0)),occupied,min_size=2,inset=.05)
        if place:draw(key,content,*place,parcel=p)
        else:issues.append(f'{content}: address too small to fit in this framing.')
    layout_occupied=list(occupied)
    for i,name in enumerate([case.street,*case.cross_streets]):
        key='street:'+road_key(name);content=display_text(settings,key,road_name(name))
        size,manual=street_style(settings,key);candidates=[]
        for road in _road_lines(data,name):
            line=pg(road).intersection(view)
            for part in ([line] if line.geom_type=='LineString' else getattr(line,'geoms',[])):
                if part.geom_type!='LineString' or part.is_empty or part.length==0:continue
                for n in range(1,20):
                    q=part.interpolate(n/20,normalized=True);d=part.project(q)
                    a=part.interpolate(max(0,d-30));b=part.interpolate(min(part.length,d+30))
                    degrees=_angle(math.degrees(math.atan2(b.y-a.y,b.x-a.x)))
                    rect=label_rectangle(content,q.x,q.y,degrees,18)
                    if view.covers(rect):
                        score=rect.intersection(unary_union(page_polys)).area*100+sum(rect.intersection(g).area*100 for _,g in layout_occupied)+q.distance(unary_union(page_polys).centroid)
                        candidates.append((score,q.x,q.y,degrees))
        if candidates:_,x,y,degrees=min(candidates)
        else:x,y,degrees=mx+mw/2,my+mh-25-25*i,0;issues.append(f'{content}: adjust the street label to its preferred position.')
        dx,dy=settings.get('label_offsets',{}).get(key,[0,0])
        layout_occupied.append((content,label_rectangle(content,x+dx,y+dy,degrees,18)))
        draw(key,content,x,y,degrees if manual is None else manual,size)
    if not view.covers(unary_union(page_polys)):issues.append('A selected parcel is outside the map frame. Zoom out or reset framing.')
    c.restoreState();c.setStrokeColor(black);c.setLineWidth(1);c.rect(*FRAME,fill=0)
    # North arrow and survey-foot scale use the same rotation and geometry scale.
    c.setFillColor(white);c.rect(mx+mw-164,my+1,161,24,stroke=0,fill=1)
    c.saveState();c.translate(mx+mw-148,my+12);c.rotate(angle);c.setStrokeColor(black);c.setLineWidth(1.2)
    c.line(0,-5,0,6);c.lines([(-3,2,0,6),(3,2,0,6)]);text(c,0,9,'N',7,'Arial',align='center');c.restoreState()
    feet=max(v*10**p for p in range(-1,6) for v in (1,2,5) if v*10**p*scale<=98)
    sx=mx+mw-116;sy=my+8;c.setLineWidth(1);c.line(sx,sy,sx+feet*scale,sy)
    for n in (0,feet/2,feet):
        c.line(sx+n*scale,sy-2,sx+n*scale,sy+2);text(c,sx+n*scale,sy+5,f'{n:g}',6,'Arial',align='center')
    text(c,sx+feet*scale+4,sy,'Feet',6,'Arial')
    # PC legend is deliberately separate from the standard boundary-map footer.
    legend=[];y=571
    wizard_key=data.get('pc_facts',{}).get('wizard_key','mls')
    building=wizard_key.startswith('mbl')
    title=('Building Line Block' if building else 'Lot Size Block')+(' Renewal' if wizard_key.endswith('_renewal') else '')
    headings=('Landuse and Support Map','Special Minimum',*(['Building Line Block','Renewal'] if title=='Building Line Block Renewal' else [title]))
    for original in headings:
        key=f'legend:{len(legend)}';content=display_text(settings,key,original)
        if pdfmetrics.stringWidth(content,'Arial-Bold',11)>154:raise MapError('Shorten the PC legend heading to fit its column.')
        text(c,26,y,content,11,'Arial-Bold');legend.append({'key':key,'text':content,'size':11,'bounds':[26,y-3,180,y+12]});y-=14
    facts=data.get('pc_facts',{});ci=facts.get('case_info',{})
    def wrap(s,width=150,size=8):
        lines=[];line=''
        for word in str(s).split():
            trial=(line+' '+word).strip()
            if line and pdfmetrics.stringWidth(trial,'Arial',size)>width:lines.append(line);line=word
            else:line=trial
        return lines+[line] if line else lines
    for value in ('Subdivision: '+(ci.get('Subdivision_Name') or '[review]'),
                  'File Number: '+case.reference,('Minimum Building Line: ' if building else 'Minimum Lot Size: ')+case.minimum,case.block+', '+case.sides):
        for line in wrap(value):text(c,26,y-8,line,8,'Arial');y-=11
        y-=3
    y-=5
    legend_entries=[]
    statuses=[('support','Property owner signed to support application'),
              ('protest','Property owner protesting application'),
              ('below','Properties with a recorded building line below the minimum' if building else 'Properties that do not meet the minimum lot size')]
    # Allocate the complete legend before drawing. Long case descriptions may
    # use more room; wrap or modestly compact rows, never drop a category.
    captions=[(kind,caption,wrap(caption,132,8)) for kind,caption in statuses]
    for kind,caption,lines in captions:
        center=y-5*len(lines)+5
        if kind in ('support','protest'):status_symbol(c,kind,33.5,center,5)
        else:
            c.setFillColor(white);c.setStrokeColor(black);c.setLineWidth(.6);c.rect(28,center-6,11,12,fill=1)
            hatch(c,box(28,center-6,39,center+6),BELOW_HATCH,spacing=3)
        for line in lines:text(c,46,y,line,8,'Arial');y-=10
        legend_entries.append({'key':kind,'text':caption});y-=6
    row_height=min(17,(y-155)/11)
    if row_height<13:raise MapError('Shorten the subdivision or case description so the complete PC legend fits.')
    for lu in LEGEND_USES:
        c.setFillColor(HexColor(COLORS[lu]));c.setStrokeColor(black);c.setLineWidth(.7);c.rect(28,y-4,11,11,fill=1)
        if lu=='TRN':hatch(c,box(28,y-4,39,y+7),TRANSPORT_HATCH,spacing=3,descending=True)
        text(c,46,y-1,NAMES[lu],8,'Arial')
        legend_entries.append({'key':lu,'text':NAMES[lu],'color':COLORS[lu],'bounds':[28,y-4,39,y+7]});y-=row_height
    c.setStrokeColor(black);c.setLineWidth(1.7);c.rect(28,y-4,11,11,fill=0)
    text(c,46,y-1,'Proposed Application Area',8,'Arial');legend_entries.append({'key':'boundary','text':'Proposed Application Area'})
    notes=['Source: Harris Central Appraisal District', f'Date: {datetime.now():%B %d, %Y}',
           'Reference: '+case.reference, 'Support / minimum: calculation sheet', 'Protests: planner-reviewed properties']
    if data.get('pc_land_use_years'):notes.append('Surrounding land use: City '+', '.join(data['pc_land_use_years']))
    if 'UNK' in used:notes.append('White parcels: land use unavailable')
    note_y=139
    for value in notes:
        for line in wrap(value,150,6):text(c,26,note_y,line,6,'Arial');note_y-=8
    disclaimer='This map is made available for reference purposes only and should not be substituted for a survey product. The City of Houston will not accept liability of any kind in conjunction with its use.'
    for line in wrap(disclaimer,150,5.5):text(c,26,note_y-3,line,5.5,'Arial');note_y-=7
    c.drawImage(city_seal(),26,23,150,max(28,min(49,note_y-27)),preserveAspectRatio=True,anchor='c')
    c.save()
    return {'pdf':output.getvalue(),'labels':labels,'editable_labels':labels+legend,'issues':list(dict.fromkeys(issues)),
            'blocking':[],'rejected_offsets':rejected,'scale':scale,'scale_feet':feet,'footer_lines':[x['text'] for x in legend],
            'selected_count':len(raw),'page_transform':matrix,'rotation_origin':origin,'rotation':angle,
            'pc_symbols':symbol_records,'pc_land_uses':land_use_records,'legend_entries':legend_entries,'page_size':PAGE}

def save_pc_map(data,settings,rendered,output):
    folder=Path(output);folder.mkdir(parents=True,exist_ok=True)
    stem=data['case']['reference'].replace(' ','_')+'_PC_Map_'+datetime.now().strftime('%Y%m%d_%H%M%S_%f')
    pdf=folder/(stem+'.pdf');written=[]
    metadata={'schema':1,'kind':'MLS_PC_MAP','case':data['case'],'settings':settings,
              'warnings':rendered['issues'],'pc_facts':data.get('pc_facts'),'sources':data['sources'],
              'pc_markings':[m.get('pc',{}) for m in data['matched']],
              'pc_land_uses':rendered['pc_land_uses'],'pc_legend_version':2}
    try:
        for suffix,content in [('.pdf',rendered['pdf']),('.png',pdf_png(rendered['pdf'],200)),('.boundary.json',json.dumps(metadata,indent=2).encode())]:
            path=folder/(stem+suffix)
            with path.open('xb') as f:written.append(path);f.write(content)
    except Exception:
        for path in written:path.unlink(missing_ok=True)
        raise
    return pdf
