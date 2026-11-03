"""Separate legal lots by equal area while retaining their shared HCAD account."""
from __future__ import annotations
import copy, io, json, math, re, zipfile
from pathlib import Path
from lxml import etree as ET
from openpyxl import load_workbook
from openpyxl.formula.translate import Translator
from openpyxl.utils import get_column_letter
from shapely.geometry import shape, mapping, LineString
from shapely.ops import unary_union, split

S='http://schemas.openxmlformats.org/spreadsheetml/2006/main'
N={'s':S}; META='_Wizard Lots'

class LotExpansionUnavailable(ValueError):
    """An older/fixed-value sheet cannot safely expand its calculation ranges."""

def clean(value):return re.sub(r'\s+',' ',str(value or '')).strip()

def parse_lots(description):
    """Only unambiguous whole lots. Partial tracts/remainders require review."""
    text=clean(description).upper().replace('–','-').replace('—','-')
    if re.search(r'\b(TRS?|TRACTS?|PTS?|PART|PORTION|LESS|EXCEPT|UND|UNDIVIDED|FT|FEET|ACRES?)\b',text):return []
    m=re.fullmatch(r'(?:LTS?|LOTS?)\s+(.+?)\s+(BLK\s+[A-Z0-9-]+(?:\s+.*)?)',text)
    if not m:return []
    expression=re.sub(r'\b(?:AND)\b','&',m[1])
    expression=re.sub(r'\b(?:THRU|THROUGH|TO)\b','-',expression)
    numbers=[]
    for part in re.split(r'\s*[,&]\s*',expression):
        part=part.strip()
        if re.fullmatch(r'\d+[A-Z]?',part):numbers.append(part)
        elif re.fullmatch(r'\d+\s*-\s*\d+',part):
            a,b=map(int,re.split(r'\s*-\s*',part))
            if b<a or b-a>200:return []
            numbers.extend(str(x) for x in range(a,b+1))
        else:return []
    if not numbers or len(set(numbers))!=len(numbers):return []
    return [{'number':n,'description':f'LT {n} {m[2]}'} for n in numbers]

def lot_key(account,description,subdivision=''):
    desc=re.sub(r'^L(?:O)?TS?\s+','LT ',clean(description).upper())
    return '|'.join((clean(account),desc,clean(subdivision).upper()))

def positive(value):
    if value is None or clean(value)=='':return None
    try:value=float(str(value).replace(',',''))
    except (TypeError,ValueError):raise ValueError('Lot areas must be numeric square feet.') from None
    if not math.isfinite(value) or value<=0:raise ValueError('Lot areas must be greater than zero.')
    return value

def proposed_lots(row):
    if row.get('lots') is not None:return copy.deepcopy(row['lots'])
    parsed=parse_lots(row.get('legal_description',''))
    if len(parsed)<2:return []
    count=len(parsed)
    area=allocated_area(row.get('land_area'),count)
    original=allocated_area(row.get('original_lot_area'),count)
    reference=f'Property area divided equally among {count} listed lots'
    return [{**x,'area':area,'land_use':row.get('land_use',''),'area_source':reference,
             'allocation':'equal','original_lot_area':original,
             'original_building_line':row.get('original_building_line'),
             'building_line':row.get('building_line')} for x in parsed]


def allocated_area(value,count):
    """Do not round individual shares and change the property's total."""
    if value in (0,'0'):return 0.0
    area=positive(value)
    return area/count if area is not None else None


def lot_area(lot):
    if lot.get('allocation')=='equal' and lot.get('area') in (0,'0'):return 0.0
    return positive(lot.get('area'))

def validate_lots(lots,total=None,allow_unknown=False):
    if not lots:raise ValueError('Enter at least one lot.')
    keys=[];areas=[]
    for lot in lots:
        desc=clean(lot.get('description'))
        if not desc:raise ValueError('Every lot needs its individual legal description.')
        keys.append(lot_key('',desc))
        area=lot_area(lot);areas.append(area)
        if area is None and not allow_unknown:raise ValueError('Enter each lot area, or save this split as a draft with unknown areas.')
    if len(set(keys))!=len(keys):raise ValueError('Each lot description must be distinct within this property.')
    if total and all(a is not None for a in areas) and abs(sum(areas)-float(total))>max(.1,float(total)*.00001):
        raise ValueError(f'Lot areas total {sum(areas):,.2f} sq ft; the property has {float(total):,.2f}. Resolve the difference or correct the property total first.')

def expand_parcels(rows):
    result=[];identities=set()
    for parent in rows:
        lots=parent.get('lots') or (proposed_lots(parent) if parent.get('include') else [])
        if not lots:
            children=[copy.deepcopy(parent)]
        else:
            validate_lots(lots,parent.get('land_area'),allow_unknown=True)
            children=[]
            for lot in lots:
                p=copy.deepcopy(parent);p.pop('lots',None)
                p.update(legal_description=clean(lot['description']),land_area=lot_area(lot),
                         land_use=lot.get('land_use',parent.get('land_use','')),
                         parent_legal_description=parent.get('legal_description',''),parent_land_area=parent.get('land_area'),
                         lot_id=lot_key(parent['account'],lot['description'],parent.get('subdivision','')),
                         area_source={'method':'Equal division' if lot.get('allocation')=='equal' else 'Reviewed individual lot',
                                      'reference':lot.get('area_source',''),'property_area':parent.get('land_area'),'lot_count':len(lots)})
                p['original_lot_area']=lot.get('original_lot_area')
                p['original_building_line']=lot.get('original_building_line')
                p['building_line']=lot.get('building_line')
                if p['land_area'] is None:p.setdefault('notes',[]).append('Individual lot area is unknown; complete Step 1 before relying on calculations.')
                children.append(p)
        for p in children:
            key=lot_key(p['account'],p.get('legal_description'),p.get('subdivision'))
            if p.get('include') and key in identities:raise ValueError('The same account and lot appears twice: '+key)
            if p.get('include'):identities.add(key)
            result.append(p)
    return result

def read_metadata(book):
    if META not in book.sheetnames:return {'schema':1,'groups':{}}
    try:
        content=''.join(str(r[0] or '') for r in book[META].iter_rows(min_row=2,max_col=1,values_only=True))
        obj=json.loads(content)
        return obj if obj.get('schema')==1 else {'schema':1,'groups':{}}
    except (ValueError,TypeError):return {'schema':1,'groups':{}}

def put_metadata(parts,metadata):
    """Embed editable provenance without adding columns or changing template layout."""
    from .application_import import sheet_paths,set_cell
    with io_zip(parts) as archive:paths=sheet_paths(archive)
    if META in paths:path=paths[META]
    else:
        wb=ET.fromstring(parts['xl/workbook.xml']);rels=ET.fromstring(parts['xl/_rels/workbook.xml.rels'])
        sheets=wb.find('s:sheets',N);sid=max(int(s.get('sheetId')) for s in sheets)+1
        rid='rIdLots'+str(sid)
        while any(r.get('Id')==rid for r in rels):rid+='x'
        path=f'xl/worksheets/wizard_lots_{sid}.xml'
        ET.SubElement(sheets,f'{{{S}}}sheet',name=META,sheetId=str(sid),state='hidden',attrib={'{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id':rid})
        ET.SubElement(rels,'{http://schemas.openxmlformats.org/package/2006/relationships}Relationship',Id=rid,Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet',Target=path[3:])
        types=ET.fromstring(parts['[Content_Types].xml'])
        ET.SubElement(types,'{http://schemas.openxmlformats.org/package/2006/content-types}Override',PartName='/'+path,ContentType='application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml')
        parts['xl/workbook.xml']=ET.tostring(wb);parts['xl/_rels/workbook.xml.rels']=ET.tostring(rels);parts['[Content_Types].xml']=ET.tostring(types)
    root=ET.Element(f'{{{S}}}worksheet',nsmap={None:S});ET.SubElement(root,f'{{{S}}}sheetData')
    set_cell(root,'A1','Wizard lot review and map geometry (EPSG:2278)')
    content=json.dumps(metadata,ensure_ascii=True,separators=(',',':'))
    for i,pos in enumerate(range(0,len(content),28000),2):set_cell(root,f'A{i}',content[pos:pos+28000])
    parts[path]=ET.tostring(root)

def io_zip(parts):
    buffer=io.BytesIO()
    with zipfile.ZipFile(buffer,'w') as out:
        for k,v in parts.items():out.writestr(k,v)
    return zipfile.ZipFile(io.BytesIO(buffer.getvalue()))

def write_parts(parts,destination):
    """New files only; invalidate calculated caches, never source workbook data."""
    from .application_import import sheet_paths
    wb=ET.fromstring(parts['xl/workbook.xml']);calc=wb.find('s:calcPr',N)
    if calc is None:calc=ET.SubElement(wb,f'{{{S}}}calcPr')
    calc.attrib.update({'calcMode':'auto','fullCalcOnLoad':'1','forceFullCalc':'1'})
    parts['xl/workbook.xml']=ET.tostring(wb)
    for key in list(parts):
        if key.startswith('xl/worksheets/') and key.endswith('.xml'):
            root=ET.fromstring(parts[key])
            for c in root.findall('.//s:c[s:f]',N):
                value=c.find('s:v',N)
                if value is not None:c.remove(value)
            parts[key]=ET.tostring(root)
    parts.pop('xl/calcChain.xml',None)
    for key,tag,attr in [('xl/_rels/workbook.xml.rels','Relationship','Type'),('[Content_Types].xml','Override','PartName')]:
        root=ET.fromstring(parts[key])
        for child in list(root):
            if 'calcChain' in child.get(attr,''):root.remove(child)
        parts[key]=ET.tostring(root)
    with Path(destination).open('xb') as stream:
        with zipfile.ZipFile(stream,'w',zipfile.ZIP_DEFLATED) as out:
            for k,v in parts.items():out.writestr(k,v)

def geometry_lots(record,parent,expected=None):
    """Reject overlap, gaps, wrong identities and geometry from another parcel."""
    if not record or not record.get('reviewed') or not clean(record.get('source')):return []
    if record.get('crs')!='EPSG:2278':raise ValueError('Lot geometry must use EPSG:2278 (US survey feet).')
    lots=record.get('lots',[]);polys=[];ids=[]
    for lot in lots:
        p=shape(lot['geometry'])
        if p.geom_type!='Polygon' or not p.is_valid or p.area<=.1:raise ValueError('Each lot must be a valid polygon.')
        polys.append(p);ids.append(lot['id'])
    if len(ids)!=len(set(ids)) or (expected is not None and set(ids)!=set(expected)):raise ValueError('Map lot identities do not match the calculation sheet. Review the divisions again.')
    if not polys:return []
    combined=unary_union(polys);tol=max(1,parent.area*.002)
    if sum(p.area for p in polys)-combined.area>tol:raise ValueError('The lot divisions overlap.')
    if parent.symmetric_difference(combined).area>tol:raise ValueError('The lot divisions do not cover this HCAD parcel. Review the plat alignment.')
    return list(zip(lots,polys))

def split_geometry(polygons,points):
    line=LineString(points)
    if not line.is_valid or line.length<.01:raise ValueError('Draw a dividing line across the property.')
    result=[]
    for p in polygons:result.extend(split(p,line).geoms)
    if len(result)==len(polygons):raise ValueError('The line must cross the lot from edge to edge. Extend both ends outside the parcel.')
    if any(p.area<.1 for p in result):raise ValueError('That line would create a sliver. Adjust its endpoints.')
    return result

def map_lot_records(data,settings):
    records=copy.deepcopy(data.get('case',{}).get('lot_metadata',{}).get('groups',{}))
    records.update(settings.get('lot_divisions',{}))
    return records

def interior_lines(data,settings):
    from .boundary_map import parcel_polygon
    result=[];notes=[]
    if not settings.get('show_platted_lot_lines',True):return result,notes
    records=map_lot_records(data,settings)
    for match in data['matched']:
        p=match['property'];record=records.get(p['account']);rows=p.get('lot_rows',[])
        expected=[r['id'] for r in rows] if rows else None
        if not record:continue
        try:
            parent=parcel_polygon(match['feature']);lots=geometry_lots(record,parent,expected)
            if not lots:continue
            edges=unary_union([poly.boundary for _,poly in lots]).difference(parent.boundary.buffer(.01))
            if not edges.is_empty:result.append(edges)
        except (ValueError,KeyError,TypeError) as exc:notes.append(p['address']+': lot divisions not drawn. '+str(exc))
    return result,notes

def line_parts(geometry):
    if geometry.geom_type in ('LineString','LinearRing'):return [geometry]
    return [line for g in getattr(geometry,'geoms',[]) for line in line_parts(g)]

def workbook_rows(path):
    from .workbooks import read_shared_bytes
    from .boundary_map import account_number
    data=read_shared_bytes(path);book=load_workbook(io.BytesIO(data),read_only=True,data_only=True)
    try:
        ws=next(s for s in book if s.title.lower().startswith('step 1'))
        heads={clean(c):i for i,c in enumerate(next(ws.iter_rows(max_row=1,values_only=True))) if c}
        s2=next((s for s in book if s.title.lower().startswith('step 2')),None)
        measurements={};first=None
        if s2 is not None:
            for header in (14,16):
                for cells in s2.iter_rows(min_row=header,max_row=header,values_only=True):
                    for col,title in enumerate(cells[:5]):
                        title=clean(title).lower()
                        key=('original_lot_area' if 'original lot size' in title else
                             'original_building_line' if 'original building line' in title else
                             'building_line' if 'building line' in title and 'ranked' not in title else None)
                        if key:measurements[key]=col;first=header+1
        step2=list(s2.iter_rows(min_row=first,values_only=True)) if first else []
        rows=[]
        for n,row in enumerate(ws.iter_rows(min_row=2,values_only=True),2):
            def value(*names):
                return next((row[heads[k]] for k in names if k in heads and row[heads[k]] not in (None,'')),None)
            raw=value('HCAD_NUM','ALT_TAX_ID','TAX_ID','HCAD_NUM_1')
            if not raw:continue
            account=account_number(raw);desc=clean(value('Legal_Dscr_1'))
            address=clean(value('Site_addr_1')) or ' '.join(clean(value(k)) for k in ('Str_num','Str_pfx','Str_name','Str_sfx') if value(k))
            rows.append({'row':n,'account':account,'address':address,'legal_description':desc,
                         'subdivision':clean(value('Legal_Dscr_2')),'land_area':value('Total_Land_Area','Lot Size'),
                         'land_use':clean(value('Land Use')),'include':True,
                         **{key:step2[n-2][col] for key,col in measurements.items() if n-2<len(step2) and col<len(step2[n-2])}})
        return rows,read_metadata(book)
    finally:book.close()

def reviewed_workbook_copy(source,destination,rows,metadata=None,wizard='mls'):
    """Expand physical rows; preserve City formulas, images and every other ZIP part."""
    from .workbooks import read_shared_bytes
    from .application_import import sheet_paths,set_cell
    raw=read_shared_bytes(source)
    with zipfile.ZipFile(io.BytesIO(raw)) as original:parts={n:original.read(n) for n in original.namelist()}
    with io_zip(parts) as original:paths=sheet_paths(original)
    s1name=next(k for k in paths if k.lower().startswith('step 1'));s2name=next(k for k in paths if k.lower().startswith('step 2'))
    book=load_workbook(io.BytesIO(raw),read_only=True,data_only=False)
    try:
        heads={clean(c.value).rstrip(' *'):c.column_letter for c in book[s1name][1] if c.value}
        source_meta=read_metadata(book)
        s2=book[s2name];first=15 if wizard.endswith('_renewal') else 17
        headings={c.column_letter:clean(c.value).lower() for c in s2[first-1] if c.value}
        rank_columns={col for col,title in headings.items() if 'ranked' in title}
        input_limit=4 if wizard=='mls' else 5
        # Actual template formula coverage, not a guessed row count.
        covered=[r[0].row for r in s2.iter_rows(min_row=first,max_col=1) if r[0].data_type=='f' and 'Step 1' in str(r[0].value)]
        capacity=max(covered)-first+1 if covered else 0
    finally:book.close()
    expanded=[]
    for row in rows:
        lots=row.get('lots') or proposed_lots(row)
        if lots:
            validate_lots(lots,row.get('land_area'),allow_unknown=True)
            for lot in lots:expanded.append((row,lot,len(lots)>1))
        else:
            expanded.append((row,{'description':row['legal_description'],'area':row['land_area'],'land_use':row['land_use']},False))
    if not capacity:raise LotExpansionUnavailable('This sheet has no Step 2 row formulas, so its lot rows were kept unchanged. Import the application into the current calculation template to split its lots automatically.')
    if len(expanded)>capacity:raise LotExpansionUnavailable(f'This template has room for {capacity} lots; the split needs {len(expanded)}. Its rows were kept unchanged. Use an expanded City template to split these lots.')
    roots={name:ET.fromstring(parts[paths[name]]) for name in (s1name,s2name)}
    old_rows={name:{int(r.get('r')):copy.deepcopy(r) for r in root.findall('s:sheetData/s:row',N)} for name,root in roots.items()}
    def shifted(row,new_index):
        out=copy.deepcopy(row);out.set('r',str(new_index))
        for c in out:
            old=c.get('r');col=re.sub(r'\d+','',old);new=col+str(new_index);c.set('r',new)
            f=c.find('s:f',N)
            if f is not None and f.text:
                f.text=Translator('='+f.text,origin=old).translate_formula(new)[1:]
                f.attrib.clear()
        return out
    for name,start,end in [(s1name,2,capacity+1),(s2name,first,first+capacity-1)]:
        sheetdata=roots[name].find('s:sheetData',N)
        for row in list(sheetdata):
            if start<=int(row.get('r'))<=end:sheetdata.remove(row)
        for j,(source_row,lot,is_split) in enumerate(expanded):
            old_index=source_row['row'] if name==s1name else first+source_row['row']-2
            target=start+j
            original=old_rows[name].get(old_index)
            if original is None:raise ValueError('The calculation template is missing a property row.')
            adjusted=shifted(original,target)
            if name==s2name:
                # Ranking/formula helper columns belong to the destination row.
                # Only the property input columns travel with the source lot.
                adjusted=copy.deepcopy(old_rows[name][target])
                moved=shifted(original,target)
                for col in [get_column_letter(n) for n in range(1,input_limit+1) if get_column_letter(n) not in rank_columns]:
                    ref=col+str(target)
                    for c in list(adjusted):
                        if c.get('r')==ref:adjusted.remove(c)
                    cell=moved.find(f's:c[@r="{ref}"]',N)
                    if cell is not None:adjusted.append(copy.deepcopy(cell))
                adjusted[:]=sorted(adjusted,key=lambda c:__import__('openpyxl').utils.cell.coordinate_to_tuple(c.get('r'))[1])
            sheetdata.append(adjusted)
        # Retain empty City formula rows to keep formula ranges and future inputs intact.
        for target in range(start+len(expanded),end+1):
            prototype=old_rows[name].get(target)
            if prototype is None:prototype=old_rows[name].get(start)
            row=shifted(prototype,target)
            for c in row:
                column=re.sub(r'\d+','',c.get('r'))
                editable=name==s1name or (column not in rank_columns and __import__('openpyxl').utils.cell.column_index_from_string(column)<=input_limit)
                if editable and c.find('s:f',N) is None:
                    for child in list(c):c.remove(child)
                    c.attrib.pop('t',None)
            sheetdata.append(row)
        sheetdata[:]=sorted(sheetdata,key=lambda r:int(r.get('r')))
    for j,(source_row,lot,is_split) in enumerate(expanded,2):
        # Preserve existing zero-area excluded rows; they are not new split lots.
        area=0 if not is_split and lot.get('area') in (0,'0') else lot_area(lot)
        values={'Legal_Dscr_1':lot['description'],'Total_Land_Area':area,
                'Acreage':area/43560 if area is not None else None,
                'Lot Size':area,'Land Use':lot.get('land_use',source_row['land_use'])}
        # Property Description can be a formula or a pasted string.
        if 'Property Description' in heads:
            ref=heads['Property Description']+str(j)
            c=roots[s1name].find(f'.//s:c[@r="{ref}"]',N)
            if c is None or c.find('s:f',N) is None:
                number=source_row['address'].split(' ',1)[0] or '0'
                values['Property Description']=f"{number} ({lot['description']})"
        for key,value in values.items():
            if key in heads:
                ref=heads[key]+str(j);c=roots[s1name].find(f'.//s:c[@r="{ref}"]',N)
                if key=='Lot Size' and c is not None and c.find('s:f',N) is not None:continue
                set_cell(roots[s1name],ref,value)
        if is_split:
            target=first+j-2
            # A property's historical area/setback is not an individual lot measurement.
            for col,title in headings.items():
                key=None
                if 'original lot size' in title:key='original_lot_area'
                elif 'current lot size' in title:key='area'
                elif 'original building line' in title:key='original_building_line'
                elif 'building line' in title and ('current' in title or wizard=='mbl'):key='building_line'
                if key:
                    ref=col+str(target);c=roots[s2name].find(f'.//s:c[@r="{ref}"]',N)
                    if key=='area' and c is not None and c.find('s:f',N) is not None:continue
                    set_cell(roots[s2name],ref,lot.get(key))
    for name,root in roots.items():parts[paths[name]]=ET.tostring(root)
    meta=copy.deepcopy(metadata or source_meta)
    meta['rows']=[{'row':j,'account':r['account'],'id':lot_key(r['account'],l['description'],r['subdivision']),
                  'description':l['description'],'area_source':l.get('area_source','Existing calculation sheet'),
                  'unknown_area':l.get('area') is None} for j,(r,l,_) in enumerate(expanded,2)]
    put_metadata(parts,meta);write_parts(parts,destination)
    return Path(destination)

def overlay_tax_lots(pdf,data,settings,extent):
    import fitz
    from shapely.affinity import affine_transform
    from shapely.geometry import box
    lines,notes=interior_lines(data,settings)
    if not lines:return pdf,notes
    scale=540/(extent[2]-extent[0]);matrix=[scale,0,0,scale,36-extent[0]*scale,36-extent[1]*scale]
    with fitz.open(stream=pdf,filetype='pdf') as doc:
        if len(doc)!=1 or abs(doc[0].rect.width-792)>1 or abs(doc[0].rect.height-612)>1:raise ValueError('Unexpected HCAD print size; lot divisions cannot be aligned.')
        for geometry in lines:
            for line in line_parts(geometry.intersection(box(*extent))):
                points=[fitz.Point(x,612-y) for x,y in affine_transform(line,matrix).coords]
                doc[0].draw_polyline(points,color=(.2,.2,.2),width=.75,dashes='[3 2] 0',overlay=True)
        return doc.tobytes(garbage=3,deflate=True),notes

def expand_map_lots(data,settings):
    """PC symbols/colors can refer to each verified legal lot, not the whole account."""
    from .boundary_map import parcel_polygon
    records=map_lot_records(data,settings);result=copy.deepcopy(data);matches=[]
    for match in result['matched']:
        p=match['property'];rows=p.get('lot_rows',[]);record=records.get(p['account'])
        if not record or len(rows)<2:matches.append(match);continue
        try:lots=geometry_lots(record,parcel_polygon(match['feature']),[r['id'] for r in rows])
        except (ValueError,KeyError,TypeError):matches.append(match);continue
        if not lots:matches.append(match);continue
        for lot,polygon in lots:
            m=copy.deepcopy(match);m['property']={k:v for k,v in next(r for r in rows if r['id']==lot['id']).items() if k!='id'}
            m['property']['lot_rows']=[]
            m['platted_lot']=True
            m['feature']['geometry']={'rings':[list(polygon.exterior.coords)]+[list(r.coords) for r in polygon.interiors],'spatialReference':{'wkid':2278}}
            matches.append(m)
    result['matched']=matches
    return result

def collapse_lot_notifications(book,owners):
    """One notification per underlying property; different accounts remain separate."""
    metadata=read_metadata(book);rows={str(r['row']):r for r in metadata.get('rows',[])}
    if not rows:return owners
    result=[];seen={}
    for owner in owners:
        info=rows.get(str(owner.get('Excel_Row','')))
        if not info:result.append(owner);continue
        key=info['account']
        if key not in seen:seen[key]=copy.deepcopy(owner);result.append(seen[key]);continue
        previous=seen[key]
        for field in ('Property_Description','Legal_Dscr_1','PropertyDescription'):
            if owner.get(field) and owner[field] not in previous.get(field,''):previous[field]=previous.get(field,'')+'; '+owner[field]
        for field in ('Property_Lot_Size','Total_Land_Area'):
            try:previous[field]=str(float(str(previous.get(field,0)).replace(',',''))+float(str(owner.get(field,0)).replace(',','')))
            except (TypeError,ValueError):pass
    return result
