"""Council districts from the City's published polygons and selected parcels."""
from datetime import datetime,timezone
from pathlib import Path
import json,re,zipfile
from shapely.ops import unary_union
from .boundary_map import PublicMapClient,parcel_polygon,MapError

SERVICE='https://mycity2.houstontx.gov/gisweb01/rest/services/HoustonMap/Administrative_Boundary/MapServer/2'
ROSTER_URL='https://www.houstontx.gov/council/'

def parse_members(content):
    from lxml import html
    tree=html.fromstring(content);members={}
    for heading in tree.xpath('//h2|//h3'):
        match=re.fullmatch(r'District\s+([A-K])',heading.text_content().strip(),re.I)
        if not match:continue
        for node in heading.xpath('following::*[self::p or self::h2 or self::h3]'):
            if node.tag in ('h2','h3'):break
            text=' '.join(node.text_content().split())
            if re.fullmatch(r"[A-Za-zÀ-ÿ.'’ -]{4,85}",text) and len(text.split())>=2 and text.upper() not in ('CONTACT CITY HALL','CITY COUNCIL'):
                members[match[1].upper()]=text.title() if text.isupper() else text;break
    if len(members)!=11:raise ValueError('The official council directory did not return all eleven district members.')
    return members

def add_members(result,cache=None,fetch=None):
    """Fresh official roster; dated fallback is explicit and never guessed."""
    from urllib.request import Request,urlopen
    cache=Path(cache)/'council_members.json' if cache else None
    try:
        if fetch:content=fetch(ROSTER_URL)
        else:
            with urlopen(Request(ROSTER_URL,headers={'User-Agent':'Houston-RCA-Wizard/1.0'}),timeout=20) as response:content=response.read(2_000_000)
        roster={'members':parse_members(content),'checked_utc':datetime.now(timezone.utc).isoformat(),'source':ROSTER_URL}
        if cache:cache.parent.mkdir(parents=True,exist_ok=True);cache.write_text(json.dumps(roster,indent=2),encoding='utf8')
    except Exception as exc:
        roster=None
        if cache and cache.is_file():
            try:roster=json.loads(cache.read_text(encoding='utf8'))
            except (ValueError,OSError):pass
        if roster:result['warnings'].append('Councilmember lookup unavailable; using roster checked '+roster.get('checked_utc','unknown date')+'. Confirm the member before issuing documents.')
        else:result['warnings'].append('Councilmember could not be verified; enter the name in Case Info. '+str(exc))
    result['member_value']='';result['members']={}
    if roster:
        result['members']={d:roster['members'].get(d,'') for d in result.get('districts',[])}
        result['member_value']='; '.join((f'District {d}: ' if len(result['members'])>1 else '')+n for d,n in result['members'].items() if n)
        result['member_source']=roster['source'];result['members_checked_utc']=roster['checked_utc']
    return result


def district_polygon(feature):
    """City boundary rings can contain touching edges; preserve their even/odd fill."""
    try:return parcel_polygon(feature)
    except MapError:
        from shapely import make_valid
        from shapely.geometry import Polygon,GeometryCollection
        result=GeometryCollection()
        def polygons(g):
            if g.geom_type=='Polygon':return [g]
            return [p for child in getattr(g,'geoms',[]) for p in polygons(child)]
        for ring in feature.get('geometry',{}).get('rings',[]):
            repaired=unary_union(polygons(make_valid(Polygon(ring))))
            result=result.symmetric_difference(repaired)
        result=unary_union(polygons(make_valid(result)))
        if result.is_empty or not result.is_valid or result.area<=0:
            raise MapError('The City council boundary could not be read.')
        return result


def resolve_districts(features,client):
    if not features:raise MapError('Select application properties before finding the council district.')
    parcels=[parcel_polygon(f) for f in features];area=unary_union(parcels)
    client.progress('Finding council districts from the selected properties...')
    records=client.query(SERVICE,extent=area.bounds,fields='OBJECTID,DISTRICT,YEAR')
    districts=[]
    for record in records:
        name=str(record.get('attributes',{}).get('DISTRICT','')).strip().upper()
        if re.fullmatch('[A-K]',name):districts.append((name,district_polygon(record)))
    matches=[];found=set();warnings=[]
    for index,(feature,parcel) in enumerate(zip(features,parcels)):
        # Shared edges have zero area. Ignore sub-square-foot overlay slivers,
        # while retaining real parcels/parts crossing more than one district.
        hits=[]
        for name,polygon in districts:
            overlap=parcel.intersection(polygon).area
            if overlap>max(.1,parcel.area*.0001):hits.append(name);found.add(name)
        attrs=feature.get('attributes',{})
        matches.append({'account':str(attrs.get('HCAD_NUM') or attrs.get('ACCOUNT') or ''),
                        'districts':sorted(set(hits)),'parcel_index':index})
        if not hits:warnings.append('A selected property did not match a City council district. Review the location and City limits.')
    covered=unary_union([g for _,g in districts]) if districts else None
    if covered is not None and area.difference(covered).area>max(1,area.area*.001):
        warnings.append('Part of the application area falls outside the returned district boundaries. Review before using the district value.')
    names=sorted(found)
    if len(names)>1:warnings.append('The application spans council districts '+', '.join(names)+'. Review both districts in Case Info.')
    return {'districts':names,'value':', '.join(names),'matches':matches,'warnings':list(dict.fromkeys(warnings)),
            'source':SERVICE,'retrieved_utc':datetime.now(timezone.utc).isoformat(),'sources':list(client.sources)}


def resolve_selected(rows,cache,progress=None,cancel=None):
    features=[r['feature'] for r in rows if r.get('include',True)]
    return add_members(resolve_districts(features,PublicMapClient(cache,progress,cancel,use_cache=False)),cache)


def write_case_district_copy(source,destination,result):
    """Change only the existing Case Info input, preserving all formulas/artwork."""
    from openpyxl import load_workbook
    from lxml import etree as ET
    from .application_import import sheet_paths,set_cell
    source,destination=Path(source),Path(destination)
    if not result.get('value'):raise MapError('No verified council district is available to fill.')
    if source.resolve()==destination.resolve() or destination.exists():raise MapError('Choose a new calculation sheet filename.')
    from .workbooks import read_shared_bytes
    import io
    raw=read_shared_bytes(source)
    book=load_workbook(io.BytesIO(raw),read_only=True,data_only=False)
    try:
        sheet=book['Case Info']
        inputs={str(row[0].value or '').strip().replace(' ','_').casefold():row[1].coordinate for row in sheet.iter_rows(min_col=1,max_col=2)}
        cell=inputs.get('council_district');member_cell=inputs.get('council_member')
        last_row=sheet.max_row
    finally:book.close()
    with zipfile.ZipFile(io.BytesIO(raw)) as original:
        path=sheet_paths(original)['Case Info'];xml=ET.fromstring(original.read(path))
        if not cell:
            last_row+=1;cell=f'B{last_row}';set_cell(xml,f'A{last_row}','Council_District')
        if not member_cell:
            last_row+=1;member_cell=f'B{last_row}';set_cell(xml,f'A{last_row}','Council_Member')
        set_cell(xml,cell,result['value']);set_cell(xml,member_cell,result.get('member_value') or None)
        dimension=xml.find('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}dimension')
        if dimension is not None:
            from openpyxl.utils.cell import range_boundaries,get_column_letter
            x0,y0,x1,y1=range_boundaries(dimension.get('ref'))
            dimension.set('ref',f'{get_column_letter(x0)}{y0}:{get_column_letter(max(2,x1))}{max(y1,last_row)}')
        with zipfile.ZipFile(destination,'x',zipfile.ZIP_DEFLATED) as output:
            for member in original.infolist():output.writestr(member,ET.tostring(xml) if member.filename==path else original.read(member.filename))
    return destination
