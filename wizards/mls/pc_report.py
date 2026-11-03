"""Template-based MLS Planning Commission draft and attachment assembly."""
from __future__ import annotations
from datetime import datetime
from pathlib import Path
import io
import json
import tempfile
from docx import Document
from docx.shared import Inches, Pt
from docx.enum.section import WD_SECTION, WD_ORIENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from core.paths import templates_dir
from .pc_data import number, date_text
from .mailout_engine import number_to_words, ensure_pdf, export_step2_staff_analysis_pdf

def words_count(value, capital=False):
    if value is None:return '[review] ([review])'
    words=number_to_words(int(value))
    return f'{words.capitalize() if capital else words} ({value})'

def report_context(facts, options):
    ci=facts['case_info'];notes=list(facts['warnings'])
    def field(key,ci_key=None):
        value=str(options.get(key,ci.get(ci_key or key,''))).strip()
        if not value:notes.append('Complete '+key.replace('_',' ')+'.');value='[review]'
        return value
    protest=number(options.get('protest_count',ci.get('Protest_Count')))
    if protest is None or protest<0 or protest!=int(protest):
        protest_sentence='[Review the number and status of protests.]';notes.append('Confirm the protest count.')
    elif protest==0:protest_sentence='No protest was filed.'
    else:protest_sentence=f'{words_count(int(protest),True)} written '+('protest was' if protest==1 else 'protests were')+' filed.'
    protested_rows=options.get('protested_rows',[])
    if protest is not None and int(protest)!=len(protested_rows):
        notes.append('The written-protest count and marked property count differ. One protest may cover multiple properties; review their relationship.')
    eligible='eligible lots' if facts['excluded_count'] else 'lots'
    sfr=f"Land uses of the properties consist of {words_count(facts['sfr_count'])} of {words_count(facts['eligible_count'])} single-family residential properties (representing {facts['sfr_lot_percent']}% of the {eligible} within the boundary area)."
    if facts['excluded_count']:
        suffix='property is' if facts['excluded_count']==1 else 'properties are'
        sfr+=f" {words_count(facts['excluded_count'],True)} excluded {suffix} not included in this percentage."
    support=f"Owners signed in support for {words_count(facts['signed_count'])} of {words_count(facts['eligible_count'])} eligible properties in the proposed SMLSB (representing {facts['support_area_percent']}% of the eligible application area). {protest_sentence}"
    face_count=facts['blockfaces'];face_noun='blockface' if face_count==1 else 'blockfaces'
    minimum=facts['minimum_text']
    history=f"The subdivision was platted in {field('plat_year','Subdivision_Plat_Year')}. "
    neighborhood=str(options.get('neighborhood',ci.get('Neighborhood_Age_Statement',''))).strip()
    if neighborhood and neighborhood.rstrip('. ').casefold()!=history.rstrip('. ').casefold():history+=neighborhood.rstrip('.')+'. '
    history+=field('structure_age','Structure_Age_Statement').rstrip('.')+'. '
    history+=str(options.get('preservation',ci.get('Preservation_Finding',''))).strip() or f'The establishment of a {minimum} sf minimum lot size will preserve the lot size character of the area.'
    notice=options.get('notice_sent',False)
    if not notice:notes.append('Public-hearing notice transmission is not confirmed; the draft contains a review placeholder.')
    context={'agenda':field('agenda','PC_Agenda_Number'), 'application':facts['application'],'boundary':facts['boundary'],
             'minimum':minimum,'blockface_noun':face_noun,'support_area_percent':facts['support_area_percent'],
             'protest_sentence':protest_sentence,'all_lots':words_count(facts['all_count']),
             'qualifying_lots':words_count(facts['qualifying_count']), 'qualifying_lots_cap':words_count(facts['qualifying_count'],True),
             'eligible_lots':words_count(facts['eligible_count']),'qualifying_area_percent':facts['qualifying_area_percent'],
             'sfr_sentence':sfr,'support_sentence':support,'history_sentence':history,
             'blockface_sentence':f"The application comprises {number_to_words(face_count)} {face_noun}, the {facts['sides']} of {facts['street']}.",
             'notice_sentence':'Public notice of the public hearing was transmitted to all property owners on the blockfaces.' if notice else '[Confirm that public notice of the hearing was transmitted to all property owners on the blockfaces.]',
             'hearing_date':date_text(field('hearing_date','PC_Hearing_Date'))}
    return context,notes

def fill_template(context, wizard_key='mls'):
    from core.templates import template_path
    template=template_path(wizard_key, 'pc_draft_template', {'mls':'MLS','mls_renewal':'MLS_REN','mbl':'MBL','mbl_renewal':'MBL_REN'}[wizard_key]+'_PC_Draft_Report.docx')
    doc=Document(template)
    # Slots are stored in individual w:t nodes; other runs, criteria and furniture remain intact.
    for part in [doc.part,*[r.target_part for r in doc.part.rels.values() if not r.is_external and ('/header' in r.reltype or '/footer' in r.reltype)]]:
        for node in part.element.iter(qn('w:t')):
            text=node.text or ''
            for key,value in context.items():text=text.replace('{{'+key+'}}',str(value))
            if '\n' in text:
                node.text=text.split('\n')[0]
                run=node.getparent()
                from docx.oxml import OxmlElement
                for line in text.split('\n')[1:]:
                    run.append(OxmlElement('w:br'));t=OxmlElement('w:t');t.text=line;run.append(t)
            else:node.text=text
    return doc

def append_pdf(doc,path):
    import fitz
    pdf=fitz.open(path)
    try:
        for page in pdf:
            landscape=page.rect.width>page.rect.height
            section=doc.add_section(WD_SECTION.NEW_PAGE)
            section.orientation=WD_ORIENT.LANDSCAPE if landscape else WD_ORIENT.PORTRAIT
            section.page_width=Inches(11 if landscape else 8.5)
            section.page_height=Inches(8.5 if landscape else 11)
            section.left_margin=Inches(.5);section.right_margin=Inches(.75)
            section.top_margin=Inches(1.1);section.bottom_margin=Inches(.8)
            section.header_distance=section.footer_distance=Inches(.5)
            section.header.is_linked_to_previous=False
            section.footer.is_linked_to_previous=False
            # python-docx creates new empty parts: copy the exact preceding header/footer.
            import copy
            prior=doc.sections[-2]
            for old,new in [(prior.header,section.header),(prior.footer,section.footer)]:
                for element in list(new._element):new._element.remove(element)
                for element in old._element:new._element.append(copy.deepcopy(element))
            # Adjust source-derived page furniture to the section width.
            width=int((section.page_width-section.left_margin-section.right_margin)/635)
            for table in section.header.tables:
                for grid in table._tbl.xpath('./w:tblGrid/w:gridCol'):grid.set(qn('w:w'),str(width//2))
                for cell in table._tbl.xpath('.//w:tcW'):cell.set(qn('w:w'),str(width//2))
            for tabs in section.footer._element.xpath('.//w:tab[@w:pos]'):
                alignment=tabs.get(qn('w:val'))
                if alignment in ('right','center'):tabs.set(qn('w:pos'),str(width if alignment=='right' else int(width*.65)))
            pixel=page.get_pixmap(matrix=fitz.Matrix(2.4,2.4),alpha=False)
            p=doc.add_paragraph();p.alignment=WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.space_before=Pt(0);p.paragraph_format.space_after=Pt(0)
            p.paragraph_format.line_spacing=1
            maxw=(section.page_width-section.left_margin-section.right_margin)/914400
            maxh=(section.page_height-section.top_margin-section.bottom_margin)/914400-.15
            factor=min(maxw/page.rect.width,maxh/page.rect.height)
            p.add_run().add_picture(io.BytesIO(pixel.tobytes('png')),width=Inches(page.rect.width*factor),height=Inches(page.rect.height*factor))
    finally:pdf.close()

def generate_pc_report(workbook,facts,options,output):
    """Return editable Word, PDF when available, and review notes; never modify Excel."""
    from core.office_convert import office_to_pdf
    wizard_key=facts.get('wizard_key','mls')
    prefix={'mls':'MLS','mls_renewal':'MLS_REN','mbl':'MBL','mbl_renewal':'MBL_REN'}[wizard_key]
    context_reader=report_context
    calculation_export=export_step2_staff_analysis_pdf
    if wizard_key=='mls_renewal':
        from wizards.mls_renewal.pc_report import report_context as context_reader
        from wizards.mls_renewal.wizard import export_step2_staff_analysis_pdf as calculation_export
    if wizard_key in ('mbl','mbl_renewal'):
        from wizards.mbl.pc_report import report_context as context_reader
        if wizard_key=='mbl':
            from wizards.mbl.mbl_mailout_engine import export_mbl_step2_staff_analysis_pdf as calculation_export
        else:
            from wizards.mbl_renewal.wizard import export_step2_staff_analysis_pdf as calculation_export
    folder=Path(output)/(prefix+'_'+facts['application']+'_PC_Draft_'+datetime.now().strftime('%Y%m%d_%H%M%S_%f'))
    folder.mkdir(parents=True,exist_ok=False)
    context,notes=context_reader(facts,options)
    if options.get('pc_map'):
        try:
            metadata=json.loads(Path(options['pc_map']).with_suffix('.boundary.json').read_text(encoding='utf8'))
            if metadata.get('kind')=='MLS_PC_MAP':
                notes.extend(metadata.get('warnings',[]))
                if metadata.get('case',{}).get('workbook_hash')!=facts['workbook_hash']:
                    notes.append('The selected PC map was made from a different saved workbook. It is attached as selected; review its facts.')
                marked={m['row'] for m in metadata.get('pc_markings',[]) if m.get('protested')}
                if marked!=set(options.get('protested_rows',[])):
                    notes.append('The selected PC map has different protest markings from this review. Generate a new PC map if the new selections should appear.')
        except (OSError,ValueError):
            notes.append('The selected PC map has no readable wizard metadata. Review its support, protest and minimum symbols.')
    with tempfile.TemporaryDirectory(prefix='rca_pc_report_') as temporary:
        temp=Path(temporary);attachments=[]
        if options.get('include_calculation',True):
            try:
                calculation=calculation_export(workbook,temp/'Calculation_Analysis.pdf')
                attachments.append(('Calculation Analysis',calculation))
            except Exception as error:notes.append('Calculation attachment could not be exported: '+str(error))
        for key,title in [('pc_map','Map of support, land use, meets minimum.'),('application_pdf','Application'),('boundary_map','Boundary Map')]:
            path=options.get(key,'')
            if path:
                try:attachments.append((title,ensure_pdf(path,temp)))
                except Exception as error:notes.append(title+' was not attached: '+str(error))
            else:notes.append(title+' was not selected and is not attached.')
        context['attachment_list']='\n'.join(f'{i}. {name}' for i,(name,_) in enumerate(attachments,1)) or 'No attachments selected.'
        doc=fill_template(context,wizard_key)
        for _,path in attachments:append_pdf(doc,path)
        document=folder/f"{prefix}_{facts['application']}_PC_Draft_Report.docx"
        doc.save(document)
        pdf=None
        try:pdf=office_to_pdf(document,folder)
        except Exception as error:notes.append('Word report saved. PDF export was unavailable: '+str(error))
    notes_path=folder/'PC_Draft_Review_Notes.txt'
    notes_path.write_text('\n'.join(notes) if notes else 'No missing report fields or attachments.',encoding='utf8')
    (folder/'PC_Draft_Data.json').write_text(json.dumps({'facts':facts,'options':options,'context':context,'warnings':notes},indent=2),encoding='utf8')
    return {'docx':str(document),'pdf':str(pdf) if pdf else '', 'notes':notes, 'folder':str(folder)}
