"""V35's automatic lot allocation and council information, shared by all wizards."""
from __future__ import annotations

import hashlib
import io
import os
from pathlib import Path
import tempfile
import threading

from .workbooks import read_shared_bytes, is_bundled_input, WorkbookReadCache


def _digest(raw):
    return hashlib.sha256(raw).hexdigest()


def _commit(source, before, updated, output_folder=None):
    """Back up then replace atomically; use a selected copy if Excel holds the file."""
    source=Path(source).resolve()
    if read_shared_bytes(source)!=before:
        raise ValueError('The calculation sheet changed during its update. Finish saving it in Excel, then retry.')
    if updated==before:return source,[]
    if not is_bundled_input(source):
        try:
            backup_dir=source.with_name(source.stem+'_sources')/'Workbook backups'
            backup_dir.mkdir(parents=True,exist_ok=True)
            backup=backup_dir/(_digest(before)+'.xlsx')
            if not backup.exists():
                with backup.open('xb') as stream:stream.write(before)
            with tempfile.TemporaryDirectory(prefix='.wizard_update_',dir=source.parent) as tmp:
                candidate=Path(tmp)/source.name;candidate.write_bytes(updated)
                if read_shared_bytes(source)!=before:
                    raise ValueError('The calculation sheet changed during its update. Finish saving it in Excel, then retry.')
                os.replace(candidate,source)
            return source,[]
        except OSError:
            # Never force-close Excel or overwrite its unsaved work.
            pass
    folder=Path(output_folder) if output_folder else source.parent
    folder.mkdir(parents=True,exist_ok=True)
    stem=source.stem+'_Updated'
    for number in range(1,10000):
        destination=folder/(stem+('' if number==1 else f'_{number}')+'.xlsx')
        if destination.exists():continue
        if read_shared_bytes(source)!=before:
            raise ValueError('The calculation sheet changed during its update. Finish saving it in Excel, then retry.')
        try:
            with destination.open('xb') as stream:stream.write(updated)
        except FileExistsError:continue
        return destination,[f'Updated calculation sheet saved as {destination.name} and selected in the wizard. The original was open, read-only, or a bundled example.']
    raise OSError('Could not choose a filename for the updated calculation sheet.')


def ensure_lot_rows(source,wizard,output_folder=None):
    """Split only unsplit whole-lot descriptions. Already separated rows stay intact."""
    from .lot_divisions import workbook_rows,proposed_lots,reviewed_workbook_copy,LotExpansionUnavailable
    source=Path(source).resolve();before=read_shared_bytes(source)
    with tempfile.TemporaryDirectory(prefix='rca_lot_allocation_') as tmp:
        snapshot=Path(tmp)/'source.xlsx';snapshot.write_bytes(before)
        rows,metadata=workbook_rows(snapshot)
        split_count=0;lot_count=0
        for row in rows:
            lots=proposed_lots(row)
            if lots:
                row['lots']=lots;split_count+=1;lot_count+=len(lots)
        if not split_count:return {'path':str(source),'notes':[],'split_properties':0}
        destination=Path(tmp)/'allocated.xlsx'
        try:reviewed_workbook_copy(snapshot,destination,rows,metadata,wizard)
        except LotExpansionUnavailable as exc:
            return {'path':str(source),'notes':[str(exc)],'split_properties':0}
        path,notes=_commit(source,before,destination.read_bytes(),output_folder)
    notes.insert(0,f'{split_count} properties separated into {lot_count} lot rows. Each property area is divided equally among its listed lots.')
    return {'path':str(path),'notes':notes,'split_properties':split_count}


def select_updated_calculation(owner,path,expected=None):
    """Keep the selected calculation and attachments in sync without resetting the page."""
    def apply():
        from .boundary_map_ui import wizard_paths
        current=wizard_paths(owner)[0]
        if expected and str(current)!=str(expected):return
        if str(current)==str(path):return
        if hasattr(owner,'vars'):
            owner.vars['excel'].set(str(path));owner.summary=None
        else:owner.files['workbook']=str(path)
        owner._workbook_read_cache=WorkbookReadCache()
        if hasattr(owner,'source'):owner.source=str(path)
        if hasattr(owner,'_save_session'):owner._save_session()
        elif hasattr(owner,'owner'):select_updated_calculation(owner.owner,path,expected)
    if threading.current_thread() is threading.main_thread():apply()
    else:owner.after(0,apply)


def prepare_owner_calculation(owner):
    from .boundary_map_ui import wizard_paths
    source,_,output=wizard_paths(owner)
    if not source:raise ValueError('Select a calculation workbook first.')
    result=ensure_lot_rows(source,owner._boundary_map_wizard_key,output)
    select_updated_calculation(owner,result['path'],source)
    owner._calculation_update_notes=result['notes']
    return result['path']


def _case_council(raw):
    from openpyxl import load_workbook
    book=load_workbook(io.BytesIO(raw),read_only=True,data_only=True)
    try:
        if 'Case Info' not in book.sheetnames:raise ValueError('The calculation sheet needs a Case Info tab for council information.')
        return {str(a or '').strip().replace(' ','_').casefold():str(b or '').strip()
                for a,b in book['Case Info'].iter_rows(max_col=2,values_only=True)}
    finally:book.close()


def save_map_council(source,result,output_folder=None,expected_hash=None):
    """Change only the two council inputs. Preserve a name if its district is unchanged."""
    from .council_district import write_case_district_copy
    source=Path(source).resolve();before=read_shared_bytes(source)
    if expected_hash and _digest(before)!=expected_hash:
        raise ValueError('The calculation sheet changed while the map was loading. Recreate the map to update Case Info.')
    if not result.get('value'):return source,[]
    existing=_case_council(before);values=dict(result)
    if not values.get('member_value') and existing.get('council_district')==values['value']:
        values['member_value']=existing.get('council_member','')
    if existing.get('council_district')==values['value'] and existing.get('council_member')==(values.get('member_value') or ''):
        return source,[]
    with tempfile.TemporaryDirectory(prefix='rca_council_update_') as tmp:
        snapshot=Path(tmp)/'source.xlsx';snapshot.write_bytes(before)
        destination=Path(tmp)/'council.xlsx'
        write_case_district_copy(snapshot,destination,values)
        return _commit(source,before,destination.read_bytes(),output_folder)


def load_map_data(source,wizard,output,progress=None,cancel=None,use_cache=True):
    """One map-load workflow for boundary, PC and tax previews, including Case Info."""
    from .boundary_map import fingerprint,read_map_case,PublicMapClient,fetch_map_data,MapError
    from .council_district import resolve_districts,add_members
    progress=progress or (lambda message:None)
    progress('Preparing the calculation sheet and individual lot rows...')
    prepared=ensure_lot_rows(source,wizard,output);source=prepared['path']
    cache=WorkbookReadCache();resolved=cache.resolve(source)
    source_hash=fingerprint(source)
    case=read_map_case(resolved,wizard,source_path=source)
    client=PublicMapClient(Path(output)/'.boundary_map_cache',progress,cancel,use_cache=use_cache)
    data=fetch_map_data(case,client)
    data['calculation_path']=source;data['calculation_updates']=list(prepared['notes'])
    if fingerprint(source)!=source_hash:
        raise MapError('The calculation sheet changed while its map was loading. Close this preview and create it again.')
    try:
        client.check_cancelled()
        progress('Finding the council district and councilmember for Case Info...')
        council_client=PublicMapClient(Path(output)/'.council_cache',progress,cancel,use_cache=False)
        result=resolve_districts([m['feature'] for m in data['matched']],council_client)
        if result.get('value'):result=add_members(result,council_client.cache_dir)
        client.check_cancelled()
        data['council']=result;data['warnings'].extend(result.get('warnings',[]))
        if result.get('value'):
            path,notes=save_map_council(source,result,output,source_hash)
            data['calculation_path']=str(path)
            data['case']['workbook_hash']=fingerprint(path)
            data['calculation_updates'].extend(notes)
            data['council_saved']=True
        else:data['warnings'].append('Council information could not be determined for the selected properties. The map is still available.')
    except Exception as exc:
        client.check_cancelled()
        data['warnings'].append('The map is available, but council information could not be saved to Case Info: '+str(exc))
    if wizard.endswith('_renewal'):
        try:
            from .designation_lookup import lookup,workbook_fields,save_workbook
            current=data['calculation_path'];expected=data['case']['workbook_hash']
            if fingerprint(current)!=expected:
                raise ValueError('The calculation sheet changed after its parcels were loaded. Renewal fields were not updated.')
            result=lookup([m['feature'] for m in data['matched']],wizard,workbook_fields(current),
                          Path(output)/'.renewal_lookup_cache',progress,cancel,refresh=not use_cache)
            client.check_cancelled()
            data['renewal_lookup']=result
            path,notes=save_workbook(current,result,output,expected)
            data['calculation_path']=str(path);data['case']['workbook_hash']=fingerprint(path)
            data['calculation_updates'].extend(notes)
            data['warnings'].extend(result['warnings'])
        except Exception as exc:
            client.check_cancelled()
            data['warnings'].append('The map is available, but previous ordinance information could not be saved: '+str(exc))
    return data


def council_summary(data):
    result=data.get('council',{})
    text=''
    if data.get('council_saved'):
        line='Council district '+result['value']
        if result.get('member_value'):line+=': '+result['member_value']
        text='\n\n'+line+'\nSaved to Case Info.'
    renewal=data.get('renewal_lookup',{})
    if renewal.get('values',{}).get('original_ordinance_no'):
        text+='\n\nPrevious ordinance '+renewal['values']['original_ordinance_no']+'\nSee renewal details below.'
    return text
