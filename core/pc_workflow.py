"""Use each maintained suite's existing calculation/map preparation workflow."""
from pathlib import Path

def prepare_owner_calculation(owner):
    try:
        from .automatic_case import prepare_owner_calculation as prepare
    except ModuleNotFoundError:
        from .boundary_map_ui import wizard_paths
        source, _, output = wizard_paths(owner)
        if not source or not Path(source).is_file():raise ValueError('Select a saved calculation sheet first.')
        if not output:raise ValueError('Select an output folder first.')
        return source
    return prepare(owner)

def load_map_data(source, wizard, output, progress=None, cancel=None, use_cache=True):
    try:
        from .automatic_case import load_map_data as load
    except ModuleNotFoundError:
        from .boundary_map import read_map_case, fetch_map_data, PublicMapClient
        from .workbooks import WorkbookReadCache
        resolved=WorkbookReadCache().resolve(source)
        case=read_map_case(resolved,wizard,source_path=source)
        client=PublicMapClient(Path(output)/'.boundary_map_cache',progress,cancel,use_cache=use_cache)
        data=fetch_map_data(case,client);data['calculation_path']=str(source)
        return data
    return load(source,wizard,output,progress,cancel,use_cache)
