from typing import TypedDict, Dict, List, Any

class State(TypedDict):
    raw_file_content : str
    parsed_content : List[Dict[str, Any]]
    summarized_content : List[Dict[str,Any]]
    design_config : Dict[str, str]
    slide_plan: List[Dict[str, Any]]
    layout_plan: List[Dict[str, Any]]
    builder_slide : Dict[str, Any]
    rendered_slide_images: List[str]
