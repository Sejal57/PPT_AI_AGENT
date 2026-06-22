import os
from tools import Parsing_data, designer_node, summary_node, add_layout_to_slide, slide_renderer_node,export_to_pptx_node
from schema import State
from langgraph.graph import StateGraph, START, END
from dotenv import load_dotenv
load_dotenv()
path_doc = os.getenv("DOC_PATH")


with open(path_doc, 'r') as f:
        markdown_data  = f.read()


def parsing_raw_data(state : State):
      """Take the markdown files raw data and formate the whole data into a structural json list where the data is in herarical formate."""
      data = Parsing_data(state["raw_file_content"])
      struct_data = data.final_list

      return {"parsed_content" : struct_data}

builder = StateGraph(State)

builder.add_node("parser", parsing_raw_data)
builder.add_node("designer",designer_node )
builder.add_node("summary", summary_node)
builder.add_node("layout", add_layout_to_slide)
builder.add_node("render_ppt", slide_renderer_node)
builder.add_node("convert_pptx", export_to_pptx_node)


builder.add_edge(START, "parser")
builder.add_edge("parser","summary")
builder.add_edge("summary","layout")
builder.add_edge("layout", "designer")
builder.add_edge("designer", "render_ppt")
builder.add_edge("render_ppt", "convert_pptx")
builder.add_edge("convert_pptx", END)

graph = builder.compile()


if __name__ == "__main__":
    

    initial_input = {"raw_file_content": markdown_data}
    final_data = graph.invoke(initial_input)

    print(final_data)