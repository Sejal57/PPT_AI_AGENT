# PPT_AI_AGENT
An open-source Python pipeline using LangGraph to orchestrate AI agents for data parsing. It features defensive web streaming with custom headers to safely fetch remote images, automated deduplication caching, and regular expressions to extract and transform Markdown tables into structured dashboard layouts.


# Agentic Markdown & Presentation Parsing Pipeline

An advanced, autonomous multi-agent pipeline powered by LangGraph and orchestrated via Streamlit to systematically transform raw, unstructured markdown text documents into polished, production-grade presentation slides (.pptx).

The system implements structural parsing, intelligent brand-identity engineering, text summarization, layout mapping, dynamic HTML rendering via Playwright, and binary output slide generation.
---

## 🚀 Key Architectural Features

The architecture uses a state-machine framework topology (StateGraph) enforced through predefined typed state schemas to pipeline contextual data sequentially:

    graph LR
        START --> parser[Data Parser]
        parser --> summary[Content Summarizer]
        summary --> layout[Layout Planner]
        layout --> designer[Brand Designer]
        designer --> render_ppt[Playwright Slide Renderer]
        render_ppt --> convert_pptx[PPTX Export Engine]
        convert_pptx --> END
- Parser Node (Parsing_data): Extracts hierarchical content partitions, processes Markdown tables, and cleanly handles local Base64/remote assets with defensive header configurations to prevent server blocks.

- Summary Node (summary_node): Condenses dense informational body paragraphs into structural, high-impact bullets tailored for executive review.

- Layout Node (add_layout_to_slide): Strategically maps individual sections onto ideal visual templates (e.g., Cover, Grid Bullets, Dashboards, Tables, Images) based on structured content weights.

- Designer Node (designer_node): Selects an industry archetype (Technical Modern, Executive Corporate, Creative Vibrant) and engineers a cohesive 60-30-10 color palette mapped to the deck's topic.

- Render PPT Node (slide_renderer_node): Injects theme configurations, dynamic HTML, and custom fonts into an isolated high-resolution instance headless viewport browser through Playwright.

- Convert PPTX Node (export_to_pptx_node): Compiles rendered visual frames directly into widescreen 16:9 .pptx presentations.
---

🛠️ Technology Stack

---> Workflow Orchestration: langgraph (StateGraph topology tracking)

---> LLM Model Providers: Groq (Llama 3.3 70B), Gemini (3.1 Flash Lite), and Local Ollama instances

---> Headless Canvas Rendering: playwright (Chromium page screenshots)

---> Output Compilation: python-pptx (Binary presentation deck compilation)

---> Frontend Controller: streamlit (Dynamic parameter execution dashboard)

---
## 📂 Project Structure

        ├── main.py          # StateGraph construction, edge mapping, and compilation loop
        ├── tools.py         # Specialized agent node functions, parsing components, and render engines
        ├── schema.py        # Centralized TypedDict state tracking declaration attributes
        ├── dashboard.py     # Streamlit web studio frontend and parameter upload manager
        ├── saved_images/    # local workspace containing decoded raw markdown Base64 image files
        └── output_images/   # Target destination layer hosting streamed remote web image metrics
        

⚙️ Installation & Workspace Setup
1. System Dependencies & Python Packages
Ensure you have Python 3.10+ installed. Clone the repository and install the framework components:

Bash
    pip install langgraph langchain groq python-pptx playwright streamlit python-dotenv requests
    
2. Playwright Core Initialization
Initialize your headless Chromium compilation runtime binary:

Bash
    playwright install chromium
    
3. Environment Variable Configuration
Create a .env file inside the root repository framework layer to store execution parameters:

Ini, TOML

    DOC_PATH="path/to/your/source/document.md"
    # Configure your respective API authentication keys required by the active client nodes
    GROQ_API_KEY="your-production-groq-key"
    GEMINI_API_KEY="your-production-gemini-key"
    
🎮 Execution Management
Running the Interactive Web Studio Dashboard
To explore configurations, custom themes, upload fresh markdown text, and preview rendered screens live:

Bash
    streamlit run dashboard.py
    
Scripted Terminal Invoke Workflow
To directly invoke compilation pipelines through terminal configurations:

Bash
    python main.py
    
📄 License
This asset studio project is open-source software licensed under the MIT License. Feel free to use, modify, and distribute according to your internal architectural requirements.
