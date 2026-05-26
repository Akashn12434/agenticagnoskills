from agno.agent import Agent
from agno.models.google import Gemini
from agno.skills import Skills, LocalSkills
from dotenv import load_dotenv
import importlib.util
from pathlib import Path
from typing import Any
from agno.team import Team
from agno.os import AgentOS

load_dotenv()


def load_function(script_path: Path, module_name: str, function_name: str):
    spec = importlib.util.spec_from_file_location(module_name, script_path)
    if spec is None or spec.loader is None:
        raise ImportError(f"Could not load {function_name} from {script_path}")

    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return getattr(module, function_name)


PPT_SCRIPT_PATH = Path(__file__).parent / "ppt-skill" / "scripts" / "create_ppt.py"
WORD_CREATE_SCRIPT_PATH = Path(__file__).parent / "word-skill" / "scripts" / "create_docx.py"
WORD_READ_SCRIPT_PATH = Path(__file__).parent / "word-skill" / "scripts" / "read_docx.py"
EXCEL_CREATE_SCRIPT_PATH = Path(__file__).parent / "excel-skill" / "scripts" / "create_xlsx.py"
EXCEL_READ_SCRIPT_PATH = Path(__file__).parent / "excel-skill" / "scripts" / "read_xlsx.py"

create_presentation = load_function(PPT_SCRIPT_PATH, "ppt_skill_create_ppt", "create_presentation")
create_document = load_function(WORD_CREATE_SCRIPT_PATH, "word_skill_create_docx", "create_document")
read_document = load_function(WORD_READ_SCRIPT_PATH, "word_skill_read_docx", "read_document")
create_workbook = load_function(EXCEL_CREATE_SCRIPT_PATH, "excel_skill_create_xlsx", "create_workbook")
read_workbook = load_function(EXCEL_READ_SCRIPT_PATH, "excel_skill_read_xlsx", "read_workbook")


def create_powerpoint_presentation(
    topic: str,
    slides: list[dict[str, Any]] | None = None,
    output_path: str | None = None,
) -> str:
    """Create and save a PowerPoint .pptx file for the requested topic.

    Args:
        topic: The presentation topic.
        slides: Optional slide list. Each slide should be a plain dict with title,
            optional subtitle, and bullets, for example:
            {"title": "Introduction", "bullets": ["Point one", "Point two"]}.
            Do not wrap slides under placeholder keys like example_key.
        output_path: Optional output .pptx path. Relative paths are saved under ppt-skill/output.

    Returns:
        The absolute path of the generated .pptx file.
    """
    path = create_presentation(topic=topic, slides=slides, output_path=output_path)
    return f"✅ PowerPoint presentation created successfully\n\n📂 Opening presentation...\n\nFile path: {path}"


def create_word_document(
    title: str,
    sections: list[dict[str, Any]] | None = None,
    subtitle: str | None = None,
    output_path: str | None = None,
) -> str:
    """Create and save a Word .docx file for the requested topic.

    Args:
        title: The document title.
        sections: Optional section list. Each section should be a plain dict with
            heading, paragraphs, bullets, and optional table.
        subtitle: Optional document subtitle.
        output_path: Optional output .docx path. Relative paths are saved under word-skill/output.

    Returns:
        The absolute path of the generated .docx file.
    """
    path = create_document(title=title, sections=sections, subtitle=subtitle, output_path=output_path)
    return f"✅ Word document created successfully\n\n📂 Opening document...\n\nFile path: {path}"


def read_word_document(path: str) -> str:
    """Read and return plain text from a Word .docx file."""
    return read_document(path)


def create_excel_workbook(
    title: str,
    sheets: list[dict[str, Any]] | None = None,
    output_path: str | None = None,
) -> str:
    """Create and save an Excel .xlsx workbook for the requested topic or data.

    Args:
        title: The workbook title.
        sheets: Optional worksheet list. Each sheet should be a plain dict with
            name, optional title, headers, rows, optional formulas, and optional chart.
        output_path: Optional output .xlsx path. Relative paths are saved under excel-skill/output.

    Returns:
        The absolute path of the generated .xlsx file.
    """
    path = create_workbook(title=title, sheets=sheets, output_path=output_path)
    return f"✅ Excel workbook created successfully\n\n📂 Opening workbook...\n\nFile path: {path}"


def read_excel_workbook(path: str, max_rows_per_sheet: int = 50) -> str:
    """Read and return sheet names and cell values from an Excel .xlsx workbook."""
    return read_workbook(path, max_rows_per_sheet=max_rows_per_sheet)


# Load skills from a directory
ppt_agent = Agent(
    name="PowerPoint Agent",
    model=Gemini(id="gemini-3.1-flash-lite"),
    skills=Skills(loaders=[LocalSkills("./ppt-skill")]),
    tools=[create_powerpoint_presentation],
    instructions=[
        "Load ppt-skill instructions and consult references/guide.md.",
        "Build professional slide decks.",
        "Use clear slide titles and bullet points.",
        "Call create_powerpoint_presentation when creating presentations.",
        "After creation, tell the user the presentation was created successfully and is opening automatically.",
    ]
)
word_agent = Agent(
    name="Word Agent",
    model=Gemini(id="gemini-3.1-flash-lite"),
    skills=Skills(loaders=[LocalSkills("./word-skill")]),
    tools=[
        create_word_document,
        read_word_document,
    ],
    instructions=[
        "Load word-skill instructions and consult references/guide.md.",
        "Create professional Word documents.",
        "Use headings, paragraphs, bullets, and tables where appropriate.",
        "Call create_word_document when creating documents.",
        "Call read_word_document when reading documents.",
        "After creation, tell the user the document was created successfully and is opening automatically.",
    ]
)
excel_agent = Agent(
    name="Excel Agent",
    model=Gemini(id="gemini-3.1-flash-lite"),
    skills=Skills(loaders=[LocalSkills("./excel-skill")]),
    tools=[
        create_excel_workbook,
        read_excel_workbook,
    ],
    instructions=[
        "Load excel-skill instructions and consult references/guide.md.",
        "Create professional spreadsheets.",
        "Use formulas, charts, conditional formatting and multiple sheets when useful.",
        "Call create_excel_workbook when creating spreadsheets.",
        "Call read_excel_workbook when reading spreadsheets.",
        "After creation, tell the user the workbook was created successfully and is opening automatically.",
    ]
)


team = Team(
    name="Document Team",
    model=Gemini(id="gemini-3.1-flash-lite"),
    members=[
        ppt_agent,
        word_agent,
        excel_agent,
    ],
    instructions=[
        "Delegate PowerPoint requests to PowerPoint Agent.",
        "Delegate Word requests to Word Agent.",
        "Delegate Excel requests to Excel Agent.",
        "For multi-document workflows, coordinate between agents."
    ]
)


agent_os = AgentOS(teams=[team])
app = agent_os.get_app()

if __name__ == "__main__":
    team.print_response(
        "can u create a excel file for marks of 5 students in 3 subjects and their total marks and percentage?",
        stream=True,
    )
