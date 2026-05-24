from agno.agent import Agent
from agno.models.google import Gemini
from agno.skills import Skills, LocalSkills
from dotenv import load_dotenv
import importlib.util
from pathlib import Path
from typing import Any

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
    return create_presentation(topic=topic, slides=slides, output_path=output_path)


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
    return create_document(title=title, sections=sections, subtitle=subtitle, output_path=output_path)


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
    return create_workbook(title=title, sheets=sheets, output_path=output_path)


def read_excel_workbook(path: str, max_rows_per_sheet: int = 50) -> str:
    """Read and return sheet names and cell values from an Excel .xlsx workbook."""
    return read_workbook(path, max_rows_per_sheet=max_rows_per_sheet)


# Load skills from a directory
agent = Agent(
    model=Gemini(id="gemini-3.1-flash-lite",),
    skills=Skills(loaders=[LocalSkills("./ppt-skill"), LocalSkills("./word-skill"), LocalSkills("./excel-skill")]),
    tools=[
        create_powerpoint_presentation,
        create_word_document,
        read_word_document,
        create_excel_workbook,
        read_excel_workbook,
    ],
    instructions=[
        "You are a professional document, presentation, and spreadsheet assistant with access to PowerPoint, Word, and Excel skills.",
        "For presentation requests, load ppt-skill instructions and consult its references/guide.md before creating or modifying presentations.",
        "For Word document requests, load word-skill instructions and consult its references/guide.md before creating, reading, or analyzing Word documents.",
        "For Excel workbook requests, load excel-skill instructions and consult its references/guide.md before creating, reading, or analyzing spreadsheets.",
        "Create well-structured documents and workbooks with clear sections, sheets, and professional formatting.",
        "Use headings, lists, and tables where appropriate.",
        "When a user requests a new PowerPoint presentation, build the slide titles and bullets, then call create_powerpoint_presentation to generate and save a .pptx file.",
        "Pass slides as a list of plain dictionaries with title, optional subtitle, and bullets. Do not nest slide data under example_key or any other wrapper key.",
        "When a user requests a new Word document, build sections with headings, at least one explanatory paragraph, useful bullets, and optional tables, then call create_word_document to generate and save a .docx file.",
        "For Word documents, do not send only bullet lists when paragraphs would make the content clearer.",
        "When a user requests to read a Word document, call read_word_document.",
        "When a user requests a new Excel workbook, build sheets with headers, rows, useful formulas, and charts where helpful, then call create_excel_workbook to generate and save a .xlsx file.",
        "Pass sheets as a list of plain dictionaries with name, optional title, headers, rows, formulas, and chart. Do not nest sheet data under example_key or any other wrapper key.",
        "When a user requests to read an Excel workbook, call read_excel_workbook.",
        "Always return the generated file path after a creation tool call succeeds.",
        "Do not say you cannot output a .pptx, .docx, or .xlsx file; use the available generation tools.",
    ]
)

if __name__ == "__main__":
    agent.print_response(
        "Create an Excel file for student records with columns for name, subjects marks scored in 3 subjects with percentages like Physics, Chemistry, and Biology percentages for 6 random students.",
        stream=True,
    )
