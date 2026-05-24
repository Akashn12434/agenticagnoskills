from agno.agent import Agent
from agno.models.google import Gemini
from agno.tools.google.slides import GoogleSlidesTools
from dotenv import load_dotenv
load_dotenv()

agent = Agent(
    model=Gemini(id="gemini-3.1-flash-lite"),
    tools=[
        GoogleSlidesTools(
            credentials_path="credentials.json",   
            token_path="token.json",
            oauth_port=8080,
        )
    ],
    instructions=[
        "You are a Google Slides assistant that helps users create and manage presentations.",
        "Always call get_presentation_metadata before modifying slides to get current slide IDs.",
        "Use slide_id values returned by the API -- never guess them.",
        "Return the presentation ID and URL after creating a presentation.",
    ],
    add_datetime_to_context=True,
    markdown=True,
)

agent.print_response(
    "Create a professional Google Slides presentation titled 'The Solar System'. "
    "Create exactly 6 slides. "
    "Slide 1: TITLE slide introducing the Solar System. "
    "Slides 2 to 5: TITLE_AND_BODY slides covering two planets per slide. "
    "Include detailed and informative bullet points for each planet, explaining their characteristics, atmosphere, size, moons, and unique features. "
    "Ensure every content slide contains meaningful information and multiple bullet points, not just titles. "
    "Slide 6: TITLE slide with title 'Thank You' and subtitle 'Questions and Discussion'.",
    stream=True,
)