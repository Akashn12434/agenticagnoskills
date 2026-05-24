from agno.agent import Agent
from agno.knowledge.chunking.document import DocumentChunking
from agno.knowledge.knowledge import Knowledge
from agno.knowledge.embedder.fastembed import FastEmbedEmbedder
from agno.knowledge.reader.pdf_reader import PDFReader
from agno.vectordb.chroma import ChromaDb
from agno.models.google import Gemini
from agno.vectordb.search import SearchType
from dotenv import load_dotenv
load_dotenv()

knowledge = Knowledge(
    vector_db=ChromaDb(
        collection="docs",
        path="tmp/chromadb",
        persistent_client=True,
        search_type=SearchType.hybrid,
        embedder=FastEmbedEmbedder(id="BAAI/bge-base-en-v1.5"),
    ),  
    max_results=3,
)

knowledge.insert(
    path="pdffiles/samplefiles.pdf",
     reader=PDFReader(
        name="Document Chunking Reader",
        chunking_strategy=DocumentChunking(),
    ),
    skip_if_exists=True)

agent = Agent(
    model=Gemini(id="gemini-3.1-flash-lite",
    cache_response=True,),
    knowledge=knowledge,
    search_knowledge=True,
    debug_mode=True,
)

agent.print_response("How much is the smart home market expected to be worth in 2025?", markdown=True)