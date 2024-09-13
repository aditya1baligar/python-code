from fastapi import APIRouter
from pydantic import BaseModel
from langchain_openai import OpenAIEmbeddings
import os
from langchain_openai import ChatOpenAI
from langchain_core.prompts import ChatPromptTemplate
from langchain.chains import create_retrieval_chain
from langchain_community.vectorstores import Chroma
from langchain_community.document_loaders import DirectoryLoader, TextLoader, PyMuPDFLoader
from langchain.text_splitter import RecursiveCharacterTextSplitter
from langchain.chains.combine_documents import create_stuff_documents_chain
from dotenv import load_dotenv
import time
import re


router = APIRouter()

FOLDERS = [
    "./controllers/chatbot_controller/procurengine/",
    "./controllers/chatbot_controller/WebsiteData/"
]

load_dotenv()
# Load OpenAI API Key
os.environ['OPENAI_API_KEY'] = os.getenv("OPENAI_API_KEY")



#to get information 
def get_document(index):
    return FOLDERS[index]

# Model for input
class Question(BaseModel):
    input: str
    history: list


# Initialize required components
llm = ChatOpenAI(api_key=os.getenv("OPENAI_API_KEY"), model_name="gpt-4o-mini")
#prompt
prompt = ChatPromptTemplate.from_template(
"""

Answer the questions related to procurement, both from general knowledge and the provided documentation.
Please provide the most accurate response based on the question.
If asked for the steps or configuration, give a detailed step-by-step process based on the documentation, promoting our procurement platform.
Explain each question in detail with all steps provided in the documentation.
Go through all the steps in the document first before answering, and provide general procurement advice if the answer is outside the scope of the documentation.
<context>
{context}
</context>
Questions: {input}
"""
)

prompt_website = ChatPromptTemplate.from_template(
    """
    
    You are a Guide for the customers to navigate through the ProcurEngine website.
    Answer the questions based on the procurement in such a way it should link in to the given context provided.
    Please provide the most accurate response based on the question.
    If asked to connect with the support team provide the contact number i.e +91 9650749600 not the steps.
    Just start with the answer and don't provide unnecessary information and if the information is not found in the context but is related to procurement then provide a basic knownledge about that question.
    Explain in detail for all the questions asked.
    Provide sufficient spaces to make the output readable.
    
    <context>
    {context}
    </context>
    Questions: {input}
    """
)


def vector_embedding(folder_path):
    embeddings = OpenAIEmbeddings()
    text_loader_kwargs = {'autodetect_encoding': True}
    
    # Load text files
    text_loader = DirectoryLoader(folder_path, glob="./*.txt", loader_cls=TextLoader, loader_kwargs=text_loader_kwargs)
    text_docs = text_loader.load()
    
    # Load PDF files
    pdf_loader = DirectoryLoader(folder_path, glob="./*.pdf", loader_cls=PyMuPDFLoader)
    pdf_docs = pdf_loader.load()
    
    # Combine text and PDF documents
    docs = text_docs + pdf_docs

    text_splitter = RecursiveCharacterTextSplitter(chunk_size=500, chunk_overlap=100)
    final_documents = text_splitter.split_documents(docs)
    
    try:
        vectors = Chroma.from_documents(final_documents, embeddings)
    except ValueError as e:
        raise Exception(f"Error initializing Chroma: {str(e)}")
    
    return vectors


vectors = vector_embedding(get_document(0))


@router.post("/ask/")
def get_response(question: Question):
    start = time.process_time()
    vectors = vector_embedding(get_document(0))
    document_chain = create_stuff_documents_chain(llm, prompt)
    retriever = vectors.as_retriever(search_k=10)
    retrieval_chain = create_retrieval_chain(retriever, document_chain)
    
    # Extract text from each dictionary in the history list where sender is "user"
    conversation_history = " ".join([entry['text'] for entry in question.history if entry['sender'] == "user"])
    
    # Use the history for context but focus the response on the current input
    response = retrieval_chain.invoke({"input": question.input, "context": conversation_history})
    
    # Add target="_blank" to any links in the response
    response_text = response['answer']
    response_text = re.sub(r'<a\s+(href="[^"]+")', r'<a \1 target="_blank"', response_text)
    
    print("Response time:", time.process_time() - start)
    
    return {"answer": response_text, "context": response["context"]}

vectors2 = vector_embedding(get_document(1))
@router.post("/askweb/")
def get_response(question: Question):
    start = time.process_time()
    
    document_chain = create_stuff_documents_chain(llm, prompt_website)
    retriever = vectors2.as_retriever(search_k=10)
    retrieval_chain = create_retrieval_chain(retriever, document_chain)
    
    # Extract text from each dictionary in the history list where sender is "user"
    conversation_history = " ".join([entry['text'] for entry in question.history if entry['sender'] == "user"])
    
    # Use the history for context but focus the response on the current input
    response = retrieval_chain.invoke({"input": question.input, "context": conversation_history})
    
    # Add target="_blank" to any links in the response
    response_text = response['answer']
    response_text = re.sub(r'<a\s+(href="[^"]+")', r'<a \1 target="_blank"', response_text)
    
    print("Response time:", time.process_time() - start)
    
    return {"answer": response_text, "context": response["context"]}

