from langchain_ollama import OllamaLLM

def get_llm():
    llm = OllamaLLM(model="llama3.2")

    return llm
