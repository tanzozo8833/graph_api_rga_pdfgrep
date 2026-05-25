import os

from langchain_openai import AzureChatOpenAI


def make_llm():
    endpoint = os.getenv("AZURE_OPENAI_ENDPOINT")
    api_key = os.getenv("AZURE_OPENAI_API_KEY")
    api_version = os.getenv("AZURE_OPENAI_API_VERSION")
    deployment = os.getenv("AZURE_OPENAI_DEPLOYMENT")

    missing = [
        name
        for name, val in [
            ("AZURE_OPENAI_ENDPOINT", endpoint),
            ("AZURE_OPENAI_API_KEY", api_key),
            ("AZURE_OPENAI_API_VERSION", api_version),
            ("AZURE_OPENAI_DEPLOYMENT", deployment),
        ]
        if not val
    ]
    if missing:
        raise RuntimeError(
            "Missing Azure OpenAI env vars: " + ", ".join(missing) + ". Set them in .env."
        )

    return AzureChatOpenAI(
        azure_endpoint=endpoint,
        api_key=api_key,
        api_version=api_version,
        azure_deployment=deployment,
        temperature=0,
        max_tokens=4096,
    )
