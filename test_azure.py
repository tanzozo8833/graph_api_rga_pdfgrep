"""Quick smoke test for Azure OpenAI credentials.

Run: venv\\Scripts\\python.exe test_azure.py
"""
import os
from dotenv import load_dotenv

load_dotenv()

endpoint = os.getenv("AZURE_OPENAI_ENDPOINT", "")
api_key = os.getenv("AZURE_OPENAI_API_KEY", "")
api_version = os.getenv("AZURE_OPENAI_API_VERSION", "")
deployment = os.getenv("AZURE_OPENAI_DEPLOYMENT", "")

print("=" * 60)
print("Azure OpenAI config (loaded from .env):")
print(f"  ENDPOINT       = {endpoint!r}")
print(f"  API_VERSION    = {api_version!r}")
print(f"  DEPLOYMENT     = {deployment!r}")
print(f"  API_KEY        = {'<set, ' + str(len(api_key)) + ' chars>' if api_key else '<EMPTY>'}")
print("=" * 60)

# Validate format
problems = []
if not endpoint.startswith("https://"):
    problems.append(f"ENDPOINT must start with 'https://', got: {endpoint!r}")
if "/openai/deployments" in endpoint:
    problems.append("ENDPOINT must NOT contain '/openai/deployments/...'. "
                    "Use only the root, e.g. 'https://my-resource.openai.azure.com/'")
if "/openai/v1" in endpoint or endpoint.rstrip("/").endswith("/openai"):
    problems.append("ENDPOINT must NOT contain '/openai' or '/openai/v1'. "
                    "Use only the resource root. The SDK appends the path.")
if not endpoint.endswith("/"):
    print("WARN: endpoint has no trailing '/'; may be OK but recommended")

if not deployment:
    problems.append("DEPLOYMENT is empty. This is the DEPLOYMENT NAME in Azure portal (e.g. 'gpt-4o'), NOT the model name")
if "/" in deployment:
    problems.append(f"DEPLOYMENT name contains '/': {deployment!r}. Use a plain name (gpt-4o, gpt-4.1-mini, ...)")

if not api_version:
    problems.append("API_VERSION is empty. E.g. '2024-10-21' or '2024-08-01-preview'")

if problems:
    print("\nConfiguration problems detected:")
    for p in problems:
        print(f"  - {p}")
    print("\nFix .env and rerun this script.")
    raise SystemExit(1)

# Compute the URL the OpenAI SDK will call
expected_url = f"{endpoint.rstrip('/')}/openai/deployments/{deployment}/chat/completions?api-version={api_version}"
print(f"\nSDK will POST to URL:\n  {expected_url}")
print("\nTesting LLM call...")

try:
    from langchain_openai import AzureChatOpenAI
    llm = AzureChatOpenAI(
        azure_endpoint=endpoint,
        api_key=api_key,
        api_version=api_version,
        azure_deployment=deployment,
        temperature=0,
        max_tokens=50,
    )
    resp = llm.invoke("Say 'hello world' and nothing else.")
    print(f"\nSUCCESS. LLM replied:\n  {resp.content}")
except Exception as e:
    print(f"\nERROR: {type(e).__name__}: {e}")
    msg = str(e).lower()
    if "404" in msg or "not found" in msg:
        print("\n-> 404 = URL/deployment/api_version is wrong. Check:")
        print("  1. Azure portal -> Azure OpenAI resource -> 'Deployments'")
        print("     -> ensure a deployment exists with EXACT name = '" + deployment + "'")
        print("  2. ENDPOINT must match the resource. Get it from 'Keys and Endpoint' in the portal.")
        print("  3. Try API_VERSION values: 2024-10-21, 2024-08-01-preview, 2024-06-01")
    elif "401" in msg or "unauthorized" in msg:
        print("\n-> 401 = API_KEY is wrong. Copy it again from 'Keys and Endpoint' in the portal.")
    elif "429" in msg:
        print("\n-> 429 = Throttled. Wait and retry, or check quota.")
    raise SystemExit(1)
