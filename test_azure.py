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
print("Azure OpenAI config (đọc từ .env):")
print(f"  ENDPOINT       = {endpoint!r}")
print(f"  API_VERSION    = {api_version!r}")
print(f"  DEPLOYMENT     = {deployment!r}")
print(f"  API_KEY        = {'<set, ' + str(len(api_key)) + ' chars>' if api_key else '<EMPTY>'}")
print("=" * 60)

# Validate format
problems = []
if not endpoint.startswith("https://"):
    problems.append(f"ENDPOINT phải bắt đầu bằng 'https://', bạn đang có: {endpoint!r}")
if "/openai/deployments" in endpoint:
    problems.append("ENDPOINT KHÔNG được chứa '/openai/deployments/...'. "
                    "Chỉ cần phần gốc, vd: 'https://my-resource.openai.azure.com/'")
if "/openai/v1" in endpoint or endpoint.rstrip("/").endswith("/openai"):
    problems.append("ENDPOINT KHÔNG được chứa '/openai' hoặc '/openai/v1'. "
                    "Chỉ cần phần gốc của resource. SDK tự thêm path.")
if not endpoint.endswith("/"):
    print("WARN: endpoint không có trailing '/', có thể OK nhưng nên có")

if not deployment:
    problems.append("DEPLOYMENT trống. Đây là TÊN DEPLOYMENT trong Azure portal (vd: 'gpt-4o'), KHÔNG phải tên model")
if "/" in deployment:
    problems.append(f"DEPLOYMENT name có '/': {deployment!r}. Chỉ là tên đơn (gpt-4o, gpt-4.1-mini...)")

if not api_version:
    problems.append("API_VERSION trống. Vd: '2024-10-21' hoặc '2024-08-01-preview'")

if problems:
    print("\nCó vấn đề về cấu hình:")
    for p in problems:
        print(f"  - {p}")
    print("\nSửa .env rồi chạy lại script này.")
    raise SystemExit(1)

# Compute the URL OpenAI SDK sẽ gọi
expected_url = f"{endpoint.rstrip('/')}/openai/deployments/{deployment}/chat/completions?api-version={api_version}"
print(f"\nSDK sẽ POST đến URL:\n  {expected_url}")
print("\nĐang test gọi LLM...")

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
    print(f"\n✅ THÀNH CÔNG. LLM trả lời:\n  {resp.content}")
except Exception as e:
    print(f"\n❌ LỖI: {type(e).__name__}: {e}")
    msg = str(e).lower()
    if "404" in msg or "not found" in msg:
        print("\n→ 404 = URL/deployment/api_version sai. Kiểm tra:")
        print("  1. Vào Azure portal → Azure OpenAI resource → 'Deployments'")
        print("     → đảm bảo có deployment với tên CHÍNH XÁC = '" + deployment + "'")
        print("  2. ENDPOINT phải khớp resource. Lấy từ 'Keys and Endpoint' trong portal.")
        print("  3. API_VERSION thử các giá trị: 2024-10-21, 2024-08-01-preview, 2024-06-01")
    elif "401" in msg or "unauthorized" in msg:
        print("\n→ 401 = API_KEY sai. Copy lại từ 'Keys and Endpoint' trong portal.")
    elif "429" in msg:
        print("\n→ 429 = Bị throttle. Đợi rồi thử lại, hoặc check quota.")
    raise SystemExit(1)
