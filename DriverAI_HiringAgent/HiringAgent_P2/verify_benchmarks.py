from pathlib import Path
from hiring_agent.extraction import extract_text_from_bytes, extract_candidate_details_smart

resumes = sorted(Path("benchmark_resumes").glob("*.pdf"))
print(f"Testing {len(resumes)} benchmark resumes with updated extract_text_from_bytes:\n")

for p in resumes:
    raw = p.read_bytes()
    text = extract_text_from_bytes(raw, p.name)
    details = extract_candidate_details_smart(text)
    print(f"[{p.name}]")
    print(f"  Length  : {len(text)} chars")
    print(f"  Name    : {details.get('full_name')}")
    print(f"  Phone   : {details.get('phone')}")
    print(f"  Location: {details.get('location')}, {details.get('country')}")
    print(f"  Skills  : {details.get('skills', '')[:80]}...")
    print("-" * 60)

