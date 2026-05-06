import os
import re
import tkinter as tk
from tkinter import filedialog, messagebox

import requests
from dotenv import load_dotenv

# ----------------------------
# Load Environment Variables
# ----------------------------
load_dotenv()

DATABRICKS_URL = os.getenv("DATABRICKS_URL")
DATABRICKS_API_KEY = os.getenv("DATABRICKS_API")

if not DATABRICKS_URL or not DATABRICKS_API_KEY:
    raise EnvironmentError("Databricks URL or API Key not set")

print("Databricks environment variables loaded successfully.")

# ----------------------------
# Helpers
# ----------------------------
PART_NUMBER_PATTERN = re.compile(r"\b[A-Z0-9]{4}-[A-Z0-9]{5}\b", re.IGNORECASE)
PN_FRAGMENT = r"[A-Z0-9]{4}-[A-Z0-9]{5}"

TITLE_MAX_CHARS = 65
TITLE_WORD_MIN = 12
TITLE_WORD_MAX = 20

ALLOWED_INTENTS = [
    "Engineering Change Proposal for",
    "Alternate Part Introduction for",
    "Design Change Proposal for",
    "Material / Component Change for",
    "Lead Time Risk Mitigation for",
]


def extract_part_numbers(text: str) -> list[str]:
    if not text:
        return []
    parts = PART_NUMBER_PATTERN.findall(text)
    seen = set()
    ordered = []
    for p in parts:
        p = p.upper()
        if p not in seen:
            seen.add(p)
            ordered.append(p)
    return ordered


def extract_bom_relationships(text: str) -> list[tuple[str, str]]:
    """
    Returns list of (parent_part_number, child_part_number).
    Rule: 'Used in Part Number' means Parent Part Number.
    """
    if not text:
        return []

    relationships: list[tuple[str, str]] = []
    seen: set[tuple[str, str]] = set()

    for raw_line in text.splitlines():
        line = raw_line.strip()
        if not line:
            continue

        parent = None
        child = None

        # Pattern 1: explicit labels
        p_match = re.search(
            rf"parent(?:\s*part(?:\s*number)?)?\s*[:=\-]?\s*({PN_FRAGMENT})",
            line,
            flags=re.IGNORECASE,
        )
        c_match = re.search(
            rf"child(?:\s*part(?:\s*number)?)?\s*[:=\-]?\s*({PN_FRAGMENT})",
            line,
            flags=re.IGNORECASE,
        )
        if p_match and c_match:
            parent = p_match.group(1).upper()
            child = c_match.group(1).upper()

        # Pattern 2: "child ... used in/where used ... parent"
        if not (parent and child):
            m = re.search(
                rf"({PN_FRAGMENT}).*?(?:used\s+in|where\s+used).*?({PN_FRAGMENT})",
                line,
                flags=re.IGNORECASE,
            )
            if m:
                child = m.group(1).upper()
                parent = m.group(2).upper()

        # Pattern 3: "Used in Part Number: <parent>" + another PN on line = child
        if not (parent and child):
            m = re.search(
                rf"used\s*in\s*part\s*number\s*[:=\-]?\s*({PN_FRAGMENT})",
                line,
                flags=re.IGNORECASE,
            )
            if m:
                parent = m.group(1).upper()
                pns = extract_part_numbers(line)
                for pn in pns:
                    if pn != parent:
                        child = pn
                        break

        # Pattern 4: arrows/verbs
        if not (parent and child):
            m = re.search(
                rf"({PN_FRAGMENT})\s*(?:->|>|contains|includes|has)\s*({PN_FRAGMENT})",
                line,
                flags=re.IGNORECASE,
            )
            if m:
                parent = m.group(1).upper()
                child = m.group(2).upper()

        # Pattern 5: fallback on BOM-like lines
        if not (parent and child):
            pns = extract_part_numbers(line)
            if len(pns) >= 2 and any(k in line.lower() for k in ("bom", "where used", "used in")):
                parent = pns[0]
                child = pns[1]

        if parent and child:
            pair = (parent, child)
            if pair not in seen:
                seen.add(pair)
                relationships.append(pair)

    return relationships


def build_bom_context(text: str) -> str:
    rels = extract_bom_relationships(text)
    if not rels:
        return ""
    lines = ["Detected BOM context (Used in Part Number = Parent):"]
    for i, (parent, child) in enumerate(rels, start=1):
        lines.append(
            f"{i}. Parent Part Number: {parent}; Child Part Number: {child}; Used in Part Number: {parent}"
        )
    return "\n".join(lines)


def clean_ai_output(text: str) -> str:
    text = re.sub(r"!\[[^\]]*\]\([^)]+\)", "", text)
    text = re.sub(r"<img[^>]*>", "", text, flags=re.IGNORECASE)

    text = re.sub(
        r"[\U0001F300-\U0001F5FF\U0001F600-\U0001F64F\U0001F680-\U0001F6FF"
        r"\U0001F700-\U0001F77F\U0001F780-\U0001F7FF\U0001F800-\U0001F8FF"
        r"\U0001F900-\U0001F9FF\U0001FA00-\U0001FAFF\U00002600-\U000026FF"
        r"\U00002700-\U000027BF]+",
        "",
        text,
    )

    text = re.sub(r"[^\w\s\-\.,:;()/%#&\n]", " ", text)
    text = re.sub(r"[ \t]+", " ", text)
    text = re.sub(r"\n{3,}", "\n\n", text)
    return text.strip()


def remove_unwanted_sections(text: str) -> str:
    blocked_headers = [
        "Part Numbers Identified",
        "BOM Relationships Identified",
        "Quality Gate Warnings",
    ]
    for header in blocked_headers:
        pattern = rf"\n*{re.escape(header)}\s*:\s*(.*?)(?=\n[A-Za-z][A-Za-z ]*:\s*|\Z)"
        text = re.sub(pattern, "", text, flags=re.IGNORECASE | re.DOTALL)
    return text.strip()


def _dedupe_lines(lines: list[str]) -> list[str]:
    seen = set()
    out = []
    for line in lines:
        key = re.sub(r"\s+", " ", line.strip().lower())
        if key and key not in seen:
            seen.add(key)
            out.append(line.strip())
    return out


def format_numbered_bullets(text: str) -> str:
    for section in ["Problem Statement", "Solution Statement"]:
        pattern = rf"({re.escape(section)}\s*:\s*)(.*?)(?=\n[A-Za-z][A-Za-z ]*:\s*|\Z)"

        def _repl(match):
            heading = match.group(1).strip()
            body = match.group(2)

            items = []
            for raw in body.splitlines():
                line = raw.strip()
                if not line:
                    continue
                line = re.sub(r"^[-*]\s+", "", line)
                line = re.sub(r"^\d+[\.\)]\s+", "", line)
                line = re.sub(r"\s+", " ", line).strip()
                if line:
                    items.append(line)

            items = _dedupe_lines(items)
            numbered = "\n".join([f"{i}. {item}" for i, item in enumerate(items, start=1)])
            return f"{heading}\n{numbered}" if numbered else f"{heading}\n1. Not specified."

        text = re.sub(pattern, _repl, text, flags=re.IGNORECASE | re.DOTALL)

    return text.strip()


def _extract_section(text: str, section_name: str) -> str:
    pattern = rf"{re.escape(section_name)}\s*:\s*(.*?)(?=\n[A-Za-z][A-Za-z ]*:\s*|\Z)"
    m = re.search(pattern, text, flags=re.IGNORECASE | re.DOTALL)
    return m.group(1).strip() if m else ""


def validate_engineering_output(output: str, source_text: str) -> list[str]:
    violations: list[str] = []

    m_title = re.search(r"^\s*Title\s*:\s*(.+)$", output, flags=re.IGNORECASE | re.MULTILINE)
    title = m_title.group(1).strip() if m_title else ""

    if not title:
        violations.append("Missing 'Title:' line.")
    else:
        if len(title) > TITLE_MAX_CHARS:
            violations.append(f"Title exceeds {TITLE_MAX_CHARS} characters.")
        wc = len(title.split())
        if wc < TITLE_WORD_MIN or wc > TITLE_WORD_MAX:
            violations.append(f"Title word count must be {TITLE_WORD_MIN}-{TITLE_WORD_MAX}.")
        if " for " not in title.lower() or " due to " not in title.lower():
            violations.append("Title should follow: [Intent] for [Scope] due to [Reason].")
        if not any(title.startswith(intent) for intent in ALLOWED_INTENTS):
            violations.append("Title must start with an approved Engineering Change intent keyword.")

    problem = _extract_section(output, "Problem Statement")
    solution = _extract_section(output, "Solution Statement")

    if not problem:
        violations.append("Missing 'Problem Statement' section.")
    if not solution:
        violations.append("Missing 'Solution Statement' section.")

    # Numbered bullets only
    for section_name in ["Problem Statement", "Solution Statement"]:
        sec = _extract_section(output, section_name)
        if sec and re.search(r"^\s*-\s+", sec, flags=re.MULTILINE):
            violations.append(f"{section_name} must use numbered bullets, not '-'.")

    if problem:
        disallowed_problem_terms = [
            r"\brecommend(?:ed|ation)?\b",
            r"\bpropos(?:e|ed|al)\b",
            r"\breplace(?:d|ment)?\b",
            r"\bintroduc(?:e|ed|tion)\b",
            r"\bshould\b",
            r"\bcan be\b",
            r"\baction\b",
        ]
        for pat in disallowed_problem_terms:
            if re.search(pat, problem, flags=re.IGNORECASE):
                violations.append("Problem Statement contains solution-oriented language.")
                break

    if re.search(r"\b(I|we|our|my|us)\b", output, flags=re.IGNORECASE):
        violations.append("Use third-person tone only.")

    source_parts = set(extract_part_numbers(source_text))
    out_parts = set(extract_part_numbers(output))
    if source_parts and not source_parts.issubset(out_parts):
        missing = sorted(source_parts - out_parts)
        violations.append(
            f"Missing part numbers from output: {', '.join(missing[:5])}"
            + ("..." if len(missing) > 5 else "")
        )

    return violations


def _call_databricks(prompt_text: str) -> str:
    headers = {
        "Authorization": f"Bearer {DATABRICKS_API_KEY}",
        "Content-Type": "application/json",
    }

    payload = {
        "messages": [
            {"role": "user", "content": prompt_text.strip()}
        ]
    }

    try:
        response = requests.post(
            DATABRICKS_URL,
            headers=headers,
            json=payload,
            timeout=60,
        )
    except requests.RequestException as e:
        return f"Error: Request failed - {e}"

    if response.status_code != 200:
        return f"Error: {response.status_code} - {response.text}"

    try:
        result = response.json()
        return result["choices"][0]["message"]["content"]
    except (ValueError, KeyError, IndexError, TypeError):
        return f"Error: Unexpected response format - {response.text}"


def _build_engineering_prompt(user_text: str) -> str:
    return f"""
You are drafting Engineering Change content for ECR/ECO/Deviation/PCN.

Non-negotiable rules:
1. Formal, technical, audit-safe.
2. Third-person, passive-neutral, fact-based.
3. No informal, emotional, urgency, or blame language.
4. Keep sentences short.
5. Do not repeat the same point with similar wording.

Output format (exact):
Title: <one line, 12-20 words, less than 65 characters>

Problem Statement:
1. <point>
2. <point>
3. <point>

Solution Statement:
1. <point>
2. <point>
3. <point>

Title constraints:
- Must follow: [Change Intent] for [Part / Assembly / Process] due to [Engineering Reason]
- Must start with one approved intent:
  Engineering Change Proposal for
  Alternate Part Introduction for
  Design Change Proposal for
  Material / Component Change for
  Lead Time Risk Mitigation for

Problem constraints:
- Describe current released condition and constraint only.
- Include impact if no action is taken.
- No solution, recommendation, alternate part, cost, or corrective action.

Solution constraints:
- Define proposed change, applicability, justification, and expected outcome.

Part number rules:
- If pattern XXXX-YYYYY appears, write: Part Number: XXXX-YYYYY
- In BOM/Where-Used context, "Used in Part Number" means Parent Part Number.

Use numbered bullets only. Do not use "-" bullets.

Input:
{user_text}
""".strip()


def reframe_problem(user_text: str) -> str:
    prompt = _build_engineering_prompt(user_text)
    draft = _call_databricks(prompt)
    if draft.startswith("Error:"):
        return draft

    draft = clean_ai_output(draft)
    draft = remove_unwanted_sections(draft)
    draft = format_numbered_bullets(draft)

    violations = validate_engineering_output(draft, user_text)

    if violations:
        fix_prompt = f"""
Revise the draft to fix all violations.

Violations:
- {"\n- ".join(violations)}

Rules:
- Keep sentences short.
- Remove repetitive lines.
- Use numbered bullets only (1., 2., 3.).
- Return only:
  Title:
  Problem Statement:
  Solution Statement:

Draft:
{draft}

Original Input:
{user_text}
""".strip()

        revised = _call_databricks(fix_prompt)
        if not revised.startswith("Error:"):
            draft = clean_ai_output(revised)
            draft = remove_unwanted_sections(draft)
            draft = format_numbered_bullets(draft)

    return draft


# ----------------------------
# UI Actions
# ----------------------------
def browse_file():
    file_path = filedialog.askopenfilename(
        filetypes=[("Text Files", "*.txt"), ("All Files", "*.*")]
    )

    if file_path:
        with open(file_path, "r", encoding="utf-8", errors="ignore") as f:
            file_content.delete("1.0", tk.END)
            file_content.insert(tk.END, f.read())


def reframe_action():
    user_input = user_text.get("1.0", tk.END).strip()
    file_text = file_content.get("1.0", tk.END).strip()

    combined_input = ""
    if user_input:
        combined_input += f"User Problem:\n{user_input}\n\n"
    if file_text:
        combined_input += f"Reference File Content:\n{file_text}\n\n"

    if not combined_input.strip():
        messagebox.showwarning("Input Required", "Please provide text or browse a file.")
        return

    bom_context = build_bom_context(combined_input)
    if bom_context:
        combined_input += bom_context

    output_text.delete("1.0", tk.END)
    output_text.insert(tk.END, "Processing...\n")
    root.update_idletasks()

    reframed_text = reframe_problem(combined_input)
    if not reframed_text.startswith("Error:"):
        reframed_text = clean_ai_output(reframed_text)
        reframed_text = remove_unwanted_sections(reframed_text)
        reframed_text = format_numbered_bullets(reframed_text)

    output_text.delete("1.0", tk.END)
    output_text.insert(tk.END, reframed_text)


# ----------------------------
# UI Layout
# ----------------------------
root = tk.Tk()
root.title("AI Problem Reframer - Databricks")
root.geometry("900x700")

tk.Label(root, text="Box 1: User Problem Input").pack(anchor="w")
user_text = tk.Text(root, height=6)
user_text.pack(fill="both", padx=10, pady=5)

tk.Label(root, text="Box 2: File Content (Optional)").pack(anchor="w")
file_content = tk.Text(root, height=8)
file_content.pack(fill="both", padx=10, pady=5)

btn_frame = tk.Frame(root)
btn_frame.pack(pady=10)

tk.Button(btn_frame, text="Browse File", command=browse_file, width=20).pack(
    side="left", padx=10
)
tk.Button(btn_frame, text="Reframe Problem", command=reframe_action, width=20).pack(
    side="left"
)

tk.Label(root, text="Box 3: AI Reframed Output").pack(anchor="w")
output_text = tk.Text(root, height=16)
output_text.pack(fill="both", padx=10, pady=5)

root.mainloop()