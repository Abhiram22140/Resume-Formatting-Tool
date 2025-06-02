# parser.py

import os
import re

from docx import Document
from PyPDF2 import PdfReader


def _extract_from_docx(path: str) -> list[str]:
    """
    Read all non-empty paragraphs from a .docx document and return as a list of lines.
    """
    doc = Document(path)
    lines = [p.text.strip() for p in doc.paragraphs if p.text.strip()]
    return lines


def _extract_from_pdf(path: str) -> list[str]:
    """
    Read all text from a PDF using PyPDF2, split into lines, and return non-empty lines.
    """
    reader = PdfReader(path)
    raw_text = []
    for page in reader.pages:
        text = page.extract_text()
        if text:
            raw_text.extend(text.splitlines())
    # strip and filter empty
    return [line.strip() for line in raw_text if line.strip()]


def _find_section(lines: list[str], heading: str) -> int:
    """
    Return the index of the first line that starts with `heading` (case-insensitive),
    or -1 if not found.
    """
    heading_lower = heading.lower()
    for idx, line in enumerate(lines):
        if line.lower().startswith(heading_lower):
            return idx
    return -1


def _parse_contact_info(lines: list[str]) -> tuple[str, str, str]:
    """
    Search entire text (lines) for an email, phone, and address.
    We do a simple regex for email and phone. For address, we look for a line containing
    a comma (e.g., “City, Country”) or numeric street. If multiple matches, pick the first.
    """
    email_regex = re.compile(r"[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}")
    phone_regex = re.compile(r"(\+?\d[\d\s\-]{7,}\d)")
    email = ""
    phone = ""
    address = ""

    for line in lines:
        if not email:
            m = email_regex.search(line)
            if m:
                email = m.group(0).strip()
        if not phone:
            m = phone_regex.search(line)
            if m:
                phone = m.group(0).strip()
        # For address: pick a line that has a comma and at least one digit or a known pattern
        if not address:
            if "," in line and re.search(r"\d", line):
                address = line.strip()

        if email and phone and address:
            break

    return email, phone, address


def _parse_name_role(lines: list[str]) -> tuple[str, str]:
    """
    Assume the very first non-empty line is “Name – Role” (or “Name – Role” with dash).
    If no dash is found, take the entire line as Name and leave Role blank.
    """
    first = lines[0] if lines else ""
    if "–" in first:
        parts = first.split("–", 1)
        return parts[0].strip(), parts[1].strip()
    if "-" in first:
        parts = first.split("-", 1)
        return parts[0].strip(), parts[1].strip()
    return first.strip(), ""


def _extract_section_text(lines: list[str], heading: str, next_headings: list[str]) -> list[str]:
    """
    Given `lines`, find the index of `heading`; then collect all lines until you hit
    any of the `next_headings`. Return that slice (excluding the heading itself).
    If heading not found, return [].
    """
    start = _find_section(lines, heading)
    if start < 0:
        return []
    result = []
    for line in lines[start + 1 :]:
        # Stop if we see any next heading
        if any(line.lower().startswith(h.lower()) for h in next_headings):
            break
        result.append(line)
    return result


def parse_resume(filepath: str) -> dict:
    """
    Main entry point. Detect file extension, extract raw lines, and then parse:
      - name, role
      - email, phone, address
      - skills (lines under “Skills”)
      - summary (lines under “Summary”)
      - education (lines under “Education”)
      - experience (lines under “Experience”)

    Returns a dict with keys:
      name, role, email, phone, address, skills, summary,
      education (list of dicts), experience (list of dicts)
    """

    ext = os.path.splitext(filepath)[1].lower()
    if ext == ".docx":
        lines = _extract_from_docx(filepath)
    elif ext == ".pdf":
        lines = _extract_from_pdf(filepath)
    else:
        raise ValueError(f"Unsupported resume format: {ext}. Only .docx and .pdf are supported.")

    # 1) Name & Role
    name, role = _parse_name_role(lines)

    # 2) Contact Info
    email, phone, address = _parse_contact_info(lines)

    # 3) Skills – extract lines under “Skills” until next heading
    skill_lines = _extract_section_text(lines, "Skills", ["Summary", "Education", "Experience"])
    # Split skills by commas if they appear on one line, otherwise treat each line as one skill
    skills = []
    for sl in skill_lines:
        if "," in sl:
            skills.extend([s.strip() for s in sl.split(",") if s.strip()])
        else:
            skills.append(sl.strip())

    # 4) Summary
    summary_lines = _extract_section_text(lines, "Summary", ["Skills", "Education", "Experience"])
    summary = "\n".join(summary_lines).strip()

    # 5) Education – each line treated as “Degree, Institution (YYYY–YYYY)” or similar
    edu_lines = _extract_section_text(lines, "Education", ["Summary", "Skills", "Experience"])
    education = []
    for entry in edu_lines:
        # Try to parse "Degree, Institution (Year–Year)"
        m = re.match(r"^(.*?),\s*(.*?)\s*\((.*?)\)$", entry)
        if m:
            degree, inst, years = m.groups()
            se = years.split("–")
            start = se[0].strip()
            end = se[1].strip() if len(se) > 1 else ""
            education.append({
                "degree":      degree.strip(),
                "institution": inst.strip(),
                "start":       start,
                "end":         end
            })
        else:
            # fallback: treat whole line as degree/institution
            education.append({
                "degree": entry.strip(),
                "institution": "",
                "start": "",
                "end": ""
            })

    # 6) Experience – collect lines, then break into entries by detecting header lines
    exp_lines = _extract_section_text(lines, "Experience", ["Summary", "Skills", "Education"])
    experience = []
    current = {"position": "", "company": "", "dates": "", "description": ""}
    for line in exp_lines:
        # If line looks like "Title, Company (Year–Year)" → new entry
        if re.match(r".+\(.*–.*\)$", line):
            # Save previous if nonempty
            if current["position"]:
                experience.append(current)
                current = {"position": "", "company": "", "dates": "", "description": ""}
            # Parse position, company, dates
            m2 = re.match(r"^(.*?),\s*(.*?)\s*\((.*?)\)$", line)
            if m2:
                current["position"] = m2.group(1).strip()
                current["company"]  = m2.group(2).strip()
                current["dates"]    = m2.group(3).strip()
            else:
                current["position"] = line.strip()
        else:
            # Treat as part of description
            if current["description"]:
                current["description"] += "\n" + ("• " + line.lstrip("• "))
            else:
                current["description"] = "• " + line.lstrip("• ")

    # Append last one
    if current["position"]:
        experience.append(current)

    return {
        "name":       name,
        "role":       role,
        "email":      email,
        "phone":      phone,
        "address":    address,
        "skills":     skills,
        "summary":    summary,
        "education":  education,
        "experience": experience
    }
