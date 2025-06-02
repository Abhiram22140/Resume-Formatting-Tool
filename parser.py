# parser.py

from resume_parser import resumeparse

def parse_resume(filepath: str) -> dict:
    """
    Parse a résumé file (Word or PDF) and return a dict with keys:
       - name       (str)
       - role       (str)       # current or desired title
       - email      (str)
       - phone      (str)
       - address    (str)
       - summary    (str)       # professional summary or objective
       - education  (List[dict])  # each dict: { degree, institution, start, end }
       - experience (List[dict])  # each dict: { position, company, dates, description }
       - skills     (List[str])
    """
    data = resumeparse.read_file(filepath)

    # Basic fields
    name    = data.get("name", "")
    email   = data.get("email", "")
    phone   = data.get("mobile_number", "")
    address = data.get("location", "")

    # “role” and “summary” (using designation/profile_summary if available)
    role    = data.get("designation", "")
    summary = data.get("profile_summary", "") or data.get("career_objective", "")

    # Skills: comma‐separated → list
    raw_skills = data.get("skills", "")
    skills = [s.strip() for s in raw_skills.split(",")] if raw_skills else []

    # Education: build a single entry if present
    education = []
    deg  = data.get("degree", "")
    inst = data.get("college_name", "")
    if deg or inst:
        education.append({
            "degree":      deg,
            "institution": inst,
            "start":       data.get("education_start", ""),
            "end":         data.get("education_end", "")
        })

    # Experience: build a single entry from “designation” / “company_names” / “experience”
    experience = []
    pos       = data.get("designation", "")
    comp      = data.get("company_names", "")
    total_exp = data.get("experience", "")  # e.g. “3 yrs 6 mos”
    if pos or comp:
        experience.append({
            "position":    pos,
            "company":     comp,
            "dates":       total_exp,
            "description": data.get("experience_summary", "")
        })

    return {
        "name":       name,
        "role":       role,
        "email":      email,
        "phone":      phone,
        "address":    address,
        "summary":    summary,
        "education":  education,
        "experience": experience,
        "skills":     skills
    }
