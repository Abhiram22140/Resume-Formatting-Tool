# ppt_merger.py

from pptx import Presentation
from pptx.dml.color import RGBColor

def _copy_font_properties(src_run, dest_run):
    """
    Copy font properties (size, name, bold, italic, color) from src_run to dest_run.
    If src_run.font.color has no .rgb (e.g. _NoneColor), skip color copying.
    """
    src_font = src_run.font
    dst_font = dest_run.font

    # Copy size & font name if present
    if src_font.size:
        dst_font.size = src_font.size
    if src_font.name:
        dst_font.name = src_font.name

    # Copy bold/italic booleans
    dst_font.bold = src_font.bold
    dst_font.italic = src_font.italic

    # Copy color if src has an actual RGB value
    if src_font.color:
        try:
            rgb = src_font.color.rgb  # may raise AttributeError if no color set
            dst_font.color.rgb = RGBColor(rgb[0], rgb[1], rgb[2])
        except AttributeError:
            pass  # leave dest color as default


def _replace_text_preserve_format(shape, new_text: str):
    """
    Replace the text of a shape while preserving font formatting:
      1) Grab the first run of the first paragraph (if exists) as a “style sample.”
      2) Clear the frame.
      3) For each line in new_text, create a new paragraph & run, then copy style.
    """
    if not shape.has_text_frame:
        return

    tf = shape.text_frame

    # Grab sample run if available
    src_run = None
    if tf.paragraphs and tf.paragraphs[0].runs:
        src_run = tf.paragraphs[0].runs[0]

    # Clear existing text
    tf.clear()

    # Insert new text, one paragraph per line
    for i, line in enumerate(new_text.split("\n")):
        if i == 0:
            p = tf.paragraphs[0]
            p.text = ""
        else:
            p = tf.add_paragraph()
        run = p.add_run()
        run.text = line
        if src_run:
            _copy_font_properties(src_run, run)


def merge_into_template(parsed: dict, template_pptx_path: str, output_path: str):
    """
    1) Load the PPT template (templates/Resume.pptx)
    2) Replace shapes at these indexes (0–6) with parsed data, preserving fonts:
         • Shape 5: Name + newline + Role
         • Shape 2: “Email · Phone · Address”
         • Shape 3: “Skills: …”
         • Shape 0: Summary
         • Shape 1: Experience
         • Shape 4: Education
         • Shape 6: Picture (cleared)
    3) Save to output_path.
    """

    prs = Presentation(template_pptx_path)
    slide = prs.slides[0]  # assume exactly one slide in the template

    # 1) Shape 5: Name + newline + Role
    name = parsed.get("name", "")
    role = parsed.get("role", "")
    if role:
        name_role_text = f"{name}\n{role}"
    else:
        name_role_text = name

    shape_idx_name = 5
    if shape_idx_name < len(slide.shapes):
        _replace_text_preserve_format(slide.shapes[shape_idx_name], name_role_text)

    # 2) Shape 2: Email · Phone · Address
    email   = parsed.get("email", "")
    phone   = parsed.get("phone", "")
    address = parsed.get("address", "")
    contact_parts = [p for p in (email, phone, address) if p]
    contact_text = " · ".join(contact_parts)  # e.g. “alice@example.com · +1234567890 · City, Country”
    shape_idx_contact = 2
    if shape_idx_contact < len(slide.shapes):
        _replace_text_preserve_format(slide.shapes[shape_idx_contact], contact_text)

    # 3) Shape 3: “Skills: …”
    skills = parsed.get("skills", [])
    if skills:
        skills_text = "Skills: " + ", ".join(skills)
    else:
        skills_text = ""
    shape_idx_skills = 3
    if shape_idx_skills < len(slide.shapes):
        _replace_text_preserve_format(slide.shapes[shape_idx_skills], skills_text)

    # 4) Shape 0: Summary
    summary = parsed.get("summary", "")
    shape_idx_summary = 0
    if shape_idx_summary < len(slide.shapes):
        _replace_text_preserve_format(slide.shapes[shape_idx_summary], summary)

    # 5) Shape 1: Experience
    exp_list = parsed.get("experience", [])
    exp_lines = []
    for exp in exp_list:
        pos   = exp.get("position", "")
        comp  = exp.get("company", "")
        dates = exp.get("dates", "")
        desc  = exp.get("description", "")
        header = f"{pos}, {comp} ({dates})".strip()
        if desc:
            exp_lines.append(header)
            for line in desc.strip().split("\n"):
                exp_lines.append(f"• {line.lstrip('• ')}")
            exp_lines.append("")  # blank line between entries
        else:
            exp_lines.append(header)
            exp_lines.append("")

    experience_text = "\n".join(exp_lines).strip()
    shape_idx_experience = 1
    if shape_idx_experience < len(slide.shapes):
        _replace_text_preserve_format(slide.shapes[shape_idx_experience], experience_text)

    # 6) Shape 4: Education
    edu_list = parsed.get("education", [])
    edu_lines = []
    for edu in edu_list:
        degree = edu.get("degree", "")
        inst   = edu.get("institution", "")
        start  = edu.get("start", "")
        end    = edu.get("end", "")
        line = f"{degree}, {inst} ({start} – {end})".strip()
        edu_lines.append(line)

    education_text = "\n".join(edu_lines).strip()
    shape_idx_education = 4
    if shape_idx_education < len(slide.shapes):
        _replace_text_preserve_format(slide.shapes[shape_idx_education], education_text)

    # 7) Shape 6: Picture – clear it
    shape_idx_picture = 6
    if shape_idx_picture < len(slide.shapes):
        _replace_text_preserve_format(slide.shapes[shape_idx_picture], "")

    prs.save(output_path)
