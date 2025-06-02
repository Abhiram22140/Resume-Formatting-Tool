# ppt_merger.py

from pptx import Presentation

def _replace_text_in_shape(shape, new_text: str):
    """
    Replace the entire contents of shape.text_frame with new_text.
    If new_text is empty, this clears the shape.
    """
    if not shape.has_text_frame:
        return

    tf = shape.text_frame
    tf.clear()  # remove existing paragraphs/runs
    p = tf.paragraphs[0]
    p.text = new_text  # insert the new text (can contain '\n')

def merge_into_template(parsed: dict, template_pptx_path: str, output_path: str):
    """
    1) Load templates/Resume.pptx
    2) Overwrite specific shapes on Slide 0 with parsed data
    3) Save to output_path
    """
    prs = Presentation(template_pptx_path)
    slide = prs.slides[0]  # The template has exactly one slide

    # ---- 1) Shape 5: “Name\nRole” ----
    name_line = parsed.get("name", "")
    role_line = parsed.get("role", "")
    if role_line:
        name_role_text = f"{name_line}\n{role_line}"
    else:
        name_role_text = name_line

    shape_idx_name = 5
    if shape_idx_name < len(slide.shapes):
        _replace_text_in_shape(slide.shapes[shape_idx_name], name_role_text)

    # ---- 2) Shape 0: “Consulting Competencies…” (summary) ----
    shape_idx_summary = 0
    summary_text = parsed.get("summary", "")
    _replace_text_in_shape(slide.shapes[shape_idx_summary], summary_text)

    # ---- 3) Shape 1: “Experience…” block ----
    shape_idx_experience = 1
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
                exp_lines.append(f"• {line}")
            exp_lines.append("")  # blank line between entries
        else:
            exp_lines.append(header)
            exp_lines.append("")

    experience_text = "\n".join(exp_lines).strip()
    _replace_text_in_shape(slide.shapes[shape_idx_experience], experience_text)

    # ---- 4) Shape 4: “Education…” block ----
    shape_idx_education = 4
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
    _replace_text_in_shape(slide.shapes[shape_idx_education], education_text)

    # ---- 5) Shape 6: “Picture” placeholder – clear it ----
    shape_idx_picture = 6
    if shape_idx_picture < len(slide.shapes):
        _replace_text_in_shape(slide.shapes[shape_idx_picture], "")

    prs.save(output_path)
