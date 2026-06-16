import os
import re

def main():
    base_dir = r"d:\Git\true-soul-light\不用版控的\世豐脆課程"
    outlines = [
        "脆第一堂課程 理論基礎與實戰1_大綱.md",
        "脆第一堂課程 理論基礎與實戰2_大綱.md",
        "脆第二堂課程案例分享上半部_大綱.md",
        "脆第二堂課案例分享下半部_大綱.md",
        "脆第三堂課程 起號_大綱.md"
    ]

    combined_markdown = []

    # Slide 0: Cover of the complete course
    combined_markdown.append("# 世豐老師：Threads (脆) 流量變現與自媒體經營實戰完整課程\n## 兩萬字課程精華大綱彙整簡報\n- 講師：世豐\n- 主題：理論基礎、案例分享與起號實戰")

    for outline_file in outlines:
        path = os.path.join(base_dir, outline_file)
        if not os.path.exists(path):
            print(f"Warning: File {path} not found.")
            continue
            
        with open(path, "r", encoding="utf-8") as f:
            content = f.read()

        # Split the file by Markdown headings or slide sections
        # Let's clean up title headers and keep sections
        lines = content.split('\n')
        slide_title = ""
        current_slide_lines = []
        in_slide = False

        i = 0
        while i < len(lines):
            line = lines[i]
            # If it's a top-level header (e.g. # 脆第一堂課程...), make it a title break
            if line.strip().startswith("# "):
                title_text = line.replace("#", "").strip()
                if combined_markdown:
                    combined_markdown.append(f"# {title_text}\n## 課程大綱架構與精華")
                i += 1
                continue
            
            # If it's a section header (e.g. ## 一、 ...), it starts a new slide
            if line.strip().startswith("## "):
                # If we have collected slide lines, save them
                if current_slide_lines:
                    combined_markdown.append("\n".join(current_slide_lines).strip())
                    current_slide_lines = []
                
                # Check if it is the long case studies section
                if "學員案例診斷與優化建議" in line:
                    # Let's parse case studies and split them into 3 slides
                    cases_lines = []
                    # read forward until next ## or end
                    i += 1
                    while i < len(lines) and not lines[i].strip().startswith("## "):
                        cases_lines.append(lines[i])
                        i += 1
                    
                    # Split cases by bullet points
                    # Case bullets start with * **1. ...** or * **2. ...**
                    cases = []
                    current_case = []
                    for cl in cases_lines:
                        if cl.strip().startswith("* **") or cl.strip().startswith("- **") or re.match(r'^\s*\*?\s*\d+\.\s+\*\*.*?\*\*', cl.strip()) or (cl.strip().startswith("* ") and "**" in cl):
                            if current_case:
                                cases.append(current_case)
                            current_case = [cl]
                        else:
                            current_case.append(cl)
                    if current_case:
                        cases.append(current_case)
                    
                    # Group into 3 slides
                    groups = [
                        ("## 四、 學員案例診斷與優化建議 (1/3)", cases[0:4]),
                        ("## 四、 學員案例診斷與優化建議 (2/3)", cases[4:8]),
                        ("## 四、 學員案例診斷與優化建議 (3/3)", cases[8:])
                    ]
                    for title, grp in groups:
                        grp_lines = []
                        grp_lines.append(title)
                        for case in grp:
                            grp_lines.extend(case)
                        combined_markdown.append("\n".join(grp_lines).strip())
                    
                    # Don't increment i again
                    continue

                # Check if it's the 3-stages start
                elif "起號三階段實戰架構" in line:
                    # We want to split it by stage: 測試階段, 發酵階段, 引爆階段
                    stage_lines = []
                    i += 1
                    while i < len(lines) and not lines[i].strip().startswith("## "):
                        stage_lines.append(lines[i])
                        i += 1
                    
                    # Split stages by ### 階段
                    stages = []
                    current_stage = []
                    stage_title = "## 二、 起號三階段實戰架構"
                    for sl in stage_lines:
                        if sl.strip().startswith("### 階段") or sl.strip().startswith("###"):
                            if current_stage:
                                stages.append((stage_title, current_stage))
                            stage_title = "## 二、 起號實戰架構 — " + sl.replace("###", "").strip()
                            current_stage = []
                        else:
                            current_stage.append(sl)
                    if current_stage:
                        stages.append((stage_title, current_stage))
                    
                    for st_title, st_body in stages:
                        combined_markdown.append(f"{st_title}\n" + "\n".join(st_body).strip())
                    continue
                
                else:
                    current_slide_lines.append(line)
            else:
                if combined_markdown or current_slide_lines:
                    current_slide_lines.append(line)
            
            i += 1

        if current_slide_lines:
            combined_markdown.append("\n".join(current_slide_lines).strip())

    # Build the final combined markdown text
    final_md = "\n\n---\n\n".join(combined_markdown)

    # Read base template
    template_path = os.path.join(r"d:\Git\true-soul-light", "世豐老師脆課程.html")
    with open(template_path, "r", encoding="utf-8") as f:
        template_html = f.read()

    # Replace the editor textarea content with our final combined markdown
    # <textarea id="editor"[^>]*>([\s\S]*?)<\/textarea>
    pattern = re.compile(r'(<textarea\s+id="editor"[^>]*>)([\s\S]*?)(<\/textarea>)')
    
    # We escape html tags just in case
    escaped_md = final_md.replace("<", "&lt;").replace(">", "&gt;")
    new_html = pattern.sub(r'\1' + escaped_md + r'\3', template_html)
    
    # Also update the title in the HTML template
    new_html = re.sub(r'<title>.*?<\/title>', r'<title>世豐老師：Threads (脆) 流量變現與自媒體經營實戰完整課程 - 可編輯簡報系統</title>', new_html)
    new_html = re.sub(r'<span id="slide-footer-title">.*?<\/span>', r'<span id="slide-footer-title">世豐老師 Threads (脆) 流量變現與自媒體完整課程</span>', new_html)

    # Save to new file
    output_html_path = os.path.join(r"d:\Git\true-soul-light", "世豐老師脆課程完整版.html")
    with open(output_html_path, "w", encoding="utf-8") as f:
        f.write(new_html)

    print(f"Successfully generated {output_html_path}")

    # Now let's update convert_slides.py to run on this new file
    slides_script_path = os.path.join(r"d:\Git\true-soul-light", "convert_slides.py")
    with open(slides_script_path, "r", encoding="utf-8") as f:
        script_content = f.read()
    
    # Update the files array
    new_script = script_content.replace(
        "        '世豐老師脆課程.html',",
        "        '世豐老師脆課程.html',\n        '世豐老師脆課程完整版.html',"
    )
    
    with open(slides_script_path, "w", encoding="utf-8") as f:
        f.write(new_script)
        
    print(f"Updated {slides_script_path}")

if __name__ == "__main__":
    main()
