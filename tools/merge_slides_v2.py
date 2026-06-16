import os
import re

def split_section(title, lines, max_bullets=4):
    """
    Splits a slide content list of lines into multiple slides if it exceeds max_bullets
    or if it has too many lines.
    """
    # Find list items at level 0 (starts with *, -, or \d+.)
    # We will split the list into groups of max_bullets
    items = []
    current_item = []
    
    # We also keep track of intro paragraphs before the list
    intro_lines = []
    has_list_started = False
    
    for line in lines:
        trimmed = line.strip()
        # Match list item: * or - or 1.
        is_bullet = trimmed.startswith('* ') or trimmed.startswith('- ') or re.match(r'^\d+\.\s+', trimmed)
        
        if is_bullet:
            has_list_started = True
            if current_item:
                items.append(current_item)
            current_item = [line]
        else:
            if has_list_started:
                current_item.append(line)
            else:
                intro_lines.append(line)
                
    if current_item:
        items.append(current_item)
        
    if not items:
        # No list items, check total lines
        if len(lines) > 15:
            # Split by lines
            chunks = [lines[i:i+12] for i in range(0, len(lines), 12)]
            slides = []
            for idx, chunk in enumerate(chunks):
                suffix = f" ({idx+1}/{len(chunks)})" if len(chunks) > 1 else ""
                slides.append(f"{title}{suffix}\n" + "\n".join(chunk))
            return slides
        else:
            return [f"{title}\n" + "\n".join(lines)]

    # If we have list items
    if len(items) <= max_bullets:
        return [f"{title}\n" + "\n".join(lines)]
        
    # Split items
    slides = []
    item_chunks = [items[i:i+max_bullets] for i in range(0, len(items), max_bullets)]
    
    for idx, chunk in enumerate(item_chunks):
        suffix = f" ({idx+1}/{len(item_chunks)})"
        chunk_lines = []
        if idx == 0 and intro_lines:
            chunk_lines.extend(intro_lines)
            
        for item in chunk:
            chunk_lines.extend(item)
            
        slides.append(f"{title}{suffix}\n" + "\n".join(chunk_lines).strip())
        
    return slides

def parse_markdown_to_slides(file_path):
    with open(file_path, "r", encoding="utf-8") as f:
        content = f.read()
        
    lines = content.split('\n')
    slides = []
    
    current_title = ""
    current_lines = []
    
    # Cover info
    cover_title = ""
    cover_lines = []
    in_cover = False
    
    for line in lines:
        trimmed = line.strip()
        
        if trimmed.startswith("# "):
            # This is a main section cover slide
            if current_title:
                slides.extend(split_section(current_title, current_lines))
                current_title = ""
                current_lines = []
            
            cover_title = line
            in_cover = True
            cover_lines = []
            continue
            
        if trimmed.startswith("## "):
            if in_cover:
                # Save the cover slide
                slides.append(f"{cover_title}\n" + "\n".join(cover_lines).strip())
                in_cover = False
                cover_lines = []
                
            if current_title:
                slides.extend(split_section(current_title, current_lines))
                
            current_title = line
            current_lines = []
            continue
            
        if in_cover:
            cover_lines.append(line)
        else:
            current_lines.append(line)
            
    if in_cover:
        slides.append(f"{cover_title}\n" + "\n".join(cover_lines).strip())
    elif current_title:
        slides.extend(split_section(current_title, current_lines))
        
    return slides

def main():
    base_dir = r"d:\Git\true-soul-light\不用版控的\世豐脆課程"
    outlines = [
        "脆第一堂課程 理論基礎與實戰1_大綱.md",
        "脆第一堂課程 理論基礎與實戰2_大綱.md",
        "脆第二堂課程案例分享上半部_大綱.md",
        "脆第二堂課案例分享下半部_大綱.md",
        "脆第三堂課程 起號_大綱.md"
    ]
    
    all_slides = []
    
    # Main Presentation Cover
    all_slides.append("# 世豐老師：Threads (脆) 流量變現與自媒體經營實戰完整課程\n## 兩萬字課程精華大綱彙整簡報\n- 講師：世豐\n- 主題：理論基礎、案例分享與起號實戰")
    
    for outline_file in outlines:
        path = os.path.join(base_dir, outline_file)
        if not os.path.exists(path):
            print(f"Warning: File {path} not found.")
            continue
            
        slides = parse_markdown_to_slides(path)
        all_slides.extend(slides)
        
    # Clean up and filter out empty slides
    cleaned_slides = []
    for s in all_slides:
        lines = [l for l in s.split('\n') if l.strip() != '---']
        # check if it has actual content other than just headings
        content_lines = [l for l in lines if l.strip() and not l.strip().startswith('#') and not l.strip().startswith('##')]
        if content_lines:
            cleaned_slides.append("\n".join(lines).strip())
        elif len(lines) > 0 and (lines[0].strip().startswith('# ') or lines[0].strip().startswith('## ')):
            # Keep covers and section titles
            cleaned_slides.append("\n".join(lines).strip())
            
    # Combine with single separator
    final_md = "\n\n---\n\n".join(cleaned_slides)
    
    # Let's replace double slides separators if any
    final_md = re.sub(r'\n\n---\n\n\n\n---\n\n', '\n\n---\n\n', final_md)
    
    # Read template
    template_path = os.path.join(r"d:\Git\true-soul-light", "世豐老師脆課程.html")
    with open(template_path, "r", encoding="utf-8") as f:
        template_html = f.read()

    # Replace the editor textarea content
    pattern = re.compile(r'(<textarea\s+id="editor"[^>]*>)([\s\S]*?)(<\/textarea>)')
    escaped_md = final_md.replace("<", "&lt;").replace(">", "&gt;")
    new_html = pattern.sub(r'\1' + escaped_md + r'\3', template_html)
    
    # Update titles in HTML
    new_html = re.sub(r'<title>.*?<\/title>', r'<title>世豐老師：Threads (脆) 流量變現與自媒體經營實戰完整課程 - 可編輯簡報系統</title>', new_html)
    new_html = re.sub(r'<span id="slide-footer-title">.*?<\/span>', r'<span id="slide-footer-title">世豐老師 Threads (脆) 流量變現與自媒體完整課程</span>', new_html)

    # Save to output file
    output_html_path = os.path.join(r"d:\Git\true-soul-light", "世豐老師脆課程完整版.html")
    with open(output_html_path, "w", encoding="utf-8") as f:
        f.write(new_html)

    print(f"Successfully generated clean and split {output_html_path}")

if __name__ == "__main__":
    main()
