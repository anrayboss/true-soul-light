import sys
from md_to_pptx_jiuyang import parse_jiuyang_markdown, get_portrait_filename

slides = parse_jiuyang_markdown("最終輸出用檔案/行銷界九陽神功_完整講稿.md")
with open("tools/parse_log.txt", "w", encoding="utf-8") as out:
    for i, s in enumerate(slides):
        out.write(f"Slide {i+1}: {s['title']}\n")
        out.write(f"  has_dual_quotes: {s['has_dual_quotes']}\n")
        if s['has_dual_quotes']:
            out.write(f"  quotes_data: {len(s['quotes_data'])} items\n")
            for q in s['quotes_data']:
                p_file = get_portrait_filename(q['name'])
                out.write(f"    Name: {q['name']} -> portrait: {p_file}\n")
print("Log written to tools/parse_log.txt")
