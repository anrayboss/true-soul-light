import re
from html.parser import HTMLParser

class GeminiMarkdownExtractor(HTMLParser):
    def __init__(self):
        super().__init__()
        self.output = []
        self.in_user = False
        self.in_model = False
        self.user_depth = 0
        self.model_depth = 0
        
        # Table state
        self.in_table = False
        self.table_rows = []
        self.current_row = []
        self.current_cell = []
        
        # List state
        self.list_stack = [] # Stack of ('ul' or 'ol', next_index)
        
        # Code state
        self.in_code = False
        
        # Noise filter
        # Avoid capturing Google UI footer texts
        self.noise_keywords = [
            "Google 隱私權政策",
            "在新視窗中開啟",
            "Google 服務條款",
            "個人隱私權與 Gemini 應用程式",
            "Gemini 可能會提供不準確的資訊",
            "前往 Gemini"
        ]

    def write_text(self, text):
        if self.in_table:
            self.current_cell.append(text)
        else:
            # If we're writing normal text, avoid redundant leading spaces if the line is empty
            self.output.append(text)

    def handle_starttag(self, tag, attrs):
        attr_dict = dict(attrs)
        cls = attr_dict.get('class', '')
        
        # --- Check for Speaker Containers ---
        if 'query-text' in cls or 'user-query-container' in cls:
            if not self.in_user:
                self.in_user = True
                self.user_depth = 1
                self.write_text("\n\n---\n\n### 👤 使用者\n\n")
            else:
                self.user_depth += 1
            return
            
        elif 'message-content' in cls or 'markdown-main-panel' in cls:
            if not self.in_model:
                self.in_model = True
                self.model_depth = 1
                self.write_text("\n\n---\n\n### 🤖 Gemini\n\n")
            else:
                self.model_depth += 1
            return

        # If we are in either speaker context, parse content tags
        if self.in_user or self.in_model:
            if self.in_user:
                self.user_depth += 1
            if self.in_model:
                self.model_depth += 1
                
            # Block-level or formatting tags
            if tag in ['strong', 'b']:
                self.write_text("**")
            elif tag in ['em', 'i']:
                self.write_text("*")
            elif tag in ['h1', 'h2', 'h3', 'h4', 'h5', 'h6']:
                level = int(tag[1])
                self.write_text(f"\n\n{'#' * level} ")
            elif tag == 'p':
                # Only add newline if we're not inside a table cell
                if not self.in_table:
                    self.write_text("\n\n")
            elif tag == 'br':
                self.write_text("\n")
            elif tag == 'pre':
                self.in_code = True
                self.write_text("\n\n```\n")
            elif tag == 'code' and not self.in_code:
                self.write_text("`")
            elif tag == 'ul':
                self.list_stack.append(('ul', 0))
                if not self.in_table:
                    self.write_text("\n")
            elif tag == 'ol':
                self.list_stack.append(('ol', 1))
                if not self.in_table:
                    self.write_text("\n")
            elif tag == 'li':
                if self.list_stack:
                    indent = len(self.list_stack) - 1
                    list_type, index = self.list_stack[-1]
                    prefix = ""
                    if list_type == 'ul':
                        prefix = "- "
                    else:
                        prefix = f"{index}. "
                        self.list_stack[-1] = ('ol', index + 1)
                    
                    self.write_text(f"\n{'  ' * indent}{prefix}")
            elif tag == 'table':
                self.in_table = True
                self.table_rows = []
                self.current_row = []
            elif tag == 'tr':
                self.current_row = []
            elif tag in ['td', 'th']:
                self.current_cell = []

    def handle_endtag(self, tag):
        # --- Handle Container Exits ---
        if self.in_user:
            self.user_depth -= 1
            if self.user_depth == 0:
                self.in_user = False
                return
                
        if self.in_model:
            self.model_depth -= 1
            if self.model_depth == 0:
                self.in_model = False
                return

        # Handle tag closure formatting
        if self.in_user or self.in_model:
            if tag in ['strong', 'b']:
                self.write_text("**")
            elif tag in ['em', 'i']:
                self.write_text("*")
            elif tag in ['h1', 'h2', 'h3', 'h4', 'h5', 'h6']:
                self.write_text("\n\n")
            elif tag == 'p':
                if not self.in_table:
                    self.write_text("\n\n")
            elif tag == 'pre':
                self.in_code = False
                self.write_text("\n```\n\n")
            elif tag == 'code' and not self.in_code:
                self.write_text("`")
            elif tag in ['ul', 'ol']:
                if self.list_stack:
                    self.list_stack.pop()
                if not self.in_table:
                    self.write_text("\n")
            elif tag == 'li':
                pass # newlines are handled at the start of next li or container end
            elif tag == 'table':
                # Render table
                if self.table_rows:
                    col_count = max(len(row) for row in self.table_rows) if self.table_rows else 0
                    if col_count > 0:
                        table_lines = []
                        # Process rows, ensure all have equal columns
                        processed_rows = []
                        for row in self.table_rows:
                            processed_rows.append(row + [""] * (col_count - len(row)))
                        
                        # Header Row
                        table_lines.append("| " + " | ".join(processed_rows[0]) + " |")
                        # Separator Row
                        table_lines.append("| " + " | ".join(["---"] * col_count) + " |")
                        # Data Rows
                        for row in processed_rows[1:]:
                            table_lines.append("| " + " | ".join(row) + " |")
                        
                        formatted_table = "\n\n" + "\n".join(table_lines) + "\n\n"
                        self.in_table = False
                        self.write_text(formatted_table)
                self.in_table = False
            elif tag == 'tr':
                self.table_rows.append(self.current_row)
            elif tag in ['td', 'th']:
                cell_text = "".join(self.current_cell).strip().replace("\n", " ")
                self.current_row.append(cell_text)

    def handle_data(self, data):
        if self.in_user or self.in_model:
            # Noise filter: if data matches any noise keywords, ignore it
            cleaned_data = data.strip()
            if any(noise in cleaned_data for noise in self.noise_keywords):
                return
                
            if self.in_code:
                self.write_text(data)
            else:
                # Collapse excessive spaces inside text chunks, but preserve chinese layout
                self.write_text(data)

def clean_markdown(text):
    # Post-process the markdown to remove trailing helper text and tidy up blank lines
    
    # 1. Normalize linebreaks
    text = re.sub(r'\n{3,}', '\n\n', text)
    
    # 2. Strip out trailing Google/Gemini navigation noise from model turns
    noise_patterns = [
        r'Google 隱私權政策\s*在新視窗中開啟',
        r'Google 服務條款\s*在新視窗中開啟',
        r'個人隱私權與 Gemini 應用程式\s*在新視窗中開啟',
        r'Gemini 可能會提供不準確的資訊.*',
        r'前往 Gemini',
        r'this\.gbar_=.*',
        r'/\*\s*Copyright Google LLC.*'
    ]
    for pattern in noise_patterns:
        text = re.compile(pattern, re.DOTALL | re.IGNORECASE).sub('', text)
        
    # Clean up empty spaces and make it tidy
    text = text.strip()
    return text

def main():
    input_file = r"d:\Git\true-soul-light\_Gemini - 直接與 Google AI 互動.htm"
    output_file = r"d:\Git\true-soul-light\_Gemini - 整理對話內容.md"
    
    print("正在讀取原始 HTML 檔案...")
    with open(input_file, 'r', encoding='utf-8', errors='replace') as f:
        html_content = f.read()
        
    print("正在解析對話內容並轉換成 Markdown 格式...")
    extractor = GeminiMarkdownExtractor()
    extractor.feed(html_content)
    
    raw_markdown = "".join(extractor.output)
    clean_md = clean_markdown(raw_markdown)
    
    # Let's insert a beautiful header
    header = """# 🌌 真靈光全球身心靈科技生態系：Gemini 對話紀錄整理

此文件整理自原始網頁前端存檔，記錄了陳老師與 Gemini 關於「真靈光全球身心靈科技生態系」之預算配置、三大系統、組織分工與時間軸規劃。

---
"""
    final_output = header + "\n" + clean_md
    
    print("正在寫入輸出檔案...")
    with open(output_file, 'w', encoding='utf-8') as f:
        f.write(final_output)
        
    print(f"整理完成！乾淨對話已儲存至：{output_file}")

if __name__ == "__main__":
    main()
