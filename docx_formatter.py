"""
文档格式整理工具 - 生成 docx
统一页面、字体与段落样式；输出以正文为主，不附加封面元数据、修订记录或占位目录。
"""

from pathlib import Path
from typing import List, Dict, Any, Iterator, Union
import re
from docx import Document
from docx.table import Table as DocxTable
from docx.text.paragraph import Paragraph as DocxParagraph
from docx.shared import Pt, Cm
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.oxml.ns import qn
import logging

# 配置日志
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

class FormattedDocxGenerator:
    def __init__(self, source_dir: str = "source", target_dir: str = "target"):
        self.source_dir = Path(source_dir)
        self.target_dir = Path(target_dir)
        self.target_dir.mkdir(exist_ok=True)
    
    def setup_document_format(self, doc: Document):
        """设置文档格式"""
        # 设置页面为A4
        section = doc.sections[0]
        section.page_width = Cm(21.0)  # A4宽度
        section.page_height = Cm(29.7)  # A4高度
        
        # 设置页边距：上下左右均为2.54厘米
        section.left_margin = Cm(2.54)
        section.right_margin = Cm(2.54)
        section.top_margin = Cm(2.54)
        section.bottom_margin = Cm(2.54)
        
        # 设置默认字体
        style = doc.styles['Normal']
        font = style.font
        font.name = '微软雅黑'
        font.size = Pt(12)  # 小四号
        
        # 设置段落格式
        paragraph_format = style.paragraph_format
        paragraph_format.line_spacing = 1.5  # 1.5倍行距
        paragraph_format.space_after = Pt(0)
        paragraph_format.space_before = Pt(0)
        paragraph_format.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY  # 两端对齐
        paragraph_format.first_line_indent = Cm(0.74)  # 首行缩进2字符（约0.74厘米）
    
    def modify_existing_styles(self, doc: Document):
        """修改现有的标题样式"""
        try:
            # 修改一级标题样式
            style1 = doc.styles['Heading 1']
            style1.font.name = '黑体'
            style1.font.size = Pt(16)  # 三号
            style1.font.bold = True
            style1.font.color.rgb = None  # 确保是黑色
            style1.paragraph_format.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT
            style1.paragraph_format.space_before = Pt(12)
            style1.paragraph_format.space_after = Pt(6)
        except:
            pass
        
        try:
            # 修改二级标题样式
            style2 = doc.styles['Heading 2']
            style2.font.name = '黑体'
            style2.font.size = Pt(14)  # 四号
            style2.font.bold = True
            style2.font.color.rgb = None  # 确保是黑色
            style2.paragraph_format.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT
            style2.paragraph_format.space_before = Pt(12)
            style2.paragraph_format.space_after = Pt(6)
        except:
            pass
        
        try:
            # 修改三级标题样式
            style3 = doc.styles['Heading 3']
            style3.font.name = '黑体'
            style3.font.size = Pt(12)  # 小四号
            style3.font.bold = True
            style3.font.color.rgb = None  # 确保是黑色
            style3.paragraph_format.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT
            style3.paragraph_format.space_before = Pt(12)
            style3.paragraph_format.space_after = Pt(6)
        except:
            pass
    
    def _apply_title_style(self, paragraph, font_size):
        """应用标题样式到段落"""
        for run in paragraph.runs:
            run.font.name = '黑体'
            run._element.rPr.rFonts.set(qn('w:eastAsia'), '黑体')  # 明确设置东亚字体
            run._element.rPr.rFonts.set(qn('w:ascii'), '黑体')     # 设置西文字体
            run._element.rPr.rFonts.set(qn('w:hAnsi'), '黑体')     # 设置ANSI字体
            run.font.size = font_size
            run.font.bold = True
            run.font.color.rgb = None
    
    def _apply_answer_content_style(self, paragraph):
        """应用答案内容样式到段落"""
        for run in paragraph.runs:
            run.font.name = '微软雅黑'
            run._element.rPr.rFonts.set(qn('w:eastAsia'), '微软雅黑')  # 明确设置东亚字体
            run._element.rPr.rFonts.set(qn('w:ascii'), '微软雅黑')     # 设置西文字体
            run._element.rPr.rFonts.set(qn('w:hAnsi'), '微软雅黑')     # 设置ANSI字体
            run.font.size = Pt(12)
            run.font.bold = False
            run.font.color.rgb = None
    
    @staticmethod
    def _clean_block_text(text: str) -> str:
        """段落文本清理（与原先 extract 内逻辑一致）。"""
        try:
            text = text.replace('□', '').replace('▢', '').replace('▪', '').replace('▫', '')
            text = text.replace('○', '').replace('●', '')
            text = ''.join(c for c in text if ord(c) >= 32 or c in '\n\r\t')
            text = text.encode('utf-8', errors='ignore').decode('utf-8')
        except Exception:
            pass
        return text.strip()
    
    @staticmethod
    def _iter_body_paragraphs_and_tables(document: Document) -> Iterator[Union[DocxParagraph, DocxTable]]:
        """按文档顺序遍历 body 下的段落与表格（仅 doc.paragraphs 会漏掉表格）。"""
        body = document.element.body
        for child in body:
            if child.tag == qn('w:p'):
                yield DocxParagraph(child, document)
            elif child.tag == qn('w:tbl'):
                yield DocxTable(child, document)
    
    @staticmethod
    def _table_to_rows(table: DocxTable) -> List[List[str]]:
        rows: List[List[str]] = []
        for row in table.rows:
            cells = []
            for cell in row.cells:
                t = cell.text.replace('\r', '\n').strip()
                t = ' '.join(line.strip() for line in t.splitlines() if line.strip())
                cells.append(t)
            rows.append(cells)
        return rows
    
    def extract_content_from_docx(self, docx_path: Path) -> List[Dict]:
        """从 docx 按顺序提取段落与表格（表格在段落之间，与 Word 视图一致）。"""
        doc = Document(docx_path)
        content_blocks: List[Dict[str, Any]] = []
        
        for block in self._iter_body_paragraphs_and_tables(doc):
            if isinstance(block, DocxParagraph):
                text = self._clean_block_text(block.text)
                if not text:
                    continue
                content_blocks.append({'type': 'paragraph', 'text': text, 'raw_text': text})
            else:
                rows = self._table_to_rows(block)
                if not rows or not any(any(c for c in r) for r in rows):
                    continue
                content_blocks.append({'type': 'table', 'rows': rows})
        
        return content_blocks
    
    @staticmethod
    def _line_ends_with_question_mark(text: str) -> bool:
        """仅当本段去掉尾部空白后以 ? 或 ？ 结尾时视为『问句结束』；句中的问号不触发新题。"""
        t = text.rstrip()
        return bool(t) and (t[-1] == "?" or t[-1] == "？")
    
    @staticmethod
    def _looks_like_chapter_heading(text: str) -> bool:
        """
        识别章节类标题（用于导航），避免把答案里的长句误判为章节。
        匹配：含「（数字…道）」、以「第…章」开头等。
        """
        t = text.strip()
        if not t or FormattedDocxGenerator._line_ends_with_question_mark(t):
            return False
        if len(t) > 56:
            return False
        if re.match(r"^(?:问题|解决|解答|注意|说明|示例|小结|优势|缺点)[：:]", t):
            return False
        if re.search(r"[（(]\s*\d+[^）)]*道\s*[）)]", t):
            return True
        if re.match(r"^第[一二三四五六七八九十百千零〇两\d]+章", t):
            return True
        if re.match(r"^第[一二三四五六七八九十百千零〇两\d]+节", t):
            return True
        return False
    
    def analyze_and_organize_content(self, blocks: List[Dict]) -> List[Dict]:
        """分析并组织内容：文首大标题/章节进导航；句末问号界定题目；答案含表格块。"""
        organized: List[Dict[str, Any]] = []
        current_question = ""
        current_answer: List[Any] = []
        question_number = 1
        phase = "preamble"
        preamble: List[str] = []
        
        def flush_preamble():
            nonlocal preamble
            if not preamble:
                return
            organized.append({"type": "article_title", "text": preamble[0].strip()})
            for line in preamble[1:]:
                organized.append({"type": "chapter_title", "text": line.strip()})
            preamble = []
        
        def flush_qa():
            nonlocal current_question, current_answer, question_number
            if not current_question:
                return
            if not self._line_ends_with_question_mark(current_question):
                organized.append({"type": "paragraph", "text": current_question.strip()})
                for seg in current_answer:
                    if isinstance(seg, str) and seg.strip():
                        organized.append({"type": "paragraph", "text": seg.strip()})
                    elif isinstance(seg, dict) and seg.get("type") == "table":
                        organized.append({"type": "table", "rows": seg["rows"]})
                current_question = ""
                current_answer = []
                return
            organized.append({
                "type": "qa_pair",
                "number": question_number,
                "question": current_question.strip(),
                "answer_parts": list(current_answer),
            })
            question_number += 1
            current_question = ""
            current_answer = []
        
        for block in blocks:
            if block.get("type") == "table":
                if phase == "preamble":
                    flush_preamble()
                if current_question and self._line_ends_with_question_mark(current_question):
                    current_answer.append({"type": "table", "rows": block["rows"]})
                else:
                    organized.append({"type": "table", "rows": block["rows"]})
                continue
            
            text = block["text"]
            
            if phase == "preamble":
                if self._line_ends_with_question_mark(text):
                    flush_preamble()
                    phase = "body"
                else:
                    preamble.append(text.strip())
                    continue
            
            is_answer_start = any(
                text.startswith(prefix)
                for prefix in [
                    "答案:", "答:", "解答:", "回答:", "解:",
                    "答案：", "答：", "解答：", "回答：", "解：",
                    "🎯", "🔍", "核心结论", "扩展追问",
                ]
            )
            
            if is_answer_start:
                body = ""
                if ":" in text:
                    body = text.split(":", 1)[1].strip()
                elif "：" in text:
                    body = text.split("：", 1)[1].strip()
                if current_question and self._line_ends_with_question_mark(current_question):
                    if body:
                        current_answer.append(body)
                elif body:
                    organized.append({"type": "paragraph", "text": text.strip()})
                continue
            
            if self._line_ends_with_question_mark(text):
                if current_question and self._line_ends_with_question_mark(current_question):
                    flush_qa()
                if current_question and not self._line_ends_with_question_mark(current_question):
                    current_question = (current_question + " " + text).strip()
                else:
                    current_question = text
                continue
            
            # 无句末问号：题干已以问号结束时视为答案；否则续接未完成题干
            if current_question:
                if self._line_ends_with_question_mark(current_question):
                    current_answer.append(text)
                else:
                    current_question = (current_question + " " + text).strip()
                continue
            
            if self._looks_like_chapter_heading(text):
                organized.append({"type": "chapter_title", "text": text.strip()})
            else:
                organized.append({"type": "paragraph", "text": text.strip()})
        
        flush_qa()
        if phase == "preamble":
            flush_preamble()
        
        return organized
    
    def _add_table_to_document(self, doc: Document, rows: List[List[str]], paragraph_styles: list) -> None:
        """在 docx 中插入表格并登记段落样式（单元格内段落按正文样式）。"""
        if not rows:
            return
        ncols = max(len(r) for r in rows)
        nrows = len(rows)
        if ncols == 0:
            return
        tbl = doc.add_table(rows=nrows, cols=ncols)
        try:
            tbl.style = "Table Grid"
        except (KeyError, ValueError):
            pass
        for ri in range(nrows):
            row_data = rows[ri] if ri < len(rows) else []
            for ci in range(ncols):
                val = row_data[ci] if ci < len(row_data) else ""
                cell = tbl.rows[ri].cells[ci]
                cell.text = val
                for para in cell.paragraphs:
                    para.paragraph_format.first_line_indent = Cm(0)
                    paragraph_styles.append((para, "answer_content", {"is_option": False}))
    
    def create_formatted_docx(self, organized_content: List[Dict]) -> Document:
        """创建格式化的 docx：统一版面与样式，不附加封面、修订记录或占位目录。"""
        doc = Document()
        
        # 设置文档格式
        self.setup_document_format(doc)
        
        # 修改现有样式而不是创建新样式
        self.modify_existing_styles(doc)
        
        # 记录需要应用样式的段落信息
        paragraph_styles = []
        
        # 添加正文内容
        for item in organized_content:
            if item['type'] == 'article_title':
                para = doc.add_paragraph(item['text'])
                paragraph_styles.append((para, 'article_heading', {'font_size': Pt(18)}))

            elif item['type'] == 'chapter_title':
                para = doc.add_paragraph(item['text'])
                paragraph_styles.append((para, 'chapter_heading', {'font_size': Pt(16)}))

            elif item['type'] == 'paragraph':
                para = doc.add_paragraph(item['text'])
                paragraph_styles.append((para, 'answer_content', {'is_option': False}))

            elif item['type'] == 'table':
                self._add_table_to_document(doc, item['rows'], paragraph_styles)
                doc.add_paragraph("")

            elif item['type'] == 'qa_pair':
                # 问答对
                question = item['question'].strip()
                
                # 清理问题格式
                question = re.sub(r'^\d+[\.、]\s*', '', question)
                question = re.sub(r'^\(\d+\)[\.、]\s*', '', question)
                
                # 问题段落
                question_text = f"{item['number']}. {question}"
                question_para = doc.add_paragraph(question_text)
                paragraph_styles.append((question_para, 'question', {'font_size': Pt(14)}))
                
                answer_parts = item.get('answer_parts') or []
                if answer_parts:
                    # 答案标题
                    answer_title_para = doc.add_paragraph("答案：")
                    paragraph_styles.append((answer_title_para, 'answer_title', {}))
                    
                    for seg in answer_parts:
                        if isinstance(seg, dict) and seg.get("type") == "table":
                            self._add_table_to_document(doc, seg["rows"], paragraph_styles)
                        elif isinstance(seg, str):
                            for line in seg.split("\n"):
                                line = line.strip()
                                if line:
                                    answer_para = doc.add_paragraph(line)
                                    is_option = bool(re.match(r'^[A-D][\.、]\s*', line))
                                    paragraph_styles.append((answer_para, 'answer_content', {'is_option': is_option}))
                else:
                    # 无答案提示
                    no_answer_para = doc.add_paragraph("（暂无答案）")
                    paragraph_styles.append((no_answer_para, 'no_answer', {}))
                # 每个问答块后空一行，便于阅读
                doc.add_paragraph("")
        
        # 统一应用所有段落样式
        self._apply_all_paragraph_styles(doc, paragraph_styles)
        
        return doc
    
    def _apply_all_paragraph_styles(self, doc, paragraph_styles):
        """统一应用所有段落样式"""
        for paragraph, style_type, style_params in paragraph_styles:
            if style_type == 'article_heading':
                paragraph.style = doc.styles['Heading 1']
                self._apply_title_style(paragraph, style_params['font_size'])
            elif style_type == 'chapter_heading':
                paragraph.style = doc.styles['Heading 2']
                self._apply_title_style(paragraph, style_params['font_size'])
            elif style_type == 'question':
                # 标题三：导航中显示题目，且低于文首标题/章节
                paragraph.style = doc.styles['Heading 3']
                self._apply_title_style(paragraph, style_params['font_size'])
                
            elif style_type == 'answer_title':
                # 「答案：」与正文统一为微软雅黑正文，仅带问号的问题行用标题字体强调
                paragraph.style = doc.styles['Normal']
                self._apply_answer_content_style(paragraph)
                
            elif style_type == 'answer_content':
                # 答案内容样式
                paragraph.style = doc.styles['Normal']
                self._apply_answer_content_style(paragraph)
                
                # 如果是选项，设置缩进
                if style_params.get('is_option', False):
                    paragraph.paragraph_format.left_indent = Cm(0.74)
                    
            elif style_type == 'no_answer':
                # 无答案提示样式
                for run in paragraph.runs:
                    run.font.name = '微软雅黑'
                    run._element.rPr.rFonts.set(qn('w:eastAsia'), '微软雅黑')  # 明确设置东亚字体
                    run._element.rPr.rFonts.set(qn('w:ascii'), '微软雅黑')     # 设置西文字体
                    run._element.rPr.rFonts.set(qn('w:hAnsi'), '微软雅黑')     # 设置ANSI字体
                    run.font.size = Pt(12)
                    run.font.italic = True
                    run.font.color.rgb = None  # 确保是黑色
    
    def process_single_file(self, docx_path: Path) -> Path:
        """处理单个docx文件"""
        try:
            logger.info(f"开始处理文件: {docx_path}")
            
            # 提取内容
            blocks = self.extract_content_from_docx(docx_path)
            
            # 分析和组织内容
            organized = self.analyze_and_organize_content(blocks)
            
            # 创建格式化的docx
            filename = docx_path.stem
            formatted_doc = self.create_formatted_docx(organized)
            
            # 保存文件 - 直接覆盖原文件名
            output_path = self.target_dir / f"{filename}_formatted.docx"
            formatted_doc.save(output_path)
            
            logger.info(f"格式化docx文件已保存: {output_path}")
            return output_path
            
        except Exception as e:
            logger.error(f"处理文件 {docx_path} 失败: {e}")
            return None
    
    def process_all_files(self) -> List[Path]:
        """处理所有docx文件"""
        processed_files = []
        docx_files = sorted(
            p for p in self.source_dir.glob("*.docx") if not p.name.startswith("~$")
        )
        
        if not docx_files:
            logger.warning(f"在 {self.source_dir} 目录中未找到docx文件")
            return processed_files
        
        logger.info(f"找到 {len(docx_files)} 个docx文件")
        
        for docx_file in docx_files:
            output_path = self.process_single_file(docx_file)
            if output_path:
                processed_files.append(output_path)
        
        return processed_files

def main():
    print("=== 文档格式整理工具 - 生成规范格式docx版本 ===")
    print("正在按照文档规范生成格式化的docx文件...")
    
    formatter = FormattedDocxGenerator()
    processed_files = formatter.process_all_files()
    
    if processed_files:
        print(f"\n处理完成！成功整理了 {len(processed_files)} 个文件:")
        for file_path in processed_files:
            print(f"  - {file_path}")
        print(f"\n输出目录: {Path('target').absolute()}")
        print("\n生成的文档为统一排版后的正文（无封面元数据、修订记录与占位目录）。")
    else:
        print("没有找到可处理的文件")

if __name__ == "__main__":
    main()
