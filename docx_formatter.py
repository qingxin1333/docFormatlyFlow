"""
文档格式整理工具 - 生成 docx
统一页面、字体与段落样式；输出以正文为主，不附加封面元数据、修订记录或占位目录。
"""

from pathlib import Path
from typing import List, Dict
import re
from docx import Document
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
    
    def _apply_answer_title_style(self, paragraph):
        """应用答案标题样式到段落"""
        for run in paragraph.runs:
            run.font.name = '黑体'
            run._element.rPr.rFonts.set(qn('w:eastAsia'), '黑体')  # 明确设置东亚字体
            run._element.rPr.rFonts.set(qn('w:ascii'), '黑体')     # 设置西文字体
            run._element.rPr.rFonts.set(qn('w:hAnsi'), '黑体')     # 设置ANSI字体
            run.font.size = Pt(12)
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
    
    def extract_content_from_docx(self, docx_path: Path) -> List[Dict]:
        """从docx文件提取内容"""
        doc = Document(docx_path)
        content_blocks = []
        
        for paragraph in doc.paragraphs:
            text = paragraph.text.strip()
            if not text:
                continue
            
            # 彻底清理字符编码问题
            try:
                # 移除所有可能的乱码字符
                text = text.replace('□', '')
                text = text.replace('▢', '')
                text = text.replace('▪', '')
                text = text.replace('▫', '')
                text = text.replace('○', '')
                text = text.replace('●', '')
                
                # 清理不可见字符
                text = ''.join(c for c in text if ord(c) >= 32 or c in '\n\r\t')
                
                # 重新编码确保正确
                text = text.encode('utf-8', errors='ignore').decode('utf-8')
                
            except Exception as e:
                # 如果清理失败，使用原始文本
                pass
            
            content_blocks.append({
                'text': text,
                'raw_text': text
            })
        
        return content_blocks
    
    @staticmethod
    def _line_has_question_mark(text: str) -> bool:
        """含英文 ? 或中文 ？ 则视为问题相关行（支持句中或句末问号）。"""
        return "?" in text or "？" in text
    
    def analyze_and_organize_content(self, blocks: List[Dict]) -> List[Dict]:
        """分析并组织内容：以问号识别问题；无问号的续行归入答案或题干。"""
        organized = []
        current_question = ""
        current_answer: List[str] = []
        question_number = 1
        
        def flush_qa():
            nonlocal current_question, current_answer, question_number
            if not current_question:
                return
            if not self._line_has_question_mark(current_question):
                organized.append({"type": "paragraph", "text": current_question.strip()})
                for line in current_answer:
                    if line.strip():
                        organized.append({"type": "paragraph", "text": line.strip()})
                current_question = ""
                current_answer = []
                return
            organized.append({
                "type": "qa_pair",
                "number": question_number,
                "question": current_question.strip(),
                "answer": "\n".join(current_answer).strip(),
            })
            question_number += 1
            current_question = ""
            current_answer = []
        
        for block in blocks:
            text = block["text"]
            
            is_answer_start = any(
                text.startswith(prefix)
                for prefix in [
                    "答案:", "答:", "解答:", "回答:", "解:",
                    "答案：", "答：", "解答：", "回答：", "解：",
                    "🎯", "🔍", "核心结论", "扩展追问",
                ]
            )
            
            is_title = (
                "引言" in text or "背景" in text or "目标" in text or "范围" in text or
                "架构" in text or "模块" in text or "接口" in text or "数据库" in text or
                "功能" in text or "安全" in text or "性能" in text or "可用" in text
            )
            
            if is_title and not self._line_has_question_mark(text):
                flush_qa()
                organized.append({"type": "section_title", "text": text.strip()})
                continue
            
            if is_answer_start:
                body = ""
                if ":" in text:
                    body = text.split(":", 1)[1].strip()
                elif "：" in text:
                    body = text.split("：", 1)[1].strip()
                if current_question and self._line_has_question_mark(current_question):
                    if body:
                        current_answer.append(body)
                elif body:
                    organized.append({"type": "paragraph", "text": text.strip()})
                continue
            
            if self._line_has_question_mark(text):
                if current_question and self._line_has_question_mark(current_question):
                    flush_qa()
                if current_question and not self._line_has_question_mark(current_question):
                    current_question = (current_question + " " + text).strip()
                else:
                    current_question = text
                continue
            
            # 无问号：已有关键问句则视为答案；否则续接未完成题干或独立段落
            if current_question:
                if self._line_has_question_mark(current_question):
                    current_answer.append(text)
                else:
                    current_question = (current_question + " " + text).strip()
                continue
            
            organized.append({"type": "paragraph", "text": text.strip()})
        
        flush_qa()
        
        return organized
    
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
            if item['type'] == 'paragraph':
                para = doc.add_paragraph(item['text'])
                paragraph_styles.append((para, 'answer_content', {'is_option': False}))

            elif item['type'] == 'section_title':
                # 章节标题
                section_text = item['text']
                para = doc.add_paragraph(section_text)
                paragraph_styles.append((para, 'section_title', {'font_size': Pt(16)}))
                
            elif item['type'] == 'qa_pair':
                # 问答对
                question = item['question'].strip()
                answer = item['answer'].strip()
                
                # 清理问题格式
                question = re.sub(r'^\d+[\.、]\s*', '', question)
                question = re.sub(r'^\(\d+\)[\.、]\s*', '', question)
                
                # 问题段落
                question_text = f"{item['number']}. {question}"
                question_para = doc.add_paragraph(question_text)
                paragraph_styles.append((question_para, 'question', {'font_size': Pt(14)}))
                
                # 答案
                if answer:
                    # 答案标题
                    answer_title_para = doc.add_paragraph("答案：")
                    paragraph_styles.append((answer_title_para, 'answer_title', {}))
                    
                    # 答案内容
                    answer_lines = answer.split('\n')
                    for line in answer_lines:
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
            if style_type == 'section_title':
                # 章节标题样式
                paragraph.style = doc.styles['Heading 1']
                self._apply_title_style(paragraph, style_params['font_size'])
                
            elif style_type == 'question':
                # 问题标题样式
                paragraph.style = doc.styles['Heading 2']
                self._apply_title_style(paragraph, style_params['font_size'])
                
            elif style_type == 'answer_title':
                # 答案标题样式
                paragraph.style = doc.styles['Normal']
                self._apply_answer_title_style(paragraph)
                
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
            output_path = self.target_dir / f"{filename}.docx"
            formatted_doc.save(output_path)
            
            logger.info(f"格式化docx文件已保存: {output_path}")
            return output_path
            
        except Exception as e:
            logger.error(f"处理文件 {docx_path} 失败: {e}")
            return None
    
    def process_all_files(self) -> List[Path]:
        """处理所有docx文件"""
        processed_files = []
        docx_files = list(self.source_dir.glob("*.docx"))
        
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
