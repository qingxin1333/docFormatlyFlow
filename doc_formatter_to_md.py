"""
文档格式整理工具 - 生成Markdown版本
从docx文件提取内容并生成格式化的markdown文件
"""

import os
import re
from pathlib import Path
from typing import List, Dict
from docx import Document
import logging
from datetime import datetime

# 配置日志
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

class MarkdownDocFormatter:
    def __init__(self, source_dir: str = "source", target_dir: str = "target"):
        self.source_dir = Path(source_dir)
        self.target_dir = Path(target_dir)
        self.target_dir.mkdir(exist_ok=True)
    
    def extract_content_from_docx(self, docx_path: Path) -> List[Dict]:
        """从docx文件提取内容"""
        doc = Document(docx_path)
        content_blocks = []
        
        for paragraph in doc.paragraphs:
            text = paragraph.text.strip()
            if not text:
                continue
            
            # 清理乱码字符
            try:
                text = text.replace('□', '').replace('▢', '').replace('▪', '').replace('▫', '')
                text = text.replace('○', '').replace('●', '')
                text = ''.join(c for c in text if ord(c) >= 32 or c in '\n\r\t')
                text = text.encode('utf-8', errors='ignore').decode('utf-8')
            except Exception:
                pass
            
            content_blocks.append({
                'text': text,
                'raw_text': text
            })
        
        return content_blocks
    
    def analyze_and_organize_content(self, blocks: List[Dict]) -> List[Dict]:
        """分析并组织内容"""
        organized = []
        current_question = ""
        current_answer = []
        question_number = 1
        
        for block in blocks:
            text = block['text']
            
            # 检查是否是题目
            question_patterns = [
                r'^\d+[\.、]\s*',
                r'^\(\d+\)[\.、]\s*',
                r'^第\d+[题问]\s*',
                r'^\d+\s*[\.、]?\s*',
            ]
            
            is_question = any(re.match(pattern, text) for pattern in question_patterns)
            
            # 检查是否是答案开始
            is_answer_start = any(text.startswith(prefix) for prefix in [
                '答案:', '答:', '解答:', '回答:', '解:', 
                '答案：', '答：', '解答：', '回答：', '解：',
                '🎯', '🔍', '核心结论', '扩展追问'
            ])
            
            # 检查是否是标题
            is_title = (
                '引言' in text or '背景' in text or '目标' in text or '范围' in text or
                '架构' in text or '模块' in text or '接口' in text or '数据库' in text or
                '功能' in text or '安全' in text or '性能' in text or '可用' in text
            )
            
            if is_title and not is_question:
                # 保存当前内容
                if current_question:
                    organized.append({
                        "type": "qa_pair",
                        "number": question_number,
                        "question": current_question.strip(),
                        "answer": "\n".join(current_answer).strip()
                    })
                    question_number += 1
                    current_question = ""
                    current_answer = []
                
                # 添加章节标题
                organized.append({
                    "type": "section_title",
                    "text": text.strip()
                })
                
            elif is_question:
                # 保存上一个问题
                if current_question:
                    organized.append({
                        "type": "qa_pair",
                        "number": question_number,
                        "question": current_question.strip(),
                        "answer": "\n".join(current_answer).strip()
                    })
                    question_number += 1
                
                # 开始新问题
                current_question = text
                current_answer = []
                
            elif is_answer_start or (current_question and not is_question):
                # 答案内容
                if is_answer_start:
                    answer_text = text.split(':', 1)[1].strip() if ':' in text else text.split('：', 1)[1].strip() if '：' in text else ""
                    if answer_text:
                        current_answer.append(answer_text)
                else:
                    current_answer.append(text)
                    
            elif current_question and not is_question:
                # 题目的继续内容
                current_question += " " + text
        
        # 保存最后一个问题
        if current_question:
            organized.append({
                "type": "qa_pair",
                "number": question_number,
                "question": current_question.strip(),
                "answer": "\n".join(current_answer).strip()
            })
        
        return organized
    
    def generate_markdown_content(self, organized_content: List[Dict], filename: str) -> str:
        """生成markdown内容"""
        md_content = []
        
        # 添加文档头部
        md_content.append(f"# {filename.replace('_', ' ')}")
        md_content.append("")
        md_content.append(f"**生成时间**: {datetime.now().strftime('%Y年%m月%d日 %H:%M:%S')}")
        md_content.append("")
        md_content.append("---")
        md_content.append("")
        
        # 添加目录
        md_content.append("## 目录")
        md_content.append("")
        
        # 生成目录
        section_count = 0
        for item in organized_content:
            if item['type'] == 'section_title':
                section_count += 1
                md_content.append(f"{section_count}. [{item['text']}](#{item['text'].replace(' ', '-')})")
        
        md_content.append("")
        md_content.append("---")
        md_content.append("")
        
        # 添加正文内容
        for item in organized_content:
            if item['type'] == 'section_title':
                # 章节标题
                md_content.append(f"## {item['text']}")
                md_content.append("")
                
            elif item['type'] == 'qa_pair':
                # 问答对
                question = item['question'].strip()
                answer = item['answer'].strip()
                
                # 清理问题格式
                question = re.sub(r'^\d+[\.、]\s*', '', question)
                question = re.sub(r'^\(\d+\)[\.、]\s*', '', question)
                
                # 问题
                md_content.append(f"### {item['number']}. {question}")
                md_content.append("")
                
                # 答案
                if answer:
                    md_content.append("**答案：**")
                    md_content.append("")
                    
                    # 处理答案内容
                    answer_lines = answer.split('\n')
                    for line in answer_lines:
                        line = line.strip()
                        if line:
                            # 处理选项格式
                            if re.match(r'^[A-D][\.、]\s*', line):
                                md_content.append(f"   {line}")
                            else:
                                md_content.append(line)
                    md_content.append("")
                else:
                    md_content.append("*（暂无答案）*")
                    md_content.append("")
        
        return '\n'.join(md_content)
    
    def process_single_file(self, docx_path: Path) -> Path:
        """处理单个docx文件"""
        try:
            logger.info(f"开始处理文件: {docx_path}")
            
            # 提取内容
            blocks = self.extract_content_from_docx(docx_path)
            
            # 分析和组织内容
            organized = self.analyze_and_organize_content(blocks)
            
            # 生成markdown内容
            filename = docx_path.stem
            md_content = self.generate_markdown_content(organized, filename)
            
            # 保存文件
            output_path = self.target_dir / f"{filename}.md"
            with open(output_path, 'w', encoding='utf-8') as f:
                f.write(md_content)
            
            logger.info(f"Markdown文件已保存: {output_path}")
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
    print("=== 文档格式整理工具 - 生成Markdown版本 ===")
    print("正在从docx文件生成格式化的markdown文件...")
    
    formatter = MarkdownDocFormatter()
    processed_files = formatter.process_all_files()
    
    if processed_files:
        print(f"\n处理完成！成功整理了 {len(processed_files)} 个文件:")
        for file_path in processed_files:
            print(f"  - {file_path}")
        print(f"\n输出目录: {Path('target').absolute()}")
        print("\n生成的markdown文件包含：")
        print("- 文档标题和生成时间")
        print("- 自动生成的目录")
        print("- 格式化的问答内容")
        print("- 支持章节标题和问题编号")
    else:
        print("没有找到可处理的文件")

if __name__ == "__main__":
    main()
