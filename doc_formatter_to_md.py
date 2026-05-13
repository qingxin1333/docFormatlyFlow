"""
文档格式整理工具 - 生成Markdown版本
从docx文件提取内容并生成格式化的markdown文件
"""

import re
from pathlib import Path
from typing import List, Dict
import logging
from datetime import datetime

from docx_formatter import FormattedDocxGenerator

# 配置日志
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

class MarkdownDocFormatter:
    def __init__(self, source_dir: str = "source", target_dir: str = "target"):
        self.source_dir = Path(source_dir)
        self.target_dir = Path(target_dir)
        self.target_dir.mkdir(exist_ok=True)
        self._docx = FormattedDocxGenerator(source_dir=source_dir, target_dir=target_dir)
    
    def extract_content_from_docx(self, docx_path: Path) -> List[Dict]:
        """与 docx 生成器一致：按顺序提取段落与表格。"""
        return self._docx.extract_content_from_docx(docx_path)
    
    def analyze_and_organize_content(self, blocks: List[Dict]) -> List[Dict]:
        """与 docx 生成器一致的问答/表格分块逻辑。"""
        return self._docx.analyze_and_organize_content(blocks)
    
    @staticmethod
    def _rows_to_markdown_table(rows: List[List[str]]) -> str:
        if not rows:
            return ""
        def esc(cell: str) -> str:
            return (cell or "").replace("\\", "\\\\").replace("|", "\\|").replace("\n", " ")
        header = rows[0]
        n = len(header)
        lines = ["| " + " | ".join(esc(h) for h in header) + " |"]
        lines.append("| " + " | ".join(":---" for _ in header) + " |")
        for r in rows[1:]:
            padded = list(r) + [""] * (n - len(r))
            padded = padded[:n]
            lines.append("| " + " | ".join(esc(c) for c in padded) + " |")
        return "\n".join(lines)
    
    def generate_markdown_content(self, organized_content: List[Dict], filename: str) -> str:
        """生成markdown内容（大纲与 Word 导航对应：# 文题 / ## 章节 / ### 题目）。"""
        md_content = []
        
        art = next((x for x in organized_content if x.get("type") == "article_title"), None)
        if art:
            md_content.append(f"# {art['text']}")
        else:
            md_content.append(f"# {filename.replace('_', ' ')}")
        md_content.append("")
        md_content.append(f"**生成时间**: {datetime.now().strftime('%Y年%m月%d日 %H:%M:%S')}")
        md_content.append("")
        md_content.append("---")
        md_content.append("")
        
        md_content.append("## 目录")
        md_content.append("")
        for item in organized_content:
            if item["type"] == "chapter_title":
                md_content.append(f"- 章节：{item['text']}")
            elif item["type"] == "qa_pair":
                q = item["question"].strip().replace("\n", " ")
                if len(q) > 56:
                    q = q[:56] + "…"
                md_content.append(f"- {item['number']}. {q}")
        md_content.append("")
        md_content.append("---")
        md_content.append("")
        
        for item in organized_content:
            if item['type'] == 'article_title':
                continue
            
            if item['type'] == 'chapter_title':
                md_content.append(f"## {item['text']}")
                md_content.append("")
            
            elif item['type'] == 'paragraph':
                md_content.append(item['text'])
                md_content.append("")
            
            elif item['type'] == 'table':
                md_content.append(self._rows_to_markdown_table(item['rows']))
                md_content.append("")
                
            elif item['type'] == 'qa_pair':
                question = item['question'].strip()
                question = re.sub(r'^\d+[\.、]\s*', '', question)
                question = re.sub(r'^\(\d+\)[\.、]\s*', '', question)
                md_content.append(f"### {item['number']}. {question}")
                md_content.append("")
                
                answer_parts = item.get('answer_parts') or []
                if answer_parts:
                    md_content.append("**答案：**")
                    md_content.append("")
                    for seg in answer_parts:
                        if isinstance(seg, dict) and seg.get("type") == "table":
                            md_content.append(self._rows_to_markdown_table(seg["rows"]))
                            md_content.append("")
                        elif isinstance(seg, str):
                            for line in seg.split("\n"):
                                line = line.strip()
                                if line:
                                    if re.match(r'^[A-D][\.、]\s*', line):
                                        md_content.append(f"   {line}")
                                    else:
                                        md_content.append(line)
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
