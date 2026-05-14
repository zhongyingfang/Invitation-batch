"""
批量处理主程序
协调Excel读取、模板填充和文件转换
"""
import os
from pathlib import Path
from excel_reader import ExcelReader
from pptx_handler import PPTXHandler
from docx_handler import DOCXHandler
from typing import List, Dict, Any, Callable, Optional


class BatchProcessor:
    """批量处理邀请函"""

    def __init__(self, excel_path: str, pptx_template: str | None, docx_template: str | None,
                 output_docs_dir: str, output_images_dir: str):
        self.excel_reader = ExcelReader(excel_path)
        self.pptx_handler = PPTXHandler(pptx_template, output_docs_dir) if pptx_template else None
        self.docx_handler = DOCXHandler(docx_template, output_docs_dir) if docx_template else None
        self.output_docs_dir = output_docs_dir
        self.output_images_dir = output_images_dir
        self.data = []
    
    def read_data(self) -> List[Dict[str, Any]]:
        """读取Excel数据"""
        self.data = self.excel_reader.read()

        for record in self.data:
            zhiwu = record.get("职务")
            if zhiwu is None or str(zhiwu).strip() == "":
                xingbie = record.get("性别")
                xingbie_str = str(xingbie).strip() if xingbie is not None else ""
                if "女" in xingbie_str:
                    record["职务"] = "女士"
                else:
                    record["职务"] = "先生"

        return self.data
    
    def process(self, generate_png: bool = True, progress_callback: Optional[Callable[[int, int, str], None]] = None):
        if not self.data:
            self.read_data()

        total = len(self.data)
        print(f"\n开始处理 {total} 条邀请函数据...")

        success_count = 0
        error_count = 0

        for i, record in enumerate(self.data, 1):
            try:
                name = record.get("姓名", f"person_{i}")

                print(f"\n[{i}/{total}] 处理: {name}")

                pptx_file = None
                docx_file = None

                if self.pptx_handler:
                    pptx_output_name = str(name)
                    pptx_file = self.pptx_handler.fill_template(record, pptx_output_name)

                if self.docx_handler:
                    docx_output_name = str(name)
                    docx_file = self.docx_handler.fill_template(record, docx_output_name)

                if generate_png:
                    if self.pptx_handler and pptx_file:
                        pptx_png_dir = os.path.join(self.output_images_dir, "电子邀请函")
                        self.pptx_handler.convert_to_png(pptx_file, pptx_png_dir)

                    if self.docx_handler and docx_file:
                        docx_png_dir = os.path.join(self.output_images_dir, "政府邀请函")
                        self.docx_handler.convert_to_png(docx_file, docx_png_dir)

                success_count += 1
                print(f"[OK] 成功处理: {name}")

                if progress_callback:
                    progress_callback(i, total, str(name))

            except Exception as e:
                error_count += 1
                print(f"[FAIL] 处理失败 ({name}): {str(e)}")
                if progress_callback:
                    progress_callback(i, total, f"{name} (失败)")
                continue
        
        # 输出处理结果统计
        print(f"\n{'='*50}")
        print(f"处理完成！")
        print(f"成功: {success_count}/{len(self.data)}")
        print(f"失败: {error_count}/{len(self.data)}")
        print(f"输出文件夹:")
        print(f"  - 文档: {self.output_docs_dir}")
        print(f"  - 图片: {self.output_images_dir}")
        print(f"{'='*50}")


def main():
    """主函数"""
    # 获取脚本所在目录
    script_dir = Path(__file__).parent.parent
    
    # 文件路径配置
    excel_file = script_dir / "中国赣州家具博览会邀请表.xlsx"
    pptx_template = script_dir / "电子邀请函模板.pptx"
    docx_template = script_dir / "（人民政府）邀请函模板.docx"
    output_docs = script_dir / "output_documents"
    output_images = script_dir / "output_images"
    
    # 验证文件存在
    for file_path in [excel_file, pptx_template, docx_template]:
        if not file_path.exists():
            print(f"错误: 文件不存在 - {file_path}")
            return
    
    # 创建处理器并执行
    processor = BatchProcessor(
        str(excel_file),
        str(pptx_template),
        str(docx_template),
        str(output_docs),
        str(output_images)
    )
    
    # 执行处理（enable_png=True以生成PNG）
    processor.process(generate_png=True)


if __name__ == "__main__":
    main()
