"""
文档转换工具模块
提供Markdown转Word等功能
"""

import os
import re
import subprocess
from typing import Optional


def convert_to_docx_basic(input_file: str, docx_output: Optional[str] = None) -> str:
    """
    将Markdown文件转换为Word文档（基本版本）
    
    Args:
        input_file (str): 输入的Markdown文件路径
        docx_output (str, optional): 输出的Word文件路径，默认使用与输入文件相同的名称但扩展名为.docx
        
    Returns:
        str: 输出文件的路径，如果转换失败则返回None
    """
    if not os.path.exists(input_file):
        print(f"[错误] 输入文件不存在: {input_file}")
        return None

    if docx_output is None:
        docx_output = os.path.splitext(input_file)[0] + '.docx'
    
    try:
        print(f"正在读取文件: {input_file}")
        with open(input_file, 'r', encoding='utf-8') as f:
            content = f.read()

        # 预处理：将表格分隔线中的全角破折号'—'替换为半角'-'，确保pandoc能识别。
        # 之前尝试通过正则表达式确保表格前有空行，但该方法会错误地在表格标题和分隔线之间插入换行符，
        # 从而破坏了表格的结构。经过分析，主要问题在于非标准的'—'字符，
        # 仅进行字符替换即可解决问题。
        processed_content = content.replace('—', '-')

        pandoc_cmd = [
            "pandoc",
            "-f", "markdown",
            "-o", docx_output,
            "--standalone",
            "--resource-path=.",
        ]
        
        print(f"正在调用Pandoc进行转换，输出至: {docx_output}")
        subprocess.run(
            pandoc_cmd, 
            input=processed_content.encode('utf-8'), 
            check=True, 
            capture_output=True
        )
        
        print(f"\n📄 Word版报告已生成: {docx_output}")
        return docx_output

    except FileNotFoundError:
        print("[错误] 未找到 'pandoc' 命令。请确保 Pandoc 已经安装并配置在系统的 PATH 环境变量中。")
    except subprocess.CalledProcessError as e:
        error_message = e.stderr.decode('gbk', errors='replace') if e.stderr else "无详细错误信息"
        print(f"[提示] pandoc转换失败。错误信息: {error_message}")
    except Exception as e:
        print(f"[提示] 转换过程中发生未知错误: {e}")
    
    return None


def convert_to_docx_with_indent(input_file: str, docx_output: Optional[str] = None) -> str:
    """
    将Markdown文件转换为Word文档（带自定义样式的高级版本）。
    使用一个预先配置好的 `reference.docx` 文件来应用所有样式，包括字体、缩进、表格边框等。

    Args:
        input_file (str): 输入的Markdown文件路径
        docx_output (str, optional): 输出的Word文件路径，默认与输入文件同名

    Returns:
        str: 输出文件的路径，如果转换失败则返回None
    """
    if not os.path.exists(input_file):
        print(f"[错误] 输入文件不存在: {input_file}")
        return None

    if docx_output is None:
        docx_output = os.path.splitext(input_file)[0] + '.docx'

    # 定义参考文档的路径。这个文件现在应该是预先配置好的。
    utils_dir = os.path.dirname(os.path.abspath(__file__))
    reference_docx = os.path.join(utils_dir, "reference.docx")

    if not os.path.exists(reference_docx):
        print(f"❌ [错误] 参考文档 'reference.docx' 不存在于: {utils_dir}")
        print("➡️ [操作建议] 请先在终端运行 `python create_reference.py` 脚本来生成默认的参考文档。")
        print("             然后您可以在Word中打开并编辑它，以定义您自己的样式。")
        print("---------------------------------------------------------------------")
        print("⚠️ [自动回退] 由于缺少样式文件，将使用无格式的基础模式进行转换。")
        return convert_to_docx_basic(input_file, docx_output)

    # --- 使用配置好的参考文档进行转换 ---
    try:
        print(f"正在读取文件: {input_file}")
        with open(input_file, 'r', encoding='utf-8') as f:
            content = f.read()

        # 预处理：将表格分隔线中的全角破折号'—'替换为半角'-'，确保pandoc能识别。
        processed_content = content.replace('—', '-')

        pandoc_cmd = [
            "pandoc",
            "-f", "markdown",
            "-o", docx_output,
            "--standalone",
            "--resource-path=.",  # 指定图片等资源的搜索路径为当前目录
            "--reference-doc", reference_docx
        ]
        
        print(f'正在使用自定义参考文档进行转换: {reference_docx}')
        print(f"正在调用Pandoc进行转换，输出至: {docx_output}")
        
        subprocess.run(
            pandoc_cmd,
            input=processed_content.encode('utf-8'),
            check=True,
            capture_output=True
        )

        print(f"\n📄 Word版报告已生成: {docx_output}")
        return docx_output

    except FileNotFoundError:
        print("[错误] 未找到 'pandoc' 命令。请确保 Pandoc 已经安装并配置在系统的 PATH 环境变量中。")
    except subprocess.CalledProcessError as e:
        error_message = e.stderr.decode('gbk', errors='replace') if e.stderr else "无详细错误信息"
        print(f"[提示] pandoc转换失败。错误信息: {error_message}")
    except Exception as e:
        print(f"[提示] 转换过程中发生未知错误: {e}")

    return None


if __name__ == "__main__":
    import os
    import sys

    # 构建到Markdown文件的绝对路径，确保无论从哪里运行脚本都能找到文件
    # 1. 获取当前脚本(doc_converter.py)所在的目录
    script_directory = os.path.dirname(os.path.abspath(__file__))
    # 2. 从脚本目录向上回溯三级，到达项目根目录
    project_root = os.path.abspath(os.path.join(script_directory, '..', '..', '..'))
    # 3. 拼接得到目标Markdown文件的完整路径 (现在位于reports目录下)
    md_file = os.path.join(project_root, 'reports', 'Industry_Research_Report.md')

    print(f"目标文件路径: {md_file}")
    convert_to_docx_with_indent(md_file)