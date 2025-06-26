import os
import subprocess
import sys

def create_reference_docx():
    """
    调用 Pandoc 来生成一份默认的 reference.docx 文件。
    这个文件之后可以被用户手动编辑，用于定义文档转换的自定义样式。
    """
    print("正在尝试创建 Pandoc 默认的 reference.docx 文件...")

    # 将 reference.docx 定义在 'src/tools/document_processing/' 目录下
    script_dir = os.path.dirname(os.path.abspath(__file__))
    output_path = os.path.join(script_dir, "src", "tools", "document_processing", "reference.docx")
    
    # 确保目标目录存在
    os.makedirs(os.path.dirname(output_path), exist_ok=True)

    pandoc_cmd = [
        "pandoc",
        "--print-default-data-file",
        "reference.docx"
    ]

    try:
        # Pandoc 会将二进制数据输出到 stdout，所以我们需要捕获它
        result = subprocess.run(
            pandoc_cmd,
            check=True,
            capture_output=True
        )
        
        # 将捕获到的二进制数据写入我们的目标文件
        with open(output_path, "wb") as f:
            f.write(result.stdout)
            
        print(f"\n✅ 已成功创建参考文档: {output_path}")
        print("\n--- 操作指南 ---")
        print("1. 请用 Microsoft Word 打开刚刚生成的 reference.docx 文件。")
        print("2. 在Word中修改您需要的样式，例如：")
        print("   - 表格样式: 在 '开始' 选项卡的 '样式' 区域找到 'Table' 样式，右键点击 -> '修改' -> '格式' -> '边框和底纹'，在这里为表格添加所有框线。")
        print('  - 正文样式: 修改 "Normal" 或 "正文" 样式的字体（如"等线"）。')
        print("   - 标题样式: 修改 'Heading 1', 'Heading 2' 等样式的字体和大小。")
        print("3. 修改完成后，直接保存并关闭 reference.docx 文件。")
        print("4. 从此以后，运行报告生成脚本时，就会自动应用您设置好的所有精美格式。")

    except FileNotFoundError:
        print("\n❌ [错误] 未找到 'pandoc' 命令。")
        print("请确保 Pandoc 已经安装并加入到您系统的 PATH 环境变量中。")
        sys.exit(1)
    except subprocess.CalledProcessError as e:
        error_message = e.stderr.decode('utf-8', errors='replace') if e.stderr else "无 stderr 信息。"
        print(f"\n❌ [错误] Pandoc 生成参考文件失败。错误信息: {error_message}")
        sys.exit(1)
    except Exception as e:
        print(f"\n❌ [错误] 发生未知错误: {e}")
        sys.exit(1)

if __name__ == "__main__":
    create_reference_docx() 