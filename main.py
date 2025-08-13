import os
import docx
from win32com.client import Dispatch
import ctypes  # 用于检查管理员权限（处理某些受保护文件夹）


def is_admin():
    """检查程序是否以管理员权限运行"""
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False


def read_doc_file(file_path):
    """读取.doc文件内容"""
    try:
        word = Dispatch("Word.Application")
        word.Visible = False
        # 处理包含特殊字符的路径
        doc = word.Documents.Open(FileName=file_path, ConfirmConversions=False)
        content = doc.Content.Text
        doc.Close(SaveChanges=0)  # 不保存关闭
        word.Quit()
        return content
    except Exception as e:
        print(f"⚠️ 读取.doc文件出错 {file_path}: {str(e)}")
        return ""


def read_docx_file(file_path):
    """读取.docx文件内容"""
    try:
        doc = docx.Document(file_path)
        full_text = []
        for para in doc.paragraphs:
            full_text.append(para.text)
        return ' '.join(full_text)
    except Exception as e:
        print(f"⚠️ 读取.docx文件出错 {file_path}: {str(e)}")
        return ""


def validate_folder_path(folder_path):
    """验证文件夹路径是否有效"""
    if not folder_path or not os.path.exists(folder_path):
        return False, "文件夹路径不存在"
    if not os.path.isdir(folder_path):
        return False, "指定路径不是一个文件夹"
    # 检查是否有访问权限
    try:
        test_file = os.path.join(folder_path, "test_access.tmp")
        with open(test_file, 'w') as f:
            f.write("test")
        os.remove(test_file)
        return True, "路径有效"
    except Exception as e:
        return False, f"没有访问权限: {str(e)}"


def search_keyword_in_word_files(folder_path, keyword):
    """搜索文件夹中包含关键词的Word文档"""
    # 验证关键词
    if not keyword or len(keyword.strip()) < 3:
        print("❌ 错误：关键词不能为空且长度不能少于3个字符")
        return []

    keyword = keyword.lower()
    results = []
    file_count = 0
    processed_count = 0

    # 统计总文件数（用于进度提示）
    for root, dirs, files in os.walk(folder_path):
        for file in files:
            if file.endswith(('.doc', '.docx')) and not file.startswith('~$'):  # 排除临时文件
                file_count += 1

    print(f"📂 发现 {file_count} 个Word文档，正在搜索中...")

    # 遍历文件夹
    for root, dirs, files in os.walk(folder_path):
        for file in files:
            # 排除Word临时文件（以~$开头）
            if file.startswith('~$'):
                continue

            # 检查文件是否为Word文档
            if file.endswith('.doc') and not file.endswith('.docx'):
                file_path = os.path.join(root, file)
                content = read_doc_file(file_path)
                processed_count += 1
                if keyword in content.lower():
                    results.append((file, file_path))
            elif file.endswith('.docx'):
                file_path = os.path.join(root, file)
                content = read_docx_file(file_path)
                processed_count += 1
                if keyword in content.lower():
                    results.append((file, file_path))

            # 显示进度
            if processed_count % 5 == 0:
                print(f"🔍 已处理 {processed_count}/{file_count} 个文件", end='\r')

    print(f"✅ 搜索完成，共处理 {processed_count} 个文件")
    return results


def main():
    # 显示程序信息
    print("=" * 50)
    print("      Word文档关键词搜索工具      ")
    print("=" * 50)

    # 检查是否以管理员权限运行（提示）
    if not is_admin():
        print("⚠️ 提示：程序未以管理员权限运行，可能无法访问某些系统文件夹")
        print()

    # 获取用户输入的文件夹路径
    while True:
        folder_path = input("请输入要搜索的文件夹路径: ").strip()
        # 处理PowerShell中可能的引号
        folder_path = folder_path.strip('"\'')
        is_valid, message = validate_folder_path(folder_path)
        if is_valid:
            print(f"📌 已确认文件夹路径: {folder_path}")
            break
        else:
            print(f"❌ {message}，请重新输入")

    # 获取用户输入的关键词
    while True:
        keyword = input("请输入要搜索的关键词（至少3个字符）: ").strip()
        if keyword and len(keyword) >= 3:
            break
        print("❌ 关键词不符合要求，请重新输入")

    # 执行搜索
    matches = search_keyword_in_word_files(folder_path, keyword)

    # 显示结果
    print("\n" + "=" * 50)
    if matches:
        print(f"🎉 找到 {len(matches)} 个包含关键词 '{keyword}' 的文件：")
        print("-" * 50)
        for i, (file_name, file_path) in enumerate(matches, 1):
            print(f"{i}. 文件名：{file_name}")
            print(f"   路径：{file_path}")
            print("-" * 50)
    else:
        print(f"🔍 未找到包含关键词 '{keyword}' 的Word文档")
    print("=" * 50)


if __name__ == "__main__":
    main()
    # 保持PowerShell窗口不关闭
    input("按任意键退出...")
