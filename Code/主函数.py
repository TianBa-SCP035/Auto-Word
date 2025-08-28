# main.py
from pathlib import Path
from export_sql import export_sql_to_excel
from fill_word import fill_word_template

def main():
    # 1. 让用户输入项目编号
    project_code = input("请输入项目编号（如 25P1156）：").strip()
    if not project_code:
        print("❌ 未输入项目编号，程序退出。")
        return

    # 2. 生成临时 Excel 文件路径
    excel_path = Path(f"{project_code}_明细.xlsx")

    # 3. 执行 SQL → 写入 Excel（竖向）
    export_sql_to_excel(project_code, excel_path)

    # 4. 模板路径（你可以改成自己的路径）
    template_path = Path(__file__).parent / "Mode1.docx"  # 模板文件
    if not template_path.exists():
        print(f"❌ 模板文件不存在：{template_path}")
        return

    # 5. 生成最终 Word 文件路径
    word_output_path = Path(f"{project_code}_项目方案.docx")

    # 6. Excel → Word 模板替换
    fill_word_template(excel_path, template_path, word_output_path)

    print(f"🎉 成功生成项目方案：{word_output_path}")
    
if __name__ == "__main__":
    main()
