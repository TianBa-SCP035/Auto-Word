from flask import Flask, request, jsonify, send_file
from flask_cors import CORS
from pathlib import Path
import tempfile
import os
from datetime import datetime
import uuid
from export_sql import export_sql_to_excel
from fill_word import fill_word_template

app = Flask(__name__)
CORS(app)  # 启用CORS支持，允许所有来源的跨域请求

# 配置上传和输出目录
UPLOAD_FOLDER = Path('temp_files')
OUTPUT_FOLDER = Path('output_files')
UPLOAD_FOLDER.mkdir(exist_ok=True)
OUTPUT_FOLDER.mkdir(exist_ok=True)

@app.route('/api/generate-project-plan', methods=['POST'])
def generate_project_plan():
    """
    生成项目方案API
    接收JSON参数：
    {
        "project_code": "25P1156",  # 项目编号
        "template_path": "成瘤-项目方案模板 - 程序.docx"  # 可选，模板文件路径
    }
    返回生成的Word文档
    """
    try:
        # 获取请求数据
        data = request.get_json()
        if not data:
            return jsonify({'error': '请提供JSON数据'}), 400
        
        project_code = data.get('project_code', '').strip()
        if not project_code:
            return jsonify({'error': '请提供项目编号'}), 400
        
        # 获取模板路径（可选）
        template_path = data.get('template_path', str(Path(__file__).parent / 'Mode1.docx'))
        
        # 检查模板文件是否存在
        template_file = Path(template_path)
        if not template_file.exists():
            return jsonify({'error': f'模板文件不存在：{template_path}'}), 400
        
        # 生成文件名（与原代码保持一致）
        excel_filename = f"{project_code}_明细.xlsx"
        word_filename = f"{project_code}_项目方案.docx"
        
        # 创建临时文件路径
        excel_path = UPLOAD_FOLDER / excel_filename
        word_output_path = OUTPUT_FOLDER / word_filename
        
        print(f"开始处理项目编号：{project_code}")
        print(f"Excel临时文件路径：{excel_path}")
        print(f"Word输出文件路径：{word_output_path}")
        
        # 1. 执行 SQL → 写入 Excel（竖向）
        export_sql_to_excel(project_code, excel_path)
        
        # 2. Excel → Word 模板替换
        fill_word_template(excel_path, template_file, word_output_path)
        
        # 3. 删除临时Excel文件
        if excel_path.exists():
            excel_path.unlink()
            print(f"已删除临时文件：{excel_path}")
        
        # 4. 返回生成的Word文件
        return send_file(
            word_output_path,
            as_attachment=True,
            download_name=word_filename,
            mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document'
        )
        
    except Exception as e:
        print(f"生成项目方案失败：{str(e)}")
        return jsonify({'error': f'生成项目方案失败：{str(e)}'}), 500

@app.route('/api/health', methods=['GET'])
def health_check():
    """健康检查接口"""
    return jsonify({'status': 'ok', 'message': 'API服务正常运行'})

@app.route('/api/test-connection', methods=['POST'])
def test_connection():
    """
    测试数据库连接
    接收JSON参数：
    {
        "project_code": "25P1156"  # 测试用的项目编号
    }
    """
    try:
        data = request.get_json()
        project_code = data.get('project_code', '').strip()
        
        if not project_code:
            return jsonify({'error': '请提供项目编号'}), 400
        
        # 创建临时文件路径测试导出功能
        with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as tmp:
            tmp_path = tmp.name
        
        try:
            # 测试SQL导出
            result_path = export_sql_to_excel(project_code, tmp_path)
            
            # 检查文件是否生成
            if Path(result_path).exists():
                file_size = Path(result_path).stat().st_size
                return jsonify({
                    'status': 'success',
                    'message': '数据库连接正常，数据导出成功',
                    'file_size': file_size,
                    'project_code': project_code
                })
            else:
                return jsonify({'error': '文件生成失败'}), 500
                
        finally:
            # 清理临时文件
            if Path(tmp_path).exists():
                Path(tmp_path).unlink()
                
    except Exception as e:
        return jsonify({'error': f'数据库连接测试失败：{str(e)}'}), 500

if __name__ == '__main__':
    print("启动项目方案生成API服务...")
    print("API端点：")
    print("  POST /api/generate-project-plan - 生成项目方案")
    print("  GET  /api/health - 健康检查")
    print("  POST /api/test-connection - 测试数据库连接")
    print("\n服务启动在 http://localhost:5000")
    app.run(debug=False, host='0.0.0.0', port=5000)