import pandas as pd
from flask import Flask, request, send_file, jsonify
from flask_cors import CORS
import io
import os

# 初始化 Flask 应用
app = Flask(__name__)
CORS(app)  # 允许跨域请求
app.config['JSON_AS_ASCII'] = False

# 解决 pandas 读取 Excel 的引擎警告，指定默认引擎
pd.set_option('mode.chained_assignment', None)

def normalize_code(code):
    """标准化编码处理，去除空格和特殊字符"""
    if pd.isna(code):
        return ""
    # 转为字符串，去除空格和换行符
    return str(code).strip()

@app.route('/upload', methods=['POST'])
def upload_files():
    """
    上传两个文件：编号列表文件和Excel文件
    返回处理后的Excel文件
    """
    try:
        # 检查文件是否上传
        if 'codes_file' not in request.files or 'excel_file' not in request.files:
            return jsonify({'error': '请上传两个文件'}), 400

        codes_file = request.files['codes_file']
        excel_file = request.files['excel_file']

        if codes_file.filename == '' or excel_file.filename == '':
            return jsonify({'error': '请选择文件'}), 400

        # 读取编号列表文件
        codes_content = codes_file.read().decode('utf-8')
        # 处理多种分隔符：逗号、换行、空格
        codes_list = []
        for line in codes_content.splitlines():
            # 按逗号分割，然后按空格分割
            parts = line.split(',')
            for part in parts:
                codes_list.extend(part.split())

        # 去除空字符串和重复，标准化编码
        normalized_codes = set()
        for code in codes_list:
            normalized = normalize_code(code)
            if normalized:
                normalized_codes.add(normalized)

        print(f"需要筛选的编号数量: {len(normalized_codes)}")
        print(f"编号列表: {sorted(normalized_codes)[:10]}...")  # 打印前10个编号

        # 读取Excel文件
        excel_data = excel_file.read()

        # 使用pandas读取Excel
        xls = pd.ExcelFile(io.BytesIO(excel_data))

        # 创建一个新的Excel writer
        output = io.BytesIO()

        # 用于存储Excel中所有的卷烟编码
        excel_codes_found = set()

        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            print("所有工作表名称:", xls.sheet_names)
            # 处理每个工作表
            for sheet_name in xls.sheet_names:
                print(f"处理工作表: '{sheet_name}'")  # 加上引号看具体内容

                # 使用 strip() 去除首尾空格
                cleaned_sheet_name = sheet_name.strip()

                # 方法1：使用模糊匹配
                if '雪茄' in cleaned_sheet_name:
                    print("匹配到雪茄表")
                    # 对于雪茄表，直接复制，不做筛选
                    df = pd.read_excel(xls, sheet_name=sheet_name, header=None)
                    df.to_excel(writer, sheet_name=sheet_name, index=False, header=False)

                elif '价位段' in cleaned_sheet_name:
                    print("匹配到价位段表")
                    # 对于按价位段表，直接复制，不做筛选
                    df = pd.read_excel(xls, sheet_name=sheet_name, header=None)
                    df.to_excel(writer, sheet_name=sheet_name, index=False, header=False)

                elif '卷烟定量' in cleaned_sheet_name:
                    print("匹配到卷烟表 - 开始处理")
                    # 对于卷烟表，进行筛选
                    # 读取数据，注意表头有3行
                    # 我们需要保留原始格式，所以读取时不设置表头
                    df = pd.read_excel(xls, sheet_name=sheet_name, header=None)

                    # 打印前几行数据查看结构
                    print("数据前5行预览:")
                    for i in range(min(5, len(df))):
                        print(f"行{i}: {df.iloc[i].tolist()[:5]}...")

                    # 查找卷烟编码所在的列
                    # 从第0行开始查找，因为不确定具体在哪一行
                    code_col_index = None
                    for i in range(len(df)):
                        row = df.iloc[i]
                        for j in range(len(row)):
                            cell_value = str(row[j])
                            if '卷烟编码' in cell_value:
                                code_col_index = j
                                print(f"在第{i + 1}行，第{j + 1}列找到'卷烟编码'")
                                break
                        if code_col_index is not None:
                            break

                    if code_col_index is not None:
                        print(f"卷烟编码列索引: {code_col_index}")

                        # 查找数据开始的行（通常是'卷烟编码'所在行之后）
                        data_start_row = i + 1  # '卷烟编码'行的下一行开始是数据

                        # 分离表头和数据
                        header_rows = df.iloc[:data_start_row]  # 表头包括'卷烟编码'这一行
                        data_rows = df.iloc[data_start_row:]  # '卷烟编码'之后的行是数据

                        # 收集Excel中所有的卷烟编码（用于后续对比）
                        for idx, row in data_rows.iterrows():
                            code_value = normalize_code(row[code_col_index])
                            if code_value:
                                excel_codes_found.add(code_value)

                        # 筛选数据行
                        filtered_data = []
                        for idx, row in data_rows.iterrows():
                            code_value = normalize_code(row[code_col_index])
                            # 打印检查每个编码
                            print(f"检查编码: '{code_value}'")
                            if code_value in normalized_codes:
                                filtered_data.append(row)
                                print(f"保留编码: {code_value}")
                            else:
                                # 打印被删除的编码（用于调试）
                                if code_value and code_value != 'nan':
                                    print(f"删除编码: {code_value}")

                        # 重新组合数据
                        if filtered_data:
                            filtered_df = pd.concat([header_rows, pd.DataFrame(filtered_data)], ignore_index=True)
                        else:
                            # 如果没有匹配的数据，只保留表头
                            filtered_df = header_rows

                        print(f"原始数据行数: {len(df)}")
                        print(f"表头行数: {len(header_rows)}")
                        print(f"原始数据行数（不含表头）: {len(data_rows)}")
                        print(f"筛选后数据行数（不含表头）: {len(filtered_data)}")

                        # 写入筛选后的数据
                        filtered_df.to_excel(writer, sheet_name=sheet_name, index=False, header=False)
                    else:
                        print("未找到'卷烟编码'列，跳过筛选")
                        # 打印列名看看
                        if len(df) > 0:
                            print("第一行内容:", df.iloc[0].tolist())
                        df.to_excel(writer, sheet_name=sheet_name, index=False, header=False)

                else:
                    print(f"其他工作表: {sheet_name} - 直接复制")
                    # 其他工作表直接复制
                    df = pd.read_excel(xls, sheet_name=sheet_name, header=None)
                    df.to_excel(writer, sheet_name=sheet_name, index=False, header=False)

        # ===== 新增：打印未匹配的编码 =====
        # 找出筛选列表中有但Excel中没有的编码
        unmatched_codes = normalized_codes - excel_codes_found
        print("\n========== 未匹配的编码 ==========")
        if unmatched_codes:
            print(f"总共有 {len(unmatched_codes)} 个编码未在Excel中找到匹配:")
            # 按字母顺序排序打印，方便查看
            for code in sorted(unmatched_codes):
                print(f"- {code}")
        else:
            print("所有筛选编码都在Excel中找到匹配！")
        print("==================================\n")

        # 准备返回文件
        output.seek(0)

        return send_file(
            output,
            as_attachment=True,
            download_name='filtered_tobacco_data.xlsx',
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )

    except Exception as e:
        import traceback
        print(f"处理出错: {str(e)}")
        traceback.print_exc()
        return jsonify({'error': f'处理失败: {str(e)}'}), 500

@app.route('/test', methods=['GET'])
def test():
    """测试接口"""
    return jsonify({'status': '服务器运行正常'})

# ========== Vercel 适配关键修改 ==========
# Vercel 的 Serverless 环境需要暴露 app 实例，不能只靠 if __name__ == '__main__' 启动
# 1. 移除原有启动逻辑的限制，确保 Vercel 能识别 app 实例
port = int(os.environ.get('PORT', 5000))

# 2. Vercel 会自动调用 app 实例，无需手动启动（保留启动逻辑仅用于本地测试）
if __name__ == '__main__':
    app.run(debug=False, host='0.0.0.0', port=port)

# 3. 关键：Vercel 需要明确导出 app 实例
application = app  # 兼容 Gunicorn 等 WSGI 服务器的命名规范
