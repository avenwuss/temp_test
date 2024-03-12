from flask import Flask, request, jsonify
from flask_httpauth import HTTPBasicAuth
import pyodbc

# 创建 Flask 应用
app = Flask(__name__)
auth = HTTPBasicAuth()

# 数据库连接配置
db_config = {
    'Driver': 'SQL Server',
    'Server': '10.164.113.132',
    'Database': 'ZZZ_TMP',
    'UID': 'sa',
    'PWD': 'Sa@1qaz2wsx3edc'
}

# 验证函数，用于检查用户名和密码
@auth.verify_password
def verify_password(username, password):
    # 在此处进行用户名和密码的验证，可以从数据库或其他存储位置进行验证
    valid_username = "wu"
    valid_password = "123"
    return username == valid_username and password == valid_password

# 处理 GET 请求的路由
@app.route('/getCT', methods=['GET'])
@auth.login_required
def get_data():
    try:
        # 获取参数
        param = request.args.get('contractNO')

        # 建立数据库连接
        conn = pyodbc.connect(**db_config)

        # 创建游标对象
        cursor = conn.cursor()

        # 调用存储过程
        cursor.execute("{CALL p_contract_query(?)}", (param,))

        # 获取存储过程的结果集
        result = cursor.fetchall()

        # 关闭游标和数据库连接
        cursor.close()
        conn.close()

        # 将查询结果转换为 JSON 格式并返回
        data = [{'ContractNO': row[0], 'CT_NO': row[1], 'CT_NAME': row[2]} for row in result]
        return jsonify(data)

    except Exception as e:
        return jsonify({'error': str(e)}), 500


# 处理 GET 请求的路由，用于调用存储过程 p_getJournalListByContract
@app.route('/get_CT_Journal', methods=['GET'])
@auth.login_required
def get_journal_list_by_contract():
    try:
        # 获取参数
        param = request.args.get('contractNO')

        # 建立数据库连接
        conn = pyodbc.connect(**db_config)

        # 创建游标对象
        cursor = conn.cursor()

        # 调用存储过程
        cursor.execute("{CALL p_getJournalListByContract(?)}", (param,))

        # 获取存储过程的结果集
        result = cursor.fetchall()

        # 关闭游标和数据库连接
        cursor.close()
        conn.close()

        # 将查询结果转换为 JSON 格式并返回
        data = [
            {
                'PERIOD': row.PERIOD,
                'JRNAL_NO': row.JRNAL_NO,
                'JRNAL_LINE': row.JRNAL_LINE,
                'TREFERENCE': row.TREFERENCE,
                'DESCRIPTN': row.DESCRIPTN,
                'ACCNT_CODE': row.ACCNT_CODE,
                'BaseAmount': row.BaseAmount,
                'TransAmount': row.TransAmount,
                'D_C': row.D_C,
                'Department': row.Department,
                'Products': row.Products,
                'Employee': row.Employee,
                'Projects': row.Projects,
                'ARAP_Item': row.ARAP_Item,
                'Cash_Flow': row.Cash_Flow,
                'Contract_Num': row.Contract_Num,
                'Transfer_Station': row.Transfer_Station,
                'BudgetDept': row.BudgetDept,
                'Other': row.Other
            }
            for row in result
        ]
        return jsonify(data)

    except Exception as e:
        return jsonify({'error': str(e)}), 500

# 运行应用
if __name__ == '__main__':
    app.run(host='0.0.0.0')
