import time,os,win32com.client,socket,pythoncom
from flask import Flask, render_template, request, redirect, url_for, session
from waitress import serve
from threading import Lock

def print_bartender_label(label_path, printer_name, copies):
    # 创建 BarTender 应用程序对象

    bt_app = win32com.client.Dispatch("BarTender.Application")

    # 使 BarTender 可见（可选）
    bt_app.Visible = False

    # 打开标签文档
    label_format = bt_app.Formats.Open(label_path, False, "")

    # 设置打印机

    # label_format.PrintSetup.PrinterName = printer_name
    label_format.Printer = printer_name

    # 设置打印的份数
    label_format.PrintSetup.IdenticalCopiesOfLabel = copies

    # 打印标签
    label_format.PrintOut(False, False)

    # 关闭标签文档
    label_format.Close(False)

    # 退出 BarTender 应用程序

def bartenderprinter(list,big):
    if big == True:
        big_a = r'大字标签\\'
    else:
        big_a = ''

    for i in list:
        if i[1]:
            current_directory = os.path.dirname(os.path.abspath(__file__))

            # print(current_directory)
            label_file_path = current_directory+f"\\标签文件\\{big_a}{i[0]}.btw"

            # 标签文件的路径
            # D:\py\text\标签打印自动化（整合）\web端\标签文件
            # label_file_path = "E:\\Users\\Documents\\产品标签\\脚本用\\包心丸.btw"

            # 打印机名称
            printer_name = 'Deli DL-888B(NEW)'



            # 打印份数
            number_of_copies = i[1]

            # 打印标签
            print_bartender_label(label_file_path, printer_name, number_of_copies)

app = Flask(__name__)
app.secret_key = 'your_secret_key'  # 设置一个密钥用于加密session数据

# Image files and initial setup
image_files = ['特制包心丸', '包心丸', '拳头母', '肉燕', '脆丸', '珍珠丸',
               '发菜丸', '排骨贡丸', '鲨鱼羹', '鲨鱼丸', '香肠', '墨鱼丸']
img_address = "static/img/"
entries = {name: 0 for name in image_files}

@app.before_request
def before_request():
    # Initialize COM for each request to avoid threading issues
    pythoncom.CoInitialize()
    if 'switch_is_on' not in session:
        session['switch_is_on'] = False

@app.route('/', methods=['GET', 'POST'])
def index():
    global img_address

    if request.method == 'POST':
        if 'switch' in request.form:
            # Toggle the switch state in the session
            session['switch_is_on'] = not session['switch_is_on']
            img_address = "/static/img/big/" if session['switch_is_on'] else "/static/img/"
        elif 'print' in request.form:
            print_time = int(request.form.get('print_time', 0))
            time.sleep(print_time)  # Consider moving to a background task
            results = [(name, int(request.form.get(name, 0))) for name in image_files]
            bartenderprinter(results, session['switch_is_on'])
        elif 'increment' in request.form:
            increment_item = request.form['increment']
            if increment_item in entries:
                entries[increment_item] += 1
        elif 'print_date' in request.form:
            bartenderprinter([['日期',1]], session['switch_is_on'])

    return render_template('index.html', image_files=image_files, entries=entries, img_address=img_address)

@app.route('/print', methods=['POST'])
def print_label():
    return "Label printed successfully"


if __name__ == '__main__':
    print('脚本已启动..')
    res = socket.gethostbyname(socket.gethostname())
    print('http://'+res+':5000')
    serve(app, host='0.0.0.0', port=5000, threads=8)

