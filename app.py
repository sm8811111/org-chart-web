import os
from flask import Flask, render_template, request, send_file
from graphviz import Digraph
from werkzeug.utils import secure_filename

# 初始化 Flask
app = Flask(__name__)

# 設定上傳資料夾
UPLOAD_FOLDER = "uploads"
OUTPUT_FOLDER = "outputs"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['OUTPUT_FOLDER'] = OUTPUT_FOLDER

# 允許上傳的檔案類型
ALLOWED_EXTENSIONS = {'txt', 'csv'}

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

# 使用 pure-Python 生成 org chart 的函數
def generate_chart(upload_path, output_filename="chart.png"):
    """
    使用 Graphviz Digraph 生成 PNG 圖片，直接用 pipe，不依賴系統 dot
    """
    dot = Digraph(format='png')
    
    # 這裡示範用簡單例子，如果你有讀取 CSV 或上傳檔案，可在這裡處理
    dot.node('A', 'Start')
    dot.node('B', 'Process')
    dot.node('C', 'End')
    dot.edge('A', 'B')
    dot.edge('B', 'C')

    # pipe() 生成 PNG bytes
    image_bytes = dot.pipe()
    output_path = os.path.join(app.config['OUTPUT_FOLDER'], output_filename)
    with open(output_path, 'wb') as f:
        f.write(image_bytes)
    
    return output_path

# 首頁
@app.route("/", methods=['GET', 'POST'])
def index():
    chart_image = None
    if request.method == 'POST':
        # 檢查是否有檔案上傳
        if 'file' not in request.files:
            return "No file part"
        file = request.files['file']
        if file.filename == '':
            return "No selected file"
        if file and allowed_file(file.filename):
            filename = secure_filename(file.filename)
            upload_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
            file.save(upload_path)

            # 生成圖表
            chart_image = generate_chart(upload_path)

    return render_template("index.html", chart_image=chart_image)

# 提供 PNG 圖檔下載
@app.route("/outputs/<filename>")
def outputs(filename):
    return send_file(os.path.join(app.config['OUTPUT_FOLDER'], filename), mimetype='image/png')

# Render 專用
if __name__ == "__main__":
    port = int(os.environ.get("PORT", 10000))
    app.run(host="0.0.0.0", port=port)