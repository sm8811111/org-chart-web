import os
from flask import Flask, render_template, request
import pandas as pd
import networkx as nx
import matplotlib.pyplot as plt
from werkzeug.utils import secure_filename
from datetime import datetime

# 初始化 Flask
app = Flask(__name__)

UPLOAD_FOLDER = "uploads"
OUTPUT_FOLDER = "outputs"

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

ALLOWED_EXTENSIONS = {'xls', 'xlsx', 'csv'}

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

# 使用 networkx 生成 PNG 組織圖
def generate_chart(file_path):
    # 讀 Excel
    if file_path.endswith('.csv'):
        df = pd.read_csv(file_path)
    else:
        df = pd.read_excel(file_path)

    # 清理欄位名稱
    df.columns = [c.strip().replace("　", "") for c in df.columns]

    if '編號' not in df.columns or '推薦人' not in df.columns or '姓名' not in df.columns:
        raise ValueError("Excel 必須包含欄位: 編號, 推薦人, 姓名")

    # 建立有向圖
    G = nx.DiGraph()

    for _, row in df.iterrows():
        node = str(row['編號'])
        name = str(row['姓名'])
        G.add_node(node, label=name)
        parent = str(row['推薦人'])
        if parent != '0' and parent in df['編號'].astype(str).values:
            G.add_edge(parent, node)

    # 畫圖
    plt.figure(figsize=(12, 8))
    pos = nx.nx_agraph.graphviz_layout(G, prog='dot')  # 需要 pygraphviz 或 pydotplus
    labels = nx.get_node_attributes(G, 'label')
    nx.draw(G, pos, labels=labels, with_labels=True,
            node_color='lightblue', node_size=2500, arrows=True,
            font_size=10, font_weight='bold')
    
    timestamp = datetime.now().strftime("%Y%m%d-%H%M%S")
    filename = f"org_{timestamp}.png"
    output_path = os.path.join(OUTPUT_FOLDER, filename)
    plt.savefig(output_path, bbox_inches='tight')
    plt.close()
    return filename

# 首頁
@app.route("/", methods=["GET", "POST"])
def index():
    chart_image = None
    error_message = None
    if request.method == "POST":
        file = request.files.get("file")
        if not file or file.filename == '':
            error_message = "請選擇檔案"
        elif not allowed_file(file.filename):
            error_message = "只允許 xls, xlsx, csv 檔案"
        else:
            try:
                filename = secure_filename(file.filename)
                upload_path = os.path.join(UPLOAD_FOLDER, filename)
                file.save(upload_path)
                chart_filename = generate_chart(upload_path)
                chart_image = "/outputs/" + chart_filename
            except Exception as e:
                error_message = f"生成失敗: {e}"

    return render_template("index.html", chart_image=chart_image, error_message=error_message)

# 提供 PNG 圖檔
@app.route("/outputs/<filename>")
def outputs(filename):
    return app.send_static_file(os.path.join(OUTPUT_FOLDER, filename))

# 啟動
if __name__ == "__main__":
    port = int(os.environ.get("PORT", 10000))
    app.run(host="0.0.0.0", port=port, debug=True)