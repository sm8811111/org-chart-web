import os
import itertools
from datetime import datetime
from flask import Flask, request, render_template, jsonify
import pandas as pd
from graphviz import Digraph

app = Flask(__name__)
UPLOAD_FOLDER = "static/output"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
MAX_HISTORY = 10

color_palette = [
    'lightblue', 'lightgreen', 'orange', 'pink', 'yellow',
    'lightcoral', 'lightgoldenrod', 'lightsalmon', 'lightseagreen'
]

def get_node_size(df, value, min_size=1.0, max_size=2.0):
    min_val = df['小組積分額'].min()
    max_val = df['小組積分額'].max()
    if max_val == min_val:
        return 1.5
    norm = (value - min_val) / (max_val - min_val)
    return round(min_size + norm * (max_size - min_size), 2)

def cleanup_old_files(folder, max_files=10):
    files = [os.path.join(folder, f) for f in os.listdir(folder) if f.endswith(".png")]
    files.sort(key=os.path.getmtime, reverse=True)
    for old_file in files[max_files:]:
        os.remove(old_file)

@app.route("/", methods=["GET", "POST"])
def index():

    if request.method == "GET":
        return render_template("index.html")

    try:
        file = request.files.get("file")
        if not file:
            return jsonify({"error": "未上傳檔案"}), 400

        df = pd.read_excel(file)
        df.columns = [c.strip().replace("　", "") for c in df.columns]

        required_cols = ['編號','姓名','層級','類型','推薦人','小組積分額','小組售貨額','%']
        for col in required_cols:
            if col not in df.columns:
                return jsonify({"error": f"缺少必要欄位: {col}"}), 400

        if '會員積分額' not in df.columns:
            df['會員積分額'] = 0

        df['層級'] = pd.to_numeric(df['層級'], errors='coerce').fillna(1)
        df['小組積分額'] = pd.to_numeric(df['小組積分額'], errors='coerce').fillna(0)
        df['小組售貨額'] = pd.to_numeric(df['小組售貨額'], errors='coerce').fillna(0)
        df['編號'] = df['編號'].astype(str)
        df['%'] = pd.to_numeric(df['%'], errors='coerce').fillna(0)
        df['會員積分額'] = pd.to_numeric(df['會員積分額'], errors='coerce').fillna(0)

        # 建立 Graphviz
        dot = Digraph(comment='Organization Chart')

        # ⭐ 全域字型設定 (解決中文亂碼)
        dot.attr(fontname='Microsoft JhengHei')
        dot.attr('node', fontname='Microsoft JhengHei')
        dot.attr('edge', fontname='Microsoft JhengHei')

        node_count = len(df)

        ranksep = "1.0"
        nodesep = "0.4"

        if node_count > 300:
            ranksep = "1.5"
            nodesep = "0.6"

        if node_count > 800:
            ranksep = "2.0"
            nodesep = "0.8"

        dot.attr(
            overlap='false',
            splines='true',
            layout='dot',
            rankdir='TB',
            ranksep=ranksep,
            nodesep=nodesep
        )

        last_nodes = {}
        node_color_map = {}

        level2_nodes = df[df['層級']==2]['編號'].tolist()
        color_cycle = itertools.cycle(color_palette)
        level2_colors = {node: next(color_cycle) for node in level2_nodes}

        for _, row in df.iterrows():

            level = int(row['層級'])
            num = row['編號']

            percent_td = ""
            if row['%'] != 0:
                percent_value = int(round(row['%'] * 100))
                percent_label = f"<B><FONT POINT-SIZE='16'>{percent_value}%</FONT></B>"
                percent_td = f"<TR><TD BGCOLOR='#D32F2F' ALIGN='CENTER'>{percent_label}</TD></TR>"

            bv_value = row['小組售貨額']
            if row['類型'] == '會員':
                bv_value += row['會員積分額']

            bv_td = f"<TR><TD ALIGN='CENTER'>BV:{bv_value}</TD></TR>" if bv_value != 0 else ""

            label = f"""
            <<TABLE BORDER='0' CELLBORDER='0' CELLSPACING='0'>
            {percent_td}
            <TR><TD ALIGN='CENTER'>{row['姓名']}</TD></TR>
            {bv_td}
            </TABLE>>
            """.strip()

            size = get_node_size(df, row['小組積分額'])

            if level == 1:
                color = 'lightgrey'

            elif level == 2:
                children = df[df['推薦人'] == num]
                color = 'lightgrey' if children.empty and row['類型'] == '會員' else level2_colors.get(num, 'lightblue')

            else:
                parent_level = level - 1
                while parent_level > 1 and parent_level not in last_nodes:
                    parent_level -= 1

                parent_node = last_nodes.get(parent_level)
                color = node_color_map.get(parent_node, 'lightblue')

            node_color_map[num] = color

            shape = 'circle' if row['類型'] == '直銷商' else 'box'

            dot.node(
                num,
                label=label,
                shape=shape,
                style='filled',
                color=color,
                width=str(size),
                fixedsize='false',
                margin='0.1'
            )

            if level == 1:
                last_nodes[level] = num
            else:
                parent = last_nodes.get(level - 1)
                if parent:
                    dot.edge(parent, num)

                last_nodes[level] = num

        timestamp = datetime.now().strftime("%Y%m%d-%H%M%S")
        output_filename = f"組織_{timestamp}.png"
        output_path = os.path.join(UPLOAD_FOLDER, output_filename)

        base_path = output_path.replace('.png','')

        dot.render(base_path, format='png', cleanup=True)
        dot.render(base_path, format='svg', cleanup=True)

        cleanup_old_files(UPLOAD_FOLDER, MAX_HISTORY)

        img_path = f"/{UPLOAD_FOLDER}/{output_filename}"
        svg_path = img_path.replace(".png",".svg")

        return jsonify({
            "img_path": img_path,
            "svg_path": svg_path
        })

    except Exception as e:
        print("Error:", e)
        return jsonify({"error": str(e)}), 500


if __name__=="__main__":
    app.run(host="0.0.0.0", port=5000, debug=True)