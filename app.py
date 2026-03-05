from flask import Flask, render_template, request, send_file
import pandas as pd
from graphviz import Digraph
import os
import itertools
from datetime import datetime
from werkzeug.utils import secure_filename

app = Flask(
    __name__,
    static_folder="outputs",
    static_url_path="/outputs"
)

UPLOAD_FOLDER = "uploads"
OUTPUT_FOLDER = "outputs"

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)


def generate_chart(file_path):

    df = pd.read_excel(file_path)

    df.columns = [c.strip().replace("　", "") for c in df.columns]

    df['層級'] = pd.to_numeric(df['層級'], errors='coerce').fillna(1)
    df['小組積分額'] = pd.to_numeric(df['小組積分額'], errors='coerce').fillna(0)
    df['小組售貨額'] = pd.to_numeric(df['小組售貨額'], errors='coerce').fillna(0)
    df['編號'] = df['編號'].astype(str)
    df['%'] = pd.to_numeric(df['%'], errors='coerce').fillna(0)

    if '會員積分額' not in df.columns:
        df['會員積分額'] = 0
    else:
        df['會員積分額'] = pd.to_numeric(df['會員積分額'], errors='coerce').fillna(0)

    def get_node_size(value, min_size=1.0, max_size=2.0):

        min_val = df['小組積分額'].min()
        max_val = df['小組積分額'].max()

        if max_val == min_val:
            return 1.5

        norm = (value - min_val) / (max_val - min_val)

        return round(min_size + norm * (max_size - min_size), 2)

    color_palette = [
        'lightblue','lightgreen','orange','pink','yellow',
        'lightcoral','lightgoldenrod','lightsalmon','lightseagreen'
    ]

    color_cycle = itertools.cycle(color_palette)

    level2_nodes = df[df['層級']==2]['編號'].tolist()

    level2_colors = {node:next(color_cycle) for node in level2_nodes}

    dot = Digraph(format='png')

    dot.attr(
        overlap='false',
        splines='true',
        layout='dot',
        rankdir='TB',
        fontname='Noto Sans CJK TC'
    )

    last_nodes={}
    node_color_map={}

    for _,row in df.iterrows():

        level=int(row['層級'])
        num=row['編號']

        percent_td=""

        if row['%']!=0:

            percent_value=int(round(row['%']*100))

            percent_td=f"""
            <TR>
            <TD BGCOLOR="#D32F2F" ALIGN="CENTER">
            <B><FONT POINT-SIZE="16">{percent_value}%</FONT></B>
            </TD>
            </TR>
            """

        bv_value=row['小組售貨額']

        if row['類型']=="會員":
            bv_value+=row['會員積分額']

        bv_td=""

        if bv_value!=0:
            bv_td=f"<TR><TD ALIGN='CENTER'>BV:{bv_value}</TD></TR>"

        label=f"""<
        <TABLE BORDER="0" CELLBORDER="0" CELLSPACING="0">
        {percent_td}
        <TR><TD ALIGN="CENTER">{row['姓名']}</TD></TR>
        {bv_td}
        </TABLE>
        >"""

        size=get_node_size(row['小組積分額'])

        if level==1:
            color="lightgrey"

        elif level==2:

            children=df[df['推薦人']==num]

            if children.empty and row['類型']=="會員":
                color="lightgrey"
            else:
                color=level2_colors.get(num,'lightblue')

        else:

            parent_level=level-1

            while parent_level>1 and parent_level not in last_nodes:
                parent_level-=1

            parent_node=last_nodes.get(parent_level)

            color=node_color_map.get(parent_node,'lightblue')

        node_color_map[num]=color

        shape='circle' if row['類型']=="直銷商" else 'box'

        dot.node(
            num,
            label=label,
            shape=shape,
            style='filled',
            color=color,
            width=str(size),
            margin='0.1'
        )

        if level==1:
            last_nodes[level]=num
        else:

            parent=last_nodes.get(level-1)

            if parent:
                dot.edge(parent,num)

            last_nodes[level]=num

    timestamp=datetime.now().strftime("%Y%m%d-%H%M%S")

    filename=f"org_{timestamp}"

    output_path=os.path.join(OUTPUT_FOLDER,filename)

    dot.render(output_path,cleanup=True)

    return filename+".png"


@app.route("/",methods=["GET","POST"])
def index():

    image=None

    if request.method=="POST":

        file=request.files["file"]

        filename=secure_filename(file.filename)

        upload_path=os.path.join(UPLOAD_FOLDER,filename)

        file.save(upload_path)

        image=generate_chart(upload_path)

    return render_template("index.html",image=image)


@app.route("/download/<filename>")
def download(filename):

    path=os.path.join(OUTPUT_FOLDER,filename)

    return send_file(path,as_attachment=True)


if __name__=="__main__":
    app.run(host="0.0.0.0",port=10000)