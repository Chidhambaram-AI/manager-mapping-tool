from flask import Flask, render_template, request, send_file
import pandas as pd
import io
from collections import defaultdict, deque

app = Flask(__name__)

# Normalize text values
def normalize(x):
    if pd.isna(x):
        return None
    x = str(x).strip()
    return x if x else None

# BFS for indirect manager check
def bfs(start, graph):
    dist = {}
    queue = deque([(start, 0)])
    dist[start] = 0
    while queue:
        node, d = queue.popleft()
        for m in graph.get(node, []):
            if m not in dist:
                dist[m] = d + 1
                queue.append((m, d + 1))
    return dist

@app.route("/")
def home():
    return render_template("index.html")


# ------------------ DOWNLOAD BLANK TEMPLATE ------------------

@app.route("/download_blank")
def download_blank():
    df = pd.DataFrame(columns=[
        "SNo", "Firstname", "Unique ID", "Usergroup",
        "Reporting Manager 1", "Reporting Manager 2"
    ])

    file = io.BytesIO()
    with pd.ExcelWriter(file, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Template")

    file.seek(0)
    return send_file(
        file,
        as_attachment=True,
        download_name="Blank_Template.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )


# ------------------ DOWNLOAD SAMPLE TEMPLATE ------------------

@app.route("/download_sample")
def download_sample():
    sample_data = {
        "SNo": [1, 2, 3],
        "Firstname": ["Ravi", "Arjun", "Keerthana"],
        "Unique ID": ["EMP001", "EMP002", "EMP003"],
        "Usergroup": ["Finance", "Sales", "HR"],
        "Reporting Manager 1": ["Manoj", "Deepa", "Manoj"],
        "Reporting Manager 2": ["Deepa", None, None]
    }

    df = pd.DataFrame(sample_data)

    file = io.BytesIO()
    with pd.ExcelWriter(file, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Sample Data")

    file.seek(0)
    return send_file(
        file,
        as_attachment=True,
        download_name="Sample_Template.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )


# ------------------ PROCESS UPLOADED FILE ------------------

@app.route("/process", methods=["POST"])
def process():
    uploaded = request.files.get("excel_file")
    if not uploaded:
        return "No file uploaded.", 400

    excel_sheets = pd.read_excel(uploaded, sheet_name=None)
    first_sheet = list(excel_sheets.keys())[0]
    df = excel_sheets[first_sheet].copy()
    df.columns = [c.strip() for c in df.columns]

    rows = []
    graph = defaultdict(list)
    managers = set()

    for _, r in df.iterrows():
        user = normalize(r.get("Firstname"))
        ug = normalize(r.get("Usergroup"))
        m1 = normalize(r.get("Reporting Manager 1"))
        m2 = normalize(r.get("Reporting Manager 2"))

        rows.append({"user": user, "ug": ug, "m1": m1, "m2": m2})

        if m1: managers.add(m1)
        if m2: managers.add(m2)

        if user:
            if m1: graph[user].append(m1)
            if m2: graph[user].append(m2)

    summary = []

    for mgr in sorted(managers):
        direct_users = [r for r in rows if r["m1"] == mgr or r["m2"] == mgr]
        direct_flag = len(direct_users) > 0

        indirect_flag = False
        for r in rows:
            u = r["user"]
            if not u:
                continue
            dist = bfs(u, graph)
            if mgr in dist and dist[mgr] >= 2:
                indirect_flag = True

        if direct_flag and indirect_flag:
            type_val = "Direct + Indirect"
        elif direct_flag:
            type_val = "Direct"
        elif indirect_flag:
            type_val = "Indirect"
        else:
            type_val = "-"

        ugroups = sorted({r["ug"] for r in direct_users if r["ug"]})
        summary.append({
            "Manager": mgr,
            "Usergroups": ", ".join(ugroups),
            "Type": type_val
        })

    summary_df = pd.DataFrame(summary)

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name=first_sheet)
        summary_df.to_excel(writer, index=False, sheet_name="Manager Summary")

    output.seek(0)
    return send_file(
        output,
        as_attachment=True,
        download_name="Processed_Manager_Mapping.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000)
