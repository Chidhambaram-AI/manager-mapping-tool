from flask import Flask, render_template, request, send_file, abort
import pandas as pd
import io
from collections import defaultdict, deque

app = Flask(__name__)

def normalize(x):
    if pd.isna(x):
        return None
    s = str(x).strip()
    return s if s else None

def collect_closure(manager, manager_to_reports):
    """
    Return set of all users who report to `manager` (directly or indirectly).
    """
    visited = set()
    stack = list(manager_to_reports.get(manager, []))
    while stack:
        u = stack.pop()
        if u in visited:
            continue
        visited.add(u)
        # extend with users who report to u (if u is also a manager)
        stack.extend(manager_to_reports.get(u, []))
    return visited

@app.route("/")
def home():
    return render_template("index.html")

@app.route("/download_blank")
def download_blank():
    df = pd.DataFrame(columns=[
        "SNo", "Firstname", "Unique ID", "Usergroup",
        "Reporting Manager 1", "Reporting Manager 2"
    ])
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Template")
    buf.seek(0)
    return send_file(buf,
                     as_attachment=True,
                     download_name="Blank_Template.xlsx",
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

@app.route("/download_sample")
def download_sample():
    data = {
        "SNo": [1,2,3,4,5,6,7],
        "Firstname": ["A","B","C","D","G","H","I"],
        "Unique ID": ["","", "","", "","",""],
        "Usergroup": ["Coding","HR","AR","Sales","dev","test","AM"],
        "Reporting Manager 1": ["D","D","","J","B","B","A"],
        "Reporting Manager 2": ["E","E","","","", "",""]
    }
    df = pd.DataFrame(data)
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Sample")
    buf.seek(0)
    return send_file(buf,
                     as_attachment=True,
                     download_name="Sample_Template.xlsx",
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

@app.route("/process", methods=["POST"])
def process():
    uploaded = request.files.get("excel_file")
    if not uploaded:
        return abort(400, "No file uploaded")

    try:
        sheets = pd.read_excel(uploaded, sheet_name=None)
    except Exception as e:
        return abort(400, f"Could not read Excel: {e}")

    first_sheet = list(sheets.keys())[0]
    df = sheets[first_sheet].copy()
    # Normalize column names
    df.columns = [str(c).strip() for c in df.columns]

    # Expected column names (case-sensitive if exactly provided). We'll accept case-insensitive mapping.
    # Map user columns
    col_map = {c.lower(): c for c in df.columns}
    required = ["firstname", "usergroup", "reporting manager 1", "reporting manager 2"]
    for r in required:
        if r not in col_map:
            # If missing required columns, return helpful error
            return abort(400, f"Required column missing in sheet '{first_sheet}': {r}")

    fn = col_map["firstname"]
    ug_col = col_map["usergroup"]
    rm1_col = col_map["reporting manager 1"]
    rm2_col = col_map["reporting manager 2"]

    # Build name -> usergroup map (for users listed as Firstname)
    name_to_group = {}
    firstname_list = []
    for _, row in df.iterrows():
        name = normalize(row.get(fn))
        if name:
            firstname_list.append(name)
            grp = normalize(row.get(ug_col))
            if grp:
                name_to_group[name] = grp

    # Build manager -> direct reports mapping (manager_name -> list of user firstnames)
    manager_to_reports = defaultdict(list)
    # Also track which names appear in RM1/RM2
    managers_set = set()
    # Keep rows stored
    all_rows = []
    for _, row in df.iterrows():
        user = normalize(row.get(fn))
        ug = normalize(row.get(ug_col))
        m1 = normalize(row.get(rm1_col))
        m2 = normalize(row.get(rm2_col))
        all_rows.append({"user": user, "ug": ug, "m1": m1, "m2": m2})
        if m1:
            managers_set.add(m1)
            if user:
                manager_to_reports[m1].append(user)
        if m2:
            managers_set.add(m2)
            if user:
                manager_to_reports[m2].append(user)

    # Only managers that actually appear in RM1 or RM2
    managers = sorted(managers_set)

    summary = []
    # Precompute closure for all managers (their full reporting trees)
    closure_cache = {}
    for m in managers:
        closure_cache[m] = collect_closure(m, manager_to_reports)

    # For RM2 special inheritance: build mapping user -> their RM1 (if any)
    user_to_rm1 = {}
    for r in all_rows:
        u = r["user"]
        if u:
            if r["m1"]:
                user_to_rm1[u] = r["m1"]

    for m in managers:
        own_group = name_to_group.get(m)  # manager might also be a user
        direct_reports = manager_to_reports.get(m, [])
        direct_groups = set()
        for u in direct_reports:
            g = name_to_group.get(u)
            if g:
                direct_groups.add(g)

        # All reports (closure) groups
        closure_users = closure_cache.get(m, set())
        closure_groups = set()
        for u in closure_users:
            g = name_to_group.get(u)
            if g:
                closure_groups.add(g)

        # RM2 inheritance: find users X where X has RM2 == m; for each such X,
        # find X's RM1 (call it r1). If r1 exists, include full closure groups of r1.
        rm2_inherit_groups = set()
        for r in all_rows:
            u = r["user"]
            if not u:
                continue
            if r["m2'] == m if False else None:
                pass
        # The above is a placeholder. We'll compute properly:
        for r in all_rows:
            u = r["user"]
            if not u:
                continue
            if r["m2"] == m:
                r1 = r["m1"]
                if r1:
                    # include closure of r1 (all groups under r1) and r1 own & direct groups
                    # r1 might not be a manager (edge case); we'll collect groups for r1's closure
                    r1_closure = closure_cache.get(r1, set())
                    # add r1's own group as well
                    if name_to_group.get(r1):
                        rm2_inherit_groups.add(name_to_group.get(r1))
                    for uu in r1_closure:
                        gg = name_to_group.get(uu)
                        if gg:
                            rm2_inherit_groups.add(gg)

        # Compose final groups:
        final_groups = []
        # Keep insertion order: own -> direct -> rm2_inherit -> closure
        seen = set()
        def add_group(g):
            if not g: 
                return
            if g not in seen:
                seen.add(g)
                final_groups.append(g)

        # Own first
        add_group(own_group)
        # Direct report groups
        for g in sorted(direct_groups):
            add_group(g)
        # RM2 inheritance groups
        for g in sorted(rm2_inherit_groups):
            add_group(g)
        # Finally closure groups (indirect deeper)
        for g in sorted(closure_groups):
            add_group(g)

        # Determine types
        types = []
        if own_group:
            types.append("Own")
        if len(direct_reports) > 0:
            types.append("Direct")
        # If closure contains users beyond the direct set, mark Indirect
        indirect_users = closure_users - set(direct_reports)
        if len(indirect_users) > 0:
            types.append("Indirect")

        # If manager has only own (no direct/indirect), still include Own or mark NoReports
        if not types:
            types_text = "NoReports"
        else:
            types_text = ",".join(types)

        summary.append({
            "Manager": m,
            "Usergroups": ", ".join(final_groups) if final_groups else "",
            "Type": types_text,
            "Own": "Yes" if own_group else "No",
            "DirectCount": len(direct_reports),
            "IndirectCount": len(indirect_users)
        })

    # Create summary dataframe
    summary_df = pd.DataFrame(summary)

    # Write output excel with original sheet + Manager Summary
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name=first_sheet)
        summary_df.to_excel(writer, index=False, sheet_name="Manager Summary")
    buf.seek(0)

    return send_file(buf,
                     as_attachment=True,
                     download_name="Processed_Manager_Mapping.xlsx",
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000, debug=False)

