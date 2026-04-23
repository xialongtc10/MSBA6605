import streamlit as st
import pandas as pd
import json
import io
import xlsxwriter
from openai import OpenAI
from difflib import get_close_matches

# =========================
# PAGE CONFIG
# =========================
st.set_page_config(page_title="AI Excel Agent", layout="wide")
st.title("📊 AI Excel Agent")

# =========================
# API KEY INPUT
# =========================
st.sidebar.header("🔐 API Key")

api_key_input = st.sidebar.text_input("Enter OpenAI API Key", type="password")
save_key = st.sidebar.checkbox("Save for session")

if api_key_input and save_key:
    st.session_state["api_key"] = api_key_input

api_key = st.session_state.get("api_key", api_key_input)

if not api_key:
    st.warning("Please enter your API key in the sidebar.")
    st.stop()

client = OpenAI(api_key=api_key)

# =========================
# UTIL FUNCTIONS
# =========================
def normalize_columns(df):
    df.columns = (
        df.columns
        .str.strip()
        .str.lower()
        .str.replace(" ", "")
    )
    return df

def match_column(col, columns):
    matches = get_close_matches(col.lower(), [c.lower() for c in columns], n=1, cutoff=0.6)
    if matches:
        for c in columns:
            if c.lower() == matches[0]:
                return c
    return None

# =========================
# AI PLAN GENERATION
# =========================
def generate_plan(df, prompt):
    system_prompt = """
Return ONLY valid JSON:

{
  "tables": [
    {"name": "Summary", "groupby": ["col"], "metric": "col"}
  ],
  "charts": [
    {"sheet": "Summary", "type": "bar", "x": "col", "y": "col", "title": "title"}
  ],
  "conditional_formatting": [
    {"sheet": "Summary", "column": "col", "rule": "negative_red"}
  ]
}

IMPORTANT:
- Use ONLY columns from dataset
- Use lowercase column names
"""

    response = client.chat.completions.create(
        model="gpt-4o-mini",
        messages=[
            {"role": "system", "content": system_prompt},
            {
                "role": "user",
                "content": f"""
User request: {prompt}

Columns: {list(df.columns)}

Sample:
{df.head(5).to_string()}
"""
            }
        ],
        temperature=0.2,
    )

    return json.loads(response.choices[0].message.content)

# =========================
# BUILD EXCEL
# =========================
def build_excel(df, plan):
    output = io.BytesIO()

    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        workbook = writer.book

        # ---------- RAW DATA ----------
        df.to_excel(writer, sheet_name="Raw_Data", index=False)

        created_sheets = {}

        # ---------- TABLES ----------
        for table in plan.get("tables", []):
            name = table["name"]

            try:
                if "groupby" in table:
                    group_cols = [match_column(c, df.columns) for c in table["groupby"]]
                    group_cols = [c for c in group_cols if c]

                    metric_col = match_column(table["metric"], df.columns)

                    if not group_cols or not metric_col:
                        st.warning(f"Skipping table '{name}' due to column mismatch")
                        continue

                    grouped = df.groupby(group_cols)[metric_col].sum().reset_index()
                else:
                    grouped = df.copy()

                grouped.to_excel(writer, sheet_name=name, index=False)
                worksheet = writer.sheets[name]
                created_sheets[name] = grouped

                # formatting
                header_format = workbook.add_format({
                    "bold": True,
                    "bg_color": "#DCE6F1",
                    "border": 1
                })

                for col_num, col_name in enumerate(grouped.columns):
                    worksheet.write(0, col_num, col_name, header_format)
                    worksheet.set_column(col_num, col_num, 18)

            except Exception as e:
                st.warning(f"Error creating table '{name}': {e}")

        # ---------- CHARTS ----------
        for chart_def in plan.get("charts", []):
            sheet = chart_def["sheet"]

            if sheet not in created_sheets:
                continue

            df_sheet = created_sheets[sheet]
            worksheet = writer.sheets[sheet]

            x_col = match_column(chart_def["x"], df_sheet.columns)
            y_col = match_column(chart_def["y"], df_sheet.columns)

            if not x_col or not y_col:
                st.warning(f"Skipping chart in '{sheet}' due to column mismatch")
                continue

            try:
                chart = workbook.add_chart({"type": chart_def["type"]})

                x_idx = df_sheet.columns.get_loc(x_col)
                y_idx = df_sheet.columns.get_loc(y_col)

                chart.add_series({
                    "name": chart_def["title"],
                    "categories": [sheet, 1, x_idx, len(df_sheet), x_idx],
                    "values": [sheet, 1, y_idx, len(df_sheet), y_idx],
                })

                chart.set_title({"name": chart_def["title"]})
                worksheet.insert_chart("G2", chart)

            except Exception as e:
                st.warning(f"Error creating chart: {e}")

        # ---------- CONDITIONAL FORMATTING ----------
        for rule in plan.get("conditional_formatting", []):
            sheet = rule["sheet"]

            if sheet not in created_sheets:
                continue

            df_sheet = created_sheets[sheet]
            worksheet = writer.sheets[sheet]

            col = match_column(rule["column"], df_sheet.columns)

            if not col:
                st.warning(f"Skipping formatting: column '{rule['column']}' not found")
                continue

            col_idx = df_sheet.columns.get_loc(col)

            try:
                if rule["rule"] == "negative_red":
                    worksheet.conditional_format(
                        1, col_idx, len(df_sheet), col_idx,
                        {
                            "type": "cell",
                            "criteria": "<",
                            "value": 0,
                            "format": workbook.add_format({"font_color": "red"})
                        }
                    )

                elif rule["rule"] == "greater_than_mean":
                    mean_val = df_sheet[col].mean()
                    worksheet.conditional_format(
                        1, col_idx, len(df_sheet), col_idx,
                        {
                            "type": "cell",
                            "criteria": ">",
                            "value": mean_val,
                            "format": workbook.add_format({"bg_color": "#C6EFCE"})
                        }
                    )

            except Exception as e:
                st.warning(f"Error applying formatting: {e}")

    output.seek(0)
    return output

# =========================
# UI
# =========================
uploaded_file = st.file_uploader("📂 Upload Excel file", type=["xlsx"])
prompt = st.text_area("💬 Enter your analysis request")

run_button = st.button("🚀 Run AI Analysis")

if run_button:
    if not uploaded_file:
        st.error("Please upload a file.")
        st.stop()

    if not prompt:
        st.error("Please enter a prompt.")
        st.stop()

    df = pd.read_excel(uploaded_file)
    df = normalize_columns(df)

    st.subheader("📄 Data Preview")
    st.dataframe(df.head())

    with st.spinner("AI is generating report..."):
        plan = generate_plan(df, prompt)

        st.subheader("🧠 AI Plan")
        st.json(plan)

        excel_file = build_excel(df, plan)

        st.success("Report ready!")

        st.download_button(
            "⬇️ Download Excel Report",
            data=excel_file,
            file_name="ai_excel_report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
