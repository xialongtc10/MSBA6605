import streamlit as st
import pandas as pd
import json
import io
import xlsxwriter
from openai import OpenAI

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
# AI PLAN GENERATION
# =========================
def generate_plan(df, prompt):
    system_prompt = """
Return ONLY JSON with structure:

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

        # 1. RAW DATA
        df.to_excel(writer, sheet_name="Raw_Data", index=False)

        # 2. TABLES
        created_sheets = {}

        for table in plan.get("tables", []):
            name = table["name"]

            if "groupby" in table:
                grouped = df.groupby(table["groupby"])[table["metric"]].sum().reset_index()
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

        # 3. CHARTS
        for chart_def in plan.get("charts", []):
            sheet = chart_def["sheet"]

            if sheet not in created_sheets:
                continue

            df_sheet = created_sheets[sheet]
            worksheet = writer.sheets[sheet]

            chart = workbook.add_chart({"type": chart_def["type"]})

            x_idx = df_sheet.columns.get_loc(chart_def["x"])
            y_idx = df_sheet.columns.get_loc(chart_def["y"])

            chart.add_series({
                "name": chart_def["title"],
                "categories": [sheet, 1, x_idx, len(df_sheet), x_idx],
                "values": [sheet, 1, y_idx, len(df_sheet), y_idx],
            })

            chart.set_title({"name": chart_def["title"]})

            worksheet.insert_chart("G2", chart)

        # 4. CONDITIONAL FORMATTING
        for rule in plan.get("conditional_formatting", []):
            sheet = rule["sheet"]

            if sheet not in created_sheets:
                continue

            df_sheet = created_sheets[sheet]
            worksheet = writer.sheets[sheet]

            col_idx = df_sheet.columns.get_loc(rule["column"])

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
                mean_val = df_sheet.iloc[:, col_idx].mean()
                worksheet.conditional_format(
                    1, col_idx, len(df_sheet), col_idx,
                    {
                        "type": "cell",
                        "criteria": ">",
                        "value": mean_val,
                        "format": workbook.add_format({"bg_color": "#C6EFCE"})
                    }
                )

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
