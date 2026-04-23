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
st.title("📊 AI Excel Agent (Tables + Charts + Formatting)")

# =========================
# 1. API KEY INPUT (UI)
# =========================
st.sidebar.header("🔐 API Configuration")

api_key = st.sidebar.text_input(
    "Enter OpenAI API Key",
    type="password"
)

save_key = st.sidebar.checkbox("Save API key for this session")

if api_key and save_key:
    st.session_state["OPENAI_API_KEY"] = api_key

# Use stored key if available
active_key = st.session_state.get("OPENAI_API_KEY", api_key)

if not active_key:
    st.warning("Please enter your OpenAI API key in the sidebar to continue.")
    st.stop()

client = OpenAI(api_key=active_key)

# =========================
# 2. AI PLAN GENERATION
# =========================
def generate_plan(df_sample, user_prompt, columns):
    system_prompt = """
You are an Excel assistant.

Return ONLY valid JSON.

You can output:
- tables (sheet + columns)
- charts (bar, line, pie, combo)
- conditional_formatting

Keep it simple and useful for business analysis.
"""

    response = client.chat.completions.create(
        model="gpt-4o-mini",
        messages=[
            {"role": "system", "content": system_prompt},
            {
                "role": "user",
                "content": f"""
User request:
{user_prompt}

Columns:
{columns}

Sample data:
{df_sample.head(5).to_string()}
"""
            }
        ],
        temperature=0.2,
    )

    return json.loads(response.choices[0].message.content)

# =========================
# 3. EXCEL BUILDER
# =========================
def build_excel(df, plan):
    output = io.BytesIO()

    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        workbook = writer.book

        # ---------- TABLES ----------
        for table in plan.get("tables", []):
            sheet = table["name"]
            cols = table.get("columns", df.columns)

            data = df[cols]
            data.to_excel(writer, sheet_name=sheet, index=False)

            worksheet = writer.sheets[sheet]

            header_format = workbook.add_format({
                "bold": True,
                "bg_color": "#DCE6F1",
                "border": 1
            })

            for col_num, value in enumerate(data.columns):
                worksheet.write(0, col_num, value, header_format)
                worksheet.set_column(col_num, col_num, 18)

        # ---------- CHARTS ----------
        for chart_def in plan.get("charts", []):
            sheet = chart_def["sheet"]
            worksheet = writer.sheets[sheet]

            chart = workbook.add_chart({"type": chart_def["type"]})

            df_sheet = df  # simplified assumption

            x_col = df_sheet.columns.get_loc(chart_def["x"])
            y_col = df_sheet.columns.get_loc(chart_def["y"])

            chart.add_series({
                "name": chart_def["title"],
                "categories": [sheet, 1, x_col, len(df_sheet), x_col],
                "values": [sheet, 1, y_col, len(df_sheet), y_col],
            })

            chart.set_title({"name": chart_def["title"]})
            worksheet.insert_chart("G2", chart)

        # ---------- CONDITIONAL FORMATTING ----------
        for rule in plan.get("conditional_formatting", []):
            sheet = rule["sheet"]
            column = rule["column"]

            worksheet = writer.sheets[sheet]
            col_idx = df.columns.get_loc(column)

            if rule["rule"] == "negative_red":
                worksheet.conditional_format(
                    1, col_idx, len(df), col_idx,
                    {
                        "type": "cell",
                        "criteria": "<",
                        "value": 0,
                        "format": workbook.add_format({"font_color": "red"})
                    }
                )

    output.seek(0)
    return output

# =========================
# 4. UI
# =========================
uploaded_file = st.file_uploader("📂 Upload Excel file", type=["xlsx"])
prompt = st.text_area("💬 What do you want to do with the data?")

if uploaded_file and prompt:
    df = pd.read_excel(uploaded_file)

    st.subheader("📄 Data Preview")
    st.dataframe(df.head())

    if st.button("🚀 Generate Excel Report"):
        with st.spinner("AI is building your Excel report..."):

            plan = generate_plan(
                df_sample=df,
                user_prompt=prompt,
                columns=list(df.columns)
            )

            st.subheader("🧠 AI Plan")
            st.json(plan)

            excel_file = build_excel(df, plan)

            st.success("Done!")

            st.download_button(
                "⬇️ Download Excel File",
                data=excel_file,
                file_name="ai_excel_report.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
