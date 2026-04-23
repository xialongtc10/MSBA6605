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
st.title("📊 AI Excel Agent (Pro Version)")

# =========================
# API KEY INPUT (UI)
# =========================
st.sidebar.header("🔐 API Key")

api_key_input = st.sidebar.text_input("Enter OpenAI API Key", type="password")
save_key = st.sidebar.checkbox("Save key for session")

if api_key_input and save_key:
    st.session_state["api_key"] = api_key_input

api_key = st.session_state.get("api_key", api_key_input)

if not api_key:
    st.warning("Please enter your API key in the sidebar.")
    st.stop()

client = OpenAI(api_key=api_key)

# =========================
# UTILITIES
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
# AI: PLAN GENERATION
# =========================
def generate_plan(df, prompt):
    system_prompt = """
You are an expert data analyst.

Return ONLY valid JSON.

Schema:
{
  "tables": [
    {
      "name": "Summary",
      "groupby": ["col"],
      "metric": "col"
    }
  ],
  "charts": [
    {
      "sheet": "Summary",
      "type": "bar",
      "x": "col",
      "y": "col",
      "title": "title"
    }
  ],
  "conditional_formatting": [
    {
      "sheet": "Summary",
      "column": "col",
      "rule": "negative_red"
    }
  ]
}

Rules:
- ONLY use columns from dataset
- Keep it minimal (1–3 tables, max 3 charts)
"""

    response = client.chat.completions.create(
        model="gpt-4o-mini",
        messages=[
            {"role": "system", "content": system_prompt},
            {
                "role": "user",
                "content": f"""
User request:
{prompt}

Columns:
{list(df.columns)}

Sample:
{df.head(5).to_string()}
"""
            }
        ],
        temperature=0.2,
    )

    return json.loads(response.choices[0].message.content)

# =========================
# AI INSIGHTS GENERATION
# =========================
def generate_insights(df_summary, prompt):
    try:
        response = client.chat.completions.create(
            model="gpt-4o-mini",
            messages=[
                {
                    "role": "system",
                    "content": """
You are a senior business analyst.

Generate 3–6 clear, data-driven insights.

Rules:
- Be specific (mention trends, comparisons)
- Avoid vague statements
- Focus on business meaning
"""
                },
                {
                    "role": "user",
                    "content": f"""
User request: {prompt}

Data:
{df_summary.head(30).to_string()}
"""
                }
            ],
            temperature=0.3,
        )

        return response.choices[0].message.content

    except Exception as e:
        return f"Insight generation failed: {e}"

# =========================
# EXCEL BUILDER
# =========================
def build_excel(df, plan):
    output = io.BytesIO()
    created_sheets = {}

    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        workbook = writer.book

        # ---------------- RAW DATA ----------------
        df.to_excel(writer, sheet_name="Raw_Data", index=False)

        # ---------------- TABLES ----------------
        for table in plan.get("tables", []):
            name = table["name"]

            try:
                group_cols = [match_column(c, df.columns) for c in table.get("groupby", [])]
                group_cols = [c for c in group_cols if c]

                metric = match_column(table["metric"], df.columns)

                if not group_cols or not metric:
                    continue

                grouped = df.groupby(group_cols)[metric].sum().reset_index()

                grouped.to_excel(writer, sheet_name=name, index=False)
                worksheet = writer.sheets[name]

                created_sheets[name] = grouped

                header_format = workbook.add_format({
                    "bold": True,
                    "bg_color": "#DCE6F1",
                    "border": 1
                })

                for i, col in enumerate(grouped.columns):
                    worksheet.write(0, i, col, header_format)
                    worksheet.set_column(i, i, 18)

            except:
                continue

        # ---------------- CHARTS ----------------
        for chart in plan.get("charts", []):
            sheet = chart["sheet"]

            if sheet not in created_sheets:
                continue

            df_sheet = created_sheets[sheet]
            worksheet = writer.sheets[sheet]

            x_col = match_column(chart["x"], df_sheet.columns)
            y_col = match_column(chart["y"], df_sheet.columns)

            if not x_col or not y_col:
                continue

            chart_obj = workbook.add_chart({"type": chart["type"]})

            x_idx = df_sheet.columns.get_loc(x_col)
            y_idx = df_sheet.columns.get_loc(y_col)

            chart_obj.add_series({
                "name": chart["title"],
                "categories": [sheet, 1, x_idx, len(df_sheet), x_idx],
                "values": [sheet, 1, y_idx, len(df_sheet), y_idx],
            })

            chart_obj.set_title({"name": chart["title"]})
            worksheet.insert_chart("G2", chart_obj)

        # ---------------- CONDITIONAL FORMATTING ----------------
        for rule in plan.get("conditional_formatting", []):
            sheet = rule["sheet"]

            if sheet not in created_sheets:
                continue

            df_sheet = created_sheets[sheet]
            worksheet = writer.sheets[sheet]

            col = match_column(rule["column"], df_sheet.columns)

            if not col:
                continue

            col_idx = df_sheet.columns.get_loc(col)

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

    output.seek(0)
    return output, created_sheets

# =========================
# UI
# =========================
uploaded_file = st.file_uploader("📂 Upload Excel file", type=["xlsx"])
prompt = st.text_area("💬 Enter your analysis request")

run = st.button("🚀 Run AI Analysis")

if run:
    if not uploaded_file or not prompt:
        st.error("Please upload file and enter prompt.")
        st.stop()

    df = pd.read_excel(uploaded_file)
    df = normalize_columns(df)

    st.subheader("📄 Data Preview")
    st.dataframe(df.head())

    with st.spinner("AI is thinking..."):
        plan = generate_plan(df, prompt)

        st.subheader("🧠 AI Plan")
        st.json(plan)

        excel_file, created_sheets = build_excel(df, plan)

        # pick first available result table for insights
        summary_df = next(iter(created_sheets.values()), df)

        insights = generate_insights(summary_df, prompt)

        st.subheader("🧠 AI Insights")
        st.write(insights)

        st.success("Report generated successfully!")

        st.download_button(
            "⬇️ Download Excel Report",
            data=excel_file,
            file_name="ai_excel_report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
