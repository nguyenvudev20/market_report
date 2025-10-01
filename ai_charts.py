import os
import json
import pandas as pd
import plotly.express as px

try:
    from openai import OpenAI
except Exception:
    OpenAI = None

SUPPORTED_CHARTS = {"bar", "column", "pie", "line"}

SYSTEM_PROMPT = """Bạn là một công cụ tạo biểu đồ từ yêu cầu tiếng Việt.
Hãy đọc mô tả của người dùng và tạo ra một JSON duy nhất mô tả biểu đồ theo cấu trúc sau (KHÔNG giải thích thêm, chỉ in JSON):
{
  "chart_type": "bar|column|pie|line",
  "x": "<tên cột dùng cho trục x hoặc danh mục>",
  "y": "<tên cột số hoặc 'Count'>",
  "color": "<tên cột phân nhóm hoặc null>",
  "agg": "count|sum|avg|max|min",
  "top_n": <số nguyên hoặc null>,
  "filters": [
    {"column": "<tên cột>", "op": "eq|in|contains|between", "value": "<giá trị hoặc [a,b]>"}
  ],
  "title": "<tiêu đề tiếng Việt>"
}
Quy ước: Nếu không có cột số, đặt y = "Count" và agg = "count".
Nếu người dùng nói "top 5", đặt "top_n": 5. Nếu không nói rõ thì để null.
Chỉ chọn các cột có khả năng xuất hiện trong dữ liệu như: Industry, Parameter/Method, Instrument Type, Manufacturer, Age, Date, Hanna Office, Hanna Rep, Customer Name, Model.
"""

def _get_client():
    api_key = os.getenv("OPENAI_API_KEY")
    if not api_key:
        raise RuntimeError("OPENAI_API_KEY chưa được thiết lập.")
    if OpenAI is None:
        raise RuntimeError("Chưa cài 'openai'.")
    return OpenAI(api_key=api_key)

def infer_spec_via_openai(user_prompt: str, model: str = "gpt-4o-mini") -> dict:
    client = _get_client()
    resp = client.responses.create(
        model=model,
        input=[
            {"role": "system", "content": SYSTEM_PROMPT},
            {"role": "user", "content": user_prompt},
        ],
    )
    try:
        text = resp.output_text
    except Exception:
        text = ""
        try:
            text = resp.output[0].content[0].text.value
        except Exception:
            pass
    text = text.strip().strip("`").strip()
    if text.startswith("json"):
        text = text[4:].strip()
    spec = json.loads(text)
    return spec

def _apply_filters(df: pd.DataFrame, filters):
    if not filters:
        return df
    out = df.copy()
    for f in filters:
        col = f.get("column")
        op = f.get("op")
        val = f.get("value")
        if col not in out.columns:
            continue
        if op == "eq":
            out = out[out[col] == val]
        elif op == "in":
            if not isinstance(val, list):
                val = [val]
            out = out[out[col].isin(val)]
        elif op == "contains":
            out = out[out[col].astype(str).str.contains(str(val), case=False, na=False)]
        elif op == "between":
            try:
                a, b = val
                out = out[(out[col] >= a) & (out[col] <= b)]
            except Exception:
                pass
    return out

def _aggregate(df: pd.DataFrame, x, y, color, agg):
    if y is None or y == "Count" or agg == "count":
        group_cols = [c for c in [x, color] if c]
        g = df.groupby(group_cols, dropna=False).size().reset_index(name="Count")
        ycol = "Count"
    else:
        group_cols = [c for c in [x, color] if c]
        if agg == "sum":
            g = df.groupby(group_cols, dropna=False)[y].sum().reset_index(name=y)
        elif agg == "avg":
            g = df.groupby(group_cols, dropna=False)[y].mean().reset_index(name=y)
        elif agg == "max":
            g = df.groupby(group_cols, dropna=False)[y].max().reset_index(name=y)
        elif agg == "min":
            g = df.groupby(group_cols, dropna=False)[y].min().reset_index(name=y)
        else:
            g = df.groupby(group_cols, dropna=False).size().reset_index(name="Count")
            y = "Count"
        ycol = y
    return g, ycol

def render_chart(df: pd.DataFrame, spec: dict):
    chart_type = (spec.get("chart_type") or "bar").lower()
    if chart_type not in SUPPORTED_CHARTS:
        chart_type = "bar"
    x = spec.get("x")
    y = spec.get("y") or "Count"
    color = spec.get("color")
    agg = (spec.get("agg") or "count").lower()
    top_n = spec.get("top_n")
    filters = spec.get("filters", [])
    title = spec.get("title") or "Biểu đồ"

    dff = _apply_filters(df, filters)

    if x and x not in dff.columns:
        raise ValueError(f"Cột '{x}' không tồn tại.")
    if color and color not in dff.columns:
        color = None
    if y not in dff.columns and y != "Count":
        y = "Count"
        agg = "count"

    g, ycol = _aggregate(dff, x, y, color, agg)

    if isinstance(top_n, int) and top_n > 0:
        g = g.sort_values(ycol, ascending=False).head(top_n)

    if chart_type in {"bar", "column"}:
        fig = px.bar(g, x=x, y=ycol, color=color, barmode="group", title=title)
    elif chart_type == "pie":
        fig = px.pie(g, names=x, values=ycol, title=title, hole=0.35)
    elif chart_type == "line":
        fig = px.line(g, x=x, y=ycol, color=color, title=title)
    else:
        fig = px.bar(g, x=x, y=ycol, color=color, barmode="group", title=title)

    return fig, g