import os
from typing import List, Dict, Any
import pandas as pd

# Use the official OpenAI Python SDK
# Docs:
# - API reference (Responses): https://platform.openai.com/docs/api-reference/responses
# - Migration guide: https://platform.openai.com/docs/guides/migrate-to-responses
try:
    from openai import OpenAI
except Exception as e:
    OpenAI = None

def _df_preview_markdown(df: pd.DataFrame, max_rows: int = 20) -> str:
    """Return a compact Markdown preview of a dataframe (top rows) for the model context."""
    if df is None or df.empty:
        return "_No data available_"
    head = df.head(max_rows)
    # Limit columns to most relevant
    keep_cols = [c for c in ["Industry", "Parameter/Method", "Instrument Type", "Manufacturer", "Age", "Date"] if c in head.columns]
    head = head[keep_cols] if keep_cols else head
    return "```\n" + head.to_csv(index=False) + "```"

def build_context(df: pd.DataFrame) -> str:
    """Build a concise context string for the model including small grouped summaries."""
    sections = []

    def group_count(by_cols: List[str]) -> pd.DataFrame:
        by_cols = [c for c in by_cols if c in df.columns]
        if not by_cols:
            return pd.DataFrame()
        return df.groupby(by_cols, dropna=False).size().reset_index(name="Count").sort_values("Count", ascending=False)

    sections.append("### Data Preview\n" + _df_preview_markdown(df))

    for title, cols in [
        ("Industry Summary", ["Industry"]),
        ("Instrument Type by Industry", ["Industry", "Instrument Type"]),
        ("Manufacturer by Industry", ["Industry", "Manufacturer"]),
        ("Age by Industry", ["Industry", "Age"]),
    ]:
        g = group_count(cols)
        if not g.empty:
            sections.append(f"### {title}\n```\n{g.head(50).to_csv(index=False)}\n```")

    return "\n\n".join(sections)

def generate_report(df: pd.DataFrame, user_request: str, model: str = "gpt-4o-mini") -> str:
    """
    Generate a structured market report as Markdown based on user's natural language request.
    Returns markdown text.
    """
    if OpenAI is None:
        raise RuntimeError("OpenAI SDK not installed. Please add 'openai' to requirements and install.")

    api_key = os.getenv("OPENAI_API_KEY")
    if not api_key:
        raise RuntimeError("OPENAI_API_KEY is not set. Provide it via Streamlit secrets or environment variable.")

    client = OpenAI(api_key=api_key)

    context = build_context(df)

    system = """You are a senior market analyst. Produce clear, structured Markdown reports.
- Keep it concise, evidence-based, and action-oriented.
- Include: Executive Summary, Key Trends, Segment Breakdown, Opportunities/Risks, Recommendations.
- Reference figures with captions (e.g., 'Figure 1: Industry share').
- If the user's request mentions exporting, include a final 'Deliverables' checklist.
- Assume charts shown in the web app are available; don't invent numbers beyond provided aggregates.
- When useful, include short SQL-like or pandas-like pseudo-queries to explain calculations.
"""

    prompt = f"""User request:\n{user_request}\n\nDataset context (summaries & preview):\n{context}\n\nNow write the report."""

    resp = client.responses.create(
        model=model,
        input=[
            {"role": "system", "content": system},
            {"role": "user", "content": prompt},
        ],
    )
    # Prefer output_text for simplicity
    try:
        return resp.output_text
    except Exception:
        # Fallback: extract from first content part
        if hasattr(resp, "output") and resp.output and hasattr(resp.output[0], "content"):
            return resp.output[0].content[0].text.value
        return "*(No response text returned by the model.)*"