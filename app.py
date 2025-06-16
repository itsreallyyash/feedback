# app.py
import os
from fastapi import FastAPI, HTTPException
from fastapi.responses import HTMLResponse
from dotenv import load_dotenv

# import all of your existing helpers here
from mailgen import (
    load_data_from_public_sharepoint,
    load_data_fallback,
    split_periods,
    summarize,
    extract_feedback_data,
    generate_html_report,
)

load_dotenv()  # so your .env with OPENAI_API_KEY etc. is loaded

app = FastAPI()

@app.get("/generate-report", response_class=HTMLResponse)
async def generate_report():
    try:
        # load
        try:
            df = load_data_from_public_sharepoint()
        except Exception:
            df = load_data_fallback()

        # process - FIXED: now unpacking all 8 values returned by split_periods
        recent_df, prev_df, start_prev, start_last, end_last, positive_count, negative_count, total_count = split_periods(df)
        sum_recent = summarize(recent_df)
        sum_prev   = summarize(prev_df)
        recent_fb  = extract_feedback_data(recent_df)
        prev_fb    = extract_feedback_data(prev_df)

        # HTML - FIXED: now passing all required parameters
        html = generate_html_report(
            sum_recent, sum_prev,
            recent_fb, prev_fb,
            (start_prev, start_last, end_last),
            positive_count, negative_count, total_count
        )
        return HTMLResponse(html)

    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))


if __name__ == "__main__":
    import uvicorn
    uvicorn.run("app:app", host="0.0.0.0", port=int(os.getenv("PORT", 8000)), reload=True)