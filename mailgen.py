
# import os
# import pandas as pd
# from datetime import datetime, timedelta
# from openai import OpenAI  # Updated import
# import requests
# from io import BytesIO
# from dotenv import load_dotenv

# # Load environment variables from .env file
# load_dotenv()

# # ——— Configuration ———
# # SharePoint public link 
# SHAREPOINT_PUBLIC_URL = "https://neogroupinfotech-my.sharepoint.com/:x:/g/personal/aashna_mehta_neo-group_in/ERmfzY7yPatEo3fni7EeZ8oBhjxYxuv17V6bcYTEfwlV7w?e=NDFhd4"

# # Try multiple SharePoint download URL formats
# def get_sharepoint_download_urls(sharepoint_url):
#     """Generate different SharePoint download URL formats to try"""
#     urls = []
    
#     # Method 1: Replace :x: with :u: and add download=1
#     url1 = sharepoint_url.replace(":x:", ":u:").replace("?e=", "&download=1&e=")
#     urls.append(url1)
    
#     # Method 2: Use direct download parameter
#     url2 = sharepoint_url + "&download=1"
#     urls.append(url2)
    
#     # Method 3: Extract file ID and use download format
#     if "ERmfzY7yPatEo3fni7EeZ8oBhjxYxuv17V6bcYTEfwlV7w" in sharepoint_url:
#         file_id = "ERmfzY7yPatEo3fni7EeZ8oBhjxYxuv17V6bcYTEfwlV7w"
#         base_url = "https://neogroupinfotech-my.sharepoint.com/personal/aashna_mehta_neo-group_in/_layouts/15/download.aspx"
#         url3 = f"{base_url}?share={file_id}"
#         urls.append(url3)
    
#     # Method 4: Try the embed format
#     url4 = sharepoint_url.replace(":x:", ":u:")
#     urls.append(url4)
    
#     return urls

# OUTPUT_HTML    = "report.html"
# CEO_NAME       = "Nitin Jain"
# SYSTEM_NAME    = "No-Reply Feedback System"
# DATE_COL_INDEX = 1  # zero-based index of the timestamp column (Start time)
# FEEDBACK_TYPE_COL = "What kind of feedback is it?"  # Column name for feedback type
# FEEDBACK_TEXT_COL = "Please describe your feedback"  # Column name for feedback text
# OPENAI_API_KEY = os.getenv("OPENAI_API_KEY") or "sk-your-actual-api-key-here"  # Fallback for testing

# # Initialize OpenAI client
# client = OpenAI(api_key=OPENAI_API_KEY)

# # ——— Helpers ———

# def load_data_from_public_sharepoint() -> pd.DataFrame:
#     """Load Excel data from public SharePoint link"""
#     download_urls = get_sharepoint_download_urls(SHAREPOINT_PUBLIC_URL)
    
#     headers = {
#         'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'
#     }
    
#     for i, url in enumerate(download_urls):
#         try:
#             print(f"📥 Trying download method {i+1}/{len(download_urls)}: {url[:80]}...")
            
#             response = requests.get(url, headers=headers, allow_redirects=True, timeout=30)
#             response.raise_for_status()
            
#             # Check content type and size
#             content_type = response.headers.get('content-type', '').lower()
#             content_length = len(response.content)
            
#             print(f"   Content-Type: {content_type}")
#             print(f"   Content-Length: {content_length} bytes")
            
#             # Skip if it's clearly HTML (too small or wrong content type)
#             if content_length < 5000 or 'html' in content_type:
#                 print(f"   ⚠️ Looks like HTML page, not Excel file")
#                 continue
            
#             # Try to read as Excel
#             try:
#                 df = pd.read_excel(BytesIO(response.content), engine='openpyxl', parse_dates=[DATE_COL_INDEX])
#                 df.columns = df.columns.str.strip()
#                 print(f"✅ Successfully loaded data using method {i+1}")
                
#                 # Save backup copy
#                 with open("backup_feedback.xlsx", "wb") as f:
#                     f.write(response.content)
#                 print("💾 Saved backup copy as 'backup_feedback.xlsx'")
                
#                 return df
                
#             except Exception as excel_error:
#                 print(f"   ❌ Failed to parse as Excel: {excel_error}")
#                 continue
                
#         except Exception as e:
#             print(f"   ❌ Download failed: {e}")
#             continue
    
#     raise Exception("All SharePoint download methods failed - file might require authentication")

# def load_data_fallback(local_path: str = "Anonymous_Feedback_Form.xlsx") -> pd.DataFrame:
#     """Fallback method - load from local file"""
#     # Try multiple possible filenames
#     possible_files = [
#         local_path,
#         "backup_feedback.xlsx",
#         "Anonymous_Feedback_Form.xlsx",
#         "feedback.xlsx"
#     ]
    
#     for filename in possible_files:
#         if os.path.exists(filename):
#             try:
#                 df = pd.read_excel(filename, engine='openpyxl', parse_dates=[DATE_COL_INDEX])
#                 df.columns = df.columns.str.strip()
#                 print(f"✅ Successfully loaded data from local file: {filename}")
#                 return df
#             except Exception as e:
#                 print(f"❌ Failed to load {filename}: {e}")
#                 continue
    
#     raise FileNotFoundError(f"No valid Excel files found. Tried: {possible_files}")

# def debug_sharepoint_response():
#     """Debug function to see what SharePoint is actually returning"""
#     print("\n🔍 Debug: Testing SharePoint response...")
    
#     headers = {
#         'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'
#     }
    
#     try:
#         response = requests.get(SHAREPOINT_PUBLIC_URL, headers=headers, allow_redirects=True)
#         print(f"Status Code: {response.status_code}")
#         print(f"Content-Type: {response.headers.get('content-type')}")
#         print(f"Content-Length: {len(response.content)}")
#         print(f"First 200 chars: {response.text[:200]}")
        
#         # Save the response to see what we're getting
#         with open("sharepoint_response.html", "w", encoding="utf-8") as f:
#             f.write(response.text)
#         print("💾 Saved response as 'sharepoint_response.html' for inspection")
        
#     except Exception as e:
#         print(f"❌ Debug request failed: {e}")

# def split_periods(df: pd.DataFrame):
#     """Split data into recent and previous 2-week periods"""
#     from datetime import datetime, timezone
#     # Count Positive, Negative, and Total feedback responses
#     feedback_columns = ["What kind of feedback is it?", "What kind of feedback is it?1"]
#     positive_count = sum(
#         df[col].str.contains("Positive", case=False, na=False).sum()
#         for col in feedback_columns
#         if col in df.columns
#     )
#     negative_count = sum(
#         df[col].str.contains("Negative", case=False, na=False).sum()
#         for col in feedback_columns
#         if col in df.columns
#     )
#     total_count = positive_count + negative_count
#     today = datetime.now(timezone.utc).date()
#     two_wks_ago = today - timedelta(weeks=2)
#     four_wks_ago = today - timedelta(weeks=4)
#     date_col = df.columns[DATE_COL_INDEX]

#     recent = df[(df[date_col].dt.date >= two_wks_ago) & (df[date_col].dt.date <= today)]
#     previous = df[(df[date_col].dt.date >= four_wks_ago) & (df[date_col].dt.date < two_wks_ago)]
#     return recent, previous, four_wks_ago, two_wks_ago, today, positive_count, negative_count, total_count

# def summarize(df: pd.DataFrame):
#     """Summarize sentiment counts from feedback type columns"""
#     total = len(df)
    
#     # Handle both main and secondary feedback columns
#     feedback_cols = ["What kind of feedback is it?", "What kind of feedback is it?1"]
    
#     positive = 0
#     negative = 0
    
#     for col in feedback_cols:
#         if col in df.columns:
#             positive += df[col].str.contains("Positive", case=False, na=False).sum()
#             negative += df[col].str.contains("Negative", case=False, na=False).sum()
    
#     return {"total": int(total), "positive": int(positive), "negative": int(negative)}

# def extract_feedback_data(df: pd.DataFrame):
#     """Extract all feedback data for AI analysis"""
#     feedback_data = []
    
#     for idx, row in df.iterrows():
#         record = {
#             "id": row.get("Id", idx),
#             "business_vertical": row.get("Which business vertical you belong to?", ""),
#             "function": row.get("What function does your feedback relate to", ""),
#             "enterprise_function": row.get("What enterprise function does your feedback relate to?", ""),
#             "feedback_type": row.get("What kind of feedback is it?", ""),
#             "issue_description": row.get("Please describe the issue", ""),
#             "feedback_description": row.get("Please describe your feedback", ""),
#             "financial_implications": row.get("Does it have any financial implications? (Please provide your remarks in the 'other' option)", ""),
#             "unethical_behavior": row.get("Is there any unethical or fraudulent behavior? (Please provide your remarks in the 'other' option)", ""),
#             "shared_with_manager": row.get("Have you shared the feedback with your reporting manager? (Please provide your remarks in the 'other' option)", ""),
#             # Secondary feedback columns
#             "function_1": row.get("What function does your feedback relate to1", ""),
#             "enterprise_function_1": row.get("What enterprise function does your feedback relate to?1", ""),
#             "feedback_type_1": row.get("What kind of feedback is it?1", ""),
#             "issue_description_1": row.get("Please describe the issue1", ""),
#             "feedback_description_1": row.get("Please describe your feedback1", ""),
#             "financial_implications_1": row.get("Does it have any financial implications? (Please provide your remarks in the 'other' option)1", ""),
#             "unethical_behavior_1": row.get("Is there any unethical or fraudulent behavior? (Please provide your remarks in the 'other' option)1", ""),
#             "shared_with_manager_1": row.get("Have you shared the feedback with your reporting manager? (Please provide your remarks in the 'other' option)1", ""),
#         }
#         feedback_data.append(record)
    
#     return feedback_data

# def generate_ai_content(recent_feedback, previous_feedback, recent_sum, prev_sum):
#     """Generate Executive Summary and Critical Flags using GPT"""
    
#     # Calculate key metrics for context
#     recent_total = recent_sum["total"]
#     recent_pos = recent_sum["positive"]
#     recent_neg = recent_sum["negative"]
    
#     prev_total = prev_sum["total"]
#     prev_pos = prev_sum["positive"]
#     prev_neg = prev_sum["negative"]
    
#     # Calculate percentages and changes
#     recent_pos_pct = (recent_pos / recent_total * 100) if recent_total > 0 else 0
#     recent_neg_pct = (recent_neg / recent_total * 100) if recent_total > 0 else 0
#     prev_pos_pct = (prev_pos / prev_total * 100) if prev_total > 0 else 0
#     prev_neg_pct = (prev_neg / prev_total * 100) if prev_total > 0 else 0
    
#     pos_change = recent_pos_pct - prev_pos_pct
#     neg_change = recent_neg_pct - prev_neg_pct
    
#     # Prepare feedback summaries for AI - improved extraction
#     recent_positive = []
#     recent_negative = []
    
#     for feedback in recent_feedback:
#         # Check all possible feedback type columns
#         feedback_types = [
#             str(feedback.get('feedback_type', '')).lower(),
#             str(feedback.get('feedback_type_1', '')).lower()
#         ]
        
#         # Get all possible text content
#         text_content = []
#         for field in ['feedback_description', 'feedback_description_1', 'issue_description', 'issue_description_1']:
#             if feedback.get(field) and str(feedback[field]).strip() and str(feedback[field]).strip().lower() != 'nan':
#                 text_content.append(str(feedback[field]).strip())
        
#         # Categorize based on feedback type
#         is_positive = any('positive' in ft for ft in feedback_types)
#         is_negative = any('negative' in ft for ft in feedback_types)
        
#         if is_positive and text_content:
#             recent_positive.append({
#                 'id': feedback.get('id', 'N/A'),
#                 'content': ' | '.join(text_content)
#             })
#         elif is_negative and text_content:
#             recent_negative.append({
#                 'id': feedback.get('id', 'N/A'),
#                 'content': ' | '.join(text_content)
#             })
    
#     # Compile themes for the prompt
#     positive_themes = [f"ID {f['id']}: {f['content'][:500]}" for f in recent_positive]
#     negative_themes = [f"ID {f['id']}: {f['content'][:800]}" for f in recent_negative]
    
#     prompt = f"""You are analyzing employee feedback for a CEO report. Generate two sections with SPECIFIC, ACTIONABLE insights:

# **SECTION 1: EXECUTIVE SUMMARY**
# Analyze the feedback data and create exactly 3 bullet points in this format:
# - **Strengths:** [Specific positive themes: mention what employees like - team culture, processes, leadership qualities, etc.]
# - **Concerns:** [Specific issues mentioned: leadership gaps, system problems, cultural issues, operational concerns, etc.]
# - **Overall Recommendation:** [Concrete actions based on the data - what specific steps should leadership take]

# **SECTION 2: CRITICAL FLAGS**
# For each negative feedback, create a detailed bullet point (200-250 chars) that lists the KEY SPECIFIC ISSUES mentioned:
# - [#ID] [List 2-3 specific concerns: e.g., "MIS transparency lacking, leadership visibility issues, revenue pressure affecting client focus"]

# REQUIREMENTS:
# - Be SPECIFIC, not vague
# - Mention concrete issues like: MIS systems, leadership visibility, cultural problems, operational gaps, etc.
# - For strengths: mention specific positive elements (onboarding, team support, communication style, etc.)
# - For concerns: list the actual problems mentioned in feedback
# - For flags: include multiple specific issues per feedback item when available

# **DATA ANALYSIS:**

# Recent Period Metrics:
# - Total responses: {recent_total}
# - Positive: {recent_pos} ({recent_pos_pct:.1f}%)
# - Negative: {recent_neg} ({recent_neg_pct:.1f}%)

# Changes from Previous Period:
# - Positive feedback: {pos_change:+.1f}% change
# - Negative feedback: {neg_change:+.1f}% change

# **POSITIVE FEEDBACK THEMES:**
# {chr(10).join([f"- {theme}" for theme in positive_themes]) if positive_themes else "No positive feedback in recent period"}

# **NEGATIVE FEEDBACK THEMES:**
# {chr(10).join([f"- {theme}" for theme in negative_themes]) if negative_themes else "No negative feedback in recent period"}

# **REQUIREMENTS:**
# 1. Be SPECIFIC and DETAILED - avoid vague language
# 2. For Executive Summary: mention concrete elements from feedback
# 3. For Critical Flags: include 2-4 specific issues per feedback item (200-250 characters each)
# 4. Use professional, executive-level language with actionable insights
# 5. Return ONLY the formatted bullet points as requested
# 6. For critical flags, include the actual ID numbers from the data

# **EXAMPLES OF GOOD SPECIFICITY:**
# - Strengths: "New hire onboarding process praised, strong peer collaboration noted, leadership accessibility appreciated"
# - Concerns: "MIS reporting transparency lacking, leadership visibility insufficient, revenue pressure compromising client advisory quality"
# - Flags: "[#135] MIS transparency gaps, top-down leadership limiting dialogue, revenue pressure affecting client-centric practices, operational benchmarking absent"

# **OUTPUT FORMAT:**
# EXECUTIVE_SUMMARY:
# <p><strong>Strengths:</strong> [specific content]</p>
# <p><strong>Concerns:</strong> [specific content]</p>
# <p><strong>Overall Recommendation:</strong> [specific content]</p>

# CRITICAL_FLAGS:
# <li>[#ID] [specific detailed content]</li>
# <li>[#ID] [specific detailed content]</li>
# (or)
# <li>No critical concerns flagged in the recent period</li>"""

#     try:
#         response = client.chat.completions.create(
#             model="gpt-4o-mini",
#             messages=[
#                 {"role": "user", "content": prompt}
#             ],
#             temperature=0.7,
#             max_tokens=1000
#         )
        
#         ai_response = response.choices[0].message.content.strip()
        
#         # Parse the response
#         if "EXECUTIVE_SUMMARY:" in ai_response and "CRITICAL_FLAGS:" in ai_response:
#             parts = ai_response.split("CRITICAL_FLAGS:")
#             executive_summary = parts[0].replace("EXECUTIVE_SUMMARY:", "").strip()
#             critical_flags = parts[1].strip()
#         else:
#             # Fallback parsing
#             lines = ai_response.split('\n')
#             exec_lines = []
#             flag_lines = []
#             in_flags = False
            
#             for line in lines:
#                 if 'critical' in line.lower() and 'flag' in line.lower():
#                     in_flags = True
#                     continue
#                 if in_flags:
#                     if line.strip().startswith('<li>'):
#                         flag_lines.append(line.strip())
#                 else:
#                     if line.strip().startswith('<p>'):
#                         exec_lines.append(line.strip())
            
#             executive_summary = '\n      '.join(exec_lines) if exec_lines else "<p><strong>Strengths:</strong> Positive employee engagement observed.</p>\n      <p><strong>Concerns:</strong> Some areas need attention based on feedback.</p>\n      <p><strong>Overall Recommendation:</strong> Continue monitoring feedback trends and address concerns promptly.</p>"
#             critical_flags = '\n        '.join(flag_lines) if flag_lines else "<li>No critical concerns flagged in the recent period</li>"
        
#         return executive_summary, critical_flags
        
#     except Exception as e:
#         print(f"⚠️ AI generation failed: {e}")
#         # Fallback content
#         executive_summary = """<p><strong>Strengths:</strong> Positive employee engagement and team collaboration noted in recent feedback.</p>
#       <p><strong>Concerns:</strong> Some operational and communication areas identified for improvement.</p>
#       <p><strong>Overall Recommendation:</strong> Continue monitoring feedback trends and implement targeted improvements.</p>"""
#         critical_flags = "<li>No critical concerns flagged in the recent period</li>"
#         return executive_summary, critical_flags

# def generate_html_report(recent_sum, prev_sum, recent_feedback, previous_feedback, dates, positive_count, negative_count, total_count):
#     """Generate HTML report using the exact template format with AI-generated content"""
#     start_prev, start_last, end_last = dates[0], dates[1], dates[2]
    
#     # Calculate percentages and changes
#     recent_total = recent_sum["total"]
#     recent_pos = recent_sum["positive"]
#     recent_neg = recent_sum["negative"]
    
#     prev_total = prev_sum["total"]
#     prev_pos = prev_sum["positive"]
#     prev_neg = prev_sum["negative"]
    
#     # Calculate percentages
#     recent_pos_pct = (recent_pos / recent_total * 100) if recent_total > 0 else 0
#     recent_neg_pct = (recent_neg / recent_total * 100) if recent_total > 0 else 0
    
#     prev_pos_pct = (prev_pos / prev_total * 100) if prev_total > 0 else 0
#     prev_neg_pct = (prev_neg / prev_total * 100) if prev_total > 0 else 0
    
#     # Calculate changes
#     pos_change = recent_pos_pct - prev_pos_pct
#     neg_change = recent_neg_pct - prev_neg_pct
    
#     # Format date strings
#     prev_start_str = start_prev.strftime("%-d %b %y")
#     prev_end_str = start_last.strftime("%-d %b %y") 
#     curr_start_str = start_last.strftime("%-d %b %y")
#     curr_end_str = end_last.strftime("%-d %b %y")
    
#     # Generate current timestamp
#     current_timestamp = datetime.now().strftime('%Y-%m-%d %H:%M')
#     current_date = datetime.now().strftime('%Y-%m-%d')
    
#     # Generate AI content
#     print("🤖 Generating AI content for Executive Summary and Critical Flags...")
#     executive_summary, critical_flags = generate_ai_content(recent_feedback, previous_feedback, recent_sum, prev_sum)
    
#     # Create the HTML using the exact template
#     html_content = f'''<!DOCTYPE html>
# <html lang="en">
# <head>
#   <meta charset="UTF-8" />
#   <meta name="viewport" content="width=device-width, initial-scale=1.0" />
#   <meta name="color-scheme" content="only light" />
#   <meta name="supported-color-schemes" content="light" />
#   <title>CEO Feedback Dashboard</title>
#   <style>
#     body {{
#       margin: 0;
#       padding: 0;
#       font-family: Arial, sans-serif;
#       background: #ffffff;
#       color: #000000;
#     }}

#     .container {{
#       max-width: 1000px;
#       margin: 0 auto;
#       padding: 20px;
#       box-sizing: border-box;
#     }}

#     header {{
#       display: flex;
#       justify-content: space-between;
#       align-items: center;
#       flex-wrap: wrap;
#       gap: 10px;
#       padding-bottom: 20px;
#       border-bottom: 1px solid #dddddd;
#     }}

#     header h1 {{
#       margin: 0;
#       font-size: 1.8em;
#       flex: 1 1 auto;
#       color: #000000;
#     }}

#     header .meta {{
#       font-size: 0.9em;
#       color: #555555;
#       text-align: right;
#       flex: 1 1 100%;
#     }}

#     .flex-container {{
#       display: flex;
#       gap: 20px;
#       flex-wrap: wrap;
#     }}

#     .card {{
#       background: #ffffff;
#       border-radius: 4px;
#       box-shadow: 0 2px 4px rgba(0, 0, 0, 0.1);
#       padding: 20px;
#       margin-bottom: 20px;
#       flex: 1 1 300px;
#       border: 1px solid #dddddd;
#       box-sizing: border-box;
#     }}

#     .period-card {{
#       flex: 1 1 300px;
#     }}

#     .card h3 {{
#       margin-top: 0;
#       color: #222222;
#       font-size: 1.2em;
#     }}

#     .card p {{
#       margin: 0.5em 0;
#       line-height: 1.5;
#     }}

#     .card ul {{
#       margin: 0.5em 0 1em 1.2em;
#     }}

#     .table-wrapper {{
#       width: 100%;
#       overflow-x: auto;
#     }}

#     table {{
#       width: 100%;
#       border-collapse: collapse;
#     }}

#     th, td {{
#       border: 1px solid #dddddd;
#       padding: 8px;
#       font-size: 0.9em;
#       word-break: normal;
#       white-space: normal;
#     }}

#     th {{
#       background: #f5f5f5;
#       text-align: left;
#     }}

#     td {{
#       text-align: right;
#     }}

#     td:first-child {{
#       text-align: left;
#     }}

#     .up {{
#       color: #000000;
#       font-weight: bold;
#     }}

#     .down {{
#       color: #000000;
#       font-weight: bold;
#     }}

#     footer {{
#       text-align: right;
#       font-size: 0.85em;
#       color: #777777;
#       border-top: 1px solid #dddddd;
#       padding-top: 20px;
#       margin-top: 20px;
#     }}

#     /* Responsive adjustments */
#    @media (max-width: 600px) {{
#   header {{
#     flex-direction: column;
#     align-items: flex-start;
#   }}

#   header .meta {{
#     text-align: left;
#   }}

#   .card, .period-card {{
#     flex: 1 1 100%;
#     max-width: none;
#   }}

#   .table-wrapper table th,
#   .table-wrapper table td {{
#     font-size: 4px !important;
#     padding: 4px;
#     white-space: normal !important;
#     word-break: break-word;
#   }}

#   header h1 {{
#     font-size: 1.4em;
#   }}
# }}

#   </style>
# </head>
# <body>
#   <div class="container">
#     <header>
#       <h1>Employee Feedback Summary</h1>
#       <div class="meta">Date: {current_timestamp}</div>
#     </header>

#     <div class="card">
#       <h3>Executive Summary</h3>
#       {executive_summary}
#     </div>

#     <div class="flex-container">
#       <div class="card period-card">
#         <h3>Period Comparison</h3>
#         <p><strong>Current Period:</strong><br>{curr_start_str} → {curr_end_str} </p>
#         <p><strong>Previous Period:</strong><br>{prev_start_str} → {prev_end_str} </p>
#       </div>

#       <div class="card">
#         <h3>Feedback Metrics</h3>
#         <div class="table-wrapper">
#           <table width="100%" cellpadding="0" cellspacing="0" border="0" style="width: 100%; border-collapse: collapse;">
#   <tr>
#     <th align="left" style="padding: 4px; font-size: 10px; border: 1px solid #dddddd;">Metric</th>
#     <th align="right" style="padding: 4px; font-size: 10px; border: 1px solid #dddddd;">Prev Count</th>
#     <th align="right" style="padding: 4px; font-size: 10px; border: 1px solid #dddddd;">Curr Count</th>
#     <th align="right" style="padding: 4px; font-size: 10px; border: 1px solid #dddddd;">Change</th>
#     <th align="right" style="padding: 4px; font-size: 10px; border: 1px solid #dddddd;">Overall</th>
#   </tr>
#   <tr>
#     <td style="padding: 4px; font-size: 10px; border: 1px solid #dddddd;">Positive</td>
#     <td align="right" style="padding: 4px; font-size: 10px; border: 1px solid #dddddd;">{prev_pos}</td>
#     <td align="right" style="padding: 4px; font-size: 10px; border: 1px solid #dddddd;">{recent_pos}</td>
#     <td align="right" style="padding: 4px; font-size: 10px; border: 1px solid #dddddd;">{'↑' if pos_change >= 0 else '↓'} {abs(pos_change):.1f}%</td>
#     <td align="right" style="padding: 4px; font-size: 10px; border: 1px solid #dddddd; font-weight: bold;">{positive_count}</td>
#   </tr>
#   <tr>
#     <td style="padding: 4px; font-size: 10px; border: 1px solid #dddddd;">Negative</td>
#     <td align="right" style="padding: 4px; font-size: 10px; border: 1px solid #dddddd;">{prev_neg}</td>
#     <td align="right" style="padding: 4px; font-size: 10px; border: 1px solid #dddddd;">{recent_neg}</td>
#     <td align="right" style="padding: 4px; font-size: 10px; border: 1px solid #dddddd;">{'↑' if neg_change >= 0 else '↓'} {abs(neg_change):.1f}%</td>
#     <td align="right" style="padding: 4px; font-size: 10px; border: 1px solid #dddddd; font-weight: bold;">{negative_count}</td>
#   </tr>
#   <tr>
#     <td style="padding: 4px; font-size: 10px; border: 1px solid #dddddd;">Total</td>
#     <td align="right" style="padding: 4px; font-size: 10px; border: 1px solid #dddddd;">{prev_total}</td>
#     <td align="right" style="padding: 4px; font-size: 10px; border: 1px solid #dddddd;">{recent_total}</td>
#     <td align="right" style="padding: 4px; font-size: 10px; border: 1px solid #dddddd;">- -</td>
#     <td align="right" style="padding: 4px; font-size: 10px; border: 1px solid #dddddd; font-weight: bold;">{total_count}</td>
#   </tr>
# </table>

#         </div>
#       </div>
#     </div>

#     <div class="card">
#       <h3>Critical Flags</h3>
#       <ul>
#         {critical_flags}
#       </ul>
#     </div>

#     <footer>
#     Report generated on {current_date}
#     <br>
#     Link to the response sheet: 
#     <a href="https://neogroupinfotech-my.sharepoint.com/:x:/g/personal/aashna_mehta_neo-group_in/ERmfzY7yPatEo3fni7EeZ8oBhjxYxuv17V6bcYTEfwlV7w?e=NDFhd4" target="_blank">
#         https://neogroupinfotech-my.sharepoint.com/...
#     </a>
#         <br>
#     Link to the feedback form: 
#     <a href="https://forms.office.com/Pages/ResponsePage.aspx?id=GVfz00I_UEW1Z0QhyDyoe7C8NMx2eRpKh0Iiqq6ujNtUMlAyMzZJS00zVk9IWFpROUxCNFc2U1AzMi4u" target="_blank">
#         https://forms.office.com/Page...
#     </a>
#     </footer>

#   </div>
# </body>
# </html>'''
    
#     return html_content
# # ——— Main ———
# if __name__ == "__main__":
#     # 0. Debug SharePoint response
#     debug_sharepoint_response()
    
#     # 1. Load data (with fallback)
#     try:
#         print("\n🔗 Loading data from SharePoint public link...")
#         df = load_data_from_public_sharepoint()
#     except Exception as e:
#         print(f"\n⚠️ Public download failed: {e}")
#         print("💾 Trying local fallback…")
#         try:
#             df = load_data_fallback()
#         except Exception as e2:
#             print(f"\n❌ Both methods failed.")
#             print(f"   SharePoint: {e}")
#             print(f"   Local:    {e2}")
#             exit(1)
    
#     # 2. Split into periods AND get overall counts
#     recent_df, prev_df, start_prev, start_last, end_last, \
#         positive_count, negative_count, total_count = split_periods(df)
    
#     # 3. Summarize sentiment
#     sum_recent = summarize(recent_df)
#     sum_prev   = summarize(prev_df)
    
#     # 4. Extract feedback for AI
#     recent_feedback   = extract_feedback_data(recent_df)
#     previous_feedback = extract_feedback_data(prev_df)
    
#     # 5. Generate the HTML report
#     html = generate_html_report(
#         sum_recent,
#         sum_prev,
#         recent_feedback,
#         previous_feedback,
#         (start_prev, start_last, end_last),
#         positive_count,
#         negative_count,
#         total_count
#     )
import os
import pandas as pd
from datetime import datetime, timedelta
from openai import OpenAI  # Updated import
import requests
from io import BytesIO
from dotenv import load_dotenv
import calendar

# Load environment variables from .env file
load_dotenv()

# ——— Configuration ———
# SharePoint public link 
SHAREPOINT_PUBLIC_URL = "https://neogroupinfotech-my.sharepoint.com/:x:/g/personal/aashna_mehta_neo-group_in/ERmfzY7yPatEo3fni7EeZ8oBhjxYxuv17V6bcYTEfwlV7w?e=NDFhd4"

# Try multiple SharePoint download URL formats
def get_sharepoint_download_urls(sharepoint_url):
    """Generate different SharePoint download URL formats to try"""
    urls = []
    
    # Method 1: Replace :x: with :u: and add download=1
    url1 = sharepoint_url.replace(":x:", ":u:").replace("?e=", "&download=1&e=")
    urls.append(url1)
    
    # Method 2: Use direct download parameter
    url2 = sharepoint_url + "&download=1"
    urls.append(url2)
    
    # Method 3: Extract file ID and use download format
    if "ERmfzY7yPatEo3fni7EeZ8oBhjxYxuv17V6bcYTEfwlV7w" in sharepoint_url:
        file_id = "ERmfzY7yPatEo3fni7EeZ8oBhjxYxuv17V6bcYTEfwlV7w"
        base_url = "https://neogroupinfotech-my.sharepoint.com/personal/aashna_mehta_neo-group_in/_layouts/15/download.aspx"
        url3 = f"{base_url}?share={file_id}"
        urls.append(url3)
    
    # Method 4: Try the embed format
    url4 = sharepoint_url.replace(":x:", ":u:")
    urls.append(url4)
    
    return urls

OUTPUT_HTML    = "report.html"
CEO_NAME       = "Nitin Jain"
SYSTEM_NAME    = "No-Reply Feedback System"
DATE_COL_INDEX = 1  # zero-based index of the timestamp column (Start time)
FEEDBACK_TYPE_COL = "What kind of feedback is it?"  # Column name for feedback type
FEEDBACK_TEXT_COL = "Please describe your feedback"  # Column name for feedback text
OPENAI_API_KEY = os.getenv("OPENAI_API_KEY") or "sk-your-actual-api-key-here"  # Fallback for testing

# Initialize OpenAI client
client = OpenAI(api_key=OPENAI_API_KEY)

# ——— Helpers ———

def load_data_from_public_sharepoint() -> pd.DataFrame:
    """Load Excel data from public SharePoint link"""
    download_urls = get_sharepoint_download_urls(SHAREPOINT_PUBLIC_URL)
    
    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'
    }
    
    for i, url in enumerate(download_urls):
        try:
            print(f"📥 Trying download method {i+1}/{len(download_urls)}: {url[:80]}...")
            
            response = requests.get(url, headers=headers, allow_redirects=True, timeout=30)
            response.raise_for_status()
            
            # Check content type and size
            content_type = response.headers.get('content-type', '').lower()
            content_length = len(response.content)
            
            print(f"   Content-Type: {content_type}")
            print(f"   Content-Length: {content_length} bytes")
            
            # Skip if it's clearly HTML (too small or wrong content type)
            if content_length < 5000 or 'html' in content_type:
                print(f"   ⚠️ Looks like HTML page, not Excel file")
                continue
            
            # Try to read as Excel
            try:
                df = pd.read_excel(BytesIO(response.content), engine='openpyxl', parse_dates=[DATE_COL_INDEX])
                df.columns = df.columns.str.strip()
                print(f"✅ Successfully loaded data using method {i+1}")
                
                # Save backup copy
                with open("backup_feedback.xlsx", "wb") as f:
                    f.write(response.content)
                print("💾 Saved backup copy as 'backup_feedback.xlsx'")
                
                return df
                
            except Exception as excel_error:
                print(f"   ❌ Failed to parse as Excel: {excel_error}")
                continue
                
        except Exception as e:
            print(f"   ❌ Download failed: {e}")
            continue
    
    raise Exception("All SharePoint download methods failed - file might require authentication")

def load_data_fallback(local_path: str = "Anonymous_Feedback_Form.xlsx") -> pd.DataFrame:
    """Fallback method - load from local file"""
    # Try multiple possible filenames
    possible_files = [
        local_path,
        "backup_feedback.xlsx",
        "Anonymous_Feedback_Form.xlsx",
        "feedback.xlsx"
    ]
    
    for filename in possible_files:
        if os.path.exists(filename):
            try:
                df = pd.read_excel(filename, engine='openpyxl', parse_dates=[DATE_COL_INDEX])
                df.columns = df.columns.str.strip()
                print(f"✅ Successfully loaded data from local file: {filename}")
                return df
            except Exception as e:
                print(f"❌ Failed to load {filename}: {e}")
                continue
    
    raise FileNotFoundError(f"No valid Excel files found. Tried: {possible_files}")

def debug_sharepoint_response():
    """Debug function to see what SharePoint is actually returning"""
    print("\n🔍 Debug: Testing SharePoint response...")
    
    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'
    }
    
    try:
        response = requests.get(SHAREPOINT_PUBLIC_URL, headers=headers, allow_redirects=True)
        print(f"Status Code: {response.status_code}")
        print(f"Content-Type: {response.headers.get('content-type')}")
        print(f"Content-Length: {len(response.content)}")
        print(f"First 200 chars: {response.text[:200]}")
        
        # Save the response to see what we're getting
        with open("sharepoint_response.html", "w", encoding="utf-8") as f:
            f.write(response.text)
        print("💾 Saved response as 'sharepoint_response.html' for inspection")
        
    except Exception as e:
        print(f"❌ Debug request failed: {e}")

def split_periods(df: pd.DataFrame):
    """Split data into recent month and previous month periods"""
    from datetime import datetime, timezone
    
    # Count Positive, Negative, and Total feedback responses
    feedback_columns = ["What kind of feedback is it?", "What kind of feedback is it?1"]
    positive_count = sum(
        df[col].str.contains("Positive", case=False, na=False).sum()
        for col in feedback_columns
        if col in df.columns
    )
    negative_count = sum(
        df[col].str.contains("Negative", case=False, na=False).sum()
        for col in feedback_columns
        if col in df.columns
    )
    total_count = positive_count + negative_count
    
    # Get current date and calculate monthly periods
    today = datetime.now(timezone.utc).date()
    
    # Get the first day of current month
    current_month_start = today.replace(day=1)
    
    # Get the first day of previous month
    if current_month_start.month == 1:
        prev_month_start = current_month_start.replace(year=current_month_start.year - 1, month=12)
    else:
        prev_month_start = current_month_start.replace(month=current_month_start.month - 1)
    
    # Get the last day of previous month
    prev_month_end = current_month_start - timedelta(days=1)
    
    # Get date column
    date_col = df.columns[DATE_COL_INDEX]

    # Filter data for recent month (current month)
    recent = df[(df[date_col].dt.date >= current_month_start) & (df[date_col].dt.date <= today)]
    
    # Filter data for previous month
    previous = df[(df[date_col].dt.date >= prev_month_start) & (df[date_col].dt.date <= prev_month_end)]
    
    print(f"📅 Current month period: {current_month_start} to {today}")
    print(f"📅 Previous month period: {prev_month_start} to {prev_month_end}")
    print(f"📊 Recent month data: {len(recent)} records")
    print(f"📊 Previous month data: {len(previous)} records")
    
    return recent, previous, prev_month_start, current_month_start, today, positive_count, negative_count, total_count

def summarize(df: pd.DataFrame):
    """Summarize sentiment counts from feedback type columns"""
    total = len(df)
    
    # Handle both main and secondary feedback columns
    feedback_cols = ["What kind of feedback is it?", "What kind of feedback is it?1"]
    
    positive = 0
    negative = 0
    
    for col in feedback_cols:
        if col in df.columns:
            positive += df[col].str.contains("Positive", case=False, na=False).sum()
            negative += df[col].str.contains("Negative", case=False, na=False).sum()
    
    return {"total": int(total), "positive": int(positive), "negative": int(negative)}

def extract_feedback_data(df: pd.DataFrame):
    """Extract all feedback data for AI analysis"""
    feedback_data = []
    
    for idx, row in df.iterrows():
        record = {
            "id": row.get("Id", idx),
            "business_vertical": row.get("Which business vertical you belong to?", ""),
            "function": row.get("What function does your feedback relate to", ""),
            "enterprise_function": row.get("What enterprise function does your feedback relate to?", ""),
            "feedback_type": row.get("What kind of feedback is it?", ""),
            "issue_description": row.get("Please describe the issue", ""),
            "feedback_description": row.get("Please describe your feedback", ""),
            "financial_implications": row.get("Does it have any financial implications? (Please provide your remarks in the 'other' option)", ""),
            "unethical_behavior": row.get("Is there any unethical or fraudulent behavior? (Please provide your remarks in the 'other' option)", ""),
            "shared_with_manager": row.get("Have you shared the feedback with your reporting manager? (Please provide your remarks in the 'other' option)", ""),
            # Secondary feedback columns
            "function_1": row.get("What function does your feedback relate to1", ""),
            "enterprise_function_1": row.get("What enterprise function does your feedback relate to?1", ""),
            "feedback_type_1": row.get("What kind of feedback is it?1", ""),
            "issue_description_1": row.get("Please describe the issue1", ""),
            "feedback_description_1": row.get("Please describe your feedback1", ""),
            "financial_implications_1": row.get("Does it have any financial implications? (Please provide your remarks in the 'other' option)1", ""),
            "unethical_behavior_1": row.get("Is there any unethical or fraudulent behavior? (Please provide your remarks in the 'other' option)1", ""),
            "shared_with_manager_1": row.get("Have you shared the feedback with your reporting manager? (Please provide your remarks in the 'other' option)1", ""),
        }
        feedback_data.append(record)
    
    return feedback_data

def generate_ai_content(recent_feedback, previous_feedback, recent_sum, prev_sum):
    """Generate Executive Summary and Critical Flags using GPT"""
    
    # Calculate key metrics for context
    recent_total = recent_sum["total"]
    recent_pos = recent_sum["positive"]
    recent_neg = recent_sum["negative"]
    
    prev_total = prev_sum["total"]
    prev_pos = prev_sum["positive"]
    prev_neg = prev_sum["negative"]
    
    # Handle case where there are no responses
    if recent_total == 0 and prev_total == 0:
        executive_summary = """<p><strong>Strengths:</strong> No feedback data available for analysis in both current and previous periods.</p>
      <p><strong>Concerns:</strong> No responses received in the specified time frame - consider reviewing feedback collection process.</p>
      <p><strong>Overall Recommendation:</strong> Implement targeted communication to encourage employee participation in feedback collection.</p>"""
        critical_flags = "<li>No responses received in the current time frame</li>"
        return executive_summary, critical_flags
    
    if recent_total == 0:
        executive_summary = """<p><strong>Strengths:</strong> No current month feedback data available for analysis.</p>
      <p><strong>Concerns:</strong> No responses received in the current month - consider reviewing feedback collection process.</p>
      <p><strong>Overall Recommendation:</strong> Investigate reasons for lack of feedback and implement measures to encourage participation.</p>"""
        critical_flags = "<li>No responses received in the current month</li>"
        return executive_summary, critical_flags
    
    # Calculate percentages and changes
    recent_pos_pct = (recent_pos / recent_total * 100) if recent_total > 0 else 0
    recent_neg_pct = (recent_neg / recent_total * 100) if recent_total > 0 else 0
    prev_pos_pct = (prev_pos / prev_total * 100) if prev_total > 0 else 0
    prev_neg_pct = (prev_neg / prev_total * 100) if prev_total > 0 else 0
    
    pos_change = recent_pos_pct - prev_pos_pct if prev_total > 0 else 0
    neg_change = recent_neg_pct - prev_neg_pct if prev_total > 0 else 0
    # Calculate cumulative totals for the two periods only

    # Prepare feedback summaries for AI - improved extraction
    recent_positive = []
    recent_negative = []
    
    for feedback in recent_feedback:
        # Check all possible feedback type columns
        feedback_types = [
            str(feedback.get('feedback_type', '')).lower(),
            str(feedback.get('feedback_type_1', '')).lower()
        ]
        
        # Get all possible text content
        text_content = []
        for field in ['feedback_description', 'feedback_description_1', 'issue_description', 'issue_description_1']:
            if feedback.get(field) and str(feedback[field]).strip() and str(feedback[field]).strip().lower() != 'nan':
                text_content.append(str(feedback[field]).strip())
        
        # Categorize based on feedback type
        is_positive = any('positive' in ft for ft in feedback_types)
        is_negative = any('negative' in ft for ft in feedback_types)
        
        if is_positive and text_content:
            recent_positive.append({
                'id': feedback.get('id', 'N/A'),
                'content': ' | '.join(text_content)
            })
        elif is_negative and text_content:
            recent_negative.append({
                'id': feedback.get('id', 'N/A'),
                'content': ' | '.join(text_content)
            })
    
    # Compile themes for the prompt
    positive_themes = [f"ID {f['id']}: {f['content'][:500]}" for f in recent_positive]
    negative_themes = [f"ID {f['id']}: {f['content'][:800]}" for f in recent_negative]
    
    prompt = f"""You are analyzing monthly employee feedback for a CEO report. Generate two sections with SPECIFIC, ACTIONABLE insights:

**SECTION 1: EXECUTIVE SUMMARY**
Analyze the feedback data and create exactly 3 bullet points in this format:
- **Strengths:** [Specific positive themes: mention what employees like - team culture, processes, leadership qualities, etc.]
- **Concerns:** [Specific issues mentioned: leadership gaps, system problems, cultural issues, operational concerns, etc.]
- **Overall Recommendation:** [Concrete actions based on the data - what specific steps should leadership take]

**SECTION 2: CRITICAL FLAGS**
For each negative feedback, create a detailed bullet point (200-250 chars) that lists the KEY SPECIFIC ISSUES mentioned:
- [#ID] [List 2-3 specific concerns: e.g., "MIS transparency lacking, leadership visibility issues, revenue pressure affecting client focus"]

REQUIREMENTS:
- Be SPECIFIC, not vague
- Mention concrete issues like: MIS systems, leadership visibility, cultural problems, operational gaps, etc.
- For strengths: mention specific positive elements (onboarding, team support, communication style, etc.)
- For concerns: list the actual problems mentioned in feedback
- For flags: include multiple specific issues per feedback item when available

**DATA ANALYSIS:**

Recent Month Metrics:
- Total responses: {recent_total}
- Positive: {recent_pos} ({recent_pos_pct:.1f}%)
- Negative: {recent_neg} ({recent_neg_pct:.1f}%)

Changes from Previous Month:
- Positive feedback: {pos_change:+.1f}% change
- Negative feedback: {neg_change:+.1f}% change

**POSITIVE FEEDBACK THEMES:**
{chr(10).join([f"- {theme}" for theme in positive_themes]) if positive_themes else "No positive feedback in recent month"}

**NEGATIVE FEEDBACK THEMES:**
{chr(10).join([f"- {theme}" for theme in negative_themes]) if negative_themes else "No negative feedback in recent month"}

**REQUIREMENTS:**
1. Be SPECIFIC and DETAILED - avoid vague language
2. For Executive Summary: mention concrete elements from feedback
3. For Critical Flags: include 2-4 specific issues per feedback item (200-250 characters each)
4. Use professional, executive-level language with actionable insights
5. Return ONLY the formatted bullet points as requested
6. For critical flags, include the actual ID numbers from the data

**EXAMPLES OF GOOD SPECIFICITY:**
- Strengths: "New hire onboarding process praised, strong peer collaboration noted, leadership accessibility appreciated"
- Concerns: "MIS reporting transparency lacking, leadership visibility insufficient, revenue pressure compromising client advisory quality"
- Flags: "[#135] MIS transparency gaps, top-down leadership limiting dialogue, revenue pressure affecting client-centric practices, operational benchmarking absent"

**OUTPUT FORMAT:**
EXECUTIVE_SUMMARY:
<p><strong>Strengths:</strong> [specific content]</p>
<p><strong>Concerns:</strong> [specific content]</p>
<p><strong>Overall Recommendation:</strong> [specific content]</p>

CRITICAL_FLAGS:
<li>[#ID] [specific detailed content]</li>
<li>[#ID] [specific detailed content]</li>
(or)
<li>No critical concerns flagged in the recent month</li>"""

    try:
        response = client.chat.completions.create(
            model="gpt-4o-mini",
            messages=[
                {"role": "user", "content": prompt}
            ],
            temperature=0.7,
            max_tokens=1000
        )
        
        ai_response = response.choices[0].message.content.strip()
        
        # Parse the response
        if "EXECUTIVE_SUMMARY:" in ai_response and "CRITICAL_FLAGS:" in ai_response:
            parts = ai_response.split("CRITICAL_FLAGS:")
            executive_summary = parts[0].replace("EXECUTIVE_SUMMARY:", "").strip()
            critical_flags = parts[1].strip()
        else:
            # Fallback parsing
            lines = ai_response.split('\n')
            exec_lines = []
            flag_lines = []
            in_flags = False
            
            for line in lines:
                if 'critical' in line.lower() and 'flag' in line.lower():
                    in_flags = True
                    continue
                if in_flags:
                    if line.strip().startswith('<li>'):
                        flag_lines.append(line.strip())
                else:
                    if line.strip().startswith('<p>'):
                        exec_lines.append(line.strip())
            
            executive_summary = '\n      '.join(exec_lines) if exec_lines else "<p><strong>Strengths:</strong> Positive employee engagement observed.</p>\n      <p><strong>Concerns:</strong> Some areas need attention based on feedback.</p>\n      <p><strong>Overall Recommendation:</strong> Continue monitoring feedback trends and address concerns promptly.</p>"
            critical_flags = '\n        '.join(flag_lines) if flag_lines else "<li>No critical concerns flagged in the recent month</li>"
        
        return executive_summary, critical_flags
        
    except Exception as e:
        print(f"⚠️ AI generation failed: {e}")
        # Fallback content
        executive_summary = """<p><strong>Strengths:</strong> Positive employee engagement and team collaboration noted in recent feedback.</p>
      <p><strong>Concerns:</strong> Some operational and communication areas identified for improvement.</p>
      <p><strong>Overall Recommendation:</strong> Continue monitoring feedback trends and implement targeted improvements.</p>"""
        critical_flags = "<li>No critical concerns flagged in the recent month</li>"
        return executive_summary, critical_flags

def generate_html_report(recent_sum, prev_sum, recent_feedback, previous_feedback, dates, positive_count, negative_count, total_count):
    """Generate HTML report using the exact template format with AI-generated content"""
    start_prev, start_last, end_last = dates[0], dates[1], dates[2]
    
    # Calculate percentages and changes
    recent_total = recent_sum["total"]
    recent_pos = recent_sum["positive"]
    recent_neg = recent_sum["negative"]
    
    prev_total = prev_sum["total"]
    prev_pos = prev_sum["positive"]
    prev_neg = prev_sum["negative"]
    
    # Calculate percentages
    recent_pos_pct = (recent_pos / recent_total * 100) if recent_total > 0 else 0
    recent_neg_pct = (recent_neg / recent_total * 100) if recent_total > 0 else 0
    
    prev_pos_pct = (prev_pos / prev_total * 100) if prev_total > 0 else 0
    prev_neg_pct = (prev_neg / prev_total * 100) if prev_total > 0 else 0
    
    # Calculate changes
    pos_change = recent_pos_pct - prev_pos_pct if prev_total > 0 else 0
    neg_change = recent_neg_pct - prev_neg_pct if prev_total > 0 else 0
    cumulative_positive = prev_pos + recent_pos
    cumulative_negative = prev_neg + recent_neg
    cumulative_total = prev_total + recent_total
    
    # Format date strings for monthly periods
    prev_start_str = start_prev.strftime("%b %Y")
    curr_start_str = start_last.strftime("%b %Y")
    
    # Generate current timestamp
    current_timestamp = datetime.now().strftime('%Y-%m-%d %H:%M')
    current_date = datetime.now().strftime('%Y-%m-%d')
    
    # Generate AI content
    print("🤖 Generating AI content for Executive Summary and Critical Flags...")
    executive_summary, critical_flags = generate_ai_content(recent_feedback, previous_feedback, recent_sum, prev_sum)
    
    # Create the HTML using the exact template
    html_content = f'''<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1.0" />
  <meta name="color-scheme" content="only light" />
  <meta name="supported-color-schemes" content="light" />
  <title>CEO Feedback Dashboard</title>
  <style>
    body {{
      margin: 0;
      padding: 0;
      font-family: Arial, sans-serif;
      background: #ffffff;
      color: #000000;
    }}

    .container {{
      max-width: 1000px;
      margin: 0 auto;
      padding: 20px;
      box-sizing: border-box;
    }}

    header {{
      display: flex;
      justify-content: space-between;
      align-items: center;
      flex-wrap: wrap;
      gap: 10px;
      padding-bottom: 20px;
      border-bottom: 1px solid #dddddd;
    }}

    header h1 {{
      margin: 0;
      font-size: 1.8em;
      flex: 1 1 auto;
      color: #000000;
    }}

    header .meta {{
      font-size: 0.9em;
      color: #555555;
      text-align: right;
      flex: 1 1 100%;
    }}

    .flex-container {{
      display: flex;
      gap: 20px;
      flex-wrap: wrap;
    }}

    .card {{
      background: #ffffff;
      border-radius: 4px;
      box-shadow: 0 2px 4px rgba(0, 0, 0, 0.1);
      padding: 20px;
      margin-bottom: 20px;
      flex: 1 1 300px;
      border: 1px solid #dddddd;
      box-sizing: border-box;
    }}

    .period-card {{
      flex: 1 1 300px;
    }}

    .card h3 {{
      margin-top: 0;
      color: #222222;
      font-size: 1.2em;
    }}

    .card p {{
      margin: 0.5em 0;
      line-height: 1.5;
    }}

    .card ul {{
      margin: 0.5em 0 1em 1.2em;
    }}

    .table-wrapper {{
      width: 100%;
      overflow-x: auto;
    }}

    table {{
      width: 100%;
      border-collapse: collapse;
    }}

    th, td {{
      border: 1px solid #dddddd;
      padding: 8px;
      font-size: 0.9em;
      word-break: normal;
      white-space: normal;
    }}

    th {{
      background: #f5f5f5;
      text-align: left;
    }}

    td {{
      text-align: right;
    }}

    td:first-child {{
      text-align: left;
    }}

    .up {{
      color: #000000;
      font-weight: bold;
    }}

    .down {{
      color: #000000;
      font-weight: bold;
    }}

    footer {{
      text-align: right;
      font-size: 0.85em;
      color: #777777;
      border-top: 1px solid #dddddd;
      padding-top: 20px;
      margin-top: 20px;
    }}

    /* Responsive adjustments */
   @media (max-width: 600px) {{
  header {{
    flex-direction: column;
    align-items: flex-start;
  }}

  header .meta {{
    text-align: left;
  }}

  .card, .period-card {{
    flex: 1 1 100%;
    max-width: none;
  }}

  .table-wrapper table th,
  .table-wrapper table td {{
    font-size: 4px !important;
    padding: 4px;
    white-space: normal !important;
    word-break: break-word;
  }}

  header h1 {{
    font-size: 1.4em;
  }}
}}

  </style>
</head>
<body>
  <div class="container">
    <header>
      <h1>Employee Feedback Summary</h1>
      <div class="meta">Date: {current_timestamp}</div>
    </header>

    <div class="card">
      <h3>Executive Summary</h3>
      {executive_summary}
    </div>

    <div class="flex-container">
      <div class="card period-card">
        <h3>Period Comparison</h3>
        <p><strong>Current Month:</strong><br>{curr_start_str}</p>
        <p><strong>Previous Month:</strong><br>{prev_start_str}</p>
      </div>

      <div class="card">
        <h3>Feedback Metrics</h3>
        <div class="table-wrapper">
          <table width="100%" cellpadding="0" cellspacing="0" border="0" style="width: 100%; border-collapse: collapse;">
  <tr>
    <th align="left" style="padding: 4px; font-size: 10px; border: 1px solid #dddddd;">Metric</th>
    <th align="right" style="padding: 4px; font-size: 10px; border: 1px solid #dddddd;">Prev Count</th>
    <th align="right" style="padding: 4px; font-size: 10px; border: 1px solid #dddddd;">Curr Count</th>
    <th align="right" style="padding: 4px; font-size: 10px; border: 1px solid #dddddd;">Change</th>
    <th align="right" style="padding: 4px; font-size: 10px; border: 1px solid #dddddd;">Overall</th>
  </tr>
  <tr>
    <td style="padding: 4px; font-size: 10px; border: 1px solid #dddddd;">Positive</td>
    <td align="right" style="padding: 4px; font-size: 10px; border: 1px solid #dddddd;">{prev_pos}</td>
    <td align="right" style="padding: 4px; font-size: 10px; border: 1px solid #dddddd;">{recent_pos}</td>
    <td align="right" style="padding: 4px; font-size: 10px; border: 1px solid #dddddd;">{('↑' if pos_change >= 0 else '↓') + ' ' + f'{abs(pos_change):.1f}%' if prev_total > 0 else 'N/A'}</td>
    <td align="right" style="padding: 4px; font-size: 10px; border: 1px solid #dddddd; font-weight: bold;">{cumulative_positive}</td>
  </tr>
  <tr>
    <td style="padding: 4px; font-size: 10px; border: 1px solid #dddddd;">Negative</td>
    <td align="right" style="padding: 4px; font-size: 10px; border: 1px solid #dddddd;">{prev_neg}</td>
    <td align="right" style="padding: 4px; font-size: 10px; border: 1px solid #dddddd;">{recent_neg}</td>
    <td align="right" style="padding: 4px; font-size: 10px; border: 1px solid #dddddd;">{('↑' if neg_change >= 0 else '↓') + ' ' + f'{abs(neg_change):.1f}%' if prev_total > 0 else 'N/A'}</td>
    <td align="right" style="padding: 4px; font-size: 10px; border: 1px solid #dddddd; font-weight: bold;">{cumulative_negative}</td>
  </tr>
  <tr>
    <td style="padding: 4px; font-size: 10px; border: 1px solid #dddddd;">Total</td>
    <td align="right" style="padding: 4px; font-size: 10px; border: 1px solid #dddddd;">{prev_total}</td>
    <td align="right" style="padding: 4px; font-size: 10px; border: 1px solid #dddddd;">{recent_total}</td>
    <td align="right" style="padding: 4px; font-size: 10px; border: 1px solid #dddddd;">- -</td>
    <td align="right" style="padding: 4px; font-size: 10px; border: 1px solid #dddddd; font-weight: bold;">{cumulative_total}</td>
  </tr>
</table>

        </div>
      </div>
    </div>

    <div class="card">
      <h3>Critical Flags</h3>
      <ul>
        {critical_flags}
      </ul>
    </div>

    <footer>
    Report generated on {current_date}
    <br>
    Link to the response sheet: 
    <a href="https://neogroupinfotech-my.sharepoint.com/:x:/g/personal/aashna_mehta_neo-group_in/ERmfzY7yPatEo3fni7EeZ8oBhjxYxuv17V6bcYTEfwlV7w?e=NDFhd4" target="_blank">
        https://neogroupinfotech-my.sharepoint.com/...
    </a>
        <br>
    Link to the feedback form: 
    <a href="https://forms.office.com/Pages/ResponsePage.aspx?id=GVfz00I_UEW1Z0QhyDyoe7C8NMx2eRpKh0Iiqq6ujNtUMlAyMzZJS00zVk9IWFpROUxCNFc2U1AzMi4u" target="_blank">
        https://forms.office.com/Page...
    </a>
    </footer>

  </div>
</body>
</html>'''
    
    return html_content

# ——— Main ———
if __name__ == "__main__":
    # 0. Debug SharePoint response
    debug_sharepoint_response()
    
    # 1. Load data (with fallback)
    try:
        print("\n🔗 Loading data from SharePoint public link...")
        df = load_data_from_public_sharepoint()
    except Exception as e:
        print(f"\n⚠️ Public download failed: {e}")
        print("💾 Trying local fallback…")
        try:
            df = load_data_fallback()
        except Exception as e2:
            print(f"\n❌ Both methods failed.")
            print(f"   SharePoint: {e}")
            print(f"   Local:    {e2}")
            exit(1)
    
    # 2. Split into monthly periods AND get overall counts
    recent_df, prev_df, start_prev, start_last, end_last, \
        positive_count, negative_count, total_count = split_periods(df)
    
    # 3. Summarize sentiment
    sum_recent = summarize(recent_df)
    sum_prev   = summarize(prev_df)
    
    # 4. Extract feedback for AI
    recent_feedback   = extract_feedback_data(recent_df)
    previous_feedback = extract_feedback_data(prev_df)
    
    # 5. Generate the HTML report
    html = generate_html_report(
        sum_recent,
        sum_prev,
        recent_feedback,
        previous_feedback,
        (start_prev, start_last, end_last),
        positive_count,
        negative_count,
        total_count
    )
    # 6. Write it out
    with open(OUTPUT_HTML, "w", encoding="utf-8") as f:
        f.write(html)
    print(f"✅ Report generated: {OUTPUT_HTML}")
