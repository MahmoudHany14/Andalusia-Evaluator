import streamlit as st
import pdfplumber
import pandas as pd
import google.generativeai as genai
import io
import json
from openpyxl.styles import PatternFill, Alignment, Font
from openpyxl.chart import BarChart, Reference

API_KEY = st.secrets["GEMINI_API_KEY"]
genai.configure(api_key=API_KEY)

def extract_full_text(file):
    text = ""
    with pdfplumber.open(file) as pdf:
        for page in pdf.pages:
            extracted = page.extract_text()
            if extracted:
                text += extracted + "\n"
    return text

st.set_page_config(page_title="Andalusia Tech Evaluator Ultimate", layout="wide")
st.title("Andalusia Hospitals - Pro Equipment Evaluator")

col1, col2 = st.columns(2)
with col1:
    std_file = st.file_uploader("Upload Standard Specs (PDF)", type="pdf")
with col2:
    offer_files = st.file_uploader("Upload Vendor Offers (PDF)", type="pdf", accept_multiple_files=True)

if st.button("🚀 Analyze & Generate Ultimate Excel"):
    if std_file and offer_files:
        with st.spinner("Analyzing documents, structuring data, and building the Excel report..."):
            
            std_text = extract_full_text(std_file)
            
            main_df = None
            company_names = []
            summaries = []

            for offer in offer_files:
                offer_text = extract_full_text(offer)
                
                prompt = f"""
                You are a Senior Biomedical Engineer. Compare the Vendor Offer against the Hospital Standard.
                Standard Specs: {std_text}
                Vendor Offer: {offer_text}
                
                Output STRICTLY a JSON object with this exact structure:
                {{
                  "Company_Name": "Extract the actual Vendor/Company name from the text",
                  "Final_Recommendation": "Accepted or Rejected",
                  "Overall_Score": <number out of 10>,
                  "Final_Reason": "English justification for the final decision",
                  "Specifications": [
                    {{
                      "Feature": "Parameter name",
                      "Standard": "Required value",
                      "Vendor_Value": "Offered value",
                      "Match": "Yes or No",
                      "Grade": <number out of 10>,
                      "Comment": "English technical comment"
                    }}
                  ]
                }}
                """
                
                model = genai.GenerativeModel(
                    'gemini-2.5-flash',
                    generation_config={"response_mime_type": "application/json"}
                )
                
                response = model.generate_content(prompt)
                
                try:
                    data = json.loads(response.text)
                    comp_name = data.get("Company_Name", "Unknown Vendor")
                    company_names.append(comp_name)
                    
                    summaries.append({
                        "Company": comp_name,
                        "Final Recommendation": data.get("Final_Recommendation", ""),
                        "Overall Score": data.get("Overall_Score", 0),
                        "Reason": data.get("Final_Reason", "")
                    })
                    
                    df = pd.DataFrame(data["Specifications"])
                    
                    cols = {
                        'Vendor_Value': f'{comp_name} Offer',
                        'Match': f'{comp_name} Match',
                        'Grade': f'{comp_name} Grade',
                        'Comment': f'{comp_name} Comment'
                    }
                    df.rename(columns=cols, inplace=True)
                    
                    if main_df is None:
                        main_df = df
                    else:
                        main_df = pd.merge(main_df, df, on=['Feature', 'Standard'], how='outer')
                        
                    st.success(f"✅ Analyzed: {comp_name}")
                    
                except Exception as e:
                    st.error(f"Error processing {offer.name}: {e}")

            if main_df is not None:
                
                ordered_cols = ['Feature', 'Standard']
                for i, comp in enumerate(company_names):
                    if i > 0:
                        sep_col = f'   _{i}' 
                        main_df[sep_col] = ""
                        ordered_cols.append(sep_col)
                    ordered_cols.extend([f'{comp} Offer', f'{comp} Match', f'{comp} Grade', f'{comp} Comment'])
                
                main_df = main_df[ordered_cols]

                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    main_df.to_excel(writer, index=False, sheet_name='Detailed Comparison')
                    workbook = writer.book
                    worksheet = writer.sheets['Detailed Comparison']
                    
                    green_fill = PatternFill(start_color='C6EFCE', end_color='C6EFCE', fill_type='solid')
                    green_font = Font(color='006100')
                    red_fill = PatternFill(start_color='FFC7CE', end_color='FFC7CE', fill_type='solid')
                    red_font = Font(color='9C0006')
                    gray_fill = PatternFill(start_color='E7E6E6', end_color='E7E6E6', fill_type='solid')
                    
                    for row in worksheet.iter_rows(min_row=1, max_col=worksheet.max_column, max_row=worksheet.max_row):
                        for cell in row:
                            col_letter = cell.column_letter
                            col_name = str(worksheet.cell(row=1, column=cell.column).value)
                            
                            if '   _' in col_name: 
                                worksheet.column_dimensions[col_letter].width = 4
                                cell.fill = gray_fill
                            else:
                                worksheet.column_dimensions[col_letter].width = 28
                                cell.alignment = Alignment(wrap_text=True, vertical='top')
                            
                            if cell.row == 1:
                                cell.font = Font(bold=True)
                                continue
                                
                            val = str(cell.value).strip().lower()
                            if val == 'yes':
                                cell.fill = green_fill
                                cell.font = green_font
                            elif val == 'no':
                                cell.fill = red_fill
                                cell.font = red_font
                    
                    df_summary = pd.DataFrame(summaries)
                    df_summary.to_excel(writer, index=False, sheet_name='Summary & Chart')
                    ws_summary = writer.sheets['Summary & Chart']
                    
                    for col in ws_summary.columns:
                        ws_summary.column_dimensions[col[0].column_letter].width = 30
                        for cell in col:
                            cell.alignment = Alignment(wrap_text=True, vertical='top')
                    
                    if len(summaries) > 0:
                        chart = BarChart()
                        chart.type = "col"
                        chart.style = 10
                        chart.title = "Vendors Overall Scores Comparison"
                        chart.y_axis.title = "Score (out of 10)"
                        chart.x_axis.title = "Company"
                        
                        data_ref = Reference(ws_summary, min_col=3, min_row=1, max_row=len(summaries)+1)
                        cats_ref = Reference(ws_summary, min_col=1, min_row=2, max_row=len(summaries)+1)
                        chart.add_data(data_ref, titles_from_data=True)
                        chart.set_categories(cats_ref)
                        chart.height = 10
                        chart.width = 15
                        ws_summary.add_chart(chart, "A6")

                dynamic_filename = "Compare_" + "_vs_".join(company_names) + ".xlsx"
                
                st.download_button(
                    label=f"📥 Download ({dynamic_filename})",
                    data=output.getvalue(),
                    file_name=dynamic_filename,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
    else:
        st.warning("Please upload the Standard PDF and at least one Offer PDF.")