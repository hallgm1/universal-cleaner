import streamlit as st
import re
import os
from datetime import datetime
from io import BytesIO
from docx import Document

st.set_page_config(
    page_title="Letter & Contract Adapter",
    page_icon="📝",
    layout="wide"
)

st.title("📝 Tool 2: Formal Letter & Contract Adapter")
st.write("Upload raw text or draft Word documents. This engine automatically sanitizes syntax, scales greetings, normalizes grammar rules, and exports formal assets.")

def parse_and_reformat_document(raw_bytes, ext, selected_style):
    if ext == '.docx':
        stream = BytesIO(raw_bytes)
        doc = Document(stream)
        raw_paras = [p.text.strip() for p in doc.paragraphs if p.text.strip()]
    else:
        raw_text = raw_bytes.decode("utf-8", errors="ignore")
        raw_paras = [p.strip() for p in raw_text.split('\n') if p.strip()]
            
    cleaned_doc = Document()
    full_text_block = "\n".join(raw_paras)
    
    # Intelligently scan structure parameters to flag layout intent
    is_letter = any(k in full_text_block.lower() for k in ["to:", "dear", "subject:", "re:", "attention"])
    
    # Structural Top Anchor Injection
    if is_letter and selected_style in ["Business", "Formal"]:
        cleaned_doc.add_paragraph("[SENDER CONTACT DETAILS]\n[Postal Address Line 1]\nDar es Salaam, Tanzania\n")
        current_date = datetime(2026, 5, 22).strftime('%B %d, %Y')
        cleaned_doc.add_paragraph(f"Date: {current_date}\n")
    else:
        h_para = cleaned_doc.add_heading(level=1)
        h_run = h_para.add_run(f"REFORMATTED ARTIFACT - STYLE: {selected_style.upper()}")
        h_run.bold = True
        
    greeting_injected = False
    
    for txt in raw_paras:
        # Standardize matching spaces out of row components
        txt = re.sub(r'\s+', ' ', txt)
        
        if selected_style in ["Business", "Formal"]:
            if txt.lower().startswith("date:"): 
                continue
            if txt.lower().startswith("to:"):
                txt = re.sub(r"\bto:\s*", "TO:\n", txt, flags=re.I)
                txt = txt.title().replace("Nssf", "NSSF").replace("Nhif", "NHIF")
                cleaned_doc.add_paragraph(txt)
                continue
                
            # Normalize casual greetings into executive anchors
            contains_greeting = any(k in txt.lower() for k in ["hey there", "hi team", "hey", "hi", "dear sir", "dear madam"])
            if contains_greeting:
                if not greeting_injected:
                    cleaned_doc.add_paragraph("Dear Sir/Madam,")
                    greeting_injected = True
                txt = re.sub(r"\bhey\s+there,?\s*|\bhi\s+team,?\s*|\bhey,?\s*|\bhi,?\s*|\bdear\s+sir/madam,?\s*", "", txt, flags=re.I)
                if not txt.strip(): 
                    continue

            # Complete Grammar Pass (Expand verbal shortcuts and contractions)
            txt = re.sub(r"\bi'm\b", "I am", txt, flags=re.I)
            txt = re.sub(r"\bcan't\b", "cannot", txt, flags=re.I)
            txt = re.sub(r"\bdon't\b", "do not", txt, flags=re.I)
            txt = re.sub(r"\basap\b", "as soon as possible", txt, flags=re.I)
            txt = re.sub(r"\bhaven't\b", "have not", txt, flags=re.I)
            txt = re.sub(r"\bwon't\b", "will not", txt, flags=re.I)
            txt = re.sub(r"\bask about\b", "inquire regarding", txt, flags=re.I)
            txt = re.sub(r"\bi've\b", "I have", txt, flags=re.I)
            txt = re.sub(r"\bi\b", "I", txt)
            txt = re.sub(r"\b(i\s)", "I ", txt)
            
            # Sentence Capitalization Sweep
            sentences = txt.split('.')
            processed_sentences = []
            for s in sentences:
                s_strip = s.strip()
                if len(s_strip) > 0: 
                    processed_sentences.append(s_strip[0].upper() + s_strip[1:])
            txt = ". ".join(processed_sentences)
            if len(txt) > 0 and not txt.endswith('.'): 
                txt += '.'
                
            txt = re.sub(r',+', ',', txt)
            txt = txt.replace("Nssf", "NSSF").replace("Nhif", "NHIF")
            
        elif selected_style == "Non-Formal":
            txt = re.sub(r"\butilize\b", "use", txt, flags=re.I)
            txt = re.sub(r"\bsubsequent to\b", "after", txt, flags=re.I)
            
        cleaned_doc.add_paragraph(txt)
        
    # Structural Bottom Signature Injection
    if is_letter and selected_style in ["Business", "Formal"]:
        cleaned_doc.add_paragraph("\nYours faithfully,\n\n\n_______________________\n[Insert Full Account Name]\nAuthorized Signatory / Claimant")
        
    out = BytesIO()
    cleaned_doc.save(out)
    out.seek(0)
    return out.getvalue()

st.sidebar.header("⚙️ Style Configurations")
doc_style = st.sidebar.selectbox("Target Communication Tone", ["Business", "Formal", "Non-Formal"])

uploaded_file = st.file_uploader("Upload document file", type=["docx", "txt"])

if uploaded_file is not None:
    filename = str(uploaded_file.name)
    ext = os.path.splitext(filename.lower())[1]
    raw_file_content = uploaded_file.getvalue()

    try:
        with st.spinner("Processing document restructuring algorithms..."):
            formatted_docx_bytes = parse_and_reformat_document(raw_file_content, ext, doc_style)
            
        st.success(f"🚀 Document structure converted successfully to matching {doc_style.upper()} framework layout profiles.")
        
        st.download_button(
            label=f"📥 Download Formatted {doc_style} Document (.docx)",
            data=formatted_docx_bytes,
            file_name=f"formatted_{doc_style.lower()}_document.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True
        )
        
    except Exception as e:
        st.error(f"Document Formatting System Fault: {str(e)}")
