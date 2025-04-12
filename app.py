import streamlit as st
from docxtpl import DocxTemplate
import io
import pandas as pd
from datetime import datetime
from api import get_data_gsheet
import tempfile
import pythoncom
from docx2pdf import convert
import os

# Initialize COM only if on Windows and not in Streamlit cloud
if os.name == 'nt' and not os.environ.get('STREAMLIT_SERVER'):
    pythoncom.CoInitialize()

# Load data
if "invoice" not in st.session_state:
    st.session_state.invoice = pd.DataFrame(get_data_gsheet("1kh5ZxsOeZjMSGIXxsBoX1dJPmvkwg8g_Amf0gk9OmWE", "Main Data List Invoice", "A:X"))
    
if "list_do" not in st.session_state:
    st.session_state.list_do = pd.DataFrame(get_data_gsheet("14j6jpMzMVUx_zWu9et0LWgwQ4hqUtPlO_9NUTziT2Yk", "List DO", "A:S"))

invoice = st.session_state.invoice
list_do = st.session_state.list_do



# Initialize session state for invoice items if not already present
if 'invoice_items' not in st.session_state:
    st.session_state.invoice_items = pd.DataFrame({
        'Trip': [''],
        'Description': [''],
        'License Plate': [''],
        'Shipping Date': [''],
        'Amount': ['']
    })

st.title("Generate Invoice KTI")

# Input fields for invoice details
col1, col2 = st.columns(2)
with col1:
    name = st.selectbox("Name", options=list(invoice['Customer'].unique()))
    invoice_number = st.selectbox("Invoice NTI/Customer", options=invoice[invoice['Customer'] == name]['Invoice NT/Cust'].unique())
with col2:
    invoice_date = st.date_input("Invoice Date")
    due_date = st.date_input("Due Date")

st.subheader("Invoice Items")

selected_do = st.multiselect("List DO", options=list_do[list_do['Invoice Name'] == name]['No SO'].unique())

# Editable table for invoice items
if selected_do:
    for_df = list_do[list_do['No SO'].isin(selected_do)]
    for_df['Trip'] = for_df['Origin'] + " - " + for_df['Destination']
    for_df['Description'] = ''
    for_df['Shipping Date'] = for_df['Date']
    for_df['Amount'] = for_df[' Price']

    for_df = for_df[['Trip', 'Description', 'License Plate', 'Shipping Date', 'Amount']]

    edited_df = st.data_editor(
        for_df,
        num_rows="dynamic",
        use_container_width=True,
        key="invoice_items_editor_do"
    )
else:
    edited_df = st.data_editor(
        st.session_state.invoice_items,
        num_rows="dynamic",
        use_container_width=True,
        key="invoice_items_editor"
    )

# Additional fields
payment_terms = st.text_area("Payment Terms")

# Generate Invoice button
if st.button("Generate Invoice"):
    # Calculate totals
    try:
        amounts = edited_df['Amount'].str.replace("Rp","").str.replace(" ","").str.replace(",","").apply(pd.to_numeric,errors='coerce').fillna(0).astype(int)
        total = amounts.sum()
        grand_total = total  # You can add tax or other calculations here if needed
    except:
        st.error("Please enter valid amounts for all items")
        st.stop()
    
    # Prepare context for template
    context = {
        'name': name,
        'invoice_number': invoice_number,
        'invoice_date': invoice_date.strftime("%d %B %Y"),
        'due_date': due_date.strftime("%d %B %Y"),
        'invoice_list': edited_df.values.tolist(),
        'total': f"Rp {total:,.0f}",
        'grand_total': f"Rp {grand_total:,.0f}",
        'payment_terms': payment_terms
    }
    
    try:
        # Load the template
        doc = DocxTemplate("template.docx")
        
        # Render the template
        doc.render(context)
        
        # Save to a temporary file
        with tempfile.NamedTemporaryFile(delete=False, suffix='.docx') as tmp:
            doc.save(tmp.name)
            tmp_path = tmp.name
        
        # Offer download
        with open(tmp_path, 'rb') as f:
            doc_bytes = f.read()
        
        st.download_button(
            label="Download Word Invoice",
            data=doc_bytes,
            file_name=f"Invoice_{invoice_number}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )

        # Clean up
        os.unlink(tmp_path)
        
    except Exception as e:
        st.error(f"Error generating invoice: {e}")