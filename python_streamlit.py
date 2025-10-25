import os
import pandas as pd
import streamlit as st

#os.add_dll_directory(r'C:\Users\YashwanthS\AppData\Local\Programs\Python\Python313\Lib\site-packages\clidriver\bin')
import ibm_db
print("ibm_db imported successfully")
# Replace with your DB2 connection details
dsn_hostname = "data.dal-cnv-prod.core.cirrus.ibm.com"
dsn_uid = "yashsince2001"
dsn_pwd = "Winter2024"
dsn_port = "31350"
dsn_database = "ITACPRD"
dsn_protocol = "TCPIP"

# Create the DSN connection string
dsn = (
    f"DATABASE={dsn_database};"
    f"HOSTNAME={dsn_hostname};"
    f"PORT={dsn_port};"
    f"PROTOCOL={dsn_protocol};"
    f"UID={dsn_uid};"
    f"PWD={dsn_pwd};"
)
try:
    conn = ibm_db.connect(dsn, "", "")
    print("Connection successful!")

    sql = "SELECT * FROM ITACTRNS.UI_DOCUMENT_STATUS"
    stmt = ibm_db.exec_immediate(conn, sql)

    # Collect rows into a list of dictionaries
    rows = []
    row = ibm_db.fetch_assoc(stmt)
    while row:
        rows.append(row)
        row = ibm_db.fetch_assoc(stmt)

    # Convert to DataFrame
    df = pd.DataFrame(rows)


    template_names = sorted(df["TEMPLATE_NAME"].dropna().unique())
    selected_template = st.selectbox("Select Template Name", template_names)

    # Filtered data
    filtered_df = df[df["TEMPLATE_NAME"] == selected_template]

    st.write(f"### Results for TEMPLATE_NAME = `{selected_template}`")
    st.dataframe(filtered_df, use_container_width=True)

    # --- Export to Excel ---
    export_filename = f"Filtered_{selected_template}.xlsx"
    if st.button("Export Filtered Data to Excel"):
        filtered_df.to_excel(export_filename, index=False)
        st.success(f"Exported to {export_filename} successfully!")

except Exception as e:
    st.error(f"Unable to connect: {e}")