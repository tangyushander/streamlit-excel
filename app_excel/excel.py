import pandas as pd
from openpyxl import load_workbook
import io
import streamlit as st  # 导入 Streamlit


# Function to process Excel file
def process_excel(input_file):
    # Read the Excel file
    xls = pd.ExcelFile(input_file)
    all_blocks = []

    # Extract rows 17-22 from each sheet and add the sheet name in the first column
    for sheet in xls.sheet_names:
        df = pd.read_excel(input_file, sheet_name=sheet, header=0)
        if df.shape[0] >= 22:  # Ensure there are enough rows to extract
            rows_to_extract = list(range(15, 21))  # Row 17 to Row 22
            extracted_block = df.iloc[rows_to_extract].copy()

            # If '类别' column already exists, drop it before adding it again
            if '类别' in extracted_block.columns:
                extracted_block.drop(columns=['类别'], inplace=True)

            # Insert sheet name in the first column (类别)
            extracted_block.insert(0, '类别', sheet)
            extracted_block.reset_index(drop=True, inplace=True)
            all_blocks.append(extracted_block)

    # Concatenate all blocks
    final_result = pd.concat(all_blocks, axis=0, ignore_index=True)

    # Save to a new Excel file in a BytesIO object
    output_file = io.BytesIO()
    with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
        final_result.to_excel(writer, index=False, sheet_name="汇总")

    output_file.seek(0)

    # Step 2: Merge cells in the first column (same value rows)
    wb = load_workbook(output_file)
    ws = wb["汇总"]

    last_value = None
    start_row = 2  # Start from row 2 as row 1 is headers
    for row in range(start_row, ws.max_row + 1):
        current_value = ws.cell(row=row, column=1).value
        if current_value == last_value:
            ws.merge_cells(start_row=row - 1, start_column=1, end_row=row, end_column=1)
        else:
            last_value = current_value

    # Save the final file with merged cells to a BytesIO object
    final_output = io.BytesIO()
    wb.save(final_output)
    final_output.seek(0)

    return final_output


# Streamlit App UI
st.title("Excel 数据处理工具")

st.write(
    """
    这个应用帮助你上传Excel文件，并执行以下操作：
    1. 提取每个工作表的第17行到第22行。
    2. 在每个提取的数据中添加对应的工作表名称（在“类别”列中）。
    3. 合并相同“类别”列中的单元格。
    """
)

# File uploader
uploaded_file = st.file_uploader("上传你的Excel文件", type=["xlsx"])

if uploaded_file:
    # Process the file when uploaded
    st.write("处理中...")

    # Call the process_excel function
    output = process_excel(uploaded_file)

    # Provide the download button
    st.download_button(
        label="下载处理后的Excel文件",
        data=output,
        file_name="处理结果.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )



