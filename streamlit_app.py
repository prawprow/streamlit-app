import streamlit as st
import pandas as pd
import re
import io

st.set_page_config(page_title="TXT to Excel Converter", page_icon="📄")
st.title("📄 แปลงไฟล์ .txt เป็น Excel")

uploaded_file = st.file_uploader("📤 อัปโหลดไฟล์ .txt", type="txt")

if uploaded_file is not None:
    raw_text = uploaded_file.read().decode("utf-8", errors="ignore")
    raw_lines = [line.strip() for line in raw_text.splitlines() if line.strip()]

    start_index = next((i for i, line in enumerate(raw_lines) if re.match(r'\d{6,7}', line)), 0)
    data_lines = raw_lines[start_index:]

    entry_groups = []
    current_group = []
    entry_no = None
    for line in data_lines:
        if re.match(r'\d{6,7}', line):
            if current_group:
                entry_groups.append((entry_no, current_group))
            entry_no = line.strip().split()[0]
            current_group = [line]
        elif current_group:
            current_group.append(line)
    if current_group:
        entry_groups.append((entry_no, current_group))

    all_rows = []

    for entry_index, (entry_no, group) in enumerate(entry_groups):
        group_text = "\n".join(group)

        base_row = {}
        match_ref = re.search(r'([A-Z]\d{3})-(?!D)(\d+)', group_text)
        import_ref = match_ref.group(1) + match_ref.group(2) if match_ref else ""
        base_row["เลขที่ใบขนเข้า"] = import_ref

        match_item = re.search(r'[A-Z]\d{3}-\d+\s+(-\d{4})', group_text)
        item_number = str(int(match_item.group(1).replace("-", ""))) if match_item else ""
        base_row["รายการเข้า"] = item_number

        match_entry = re.search(r'^(\d{6,7})', group_text)
        if match_entry:
            base_row["เลขชำระ"] = str(int(match_entry.group(1).lstrip("0")))

        match_date = re.search(r'\b(\d{2})/(\d{2})/(\d{2})\b', group_text)
        if match_date:
            base_row["วันชำระ"] = f"{int(match_date.group(1))}/{int(match_date.group(2))}/23"

        match_import = re.search(r'\((\d{2})/(\d{2})/(\d{2}),(\d{2})/(\d{2})/(\d{2})\)', group_text)
        if match_import:
            base_row["วันนำเข้า"] = f"{int(match_import.group(1))}/{int(match_import.group(2))}/23"
            base_row["วันdelivery"] = f"{int(match_import.group(4))}/{int(match_import.group(5))}/23"

        unit_price = ""
        duty_price = ""
        for line in group:
            m = re.search(r'[A-Z]\d{3}-\d+\s+-\d{4}.*?(\d{1,3}(?:,\d{3})*\.\d+)\s+(\d{1,3}(?:,\d{3})*\.\d+)', line)
            if m:
                unit_price = m.group(1).replace(",", "")
                duty_price = m.group(2).replace(",", "")
                break
        base_row["ราคาต่อหน่วย"] = unit_price
        base_row["อากร.ต่อหน่วย"] = duty_price

        match_material_code = re.search(r'\d{6,7}\s+\d{2}/\d{2}/\d{2}\s+(\d+)', group_text)
        if match_material_code:
            code = match_material_code.group(1).lstrip("0")
            base_row["รหัสวัตถุดิบ"] = code
            material_name = ""
            for line in group:
                match = re.match(rf"^\d{{6,7}}\s+\d{{2}}/\d{{2}}/\d{{2}}\s+0*{code}\b\s+(.*)", line)
                if match:
                    after_code = match.group(1)
                    material_name = re.split(r"\s{2,}", after_code)[0].strip()
                    break
            base_row["ชื่อวัตถุดิบ"] = material_name

        qty_match = re.search(r'(\d{1,3}(?:,\d{3})*\.\d{3})\s+\d{1,3}(?:,\d{3})*\.\d{2}', group_text)
        if qty_match:
            base_row["ปริมาณนำเข้า"] = qty_match.group(1)

        duty = ""
        for line in group:
            if re.match(r'\d{6,7}', line):
                matches = re.findall(r'\d{1,3}(?:,\d{3})*\.\d{2}', line)
                if matches:
                    duty = matches[-1]
                break
        base_row["อากรที่ชำระ"] = duty

        base_row.update({
            "เลขที่ใบขนออก": "",
            "รายการออก": "",
            "วันผ่านพิธีการ": "",
            "วันload": "",
            "วันตรวจปล่อย": "",
            "หน่วยวัตถุดิบ": "",
            "ปริมาณที่มาตัด": "",
            "เป็นอากร": "",
            "สถานะยกไป": "NO MOVEMENT",
            "_entry_index": entry_index,
            "_suborder": 0
        })
        all_rows.append(base_row)

        suborder = 1
        for line in group:
            match = re.search(
                r'(\d{2}/\d{2}/\d{2})\s+([A-Z]\d{3}-D\d+)\s+(-\d{4})\s+(\d{2}/\d{2}/\d{2})\s+(\d{2}/\d{2}/\d{2})\s+(\d+)\s+([\d,]+\.\d{3})\s+([\d,]+\.\d{2})',
                line)
            if match:
                export_row = base_row.copy()
                export_row.update({
                    "เลขที่ใบขนออก": match.group(2).replace("-", ""),
                    "รายการออก": str(int(match.group(3).replace("-", ""))),
                    "วันผ่านพิธีการ": f"{int(match.group(1)[:2])}/{int(match.group(1)[3:5])}/24",
                    "วันload": f"{int(match.group(4)[:2])}/{int(match.group(4)[3:5])}/24",
                    "วันตรวจปล่อย": f"{int(match.group(5)[:2])}/{int(match.group(5)[3:5])}/24",
                    "หน่วยวัตถุดิบ": match.group(6),
                    "ปริมาณที่มาตัด": str(int(float(match.group(7).replace(",", "")))),
                    "เป็นอากร": match.group(8),
                    "สถานะยกไป": "C/F",
                    "_suborder": suborder
                })
                suborder += 1
                all_rows.append(export_row)

    df_combined = pd.DataFrame(all_rows)
    df_combined = df_combined.sort_values(by=["_entry_index", "_suborder"]).drop(columns=["_entry_index", "_suborder"])

    has_export_keys = df_combined[df_combined["เลขที่ใบขนออก"] != ""][["เลขที่ใบขนเข้า", "รายการเข้า"]].drop_duplicates()
    mask_cleaned = ~(
        (df_combined["เลขที่ใบขนออก"] == "") &
        (df_combined[["เลขที่ใบขนเข้า", "รายการเข้า"]].apply(tuple, axis=1).isin(
            has_export_keys.apply(tuple, axis=1)
        ))
    )
    df_cleaned_final = df_combined[mask_cleaned]

    st.success("✅ ประมวลผลสำเร็จแล้ว")
    st.dataframe(df_cleaned_final)

    @st.cache_data
    def convert_df(df):
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df.to_excel(writer, index=False)
        return output.getvalue()

    st.download_button(
        label="📅 ดาวน์โหลดเป็น Excel",
        data=convert_df(df_cleaned_final),
        file_name="result.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
