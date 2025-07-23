import streamlit as st
import pandas as pd
import re
import io
import chardet

st.set_page_config(page_title="TXT to Excel Converter", page_icon="📄")
st.title("📄 แปลงไฟล์ .txt เป็น Excel (รองรับ format ผิดเพี้ยน)")

uploaded_file = st.file_uploader("📤 อัปโหลดไฟล์ .txt", type="txt")

if uploaded_file is not None:
    raw_bytes = uploaded_file.read()
    detected = chardet.detect(raw_bytes)
    encoding = detected['encoding'] or 'utf-8'
    raw_text = raw_bytes.decode(encoding, errors="ignore")

    raw_lines = [line.rstrip() for line in raw_text.splitlines() if line.strip()]

    # --- Group Entries ---
    entry_groups = []
    current_group = []
    entry_no = None

    for line in raw_lines:
        if re.search(r'\b\d{7}\b', line) and "PART FOR" not in line and "CARBURETOR" not in line:
            if current_group:
                entry_groups.append((entry_no, current_group))
            entry_no_match = re.search(r'\b\d{7}\b', line)
            entry_no = entry_no_match.group() if entry_no_match else f"UNK-{len(entry_groups)}"
            current_group = [line]
        elif current_group:
            current_group.append(line)
    if current_group:
        entry_groups.append((entry_no, current_group))

    st.info(f"📦 เจอทั้งหมด {len(entry_groups)} กลุ่ม")

    all_rows = []
    for entry_index, (entry_no, group) in enumerate(entry_groups):
        group_text = "\n".join(group)
        row = {"เลขชำระ": entry_no}

        m_ref = re.search(r'(A\d{3})-(\d+)', group_text)
        row["เลขที่ใบขนเข้า"] = m_ref.group(1) + m_ref.group(2) if m_ref else ""

        m_date = re.search(r'(\d{2}/\d{2}/\d{2})', group_text)
        row["วันชำระ"] = m_date.group(1) if m_date else ""

        m_import = re.search(r'\((\d{2}/\d{2}/\d{2}),(\d{2}/\d{2}/\d{2})\)', group_text)
        if m_import:
            row["วันนำเข้า"] = m_import.group(1)
            row["วันdelivery"] = m_import.group(2)

        for l in group:
            m_price = re.search(r'(\d{1,3}(?:,\d{3})*\.\d+)\s+(\d{1,3}(?:,\d{3})*\.\d+)', l)
            if m_price:
                row["ราคาต่อหน่วย"] = m_price.group(1).replace(",", "")
                row["อากร.ต่อหน่วย"] = m_price.group(2).replace(",", "")
                break

        m_qty = re.search(r'(\d{1,3}(?:,\d{3})*\.\d{3})', group_text)
        if m_qty:
            row["ปริมาณนำเข้า"] = m_qty.group(1).replace(",", "")

        duty = ""
        for l in group:
            if re.search(r'\b\d{7}\b', l):
                duties = re.findall(r'\d{1,3}(?:,\d{3})*\.\d{2}', l)
                if duties:
                    duty = duties[-1]
                break
        row["อากรที่ชำระ"] = duty

        m_code = re.search(r'\b\d{7}\s+\d{2}/\d{2}/\d{2}\s+(\d+)', group_text)
        if m_code:
            code = m_code.group(1).lstrip("0")
            row["รหัสวัตถุดิบ"] = code
            for l in group:
                m_name = re.match(rf"^\d{{7}}\s+\d{{2}}/\d{{2}}/\d{{2}}\s+0*{code}\b\s+(.*)", l)
                if m_name:
                    row["ชื่อวัตถุดิบ"] = re.split(r"\s{2,}", m_name.group(1))[0].strip()
                    break

        row.update({
            "เลขที่ใบขนออก": "",
            "รายการออก": "",
            "วันผ่านพิธีการ": "",
            "วันload": "",
            "วันตรวจปล่อย": "",
            "หน่วยวัตถุดิบ": "",
            "ปริมาณที่มาตัด": "",
            "เป็นอากร": "",
            "สถานะยกไป": "NO MOVEMENT",
        })

        all_rows.append(row)

    df = pd.DataFrame(all_rows)
    st.success("✅ ประมวลผลสำเร็จแล้ว")
    st.dataframe(df)

    @st.cache_data
    def convert_df(df):
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df.to_excel(writer, index=False)
        return output.getvalue()

    st.download_button(
        label="📥 ดาวน์โหลดเป็น Excel",
        data=convert_df(df),
        file_name="converted_result.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )    
