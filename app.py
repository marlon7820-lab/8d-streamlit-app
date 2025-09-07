# Collect answers into a stable list before any button press
data_rows = []
for step, *_ in npqp_steps:
    ans = st.session_state.answers.get(step, "")
    extra = st.session_state.extra_info.get(step, "")
    data_rows.append((step, ans, extra))

# Save button
if st.button("Save Full 8D Report"):
    if not any(ans for _, ans, _ in data_rows):
        st.error("⚠️ No answers filled in yet. Please complete some fields before saving.")
    else:
        # --- XLSX file ---
        from openpyxl import Workbook
        wb = Workbook()
        ws = wb.active
        ws.title = "NPQP 8D Report"

        row = 1
        ws[f"A{row}"] = "NPQP 8D Report"
        row += 2
        ws[f"A{row}"] = "Report Date"
        ws[f"B{row}"] = str(st.session_state.report_date)
        row += 1
        ws[f"A{row}"] = "Prepared By"
        ws[f"B{row}"] = st.session_state.prepared_by
        row += 2

        ws[f"A{row}"] = "Step"
        ws[f"B{row}"] = "Answer"
        ws[f"C{row}"] = "Extra Info"
        row += 1

        for step, ans, extra in data_rows:
            ws[f"A{row}"] = step
            ws[f"B{row}"] = ans
            ws[f"C{row}"] = str(extra)
            row += 1

        xlsx_file = "NPQP_8D_Report.xlsx"
        wb.save(xlsx_file)

        # --- CSV file ---
        import csv
        csv_file = "NPQP_8D_Report.csv"
        with open(csv_file, "w", newline="") as f:
            writer = csv.writer(f)
            writer.writerow(["NPQP 8D Report"])
            writer.writerow(["Report Date", st.session_state.report_date])
            writer.writerow(["Prepared By", st.session_state.prepared_by])
            writer.writerow([])
            writer.writerow(["Step", "Answer", "Extra Info"])
            writer.writerows(data_rows)

        # --- Downloads ---
        st.success("✅ NPQP 8D Report saved successfully.")
        with open(xlsx_file, "rb") as f:
            st.download_button("📥 Download XLSX (desktop)", f, file_name=xlsx_file)
        with open(csv_file, "rb") as f:
            st.download_button("📥 Download CSV (iPhone-friendly)", f, file_name=csv_file)
