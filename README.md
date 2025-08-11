def post_process_bbv(dst: Path, input_sheet_name="Inputfile", bbv_sheet_name="BBV Vorlage"):
    wb = load_workbook(dst)

    if input_sheet_name not in wb.sheetnames:
        raise ValueError(f"Sheet '{input_sheet_name}' not found in {dst}")

    ws_in = wb[input_sheet_name]

    if bbv_sheet_name in wb.sheetnames:
        ws_bbv = wb[bbv_sheet_name]
        if ws_bbv.max_row >= 3:
            ws_bbv.delete_rows(3, ws_bbv.max_row - 2)
    else:
        ws_bbv = wb.create_sheet(title=bbv_sheet_name)

    last_row_in = ws_in.max_row

    IN_J = col_letter_to_index1("J")
    IN_B = col_letter_to_index1("B")
    IN_I = col_letter_to_index1("I")
    IN_C = col_letter_to_index1("C")
    IN_E = col_letter_to_index1("E")
    IN_K = col_letter_to_index1("K")

    OUT_B = col_letter_to_index1("B")
    OUT_C = col_letter_to_index1("C")
    OUT_G = col_letter_to_index1("G")
    OUT_K = col_letter_to_index1("K")
    OUT_L = col_letter_to_index1("L")
    OUT_Q = col_letter_to_index1("Q")

    OUT_A = col_letter_to_index1("A")
    OUT_D = col_letter_to_index1("D")
    OUT_H = col_letter_to_index1("H")
    OUT_J = col_letter_to_index1("J")
    OUT_N = col_letter_to_index1("N")
    OUT_E = col_letter_to_index1("E")  # New target for C→E copy

    out_row = 3
    for r in range(2, last_row_in + 1):
        v_J = ws_in.cell(row=r, column=IN_J).value
        v_B = ws_in.cell(row=r, column=IN_B).value
        v_I = ws_in.cell(row=r, column=IN_I).value
        v_C = ws_in.cell(row=r, column=IN_C).value
        v_E = ws_in.cell(row=r, column=IN_E).value
        v_K = ws_in.cell(row=r, column=IN_K).value

        ws_bbv.cell(row=out_row, column=OUT_B, value=v_J)
        ws_bbv.cell(row=out_row, column=OUT_C, value=v_B)
        ws_bbv.cell(row=out_row, column=OUT_G, value=v_I)
        ws_bbv.cell(row=out_row, column=OUT_K, value=v_C)
        ws_bbv.cell(row=out_row, column=OUT_L, value=v_E)
        ws_bbv.cell(row=out_row, column=OUT_Q, value=v_K)

        ws_bbv.cell(row=out_row, column=OUT_A, value=2098)
        ws_bbv.cell(row=out_row, column=OUT_D, value="SA")
        ws_bbv.cell(row=out_row, column=OUT_H, value="EUR")
        ws_bbv.cell(row=out_row, column=OUT_J, value="S")
        ws_bbv.cell(row=out_row, column=OUT_N, value="V7")

        out_row += 1

    # NEW: copy C → E instead of C → D
    for r in range(3, out_row):
        ws_bbv.cell(row=r, column=OUT_E, value=ws_bbv.cell(row=r, column=OUT_C).value)

    wb.save(dst)
