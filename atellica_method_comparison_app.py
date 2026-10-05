import io
import re
import copy
import fitz
import openpyxl
import streamlit as st

st.set_page_config(
    page_title="Atellica Method Comparison",
    layout="wide"
)

# ============================================================
# METHOD COMPARISON RULE
# ============================================================
# TWO PDFs are used:
#   OLD PDF = Atellica 2 / Old Instrument
#   NEW PDF = Atellica 1 / New Instrument
#
# SAME parameter is compared with SAME parameter:
#   Alb    -> Alb
#   ALB_A  -> ALB_A
#   DBil_2 -> DBil_2
#   DBIL_A -> DBIL_A
#
# THERE IS NO Alb -> ALB_A conversion.
# Only parameters present in BOTH PDFs are included.
# ============================================================


RESULT_RE = re.compile(
    r"^([<>]?\s*[0-9]+(?:\.[0-9]+)?)\s*"
    r"(g/dL|mg/dL|ug/dL|µg/dL|ng/dL|U/L|mEq/L|mmol/L|µmol/L)?$",
    re.I
)

SID_RE = re.compile(
    r"(?i)^b2b[0-9]+(?:\.\.)?$"
)


# ============================================================
# PDF RESULT PARSING
# ============================================================

def parse_result(line):
    match = RESULT_RE.match(line.strip())

    if not match:
        return None, None

    number = match.group(1).replace(" ", "")
    unit = match.group(2) or ""

    raw = f"{number} {unit}".strip()

    # < / > results are retained as text.
    if number.startswith("<") or number.startswith(">"):
        return raw, None

    try:
        return raw, float(number)
    except ValueError:
        return raw, None


def extract_pdf(pdf_bytes):
    """
    Extract assay sections from the Siemens report.

    Returns:
        {
            "Alb": [
                {
                    "sid": "...",
                    "key": "...",
                    "raw": "...",
                    "value": ...
                }
            ],
            ...
        }
    """

    document = fitz.open(
        stream=pdf_bytes,
        filetype="pdf"
    )

    lines = []

    for page_number, page in enumerate(document, start=1):
        for raw_line in page.get_text().splitlines():

            line = raw_line.strip()

            if line:
                lines.append(
                    (page_number, line)
                )

    # --------------------------------------------------------
    # Detect assay names.
    # In this report structure, assay name is immediately
    # before "Patient Name".
    # --------------------------------------------------------

    headings = []

    for i, (_, line) in enumerate(lines):

        if line == "Patient Name" and i > 0:

            heading = lines[i - 1][1]

            if re.fullmatch(
                r"[A-Za-z][A-Za-z0-9_]{1,30}",
                heading
            ):

                if heading not in headings:
                    headings.append(heading)

    data = {}

    # --------------------------------------------------------
    # Extract each assay block.
    # --------------------------------------------------------

    for i, (_, line) in enumerate(lines):

        if line != "Patient Name" or i == 0:
            continue

        assay = lines[i - 1][1]

        if assay not in headings:
            continue

        # Find next assay section.
        j = i + 1

        while j < len(lines):

            if (
                lines[j][1] == "Patient Name"
                and j > 0
                and lines[j - 1][1] in headings
            ):
                break

            j += 1

        block = lines[i:j]

        rows = []

        last_result_index = -1

        for k, (_, text) in enumerate(block):

            raw, value = parse_result(text)

            if raw is None:
                continue

            sid = ""

            # Search upward for the nearest SID.
            for p in range(
                k - 1,
                max(-1, last_result_index),
                -1
            ):

                candidate = block[p][1]

                if SID_RE.match(candidate):
                    sid = candidate
                    break

            rows.append(
                {
                    "sid": sid,
                    "key": re.sub(
                        r"\s+",
                        "",
                        sid
                    ).upper(),
                    "raw": raw,
                    "value": value
                }
            )

            last_result_index = k

        if rows:
            data.setdefault(
                assay,
                []
            ).extend(rows)

    document.close()

    return data


# ============================================================
# SAME-NAME PARAMETER MATCHING
# ============================================================

def normalize_assay_name(name):
    return re.sub(
        r"\s+",
        "",
        name
    ).upper()


def common_parameters(old_data, new_data):

    old_normalized = {}
    new_normalized = {}

    for name in old_data:
        old_normalized[
            normalize_assay_name(name)
        ] = name

    for name in new_data:
        new_normalized[
            normalize_assay_name(name)
        ] = name

    common = []

    # Preserve OLD PDF order.
    for normalized_name, old_name in old_normalized.items():

        if normalized_name in new_normalized:

            new_name = new_normalized[
                normalized_name
            ]

            common.append(
                (
                    old_name,
                    new_name
                )
            )

    return common


# ============================================================
# EXCEL TEMPLATE HANDLING
# ============================================================

def safe_sheet_name(name, existing_names):

    base = re.sub(
        r'[:\\/?*\[\]]',
        "_",
        name or "Parameter"
    )[:31]

    if not base:
        base = "Parameter"

    candidate = base
    counter = 2

    while candidate in existing_names:

        suffix = f"_{counter}"

        candidate = (
            base[:31 - len(suffix)]
            + suffix
        )

        counter += 1

    return candidate


def duplicate_template_sheet(
    workbook,
    source_sheet,
    new_sheet_name
):
    """
    Duplicate the actual Excel template sheet inside the
    original workbook.

    This is deliberately NOT a manually reconstructed workbook.
    The original workbook remains the base workbook so its
    formatting, formulas, widths, heights, print settings and
    sheet structure are preserved as much as openpyxl allows.
    """

    new_sheet = workbook.copy_worksheet(
        source_sheet
    )

    new_sheet.title = new_sheet_name

    return new_sheet


def populate_parameter_sheet(
    ws,
    old_assay,
    new_assay,
    old_data,
    new_data,
    metadata
):

    # --------------------------------------------------------
    # HEADER AREA FROM THE USER'S TEMPLATE
    # --------------------------------------------------------

    ws["B1"] = old_assay

    ws["B2"] = metadata["instrument"]

    ws["B3"] = metadata["serial"]

    ws["B4"] = metadata["date"]

    # --------------------------------------------------------
    # REAGENT / CALIBRATOR / QC INFORMATION
    # --------------------------------------------------------

    ws["A5"] = metadata["reagent"]
    ws["B5"] = metadata["reagent_lot"]
    ws["C5"] = metadata["reagent_expiry"]

    ws["D5"] = metadata["calibrator"]
    ws["E5"] = metadata["calibrator_lot"]
    ws["F5"] = metadata["calibrator_expiry"]

    ws["G5"] = metadata["qc_lot"]
    ws["H5"] = metadata["qc_expiry"]

    # --------------------------------------------------------
    # TEMPLATE COMPARISON HEADERS
    # --------------------------------------------------------

    ws["C8"] = old_assay
    ws["D8"] = new_assay
    ws["E8"] = "Y-X"
    ws["F8"] = "%(Y-X)/X"

    ws["C9"] = old_assay
    ws["D9"] = new_assay
    ws["E9"] = "Bias"
    ws["F9"] = "%Bias"

    # --------------------------------------------------------
    # DO NOT REBUILD THE TEMPLATE CALCULATION AREA.
    #
    # The uploaded template already contains:
    # H9 = AVERAGE(C10:C29)
    # I9 = AVERAGE(D10:D29)
    # J9 = average bias
    # H10 = SLOPE(...)
    # H11 = INTERCEPT(...)
    # H12 = RSQ(...)
    #
    # Therefore we leave those formulas intact.
    # --------------------------------------------------------

    old_rows = old_data.get(
        old_assay,
        []
    )

    new_rows = new_data.get(
        new_assay,
        []
    )

    # --------------------------------------------------------
    # MATCH BY SAMPLE ID
    # --------------------------------------------------------

    old_lookup = {
        row["key"]: row
        for row in old_rows
        if row["key"]
    }

    new_lookup = {
        row["key"]: row
        for row in new_rows
        if row["key"]
    }

    ordered_sample_ids = []

    seen = set()

    # Preserve OLD PDF sample order.
    for row in old_rows:

        key = row["key"]

        if key and key not in seen:

            ordered_sample_ids.append(key)

            seen.add(key)

    # Add NEW-only IDs after OLD IDs.
    for row in new_rows:

        key = row["key"]

        if key and key not in seen:

            ordered_sample_ids.append(key)

            seen.add(key)

    # --------------------------------------------------------
    # DYNAMIC SAMPLE ROWS
    #
    # The supplied template has 20 prepared sample rows
    # (10-29), but the app must NOT stop at 20.
    #
    # If a parameter has 25, 40, 100, etc. matched samples,
    # additional rows are created automatically by copying the
    # formatting/formulas from the last template sample row.
    # --------------------------------------------------------

    required_rows = max(20, len(ordered_sample_ids))
    first_data_row = 10
    template_last_row = 29

    # Create additional rows only when more than 20 matched
    # sample IDs exist.
    if required_rows > 20:

        extra_rows = required_rows - 20

        # Insert before the next existing content, if any.
        ws.insert_rows(
            template_last_row + 1,
            amount=extra_rows
        )

        # Copy the template's row-29 formatting into each
        # newly created row. Values are populated below.
        for new_row in range(
            template_last_row + 1,
            required_rows + 10
        ):

            source_row = template_last_row

            ws.row_dimensions[new_row].height = (
                ws.row_dimensions[source_row].height
            )

            for col in range(
                1,
                max(ws.max_column, 10) + 1
            ):

                source_cell = ws.cell(
                    source_row,
                    col
                )

                target_cell = ws.cell(
                    new_row,
                    col
                )

                if source_cell.has_style:
                    target_cell._style = copy.copy(
                        source_cell._style
                    )

                if source_cell.number_format:
                    target_cell.number_format = (
                        source_cell.number_format
                    )

                if source_cell.alignment:
                    target_cell.alignment = copy.copy(
                        source_cell.alignment
                    )

                if source_cell.protection:
                    target_cell.protection = copy.copy(
                        source_cell.protection
                    )

                # Do not copy row-29 data values.
                # Data/formulas are populated below.
                target_cell.value = None

    last_data_row = (
        first_data_row + len(ordered_sample_ids) - 1
        if ordered_sample_ids
        else template_last_row
    )

    # Populate all matched samples, not just the first 20.
    for sample_index, sample_key in enumerate(
        ordered_sample_ids
    ):

        excel_row = first_data_row + sample_index

        old_row = old_lookup.get(
            sample_key
        )

        new_row = new_lookup.get(
            sample_key
        )

        # Serial number
        ws.cell(
            excel_row,
            1
        ).value = sample_index + 1

        # Clear only the data cells owned by the app.
        ws.cell(
            excel_row,
            2
        ).value = None

        ws.cell(
            excel_row,
            3
        ).value = None

        ws.cell(
            excel_row,
            4
        ).value = None

        # Sample ID
        if old_row:
            ws.cell(
                excel_row,
                2
            ).value = old_row["sid"]
        elif new_row:
            ws.cell(
                excel_row,
                2
            ).value = new_row["sid"]

        # OLD instrument
        if old_row:
            ws.cell(
                excel_row,
                3
            ).value = (
                old_row["value"]
                if old_row["value"] is not None
                else old_row["raw"]
            )

        # NEW instrument
        if new_row:
            ws.cell(
                excel_row,
                4
            ).value = (
                new_row["value"]
                if new_row["value"] is not None
                else new_row["raw"]
            )

        # Bias
        ws.cell(
            excel_row,
            5
        ).value = (
            f'=IF(AND(ISNUMBER(C{excel_row}),'
            f'ISNUMBER(D{excel_row})),'
            f'D{excel_row}-C{excel_row},"")'
        )

        # % Bias
        ws.cell(
            excel_row,
            6
        ).value = (
            f'=IF(AND(ISNUMBER(C{excel_row}),'
            f'ISNUMBER(D{excel_row}),'
            f'C{excel_row}<>0),'
            f'(D{excel_row}-C{excel_row})/'
            f'C{excel_row}*100,"")'
        )

    # --------------------------------------------------------
    # Extend the template's summary formulas to the actual
    # number of matched samples.
    #
    # The original template uses rows 10:29. When there are
    # >20 samples, the ranges become 10:last_data_row.
    # --------------------------------------------------------

    ws["H9"] = (
        f'=IFERROR(AVERAGE(C10:C{last_data_row}),"")'
    )

    ws["I9"] = (
        f'=IFERROR(AVERAGE(D10:D{last_data_row}),"")'
    )

    ws["J9"] = (
        f'=IFERROR((I9-H9)/H9*100,"")'
    )

    ws["H10"] = (
        f'=IFERROR(SLOPE(D10:D{last_data_row},'
        f'C10:C{last_data_row}),"")'
    )

    ws["H11"] = (
        f'=IFERROR(INTERCEPT(D10:D{last_data_row},'
        f'C10:C{last_data_row}),"")'
    )

    ws["H12"] = (
        f'=IFERROR(RSQ(D10:D{last_data_row},'
        f'C10:C{last_data_row}),"")'
    )


# ============================================================
# STREAMLIT USER INTERFACE
# ============================================================

st.title(
    "Atellica 2 vs Atellica 1 — Method Comparison"
)

st.info(
    "Upload two PDFs and your existing Excel template. "
    "The same parameter is compared with the same parameter. "
    "Example: Alb vs Alb and ALB_A vs ALB_A. "
    "There is no _2 to _A conversion."
)

old_file = st.file_uploader(
    "1. OLD Instrument PDF — Atellica 2",
    type=["pdf"]
)

new_file = st.file_uploader(
    "2. NEW Instrument PDF — Atellica 1",
    type=["pdf"]
)

template_file = st.file_uploader(
    "3. Your Method Comparison Excel Template",
    type=["xlsx", "xlsm"]
)


if (
    old_file
    and new_file
    and template_file
):

    with st.spinner(
        "Extracting parameters from both PDFs..."
    ):

        old_data = extract_pdf(
            old_file.getvalue()
        )

        new_data = extract_pdf(
            new_file.getvalue()
        )

        common_pairs = common_parameters(
            old_data,
            new_data
        )

    st.subheader(
        "Parameters to be compared"
    )

    comparison_table = []

    for old_assay, new_assay in common_pairs:

        comparison_table.append(
            {
                "OLD PDF": old_assay,
                "NEW PDF": new_assay,
                "OLD values":
                    len(
                        old_data.get(
                            old_assay,
                            []
                        )
                    ),
                "NEW values":
                    len(
                        new_data.get(
                            new_assay,
                            []
                        )
                    ),
                "Comparison":
                    f"{old_assay} vs {new_assay}"
            }
        )

    st.dataframe(
        comparison_table,
        use_container_width=True,
        hide_index=True
    )

    st.success(
        f"{len(common_pairs)} same-name "
        "parameters found in both PDFs."
    )

    if not common_pairs:

        st.error(
            "No same-name parameters were found "
            "in both PDFs. No Excel will be generated."
        )

        st.stop()

    # ========================================================
    # METADATA
    # ========================================================

    st.subheader(
        "Excel template information"
    )

    col1, col2, col3 = st.columns(3)

    instrument = col1.text_input(
        "Instrument Name",
        "Atellica 2 / Atellica 1"
    )

    serial = col2.text_input(
        "Serial No.",
        ""
    )

    report_date = col3.text_input(
        "Date",
        ""
    )

    col1, col2, col3 = st.columns(3)

    reagent = col1.text_input(
        "Reagent Name",
        ""
    )

    reagent_lot = col2.text_input(
        "Reagent Lot No",
        ""
    )

    reagent_expiry = col3.text_input(
        "Reagent Expiry Date",
        ""
    )

    col1, col2, col3 = st.columns(3)

    calibrator = col1.text_input(
        "Calibrator Name",
        ""
    )

    calibrator_lot = col2.text_input(
        "Calibrator Lot No",
        ""
    )

    calibrator_expiry = col3.text_input(
        "Calibrator Expiry Date",
        ""
    )

    col1, col2 = st.columns(2)

    qc_lot = col1.text_input(
        "QC Lot No",
        ""
    )

    qc_expiry = col2.text_input(
        "QC Expiry Date",
        ""
    )

    # ========================================================
    # GENERATE
    # ========================================================

    if st.button(
        "POPULATE ATTACHED EXCEL TEMPLATE",
        type="primary",
        use_container_width=True
    ):

        metadata = {
            "instrument": instrument,
            "serial": serial,
            "date": report_date,
            "reagent": reagent,
            "reagent_lot": reagent_lot,
            "reagent_expiry": reagent_expiry,
            "calibrator": calibrator,
            "calibrator_lot": calibrator_lot,
            "calibrator_expiry": calibrator_expiry,
            "qc_lot": qc_lot,
            "qc_expiry": qc_expiry
        }

        with st.spinner(
            "Populating the original Excel template..."
        ):

            # ------------------------------------------------
            # IMPORTANT:
            # Load the USER'S ORIGINAL workbook directly.
            # We do NOT create a new Workbook().
            # ------------------------------------------------

            workbook = openpyxl.load_workbook(
                io.BytesIO(
                    template_file.getvalue()
                )
            )

            template_sheet = workbook[
                workbook.sheetnames[0]
            ]

            # We need one output sheet per common parameter.
            # Keep the first original template sheet for
            # the first parameter and duplicate it for
            # subsequent parameters.
            #
            # This preserves the original workbook instead
            # of reconstructing it.
            # ------------------------------------------------

            first_parameter = True

            for old_assay, new_assay in common_pairs:

                sheet_name = safe_sheet_name(
                    old_assay,
                    set(workbook.sheetnames)
                )

                if first_parameter:

                    # Use the original Sheet1/template sheet.
                    output_sheet = template_sheet

                    output_sheet.title = sheet_name

                    first_parameter = False

                else:

                    output_sheet = duplicate_template_sheet(
                        workbook,
                        template_sheet,
                        sheet_name
                    )

                populate_parameter_sheet(
                    output_sheet,
                    old_assay,
                    new_assay,
                    old_data,
                    new_data,
                    metadata
                )

            # ------------------------------------------------
            # Remove any extra original template sheets only
            # if they are not output sheets.
            #
            # For the supplied comp(2).xlsx there is one
            # template sheet, so nothing extra is removed.
            # ------------------------------------------------

            # Ask Excel to recalculate the formulas when the
            # generated workbook is opened.
            try:
                workbook.calculation.fullCalcOnLoad = True
                workbook.calculation.forceFullCalc = True
                workbook.calculation.calcMode = "auto"
            except Exception:
                pass

            output = io.BytesIO()

            workbook.save(output)

            output.seek(0)

        st.success(
            f"Completed successfully: "
            f"{len(common_pairs)} parameter sheet(s) created. "
            "All matched samples were populated, including samples beyond 20."
        )

        st.download_button(
            "DOWNLOAD POPULATED METHOD COMPARISON EXCEL",
            data=output.getvalue(),
            file_name=(
                "Atellica_Method_Comparison.xlsx"
            ),
            mime=(
                "application/vnd.openxmlformats-officedocument."
                "spreadsheetml.sheet"
            ),
            use_container_width=True
        )

        st.warning(
            "Open the downloaded Excel directly. "
            "The app uses your uploaded template as the base workbook "
            "and does not rebuild the workbook from scratch."
        )
