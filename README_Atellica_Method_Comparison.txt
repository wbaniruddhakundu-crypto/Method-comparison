ATELLICA 2 vs ATELLICA 1 METHOD COMPARISON

IMPORTANT METHOD RULE:
There is NO Alb -> ALB_A mapping.

The two PDFs represent two instruments/runs:
    OLD PDF = Atellica 2
    NEW PDF = Atellica 1

The SAME parameter is compared with the SAME parameter:
    Alb    vs Alb
    ALB_A  vs ALB_A
    CA_2   vs CA_2
    CAAZ_A vs CAAZ_A
    DBil_2 vs DBil_2
    DBIL_A vs DBIL_A

The app automatically detects parameters present in BOTH PDFs and compares
only those same-name parameters.

The attached Excel template is used as the master format. Each common
parameter receives a copy of that template and the existing sample/result
cells are populated.

Run:
    pip install -r requirements.txt
    streamlit run atellica_method_comparison_app.py
