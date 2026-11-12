# RQ2 - Generative AI vulnerability analysis (reproducible build)

Run `python build_rq2_all.py` (needs pandas, numpy, matplotlib, openpyxl, python-docx).
It rebuilds EVERY output below from the RAW_RECORDS / PRODUCT_CATALOG_LIST / CWE_REFERENCE lists at the top of the script.

To change the data: edit RAW_RECORDS (one tuple per CVE: id, product, published, cwe, cvss, C, I, A, status),
EXCLUDED_FALSE_POSITIVES, DUPLICATES_REMOVED or V4_ONLY_RECORDS, then re-run. All counts, shares, tables, charts,
the Excel workbook, the notebook outputs and the Word report update together; the build stops with an error if the
denominators do not add up.

Outputs: 1_Product_Metadata_Catalog_Step1.csv, 2_GenAI_Vulnerabilities_Clean_Dataset_Step4.csv,
3_RQ2_Top_Valid_CWE_Distribution.csv, 4_RQ2_Records_Without_CWE_NoInfo_List.csv,
5_RQ2_Missing_and_NoInfo_Summary_Table.csv, 6_Step5_Data_Audit_and_Denominators.csv,
RQ2_Complete_Analysis_Workbook.xlsx (live formulas), RQ2_Complete_Analysis_Notebook.ipynb,
4 PNG charts, RQ2_Full_Methodology_and_Analysis_Report.docx.
