<div align="center">

# Comprehensive Security Analysis of Generative AI Ecosystems

![Python](https://img.shields.io/badge/Python-3.14-blue?style=for-the-badge&logo=python&logoColor=white)
![Cybersecurity](https://img.shields.io/badge/Cybersecurity-GenAI-red?style=for-the-badge)
![Data Analysis](https://img.shields.io/badge/Data%20Analysis-Jupyter-orange?style=for-the-badge&logo=jupyter&logoColor=white)

*An empirical investigation into the vulnerability landscape of prominent GenAI frameworks, including LangChain, LlamaIndex, AutoGPT, and Ollama (2023-2026).*

</div>

---

## Abstract

This repository contains the datasets, analytical codebase, and methodological reports for our cybersecurity research initiative. The objective is to systematically extract, validate, and analyze Common Vulnerabilities and Exposures (CVEs) and Common Weakness Enumerations (CWEs) inherent in modern Generative AI development stacks.

By processing fragmented security reports into a unified, clean dataset, this project provides actionable insights into the threat vectors currently affecting AI integrations.

## Repository Architecture

The project is structured to ensure strict data integrity and analytical reproducibility:

```text
GenAI_Vulnerability_Analysis/
├── data/
│   ├── 1_Product_Metadata_Catalog_Step1.csv
│   ├── 2_GenAI_Vulnerabilities_Clean_Dataset_Step4.csv
│   ├── 3_Top_Valid_CWE_Distribution.csv
│   ├── 4_Records_Without_CWE_NoInfo_List.csv
│   ├── 5_Missing_and_NoInfo_Summary_Table.csv
│   └── 6_Step5_Data_Audit_and_Denominators.csv
├── figures/
│   ├── CWE_Product_Severity_Breakdown.png
│   ├── Top_CWE_BarChart_Academic_Light.png
│   ├── Top_CWE_BarChart_Dark.png
│   └── Yearly_Evolution_2023_2026.png
├── notebooks/
│   └── Complete_Analysis_Notebook.ipynb
├── src/
│   └── main.py
├── README.md
├── Complete_Analysis_Workbook.xlsx
└── Full_Methodology_and_Analysis_Report.docx
Key Analytical Dimensions

    1. Vulnerability Trajectory: Tracking the yearly evolution and frequency of CVEs in GenAI tools from 2023 to 2026.

    2. CWE Density & Distribution: Identifying the architectural weaknesses most frequently exploited or reported within these frameworks.

    3. Severity Impact Assessment: Cross-referencing product data with vulnerability severity scores to highlight critical risk areas.

    4. Data Auditing Mechanics: A rigorous methodology for handling incomplete records and missing CWE definitions.
