#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
RQ2 (most frequent CWE types in Generative AI products, 2023-2026) - single reproducible build.
Regenerates ALL deliverables (CSV 1-6, Excel workbook, notebook, charts, Word report) from the data lists below.
Usage:  python build_rq2_all.py
"""
import os, re, json
import numpy as np
import pandas as pd
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.chart import BarChart, Reference

# ============================== DATA (edit here) ==============================
RAW_RECORDS = [
    ('CVE-2026-21869', 'llama_cpp', '2026-01-08', 'CWE-787', 9.8, 'HIGH', 'HIGH', 'HIGH', 'NVD record verified (User + NVD 2026)'),
    ('CVE-2026-27795', 'langchain', '2026-02-25', 'CWE-918', 7.4, 'HIGH', 'LOW', 'NONE', 'NVD record verified (NVD 2026)'),
    ('CVE-2026-34070', 'langchain', '2026-03-31', 'CWE-22', 7.5, 'HIGH', 'NONE', 'NONE', 'NVD record verified (NVD 2026)'),
    ('CVE-2026-40087', 'langchain', '2026-04-09', 'CWE-1336', 5.3, 'LOW', 'NONE', 'NONE', 'NVD record verified (NVD 2026)'),
    ('CVE-2026-55443', 'langchain', '2026-06-22', 'CWE-22', 5.5, 'HIGH', 'NONE', 'NONE', 'NVD record verified (NVD 2026)'),
    ('CVE-2026-5757', 'ollama', '2026-06-26', 'CWE-125', 7.5, 'NONE', 'NONE', 'HIGH', 'NVD record verified (User + NVD 2026)'),
    ('CVE-2026-70473', 'flowise', '2026-08-04', 'CWE-200', 7.5, 'HIGH', 'NONE', 'NONE', 'NVD record verified (NVD 2026)'),
    ('CVE-2026-70471', 'flowise', '2026-08-04', 'CWE-200', 7.5, 'HIGH', 'NONE', 'NONE', 'NVD record verified (NVD 2026)'),
    ('CVE-2026-8478', 'langflow', '2026-08-05', 'CWE-94', 8.8, 'HIGH', 'HIGH', 'HIGH', 'NVD record verified (User + NVD 2026)'),
    ('CVE-2026-73678', 'mindsdb', '2026-08-14', 'CWE-94', 10.0, 'HIGH', 'HIGH', 'HIGH', 'NVD record verified (User + NVD 2026)'),
    ('CVE-2026-87016', 'open_webui', '2026-09-09', 'CWE-155', 8.1, 'HIGH', 'HIGH', 'NONE', 'NVD record verified (NVD 2026)'),
    ('CVE-2026-101861', 'langflow', '2026-09-28', 'CWE-94', 4.1, 'LOW', 'LOW', 'LOW', 'NVD record verified (NVD 2026)'),
    ('CVE-2026-105697', 'langflow', '2026-10-05', 'CWE-78', 9.9, 'HIGH', 'HIGH', 'HIGH', 'NVD record verified (NVD 2026)'),
    ('CVE-2026-105740', 'langflow', '2026-10-06', 'CWE-78', 9.9, 'HIGH', 'HIGH', 'HIGH', 'NVD record verified (User + NVD 2026)'),
    ('CVE-2026-105760', 'vllm', '2026-10-06', 'CWE-400', 5.3, 'NONE', 'NONE', 'LOW', 'NVD record verified (User + NVD 2026)'),
    ('CVE-2026-97680', 'langflow', '2026-10-07', 'CWE-284', 8.3, 'HIGH', 'HIGH', 'LOW', 'NVD record verified (User + NVD 2026)'),
    ('CVE-2026-108836', 'ollama', '2026-10-12', 'CWE-918', 7.3, 'LOW', 'LOW', 'LOW', 'NVD record checked (CNA: VulDB CVSS v3.1; NVD analysis pending at extraction)'),
    ('CVE-2025-29783', 'vllm', '2025-03-19', 'CWE-502', 9.8, 'HIGH', 'HIGH', 'HIGH', 'NVD record verified (NVD 2025)'),
    ('CVE-2025-3248', 'langflow', '2025-04-07', 'CWE-306', 9.8, 'HIGH', 'HIGH', 'HIGH', 'NVD record verified (NVD 2025 - CISA KEV)'),
    ('CVE-2025-48956', 'vllm', '2025-08-21', 'CWE-400', 7.5, 'NONE', 'NONE', 'HIGH', 'NVD record verified (NVD 2025)'),
    ('CVE-2025-59528', 'flowise', '2025-09-22', 'CWE-94', 9.8, 'HIGH', 'HIGH', 'HIGH', 'NVD record verified (NVD 2025)'),
    ('CVE-2025-61687', 'flowise', '2025-10-06', 'CWE-434', 8.8, 'HIGH', 'HIGH', 'HIGH', 'NVD record verified (NVD 2025)'),
    ('CVE-2025-65106', 'langchain', '2025-11-21', 'CWE-1336', 8.1, 'HIGH', 'HIGH', 'NONE', 'NVD record verified (NVD 2025)'),
    ('CVE-2025-65959', 'open_webui', '2025-12-04', 'CWE-79', 7.2, 'HIGH', 'HIGH', 'NONE', 'NVD record verified (NVD 2025)'),
    ('CVE-2025-34291', 'langflow', '2025-12-05', 'CWE-346', 9.4, 'HIGH', 'HIGH', 'HIGH', 'NVD record verified (NVD 2025)'),
    ('CVE-2025-68477', 'langflow', '2025-12-19', 'CWE-918', 7.5, 'HIGH', 'NONE', 'NONE', 'NVD record verified (NVD 2025)'),
    ('CVE-2025-68664', 'langchain', '2025-12-23', 'CWE-502', 8.2, 'HIGH', 'HIGH', 'NONE', 'NVD record verified (NVD 2025)'),
    ('CVE-2025-15514', 'ollama', '2026-01-12', 'CWE-476', 7.5, 'NONE', 'NONE', 'HIGH', 'NVD record verified (CVE-2025 ID / Pub Jan 2026)'),
    ('CVE-2024-23751', 'llama_index', '2024-01-22', 'CWE-89', 9.8, 'HIGH', 'HIGH', 'HIGH', 'NVD record verified (Base + NVD 2024)'),
    ('CVE-2024-0243', 'langchain', '2024-02-26', 'CWE-918', 8.1, 'HIGH', 'LOW', 'NONE', 'NVD record verified (User + Base 2024)'),
    ('CVE-2024-28088', 'langchain', '2024-03-04', 'CWE-22', 8.1, 'HIGH', 'HIGH', 'NONE', 'NVD record verified (User + Base 2024 - Deduped CWE-22/CWE-31)'),
    ('CVE-2024-1455', 'langchain', '2024-03-26', 'CWE-776', 5.9, 'NONE', 'NONE', 'HIGH', 'NVD record verified (Base + NVD 2024)'),
    ('CVE-2024-3571', 'langchain', '2024-04-16', 'CWE-22', 8.8, 'HIGH', 'HIGH', 'NONE', 'NVD record verified (User + Base 2024)'),
    ('CVE-2024-3271', 'llama_index', '2024-04-16', 'CWE-77', 9.8, 'HIGH', 'HIGH', 'HIGH', 'NVD record verified (User + Base 2024)'),
    ('CVE-2024-30256', 'open_webui', '2024-04-16', 'CWE-918', 6.4, 'LOW', 'LOW', 'NONE', 'NVD record verified (NVD 2024)'),
    ('CVE-2024-31621', 'flowise', '2024-04-29', 'CWE-94', 7.6, 'HIGH', 'HIGH', 'HIGH', 'NVD record verified (User + NVD 2024 - Date normalized)'),
    ('CVE-2024-37032', 'ollama', '2024-05-31', 'CWE-22', 8.8, 'HIGH', 'HIGH', 'HIGH', 'NVD record verified (NVD 2024)'),
    ('CVE-2024-2965', 'langchain', '2024-06-06', 'CWE-674', 4.7, 'NONE', 'NONE', 'LOW', 'NVD record verified (User + Base 2024 - Date normalized)'),
    ('CVE-2024-1879', 'autogpt', '2024-06-06', 'CWE-352', 8.8, 'HIGH', 'HIGH', 'HIGH', 'NVD record verified (User + Base 2024)'),
    ('CVE-2024-3095', 'langchain', '2024-06-06', 'CWE-918', 7.7, 'HIGH', 'NONE', 'NONE', 'NVD record verified (User + NVD 2024)'),
    ('CVE-2024-36420', 'flowise', '2024-07-01', 'CWE-74', 7.5, 'HIGH', 'NONE', 'NONE', 'NVD record verified (NVD 2024)'),
    ('CVE-2024-37146', 'flowise', '2024-07-01', 'CWE-79', 6.1, 'LOW', 'LOW', 'NONE', 'NVD record verified (NVD 2024)'),
    ('CVE-2024-7297', 'langflow', '2024-07-30', 'CWE-913', 9.8, 'HIGH', 'HIGH', 'HIGH', 'NVD record verified (NVD 2024)'),
    ('CVE-2024-9277', 'langflow', '2024-09-27', 'CWE-1333', 5.3, 'NONE', 'NONE', 'LOW', 'NVD record verified (NVD 2024)'),
    ('CVE-2024-8309', 'langchain', '2024-10-29', 'CWE-74', 9.8, 'HIGH', 'HIGH', 'HIGH', 'NVD record verified (NVD 2024)'),
    ('CVE-2024-7774', 'langchain', '2024-10-29', 'CWE-22', 7.5, 'HIGH', 'HIGH', 'NONE', 'NVD record verified (NVD 2024)'),
    ('CVE-2024-48061', 'langflow', '2024-11-04', 'CWE-94', 9.8, 'HIGH', 'HIGH', 'HIGH', 'NVD record verified (NVD 2024)'),
    ('CVE-2023-29374', 'langchain', '2023-04-05', 'CWE-74', 9.8, 'HIGH', 'HIGH', 'HIGH', 'NVD record verified (Base + NVD 2023)'),
    ('CVE-2023-1011', 'wpbot', '2023-05-08', 'NVD-CWE-noinfo', 6.1, 'LOW', 'LOW', 'NONE', 'NVD record verified (User Proposed 2023)'),
    ('CVE-2023-1649', 'wpbot', '2023-05-08', 'CWE-79', 4.8, 'LOW', 'LOW', 'NONE', 'NVD record verified (User Proposed 2023)'),
    ('CVE-2023-1651', 'wpbot', '2023-05-08', 'NVD-CWE-noinfo', 5.4, 'LOW', 'LOW', 'NONE', 'NVD record verified (User Proposed 2023)'),
    ('CVE-2023-34540', 'langchain', '2023-06-14', 'NVD-CWE-noinfo', 9.8, 'HIGH', 'HIGH', 'HIGH', 'NVD record verified (Base + NVD 2023)'),
    ('CVE-2023-34541', 'langchain', '2023-06-20', 'NVD-CWE-noinfo', 9.8, 'HIGH', 'HIGH', 'HIGH', 'NVD record verified (Base 2023)'),
    ('CVE-2023-36258', 'langchain', '2023-07-04', 'CWE-94', 9.8, 'HIGH', 'HIGH', 'HIGH', 'NVD record verified (Base + NVD 2023)'),
    ('CVE-2023-36188', 'langchain', '2023-07-06', 'CWE-74', 9.8, 'HIGH', 'HIGH', 'HIGH', 'NVD record verified (Base + NVD 2023)'),
    ('CVE-2023-36189', 'langchain', '2023-07-06', 'CWE-89', 7.5, 'HIGH', 'NONE', 'NONE', 'NVD record verified (Base + NVD 2023)'),
    ('CVE-2023-3175', 'wpbot', '2023-07-10', 'NVD-CWE-noinfo', 4.8, 'LOW', 'LOW', 'NONE', 'NVD record verified (User Proposed 2023)'),
    ('CVE-2023-37273', 'autogpt', '2023-07-14', 'CWE-94', 8.8, 'HIGH', 'HIGH', 'HIGH', 'NVD record verified (User + Base 2023)'),
    ('CVE-2023-37274', 'autogpt', '2023-07-14', 'CWE-94', 7.8, 'HIGH', 'HIGH', 'HIGH', 'NVD record verified (User + Base 2023)'),
    ('CVE-2023-37275', 'autogpt', '2023-07-14', 'CWE-117', 4.3, 'NONE', 'LOW', 'NONE', 'NVD record verified (User + Base 2023)'),
    ('CVE-2023-36095', 'langchain', '2023-08-05', 'CWE-94', 9.8, 'HIGH', 'HIGH', 'HIGH', 'NVD record verified (Base + NVD 2023)'),
    ('CVE-2023-38896', 'langchain', '2023-08-15', 'CWE-74', 9.8, 'HIGH', 'HIGH', 'HIGH', 'NVD record verified (Base 2023)'),
    ('CVE-2023-39659', 'langchain', '2023-08-15', 'CWE-74', 9.8, 'HIGH', 'HIGH', 'HIGH', 'NVD record verified (Base + NVD 2023)'),
    ('CVE-2023-39662', 'llama_index', '2023-08-15', 'CWE-74', 9.8, 'HIGH', 'HIGH', 'HIGH', 'NVD record verified (Base 2023)'),
    ('CVE-2023-36281', 'langchain', '2023-08-22', 'CWE-94', 9.8, 'HIGH', 'HIGH', 'HIGH', 'NVD record verified (Base + NVD 2023)'),
    ('CVE-2023-39631', 'langchain', '2023-09-01', 'CWE-94', 9.8, 'HIGH', 'HIGH', 'HIGH', 'NVD record verified (Base + NVD 2023)'),
    ('CVE-2023-44467', 'langchain_experimental', '2023-10-09', 'NVD-CWE-noinfo', 9.8, 'HIGH', 'HIGH', 'HIGH', 'NVD record verified (Base + NVD 2023)'),
    ('CVE-2023-45063', 'ai_content_writing_assistant', '2023-10-12', 'CWE-352', 8.8, 'HIGH', 'HIGH', 'HIGH', 'NVD record verified (User Proposed 2023)'),
    ('CVE-2023-46229', 'langchain', '2023-10-19', 'CWE-918', 8.8, 'HIGH', 'HIGH', 'HIGH', 'NVD record verified (Base 2023)'),
    ('CVE-2023-32786', 'langchain', '2023-10-21', 'CWE-74', 7.5, 'NONE', 'NONE', 'HIGH', 'NVD record verified (Base 2023)'),
    ('CVE-2023-5606', 'wpbot', '2023-11-02', 'CWE-79', 4.8, 'LOW', 'LOW', 'NONE', 'NVD record verified (User Proposed 2023)'),
    ('CVE-2023-51419', 'bertha_ai', '2023-12-29', 'CWE-434', 9.8, 'HIGH', 'HIGH', 'HIGH', 'NVD record verified (User Proposed 2023)'),
]

PRODUCT_CATALOG_LIST = [
    ('langchain', 'LangChain', 'LangChain, Inc.', 'langchain, langchain-ai, langchain-core, @langchain/community', 'https://github.com/langchain-ai/langchain', 'LLM Application / Orchestration Framework', 'cpe:2.3:a:langchain:langchain', 'Included (Core Framework - 27 CVEs)'),
    ('langchain_experimental', 'LangChain Experimental', 'LangChain, Inc.', 'langchain-experimental, langchain_experimental', 'https://github.com/langchain-ai/langchain-experimental', 'LLM Application / Orchestration Framework', 'cpe:2.3:a:langchain:langchain_experimental', 'Included (Sub-package of LangChain - 1 CVE)'),
    ('langflow', 'Langflow (IBM Langflow OSS)', 'Langflow / DataStax / IBM', 'langflow, langflow-ai, IBM Langflow OSS', 'https://github.com/langflow-ai/langflow', 'Visual LLM Workflow / Agent Builder', 'cpe:2.3:a:langflow:langflow', 'Included (Visual Agent Platform - 11 CVEs)'),
    ('flowise', 'Flowise', 'FlowiseAI, Inc.', 'flowise, flowiseai', 'https://github.com/FlowiseAI/Flowise', 'Visual LLM Workflow / Agent Builder', 'cpe:2.3:a:flowiseai:flowise', 'Included (Visual Agent Platform - 7 CVEs)'),
    ('autogpt', 'AutoGPT', 'Significant Gravitas', 'Auto-GPT, autogpt, Significant-Gravitas/AutoGPT', 'https://github.com/Significant-Gravitas/AutoGPT', 'AI Agent Platform', 'cpe:2.3:a:significant-gravitas:autogpt', 'Included (Autonomous AI Agent - 4 CVEs)'),
    ('ollama', 'Ollama', 'Ollama', 'ollama, ollama/ollama', 'https://github.com/ollama/ollama', 'Local LLM Runtime & Inference Engine', 'cpe:2.3:a:ollama:ollama', 'Included (LLM Inference Runtime - 4 CVEs)'),
    ('vllm', 'vLLM', 'vLLM Project', 'vllm, vllm-project/vllm', 'https://github.com/vllm-project/vllm', 'High-Throughput LLM Inference Engine', 'cpe:2.3:a:vllm:vllm', 'Included (LLM Serving Engine - 3 CVEs)'),
    ('open_webui', 'Open WebUI', 'Open WebUI Community', 'open-webui, ollama-webui', 'https://github.com/open-webui/open-webui', 'GenAI Web Interface & Orchestration UI', 'cpe:2.3:a:openwebui:open_webui', 'Included (GenAI UI Platform - 3 CVEs)'),
    ('llama_index', 'LlamaIndex', 'LlamaIndex', 'llama_index, llama-index, gpt_index', 'https://github.com/run-llama/llama_index', 'LLM Application / Data Framework', 'cpe:2.3:a:llamaindex:llamaindex', 'Included (RAG / Data Framework - 3 CVEs)'),
    ('llama_cpp', 'llama.cpp', 'ggml-org (Georgi Gerganov)', 'llama.cpp, ggml-org/llama.cpp', 'https://github.com/ggml-org/llama.cpp', 'Local LLM Runtime & Inference Engine', 'cpe:2.3:a:ggml:llama.cpp', 'Included (C/C++ LLM Inference - 1 CVE)'),
    ('mindsdb', 'MindsDB Minds Platform', 'MindsDB, Inc.', 'mindsdb, MindsDB Minds Platform', 'https://github.com/mindsdb/mindsdb', 'Enterprise AI / ML Orchestration Platform', 'cpe:2.3:a:mindsdb:mindsdb', 'Included (AI Data Platform - 1 CVE)'),
    ('wpbot', 'AI ChatBot (WPBot)', 'QuantumCloud', 'wpbot, AI ChatBot for WordPress', 'https://wordpress.org/plugins/chatbot/', 'GenAI CMS Plugin / Chatbot', 'cpe:2.3:a:quantumcloud:ai_chatbot', 'Included (User Verified NVD 2023 - 5 CVEs)'),
    ('ai_content_writing_assistant', 'AI Content Writing Assistant', 'CodeSolz', 'ai-content-writing-assistant', 'https://wordpress.org/plugins/ai-content-writing-assistant/', 'GenAI CMS Plugin / Content Assistant', 'cpe:2.3:a:codesolz:ai_content_writing_assistant', 'Included (User Verified NVD 2023 - 1 CVE)'),
    ('bertha_ai', 'BERTHA AI', 'BERTHA AI', 'bertha-ai, bertha-ai-free', 'https://wordpress.org/plugins/bertha-ai-free/', 'GenAI CMS Plugin / Content Assistant', 'cpe:2.3:a:bertha:bertha_ai', 'Included (User Verified NVD 2023 - 1 CVE)'),
]

CWE_REFERENCE = {
    'CWE-94': ("Improper Control of Generation of Code ('Code Injection')", 'تزریق کد (Code Injection)', 'Code Execution / Injection'),
    'CWE-74': ("Improper Neutralization of Special Elements in Output Used by a Downstream Component ('Injection')", 'تزریق در کامپوننت پایین\u200cدستی (Downstream Injection)', 'Injection / Output Neutralization'),
    'CWE-22': ("Improper Limitation of a Pathname to a Restricted Directory ('Path Traversal')", 'پیمایش مسیر (Path Traversal)', 'File / Path Access'),
    'CWE-918': ('Server-Side Request Forgery (SSRF)', 'جعل درخواست سمت سرور (SSRF)', 'Network / Request Forgery'),
    'CWE-79': ("Improper Neutralization of Input During Web Page Generation ('Cross-site Scripting')", 'اسکریپت\u200cنویسی میان\u200cسایتی (XSS)', 'Web / Client-Side Injection'),
    'CWE-78': ("Improper Neutralization of Special Elements used in an OS Command ('OS Command Injection')", 'تزریق دستور سیستم\u200cعامل (OS Command Injection)', 'Command Execution / Injection'),
    'CWE-434': ('Unrestricted Upload of File with Dangerous Type', 'آپلود بدون محدودیت فایل خطرناک (Unrestricted File Upload)', 'File / Upload Security'),
    'CWE-502': ('Deserialization of Untrusted Data', 'دی\u200cسریالایز کردن داده\u200cهای غیرقابل اعتماد (Unsafe Deserialization)', 'Code Execution / Serialization'),
    'CWE-352': ('Cross-Site Request Forgery (CSRF)', 'جعل درخواست میان\u200cسایتی (CSRF)', 'Session / Web Access'),
    'CWE-89': ("Improper Neutralization of Special Elements used in an SQL Command ('SQL Injection')", 'تزریق SQL (SQL Injection)', 'Database Injection'),
    'CWE-200': ('Exposure of Sensitive Information to an Unauthorized Actor', 'افشای اطلاعات حساس (Information Exposure)', 'Information Disclosure / Access Control'),
    'CWE-1336': ('Improper Neutralization of Special Elements Used in a Template Engine', 'تزریق قالب سمت سرور (Template Injection / SSTI)', 'Template / Code Injection'),
    'CWE-400': ('Uncontrolled Resource Consumption', 'مصرف کنترل\u200cنشده منابع (DoS / Resource Exhaustion)', 'Resource / Availability'),
    'CWE-306': ('Missing Authentication for Critical Function', 'فقدان احراز هویت برای تابع حیاتی (Missing Authentication)', 'Authentication / Access Control'),
    'CWE-77': ("Improper Neutralization of Special Elements used in a Command ('Command Injection')", 'تزریق فرمان (Command Injection)', 'Command Execution / Injection'),
    'CWE-787': ('Out-of-bounds Write', 'نوشتن خارج از محدوده حافظه (Out-of-bounds Write)', 'Memory Buffer Errors'),
    'CWE-913': ('Improper Control of Dynamically-Managed Code Resources', 'کنترل نامناسب منابع کد پویا / ارتقای سطح دسترسی', 'Privilege / Code Control'),
    'CWE-346': ('Origin Validation Error', 'خطای اعتبارسنجی منشأ (CORS / Origin Validation Error)', 'Session / Web Access'),
    'CWE-284': ('Improper Access Control', 'کنترل دسترسی نامناسب (Improper Access Control)', 'Authentication / Access Control'),
    'CWE-155': ('Improper Neutralization of Wildcards or Matching Symbols', 'عدم خنثی\u200cسازی کاراکترهای Wildcard در احراز هویت', 'Authentication / Injection'),
    'CWE-125': ('Out-of-bounds Read', 'خواندن خارج از محدوده حافظه (Out-of-bounds Read)', 'Memory Buffer Errors'),
    'CWE-476': ('NULL Pointer Dereference', 'ارجاع به اشاره\u200cگر تهی (NULL Pointer Dereference)', 'Memory / Crash (DoS)'),
    'CWE-776': ("Improper Restriction of Recursive Entity References in DTDs ('XML Entity Expansion')", 'گسترش بازگشتی موجودیت XML (XML Entity Expansion / Billion Laughs)', 'XML / Resource Consumption'),
    'CWE-1333': ('Inefficient Regular Expression Complexity', 'پیچیدگی ناکارآمد عبارات باقاعده (ReDoS)', 'Resource / Availability'),
    'CWE-674': ('Uncontrolled Recursion', 'بازگشت بی\u200cرویه (Uncontrolled Recursion / DoS)', 'Resource / Availability'),
    'CWE-117': ('Improper Output Neutralization for Logs', 'عدم خنثی\u200cسازی خروجی در لاگ (Log Injection)', 'Logging / Output Neutralization'),
    'NVD-CWE-noinfo': ('Insufficient Information (Unspecified Weakness in NVD)', 'اطلاعات ناکافی / بدون کد CWE مشخص در NVD', 'Unspecified / Missing CWE'),
}

# ----------------------------------------------------------------------------
# Audit / methodology constants (edit here if the raw data changes)
# ----------------------------------------------------------------------------
EXTRACTION_DATE = "2026-11-12"
DUPLICATES_REMOVED = [("CVE-2024-28088", "appeared twice (CWE-22 and CWE-31); kept primary CWE-22")]
EXCLUDED_FALSE_POSITIVES = [
    ("CVE-2024-4701", "genie", "Netflix Genie (Hadoop/Spark big-data job orchestration) - not a Generative AI product"),
]
V4_ONLY_RECORDS = {"CVE-2025-34291"}  # score available here only as CVSS v4.0 (kept, flagged, sensitivity reported)
INITIAL_CANDIDATE_ROWS = len(RAW_RECORDS) + len(DUPLICATES_REMOVED) + len(EXCLUDED_FALSE_POSITIVES)

NOINFO = "NVD-CWE-noinfo"
INJECTION_FAMILY = ["CWE-94", "CWE-74", "CWE-79", "CWE-78", "CWE-502", "CWE-89", "CWE-1336", "CWE-77"]


def classify_severity(score):
    if score >= 9.0: return "CRITICAL"
    if score >= 7.0: return "HIGH"
    if score >= 4.0: return "MEDIUM"
    if score > 0: return "LOW"
    return "NONE"


def fa(x):
    """Latin digits -> Persian digits (for Word prose)."""
    return str(x).translate(str.maketrans("0123456789", "۰۱۲۳۴۵۶۷۸۹"))


# ----------------------------------------------------------------------------
# Data frames
# ----------------------------------------------------------------------------
def build_frames():
    cat = pd.DataFrame(PRODUCT_CATALOG_LIST, columns=[
        "product_id", "display_name", "vendor_org", "alternative_names",
        "official_repo", "group", "cpe_prefix", "inclusion_status"])
    df = pd.DataFrame(RAW_RECORDS, columns=[
        "cve_id", "product", "published", "cwe", "cvss_score",
        "confidentiality_impact", "integrity_impact", "availability_impact", "validation_status"])
    gm = dict(zip(cat.product_id, cat.group))
    df["group"] = df["product"].map(gm)
    assert df["group"].notna().all(), "product missing from catalog"
    assert not df.cve_id.duplicated().any(), "duplicate CVE id"
    df["published_year"] = pd.to_datetime(df["published"]).dt.year
    for i, col in enumerate(["cwe_name_en", "cwe_name_fa", "cwe_category"]):
        df[col] = df["cwe"].map(lambda x: CWE_REFERENCE[x][i])
    df["is_valid_cwe"] = df["cwe"] != NOINFO
    df["cvss_version"] = np.where(df["cve_id"].isin(V4_ONLY_RECORDS), "4.0", "3.1")
    df["severity"] = df["cvss_score"].map(classify_severity)
    df["C/I/A"] = (df.confidentiality_impact.str[0] + ":" + df.integrity_impact.str[0] + "/" + df.availability_impact.str[0])
    df["epss"] = "N/A"; df["epss_date"] = "N/A"
    df["advisory_url"] = "https://nvd.nist.gov/vuln/detail/" + df["cve_id"]
    cols = ["cve_id", "product", "group", "published", "published_year", "cwe", "cwe_name_en", "cwe_name_fa",
            "cwe_category", "is_valid_cwe", "cvss_version", "cvss_score", "severity", "confidentiality_impact",
            "integrity_impact", "availability_impact", "C/I/A", "epss", "epss_date", "advisory_url", "validation_status"]
    df = df[cols].sort_values(["published", "cve_id"], ascending=[False, False]).reset_index(drop=True)

    # catalog inclusion_status: keep the label, make the CVE count dynamic
    cnt = df["product"].value_counts().to_dict()
    def status(r):
        label = re.sub(r"\s*-\s*\d+ CVEs?\)$", "", r.inclusion_status).rstrip(")")
        label = label.replace("Included (", "")
        n = cnt.get(r.product_id, 0)
        return f"Included ({label} - {n} CVE{'s' if n != 1 else ''})"
    cat["inclusion_status"] = cat.apply(status, axis=1)
    return cat, df


def build_tables(df):
    valid, noinfo = df[df.is_valid_cwe], df[~df.is_valid_cwe]
    N, NV, NN = len(df), len(valid), len(noinfo)
    g = (valid.groupby(["cwe", "cwe_name_en", "cwe_name_fa", "cwe_category"], as_index=False)
         .agg(count=("cve_id", "count"), mean_cvss=("cvss_score", "mean"), median_cvss=("cvss_score", "median"),
              min_cvss=("cvss_score", "min"), max_cvss=("cvss_score", "max"),
              products_affected=("product", lambda s: ", ".join(sorted(s.unique()))),
              cve_list=("cve_id", lambda s: ", ".join(s)))
         .sort_values(["count", "mean_cvss", "cwe"], ascending=[False, False, True]).reset_index(drop=True))
    g["rank"] = g.index + 1
    g["share_of_valid_cwe_pct"] = (g["count"] / NV * 100).round(2)
    g["share_of_total_cve_pct"] = (g["count"] / N * 100).round(2)
    g["cumulative_valid_pct"] = (g["count"].cumsum() / NV * 100).round(2)
    for c in ["mean_cvss", "median_cvss"]: g[c] = g[c].round(2)
    g = g[["rank", "cwe", "cwe_name_en", "cwe_name_fa", "cwe_category", "count", "share_of_valid_cwe_pct",
           "share_of_total_cve_pct", "cumulative_valid_pct", "mean_cvss", "median_cvss", "min_cvss", "max_cvss",
           "products_affected", "cve_list"]]
    noinfo_t = noinfo[["cve_id", "product", "group", "published", "cwe", "cwe_name_en", "cvss_score", "severity",
                       "advisory_url"]].reset_index(drop=True)
    m = lambda s: round(float(s.mean()), 2) if len(s) else None
    summary = pd.DataFrame([
        ["Valid Specific CWE Assigned",
         f"{g.shape[0]} Distinct Valid CWEs (Top: " + ", ".join(g.cwe.head(5)) + ", ...)", NV, N,
         round(NV / N * 100, 2), m(valid.cvss_score), f"Included in Top CWE valid denominator (n_valid = {NV})"],
        ["Unspecified / Insufficient Info (NVD-CWE-noinfo)", NOINFO, NN, N, round(NN / N * 100, 2), m(noinfo.cvss_score),
         "Reported separately per RQ2 methodology (" + ", ".join(f"{v} {k}" for k, v in noinfo["product"].value_counts().items()) + ")"],
        ["Completely Empty / Null CWE Field", "NULL / NaN", int(df.cwe.isna().sum()), N, round(df.cwe.isna().sum() / N * 100, 2), None,
         "Zero blank cells; all unspecified CWEs are explicitly tagged as NVD-CWE-noinfo"],
        ["Total Dataset (All Analyzed CVEs)", f"All Verified Records ({df.published_year.min()}-{df.published_year.max()})", N, N, 100.0,
         m(df.cvss_score), f"Full verified sample across {df['product'].nunique()} GenAI products/components (N = {N})"],
    ], columns=["Category", "CWE_Code_Status", "Record_Count", "Denominator_Total_N", "Share_of_Total_Pct", "Mean_CVSS", "Notes"])
    rows = [["Initial Candidate Rows Examined (2023-2026)", INITIAL_CANDIDATE_ROWS, INITIAL_CANDIDATE_ROWS,
             "All candidate rows from base dataset, user lists and NVD extraction"],
            ["Rejected / False-Positive CVEs Excluded (Step 3)", len(EXCLUDED_FALSE_POSITIVES), INITIAL_CANDIDATE_ROWS,
             "; ".join(f"{c} ({p}): {why}" for c, p, why in EXCLUDED_FALSE_POSITIVES)],
            ["Duplicate CVE Rows Deduplicated (Step 3)", len(DUPLICATES_REMOVED), INITIAL_CANDIDATE_ROWS,
             "; ".join(f"{c} {why}" for c, why in DUPLICATES_REMOVED)],
            ["Final Clean Analytical Sample (N)", N, N, "Denominator for overall dataset metrics (2023-2026)"],
            ["Records with Valid Specific CWE (n_valid)", NV, N, f"Primary denominator for RQ2 share calculation ({NV} records across {g.shape[0]} distinct CWEs)"],
            ["Records without Specific CWE (NVD-CWE-noinfo)", NN, N, f"Reported separately; excluded from valid CWE denominator ({NN} records)"]]
    audit = pd.DataFrame([[a, c, d, f"{c / d * 100:.2f}%", EXTRACTION_DATE, desc] for a, c, d, desc in rows],
                         columns=["Audit_Metric", "Count", "Denominator", "Percentage", "Extraction_Date", "Description"])
    # sanity checks (the build fails loudly instead of shipping inconsistent numbers)
    assert INITIAL_CANDIDATE_ROWS - len(EXCLUDED_FALSE_POSITIVES) - len(DUPLICATES_REMOVED) == N
    assert g["count"].sum() == NV and NV + NN == N
    return g, noinfo_t, summary, audit


def yearly_table(df):
    rows = []
    for y, s in df.groupby("published_year"):
        v = s[s.is_valid_cwe]
        top = v.cwe.value_counts()
        dom = ", ".join(f"{c} ({n})" for c, n in top.head(5).items() if n >= 1)[:200]
        rows.append([int(y), len(s), len(v), int((~s.is_valid_cwe).sum()), round(len(s) / len(df) * 100, 2),
                     round(float(s.cvss_score.mean()), 2), dom])
    return pd.DataFrame(rows, columns=["Published Year", "Total CVEs", "Valid CWE Count", "NVD-CWE-noinfo Count",
                                       "Share of Total (%)", "Mean CVSS", "Dominant CWEs in Year"])


def derived_stats(df, g):
    N = len(df); valid = df[df.is_valid_cwe]; NV = len(valid); noinfo = df[~df.is_valid_cwe]
    top5 = g.head(5); rec = g[g["count"] >= 2]
    inj = g[g.cwe.isin(INJECTION_FAMILY)]
    nov4 = df[df.cvss_version == "3.1"]
    s = dict(N=N, NV=NV, NN=len(noinfo), ndist=len(g),
             top5_n=int(top5["count"].sum()), top5_pct=round(top5["count"].sum() / NV * 100, 2),
             top4_n=int(g.head(4)["count"].sum()), top4_pct=round(g.head(4)["count"].sum() / NV * 100, 2),
             rec_n_cwe=len(rec), rec_n=int(rec["count"].sum()), rec_pct=round(rec["count"].sum() / NV * 100, 2),
             inj_n=int(inj["count"].sum()), inj_pct=round(inj["count"].sum() / NV * 100, 2),
             mean_all=round(df.cvss_score.mean(), 2), mean_valid=round(valid.cvss_score.mean(), 2),
             mean_noinfo=round(noinfo.cvss_score.mean(), 2),
             mean_all_v31=round(nov4.cvss_score.mean(), 2), mean_valid_v31=round(nov4[nov4.is_valid_cwe].cvss_score.mean(), 2),
             n_v4=int((df.cvss_version == "4.0").sum()),
             tie_count=int((g["count"] == 2).sum()))
    return s


# ----------------------------------------------------------------------------
# Charts
# ----------------------------------------------------------------------------
PALETTE = ["#00F5D4", "#00BBF9", "#F15BB5", "#FEE440", "#9B5DE5", "#FF9F1C", "#38B000", "#FF006E", "#4EA8DE", "#94A3B8"]


def _style_dark(ax):
    ax.set_facecolor("#181B28")
    for s in ax.spines.values(): s.set_color("#2A2F45")
    ax.tick_params(colors="#E5E7EB"); ax.grid(axis="x", ls="--", alpha=.25)


def chart_bars(g, st, df, out_dark, out_light):
    rec = g[g["count"] >= 2].copy()
    lab = [f"#{r['rank']} {r.cwe} | {r.cwe_category}" for _, r in rec.iterrows()]
    colors = ["#00F5D4" if r <= 5 else "#F15BB5" for r in rec["rank"]]
    N, NV, NN = st["N"], st["NV"], st["NN"]
    title = f"RQ2: Top 5 + All Recurring Weakness Types (CWEs with >=2 CVEs: {st['rec_n_cwe']} CWEs) | Valid n = {NV}"
    # ---- dark (presentation) ----
    fig = plt.figure(figsize=(17, 8.5), facecolor="#10131C")
    gs = fig.add_gridspec(1, 2, width_ratios=[3.2, 1], wspace=.12)
    ax = fig.add_subplot(gs[0]); _style_dark(ax)
    y = np.arange(len(rec))[::-1]
    ax.barh(y, rec["count"], color=colors, edgecolor="white", lw=.6, height=.65)
    for yi, (_, r) in zip(y, rec.iterrows()):
        ax.text(r["count"] + .25, yi, f"{r['count']} CVEs ({r.share_of_valid_cwe_pct:.2f}% valid | {r.share_of_total_cve_pct:.2f}% total | Cum: {r.cumulative_valid_pct:.2f}% | Mean CVSS: {r.mean_cvss:.2f})",
                va="center", color="white", fontsize=8.5, fontweight="bold")
    ax.set_yticks(y); ax.set_yticklabels(lab, color="#E5E7EB", fontsize=9)
    ax.set_xlim(0, max(rec["count"]) * 2.1)
    ax.set_xlabel(f"Number of Verified CVEs (Valid CWE Sample n = {NV} | Total N = {N})", color="#00F5D4", fontweight="bold")
    ax.set_title(title, color="white", fontweight="bold", loc="left", fontsize=12)
    ax.legend(handles=[mpatches.Patch(color="#00F5D4", label=f"Top 5 dominant ({st['top5_n']} CVEs = {st['top5_pct']:.2f}% of valid)"),
                       mpatches.Patch(color="#F15BB5", label=f"Recurring, 2 CVEs each (ties ordered by mean CVSS)")],
              loc="lower right", facecolor="#181B28", labelcolor="#E5E7EB", fontsize=9)
    ax2 = fig.add_subplot(gs[1]); _style_dark(ax2); ax2.grid(axis="y", ls="--", alpha=.25); ax2.grid(axis="x", visible=False)
    ax2.bar([0, 1], [NV, NN], color=["#00B4EB", "#EC0078"], edgecolor="white", width=.5)
    ax2.text(0, NV + 1, f"n = {NV}\n({NV / N * 100:.2f}%)", ha="center", color="white", fontweight="bold")
    ax2.text(1, NN + 1, f"n = {NN}\n({NN / N * 100:.2f}%)", ha="center", color="white", fontweight="bold")
    ax2.set_xticks([0, 1]); ax2.set_xticklabels([f"Valid Specific CWE\n({st['ndist']} distinct)", "Unspecified\n(NVD-CWE-noinfo)"], color="#E5E7EB")
    ax2.set_ylim(0, NV * 1.18); ax2.set_ylabel(f"Number of CVE Records (N = {N})", color="#00BBF9", fontweight="bold")
    ax2.set_title(f"CWE Coverage & Missingness\n(Total N = {N})", color="white", fontweight="bold")
    fig.suptitle("Generative AI Vulnerability Analysis (2023-2026) - RQ2: Top CWE Distribution & Completeness Audit",
                 color="#00F5D4", fontsize=14, fontweight="bold", y=.985)
    fig.text(.5, .01, f"Source: Verified NVD records | N = {N} | n_valid = {NV} ({NV / N * 100:.2f}%) | n_noinfo = {NN} ({NN / N * 100:.2f}%) | Extraction date: {EXTRACTION_DATE}",
             ha="center", color="#94A3B8", fontsize=9)
    fig.savefig(out_dark, dpi=170, facecolor=fig.get_facecolor(), bbox_inches="tight"); plt.close(fig)
    # ---- light (print) ----
    fig, ax = plt.subplots(figsize=(13, 7.2))
    ax.barh(y, rec["count"], color=["#1F4E79" if r <= 5 else "#8FAADC" for r in rec["rank"]], edgecolor="black", lw=.5, height=.65)
    for yi, (_, r) in zip(y, rec.iterrows()):
        ax.text(r["count"] + .2, yi, f"{r['count']} ({r.share_of_valid_cwe_pct:.1f}%)", va="center", fontsize=9)
    ax.set_yticks(y); ax.set_yticklabels([f"{r.cwe} - {r.cwe_name_en.split(' (')[0][:48]}" for _, r in rec.iterrows()], fontsize=9)
    ax.set_xlim(0, max(rec["count"]) * 1.2); ax.grid(axis="x", ls=":", alpha=.5)
    ax.set_xlabel(f"Number of CVEs (share of n_valid = {NV})"); ax.set_title(f"RQ2: Top 5 + all recurring CWEs in Generative AI products, 2023-2026 (n_valid = {NV}, N = {N})", fontweight="bold")
    fig.text(.01, .005, f"{NN} records with NVD-CWE-noinfo are excluded from the valid-CWE denominator. Dark = Top 5, light = other CWEs with 2 CVEs (ties). Extraction: {EXTRACTION_DATE}.", fontsize=8, color="#444")
    fig.tight_layout(rect=(0, .02, 1, 1)); fig.savefig(out_light, dpi=170); plt.close(fig)


def chart_breakdown(g, st, df, out):
    rec = g[g["count"] >= 2]; valid = df[df.is_valid_cwe]
    top_prod = list(valid["product"].value_counts().head(9).index)
    N, NV, NN = st["N"], st["NV"], st["NN"]
    fig = plt.figure(figsize=(17, 10), facecolor="#10131C")
    gs = fig.add_gridspec(2, 2, height_ratios=[4.2, 1], width_ratios=[1.6, 1], hspace=.38, wspace=.25)
    a1, a2 = fig.add_subplot(gs[0, 0]), fig.add_subplot(gs[0, 1])
    a3, a4 = fig.add_subplot(gs[1, 0]), fig.add_subplot(gs[1, 1])
    for a in (a1, a2, a3, a4): _style_dark(a)
    cmap = {p: PALETTE[i] for i, p in enumerate(top_prod)}; cmap["other"] = "#94A3B8"
    y = np.arange(len(rec))[::-1]
    for yi, cw in zip(y, rec.cwe):
        s = valid[valid.cwe == cw]["product"].map(lambda p: p if p in top_prod else "other").value_counts()
        left = 0
        for p in top_prod + ["other"]:
            if p in s: a1.barh(yi, s[p], left=left, color=cmap[p], edgecolor="#10131C", height=.62); left += s[p]
        a1.text(left + .2, yi, f"{left} CVEs", va="center", color="white", fontweight="bold", fontsize=9)
    a1.set_yticks(y); a1.set_yticklabels(rec.cwe, color="#E5E7EB"); a1.set_xlim(0, rec["count"].max() * 1.25)
    a1.set_xlabel(f"CVE Count by Product ({st['rec_n_cwe']} recurring CWEs: {st['rec_n']} of {NV} valid CVEs)", color="#00F5D4", fontweight="bold")
    a1.set_title("Top 5 + recurring CWEs by product (2023-2026)", color="white", fontweight="bold", loc="left")
    a1.legend(handles=[mpatches.Patch(color=cmap[p], label=p) for p in top_prod + ["other"]], title="Product", ncol=2,
              facecolor="#181B28", labelcolor="#E5E7EB", fontsize=8, loc="lower right")
    a2.barh(y, rec["mean_cvss"], color="#00A8E0", edgecolor="white", height=.62)
    for yi, v in zip(y, rec["mean_cvss"]): a2.text(v + .1, yi, f"{v:.2f}", va="center", color="white", fontweight="bold", fontsize=9)
    a2.axvline(9, color="#F15BB5", ls="--"); a2.set_yticks(y); a2.set_yticklabels(rec.cwe, color="#E5E7EB"); a2.set_xlim(0, 11.5)
    a2.set_xlabel("Mean CVSS Base Score", color="#00BBF9", fontweight="bold")
    a2.set_title("Mean CVSS across recurring CWEs", color="white", fontweight="bold", loc="left")
    a2.legend(handles=[plt.Line2D([], [], color="#F15BB5", ls="--", label="Critical threshold (9.0)")], facecolor="#181B28", labelcolor="#E5E7EB", fontsize=8, loc="lower right")
    nn = df[~df.is_valid_cwe]; left = 0
    for i, (p, c) in enumerate(nn["product"].value_counts().items()):
        a3.barh(0, c, left=left, color=PALETTE[i], edgecolor="#10131C"); left += c
    a3.text(left + .2, 0, f"{NN} CVEs (" + ", ".join(f"{c} {p}" for p, c in nn["product"].value_counts().items()) + ")", va="center", color="#FEE440", fontsize=9, fontweight="bold")
    a3.set_yticks([0]); a3.set_yticklabels(["NVD-CWE-noinfo\n(Unspecified)"], color="#FEE440"); a3.set_xlim(0, rec["count"].max() * 1.25)
    a3.set_xlabel(f"CVE Count (Unspecified CWE sample: n = {NN})", color="#FEE440", fontweight="bold")
    a3.set_title(f"Separated unspecified records (excluded from ranking, n = {NN})", color="#FEE440", fontweight="bold", loc="left")
    a4.barh(0, st["mean_noinfo"], color="#E6007A", edgecolor="white"); a4.axvline(9, color="#F15BB5", ls="--")
    a4.text(st["mean_noinfo"] + .1, 0, f"{st['mean_noinfo']:.2f}", va="center", color="white", fontweight="bold")
    a4.set_xlim(0, 11.5); a4.set_yticks([0]); a4.set_yticklabels(["NVD-CWE-noinfo"], color="#FEE440")
    a4.set_xlabel("Mean CVSS Base Score", color="#E6007A", fontweight="bold"); a4.set_title(f"Mean CVSS (noinfo, n = {NN})", color="#FEE440", fontweight="bold", loc="left")
    fig.suptitle(f"RQ2 Extended View: Recurring Valid CWEs (n = {NV}) vs. Separated Unspecified CWEs (n = {NN}) by Product & Severity",
                 color="#00F5D4", fontsize=14, fontweight="bold", y=.97)
    fig.savefig(out, dpi=150, facecolor=fig.get_facecolor(), bbox_inches="tight"); plt.close(fig)


def chart_yearly(g, st, df, out):
    N, NV, NN = st["N"], st["NV"], st["NN"]
    yrs = sorted(df.published_year.unique()); top5 = list(g.cwe.head(5))
    fig, (a1, a2) = plt.subplots(1, 2, figsize=(15, 7.8), facecolor="#10131C")
    for a in (a1, a2): _style_dark(a); a.grid(axis="y", ls="--", alpha=.25); a.grid(axis="x", visible=False)
    vc = [int(((df.published_year == y) & df.is_valid_cwe).sum()) for y in yrs]
    nc = [int(((df.published_year == y) & ~df.is_valid_cwe).sum()) for y in yrs]
    x = np.arange(len(yrs))
    a1.bar(x, vc, .5, color="#00F5D4", edgecolor="white", label=f"Valid Specific CWE (n={NV})")
    a1.bar(x, nc, .5, bottom=vc, color="#FF006E", edgecolor="white", label=f"NVD-CWE-noinfo (n={NN})")
    for xi, v, n in zip(x, vc, nc): a1.text(xi, v + n + .5, f"Total: {v + n}\n(Valid: {v})", ha="center", color="white", fontweight="bold", fontsize=10)
    a1.set_xticks(x); a1.set_xticklabels(yrs, color="#E5E7EB"); a1.set_ylim(0, max(np.add(vc, nc)) * 1.3)
    a1.set_ylabel(f"Number of Verified CVEs (N = {N})", color="#00F5D4", fontweight="bold"); a1.set_xlabel("NVD Published Year", color="#00F5D4", fontweight="bold")
    a1.set_title(f"Temporal Distribution of Verified CVEs by Year (N = {N})", color="white", fontweight="bold"); a1.legend(facecolor="#181B28", labelcolor="#E5E7EB")
    bottoms = np.zeros(len(yrs))
    for i, cw in enumerate(top5):
        vals = np.array([int(((df.published_year == y) & (df.cwe == cw)).sum()) for y in yrs])
        a2.bar(x, vals, .5, bottom=bottoms, color=PALETTE[i], edgecolor="#10131C", label=cw); bottoms += vals
    for xi, b in zip(x, bottoms): a2.text(xi, b + .2, f"{int(b)} CVEs", ha="center", color="white", fontweight="bold")
    a2.set_xticks(x); a2.set_xticklabels(yrs, color="#E5E7EB"); a2.set_ylim(0, max(bottoms) * 1.3)
    a2.set_ylabel(f"CVE count across Top 5 CWEs (= {st['top5_n']} of {NV})", color="#00BBF9", fontweight="bold"); a2.set_xlabel("NVD Published Year", color="#00BBF9", fontweight="bold")
    a2.set_title("Evolution of the Top 5 Dominant Weakness Types", color="white", fontweight="bold"); a2.legend(facecolor="#181B28", labelcolor="#E5E7EB")
    fig.suptitle(f"Multi-Year Evolution of Generative AI Vulnerabilities (2023-2026 | N = {N})", color="#00F5D4", fontsize=15, fontweight="bold")
    fig.savefig(out, dpi=150, facecolor=fig.get_facecolor(), bbox_inches="tight"); plt.close(fig)


# ----------------------------------------------------------------------------
# Excel workbook (formulas are live; recalculate with LibreOffice/Excel)
# ----------------------------------------------------------------------------
def build_xlsx(path, cat, df, g, noinfo_t, summary, audit, yearly, st):
    wb = openpyxl.Workbook()
    F = "Arial"; hdr_fill = PatternFill("solid", fgColor="1F4E79"); thin = Side(style="thin", color="BBBBBB")
    bd = Border(left=thin, right=thin, top=thin, bottom=thin)
    def head(ws, row, ncol):
        for c in range(1, ncol + 1):
            cell = ws.cell(row, c); cell.font = Font(name=F, bold=True, color="FFFFFF"); cell.fill = hdr_fill
            cell.alignment = Alignment(wrap_text=True, vertical="center", horizontal="center"); cell.border = bd
    def body(ws, r0, r1, ncol):
        for r in range(r0, r1 + 1):
            for c in range(1, ncol + 1):
                cell = ws.cell(r, c); cell.border = bd; cell.font = Font(name=F, size=10)
                cell.alignment = Alignment(vertical="center", wrap_text=isinstance(cell.value, str) and len(str(cell.value)) > 30)
    N = len(df)
    # --- dataset sheet first (others reference it) ---
    wsd = wb.active; wsd.title = "Cleaned_Dataset_Step4"
    wsd["A1"] = f"Step 4: Verified & Enriched Generative AI CVE Dataset (2023-2026 | N = {N}, Extraction Date: {EXTRACTION_DATE})"
    wsd["A1"].font = Font(name=F, bold=True, size=12)
    cols = list(df.columns)
    for j, c in enumerate(cols, 1): wsd.cell(3, j, c)
    head(wsd, 3, len(cols))
    for i, row in enumerate(df.itertuples(index=False), 4):
        for j, v in enumerate(row, 1): wsd.cell(i, j, v if not isinstance(v, (np.bool_,)) else bool(v))
    last = 3 + N; body(wsd, 4, last, len(cols))
    for j, w in enumerate([17, 24, 36, 12, 8, 15, 40, 40, 28, 9, 8, 9, 10, 12, 12, 12, 9, 7, 9, 44, 46], 1):
        wsd.column_dimensions[get_column_letter(j)].width = w
    wsd.freeze_panes = "B4"
    rng = lambda col: f"Cleaned_Dataset_Step4!${col}$4:${col}${last}"
    # --- Top valid CWEs ---
    ws = wb.create_sheet("RQ2_Top_Valid_CWEs", 0)
    ws["A1"] = f"RQ2: Weakness distribution & most frequent CWEs in GenAI products (2023-2026 | n_valid = {st['NV']} | N = {N})"
    ws["A1"].font = Font(name=F, bold=True, size=12)
    ws["A2"] = ("Counts, shares, means, min and max are live formulas over sheet Cleaned_Dataset_Step4. Rank order and median are written by the build script "
                "(sort: count desc, mean CVSS desc, CWE id asc). Ties at 2 CVEs are all listed.")
    ws["A2"].font = Font(name=F, italic=True, size=9)
    hd = ["Rank", "CWE ID", "MITRE CWE Name (English)", "عنوان فارسی", "Category", "CVE Count", "Share of Valid CWEs (%)",
          "Share of Total CVEs (%)", "Cumulative Valid Share (%)", "Mean CVSS", "Median CVSS (static)", "Min CVSS", "Max CVSS", "Products", "Tier"]
    for j, h in enumerate(hd, 1): ws.cell(3, j, h)
    head(ws, 3, len(hd))
    r0 = 4; r1 = r0 + len(g) - 1; tot = r1 + 1
    for i, r in enumerate(g.itertuples(index=False)):
        rr = r0 + i
        ws.cell(rr, 1, int(r.rank)); ws.cell(rr, 2, r.cwe); ws.cell(rr, 3, r.cwe_name_en); ws.cell(rr, 4, r.cwe_name_fa); ws.cell(rr, 5, r.cwe_category)
        ws.cell(rr, 6, f"=COUNTIFS({rng('F')},B{rr})")
        ws.cell(rr, 7, f"=ROUND(F{rr}/$F${tot}*100,2)")
        ws.cell(rr, 8, f"=ROUND(F{rr}/COUNTA({rng('A')})*100,2)")
        ws.cell(rr, 9, f"=ROUND(SUM($F${r0}:F{rr})/$F${tot}*100,2)")
        ws.cell(rr, 10, f"=ROUND(AVERAGEIFS({rng('L')},{rng('F')},B{rr}),2)")
        ws.cell(rr, 11, float(r.median_cvss))
        ws.cell(rr, 12, f"=_xlfn.MINIFS({rng('L')},{rng('F')},B{rr})")
        ws.cell(rr, 13, f"=_xlfn.MAXIFS({rng('L')},{rng('F')},B{rr})")
        ws.cell(rr, 14, r.products_affected)
        ws.cell(rr, 15, "Top 5" if r.rank <= 5 else ("Recurring (>=2)" if r.count >= 2 else "Single"))
    ws.cell(tot, 2, "TOTAL VALID"); ws.cell(tot, 3, f"All records with a valid specific CWE ({len(g)} distinct CWEs)")
    ws.cell(tot, 6, f"=SUM(F{r0}:F{r1})"); ws.cell(tot, 7, f"=ROUND(F{tot}/$F${tot}*100,2)")
    ws.cell(tot, 8, f"=ROUND(F{tot}/COUNTA({rng('A')})*100,2)")
    ws.cell(tot, 10, f"=ROUND(AVERAGEIFS({rng('L')},{rng('F')},\"<>{NOINFO}\"),2)")
    body(ws, r0, tot, len(hd))
    for c in range(1, len(hd) + 1): ws.cell(tot, c).font = Font(name=F, bold=True, size=10)
    for j, w in enumerate([6, 11, 50, 34, 28, 9, 12, 12, 13, 9, 11, 8, 8, 34, 16], 1): ws.column_dimensions[get_column_letter(j)].width = w
    ws.freeze_panes = "C4"
    ch = BarChart(); ch.type = "bar"; ch.title = "CVE count per recurring CWE"; ch.style = 10; ch.height = 9; ch.width = 18
    nrec = int((g["count"] >= 2).sum())
    ch.add_data(Reference(ws, min_col=6, min_row=3, max_row=r0 + nrec - 1), titles_from_data=True)
    ch.set_categories(Reference(ws, min_col=2, min_row=r0, max_row=r0 + nrec - 1)); ch.legend = None
    ch.y_axis.scaling.orientation = "minMax"; ch.x_axis.scaling.orientation = "maxMin"
    ws.add_chart(ch, f"B{tot + 3}")
    # --- Missing / noinfo ---
    wm = wb.create_sheet("Missing_NoInfo_CWEs", 1)
    wm["A1"] = f"RQ2 & Step 5 audit: CWE completeness and separate report of unspecified CWEs (Extraction Date: {EXTRACTION_DATE})"
    wm["A1"].font = Font(name=F, bold=True, size=12)
    wm["A3"] = "Part A: CWE completeness (live formulas over Cleaned_Dataset_Step4)"; wm["A3"].font = Font(name=F, bold=True)
    for j, h in enumerate(["Category", "CWE_Code_Status", "Record_Count", "Denominator_Total_N", "Share_of_Total_Pct", "Mean_CVSS"], 1): wm.cell(4, j, h)
    head(wm, 4, 6)
    A = rng("A"); Fc = rng("F"); L = rng("L")
    rowsA = [("Valid Specific CWE Assigned", f"{len(g)} distinct valid CWEs", f"=COUNTA({A})-COUNTIFS({Fc},\"{NOINFO}\")", f"=ROUND(AVERAGEIFS({L},{Fc},\"<>{NOINFO}\"),2)"),
             ("Unspecified (NVD-CWE-noinfo)", NOINFO, f"=COUNTIFS({Fc},\"{NOINFO}\")", f"=ROUND(AVERAGEIFS({L},{Fc},\"{NOINFO}\"),2)"),
             ("Total dataset", "All verified records", f"=COUNTA({A})", f"=ROUND(AVERAGE({L}),2)")]
    for i, (a, b, cnt_f, mean_f) in enumerate(rowsA, 5):
        wm.cell(i, 1, a); wm.cell(i, 2, b); wm.cell(i, 3, cnt_f); wm.cell(i, 4, "=$C$7"); wm.cell(i, 5, f"=ROUND(C{i}/D{i}*100,2)"); wm.cell(i, 6, mean_f)
    body(wm, 5, 7, 6)
    wm["A10"] = f"Part B: records without a specific CWE (NVD-CWE-noinfo, n = {len(noinfo_t)})"; wm["A10"].font = Font(name=F, bold=True)
    for j, h in enumerate(noinfo_t.columns, 1): wm.cell(11, j, h)
    head(wm, 11, len(noinfo_t.columns))
    for i, row in enumerate(noinfo_t.itertuples(index=False), 12):
        for j, v in enumerate(row, 1): wm.cell(i, j, v)
    eB = 11 + len(noinfo_t); body(wm, 12, eB, len(noinfo_t.columns))
    pa = eB + 3
    wm.cell(pa - 1, 1, "Part C: Step 5 data audit and denominator disclosure").font = Font(name=F, bold=True)
    for j, h in enumerate(audit.columns, 1): wm.cell(pa, j, h)
    head(wm, pa, len(audit.columns))
    for i, row in enumerate(audit.itertuples(index=False), pa + 1):
        for j, v in enumerate(row, 1): wm.cell(i, j, v)
    body(wm, pa + 1, pa + len(audit), len(audit.columns))
    for j, w in enumerate([46, 34, 14, 22, 16, 30, 40, 14, 44], 1): wm.column_dimensions[get_column_letter(j)].width = w
    # --- Catalog ---
    wc = wb.create_sheet("Product_Catalog_Step1")
    wc["A1"] = f"Step 1: Selected Generative AI & AI Agent products metadata catalog ({len(cat)} products / components)"; wc["A1"].font = Font(name=F, bold=True, size=12)
    for j, h in enumerate(cat.columns, 1): wc.cell(3, j, h)
    head(wc, 3, len(cat.columns))
    for i, row in enumerate(cat.itertuples(index=False), 4):
        for j, v in enumerate(row, 1): wc.cell(i, j, v)
    body(wc, 4, 3 + len(cat), len(cat.columns))
    for j, w in enumerate([24, 28, 26, 40, 48, 38, 40, 46], 1): wc.column_dimensions[get_column_letter(j)].width = w
    # --- Yearly ---
    wy = wb.create_sheet("Yearly_Summary")
    wy["A1"] = f"Multi-year trend summary of GenAI vulnerabilities (Total N = {N})"; wy["A1"].font = Font(name=F, bold=True, size=12)
    for j, h in enumerate(yearly.columns, 1): wy.cell(3, j, h)
    head(wy, 3, len(yearly.columns))
    E = rng("E")
    for i, r in enumerate(yearly.itertuples(index=False), 4):
        wy.cell(i, 1, int(r[0]))
        wy.cell(i, 2, f"=COUNTIFS({E},A{i})")
        wy.cell(i, 3, f"=COUNTIFS({E},A{i},{Fc},\"<>{NOINFO}\")")
        wy.cell(i, 4, f"=COUNTIFS({E},A{i},{Fc},\"{NOINFO}\")")
        wy.cell(i, 5, f"=ROUND(B{i}/COUNTA({A})*100,2)")
        wy.cell(i, 6, f"=ROUND(AVERAGEIFS({L},{E},A{i}),2)")
        wy.cell(i, 7, r[6])
    body(wy, 4, 3 + len(yearly), 7)
    for j, w in enumerate([14, 12, 15, 20, 16, 11, 70], 1): wy.column_dimensions[get_column_letter(j)].width = w
    wb.save(path)


# ----------------------------------------------------------------------------
# Word report (Persian, RTL)
# ----------------------------------------------------------------------------
def build_docx(path, cat, df, g, noinfo_t, audit, yearly, st, chart_files):
    from docx import Document
    from docx.shared import Pt, Cm, RGBColor
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.enum.table import WD_TABLE_ALIGNMENT
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement
    N, NV, NN = st["N"], st["NV"], st["NN"]
    doc = Document()
    sec = doc.sections[0]; sec.page_width, sec.page_height = Cm(21), Cm(29.7)
    for a in ("left_margin", "right_margin"): setattr(sec, a, Cm(1.8))
    sec.top_margin = sec.bottom_margin = Cm(2)
    FONT = "Arial"
    stl = doc.styles["Normal"]; stl.font.name = FONT; stl.font.size = Pt(10.5)
    stl.element.rPr.rFonts.set(qn("w:cs"), FONT)

    def rtl_par(p, align=WD_ALIGN_PARAGRAPH.RIGHT):
        pPr = p._p.get_or_add_pPr(); b = OxmlElement("w:bidi"); pPr.append(b); p.alignment = align
    def run(p, text, bold=False, size=None, color=None, italic=False):
        r = p.add_run(text); r.bold = bold; r.italic = italic
        rPr = r._r.get_or_add_rPr(); rt = OxmlElement("w:rtl"); rPr.append(rt)
        rf = rPr.find(qn("w:rFonts"))
        if rf is None: rf = OxmlElement("w:rFonts"); rPr.append(rf)
        rf.set(qn("w:cs"), FONT); rf.set(qn("w:ascii"), FONT); rf.set(qn("w:hAnsi"), FONT)
        if size: r.font.size = Pt(size)
        if color: r.font.color.rgb = RGBColor.from_string(color)
        if size or bold:
            for tag in ("w:bCs", "w:szCs"):
                pass
        return r
    def para(text, bold=False, size=None, color=None, align=WD_ALIGN_PARAGRAPH.RIGHT, space_after=6, bullet=False):
        p = doc.add_paragraph(); rtl_par(p, align); p.paragraph_format.space_after = Pt(space_after)
        if bullet: text = "• " + text
        run(p, text, bold, size, color); return p
    def h1(text): 
        p = para(text, True, 14, "1F4E79", space_after=8); p.paragraph_format.space_before = Pt(14); p.paragraph_format.keep_with_next = True
        p._p.get_or_add_pPr().append(OxmlElement("w:outlineLvl")); p._p.pPr.findall(qn("w:outlineLvl"))[0].set(qn("w:val"), "0"); return p
    def shade(cell, color):
        tcPr = cell._tc.get_or_add_tcPr(); sh = OxmlElement("w:shd"); sh.set(qn("w:val"), "clear"); sh.set(qn("w:color"), "auto"); sh.set(qn("w:fill"), color); tcPr.append(sh)
    def table(headers, rows, widths, font=8.5):
        t = doc.add_table(rows=1, cols=len(headers)); t.style = "Table Grid"; t.alignment = WD_TABLE_ALIGNMENT.CENTER
        tblPr = t._tbl.tblPr; bv = OxmlElement("w:bidiVisual"); tblPr.find(qn("w:tblStyle")).addnext(bv)
        for i, h in enumerate(headers):
            c = t.rows[0].cells[i]; c.text = ""; p = c.paragraphs[0]; rtl_par(p, WD_ALIGN_PARAGRAPH.CENTER); run(p, h, True, font, "FFFFFF"); shade(c, "1F4E79")
        for r in rows:
            cells = t.add_row().cells
            for i, v in enumerate(r):
                cells[i].text = ""; p = cells[i].paragraphs[0]
                latin = isinstance(v, (int, float)) or str(v).replace(".", "").replace("%", "").replace("-", "").replace(",", "").isalnum() and str(v).isascii()
                rtl_par(p, WD_ALIGN_PARAGRAPH.CENTER if latin else WD_ALIGN_PARAGRAPH.RIGHT); run(p, str(v), False, font)
        t.autofit = False
        for i, w in enumerate(widths): t.columns[i].width = Cm(w)
        for row in t.rows:
            for i, w in enumerate(widths): row.cells[i].width = Cm(w)
        trPr = t.rows[0]._tr.get_or_add_trPr(); th = OxmlElement("w:tblHeader"); trPr.append(th)
        doc.add_paragraph().paragraph_format.space_after = Pt(4); return t

    top5 = g.head(5); nrec = st["rec_n_cwe"]
    para("پروژه تحلیل آماری آسیب‌پذیری‌های Generative AI و AI Agent (گروه ۵)", True, 17, "1F4E79", WD_ALIGN_PARAGRAPH.CENTER, 4)
    para(f"گزارش روش‌شناختی و آماری سؤال پژوهشی دوم (RQ2): پرتکرارترین انواع ضعف امنیتی (CWE) در بازه ۲۰۲۳ تا ۲۰۲۶ (N = {fa(N)})", True, 12, None, WD_ALIGN_PARAGRAPH.CENTER, 10)

    h1("۱. مشخصات پایه، ممیزی داده‌ها و گزارش مخرج‌ها (گام پنجم)")
    para("مطابق راهنمای پروژه، هیچ عددی بدون مخرج و تاریخ برداشت گزارش نمی‌شود. تمام اعداد این گزارش مستقیماً از فایل دیتاست نهایی (فایل ۲) و اسکریپت بازتولیدپذیر پروژه محاسبه شده‌اند.")
    para(f"تاریخ برداشت و جمع‌بندی داده: ۱۲ نوامبر ۲۰۲۶ ({EXTRACTION_DATE})", bullet=True)
    para(f"سطرهای کاندید اولیه: {fa(INITIAL_CANDIDATE_ROWS)} سطر (دیتاست پایه، فهرست‌های پیشنهادی و استخراج NVD)", bullet=True)
    for c, p_, why in EXCLUDED_FALSE_POSITIVES:
        para(f"حذف رکورد نامرتبط (False Positive) در گام سوم: {c} ({p_}) مربوط به Netflix Genie است، یک سامانه هماهنگ‌سازی کارهای کلان‌داده (Hadoop/Spark) که محصول هوش مصنوعی مولد نیست؛ بنابراین از دیتاست و کاتالوگ حذف شد.", bullet=True)
    for c, why in DUPLICATES_REMOVED:
        para(f"حذف سطر تکراری در گام سوم: {c} یک‌بار با CWE-22 و یک‌بار با CWE-31 ثبت شده بود و به یک رکورد یکتا با CWE اصلی CWE-22 تجمیع شد. همچنین فرمت تاریخ CVE-2024-31621 و CVE-2024-2965 استانداردسازی شد.", bullet=True)
    para(f"حجم نمونه نهایی: N = {fa(N)} رکورد یکتا در {fa(df['product'].nunique())} محصول/ابزار GenAI ({fa(INITIAL_CANDIDATE_ROWS)} − {fa(len(EXCLUDED_FALSE_POSITIVES))} − {fa(len(DUPLICATES_REMOVED))} = {fa(N)})", bullet=True)
    yc = df.groupby("published_year").size(); idc = df.cve_id.str.extract(r"CVE-(\d{4})")[0].astype(int).value_counts().sort_index()
    pub = "، ".join(f"{fa(y)}: {fa(n)}" for y, n in yc.items()); idy = "، ".join(f"{fa(y)}: {fa(n)}" for y, n in idc.items())
    odd = df[df.cve_id.str.extract(r"CVE-(\d{4})")[0].astype(int) != df.published_year]
    para(f"توزیع بر اساس سال انتشار در NVD (published): {pub} (جمع = {fa(int(yc.sum()))}).", bullet=True)
    if len(odd):
        para(f"توزیع بر اساس سال شناسه CVE: {idy}. تفاوت فقط ناشی از {', '.join(odd.cve_id)} است که شناسه‌اش مربوط به ۲۰۲۵ ولی تاریخ انتشار آن {odd.published.iloc[0]} است؛ این رکورد فقط یک‌بار و در سال انتشار شمرده شده است.", bullet=True)
    para(f"رکوردهای دارای CWE معتبر: n_valid = {fa(NV)} ({fa(f'{NV / N * 100:.2f}')}٪ از N) در {fa(st['ndist'])} کد CWE متمایز؛ رکوردهای NVD-CWE-noinfo: {fa(NN)} ({fa(f'{NN / N * 100:.2f}')}٪)؛ سلول خالی: ۰.", bullet=True)

    h1(f"۲. گام اول: کاتالوگ {fa(len(cat))} محصول و ابزار")
    table(["شناسه محصول", "نام و سازنده", "گروه", "مخزن / مرجع", "وضعیت و تعداد CVE"],
          [[r.product_id, f"{r.display_name} ({r.vendor_org})", r.group, r.official_repo, r.inclusion_status] for r in cat.itertuples()],
          [2.6, 3.8, 3.8, 4.2, 3.4], 7.5)

    h1("۳. پاسخ به RQ2: پرتکرارترین CWEها")
    para(f"روش شمارش: برای هر CVE یکتا یک کد ضعف اصلی ثبت شده است (نگاشت ۱ به ۱ پس از ادغام رکورد تکراری). {fa(NN)} رکورد NVD-CWE-noinfo از مخرج CWEهای معتبر جدا شده و سهم هر CWE بر مبنای n_valid = {fa(NV)} (و در ستونی جدا بر مبنای N = {fa(N)}) محاسبه شده است. درصد تجمعی از شمارش‌های خام محاسبه و فقط در نمایش نهایی با دو رقم اعشار گرد شده است.")
    para(f"قاعده رتبه‌بندی و شکستن تساوی: ابتدا تعداد CVE (نزولی)، سپس میانگین CVSS (نزولی) و در نهایت شناسه CWE (صعودی). از آنجا که {fa(st['tie_count'])} CWE هر کدام دقیقاً ۲ رکورد دارند، مرز «۱۰ مورد برتر» دلخواه می‌شد؛ بنابراین نتایج در دو سطح گزارش می‌شوند:")
    para(f"{fa(5)} ضعف برتر ({', '.join(top5.cwe)}): مجموعاً {fa(st['top5_n'])} رکورد از {fa(NV)} رکورد معتبر ({fa(f'{st['top5_pct']:.2f}')}٪).", bullet=True)
    para(f"همه {fa(nrec)} ضعف تکرارشونده (دست‌کم ۲ رکورد): مجموعاً {fa(st['rec_n'])} رکورد ({fa(f'{st['rec_pct']:.2f}')}٪ از رکوردهای معتبر).", bullet=True)
    table(["رتبه", "CWE", "عنوان ضعف", "تعداد", f"سهم از معتبر (n={fa(NV)})", f"سهم از کل (N={fa(N)})", "تجمعی معتبر", "میانگین CVSS"],
          [[r.rank, r.cwe, f"{r.cwe_name_fa}\n({r.cwe_name_en})", r.count, f"{r.share_of_valid_cwe_pct:.2f}%", f"{r.share_of_total_cve_pct:.2f}%",
            f"{r.cumulative_valid_pct:.2f}%", f"{r.mean_cvss:.2f}"] for r in g.itertuples()] +
          [["-", "جمع", f"مجموع رکوردهای دارای CWE معتبر ({st['ndist']} کد)", NV, "100.00%", f"{NV / N * 100:.2f}%", "100.00%", f"{st['mean_valid']:.2f}"]],
          [1.3, 1.8, 6.6, 1.3, 2.1, 2.0, 1.9, 1.7], 8)

    h1(f"۴. رکوردهای بدون CWE مشخص (NVD-CWE-noinfo, n = {fa(NN)})")
    nl = noinfo_t[noinfo_t["product"].str.startswith("langchain")]; wp = noinfo_t[noinfo_t["product"] == "wpbot"]
    para(f"طبق راهنمای RQ2 این رکوردها با CWEهای معتبر ترکیب نشده‌اند. از {fa(N)} رکورد، {fa(NN)} رکورد ({fa(f'{NN / N * 100:.2f}')}٪) این وضعیت را دارند: {fa(len(nl))} رکورد LangChain/LangChain Experimental (میانگین CVSS {fa(f'{nl.cvss_score.mean():.2f}')}) و {fa(len(wp))} رکورد افزونه WPBot (میانگین {fa(f'{wp.cvss_score.mean():.2f}')}). میانگین هر {fa(NN)} رکورد برابر {fa(f'{st['mean_noinfo']:.2f}')} است.")
    table(["شناسه CVE", "محصول", "تاریخ انتشار", "وضعیت CWE در NVD", "امتیاز CVSS", "شدت"],
          [[r.cve_id, r.product, r.published, r.cwe, f"{r.cvss_score:.1f}", r.severity] for r in noinfo_t.itertuples()], [3.2, 3.6, 2.8, 3.6, 2.2, 2.4], 9)

    h1(f"۵. نمودارهای تحلیلی (N = {fa(N)})")
    caps = [f"نمودار ۱ (تم تیره): ۵ ضعف برتر و همه {fa(nrec)} ضعف تکرارشونده در کنار پوشش CWE و رکوردهای بدون CWE مشخص",
            "نمودار ۲ (تم روشن، مناسب چاپ): همان توزیع با رنگ‌بندی مناسب گزارش چاپی",
            "نمودار ۳: ضعف‌های تکرارشونده به تفکیک محصول، میانگین شدت CVSS و جداسازی رکوردهای بدون CWE مشخص",
            "نمودار ۴: روند زمانی رکوردها و ۵ ضعف غالب در چهار سال ۲۰۲۳ تا ۲۰۲۶"]
    for cap, f_ in zip(caps, chart_files):
        para(cap, True, 10, space_after=3).paragraph_format.keep_with_next = True
        p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER; p.add_run().add_picture(f_, width=Cm(17))

    h1("۶. یافته‌های کلیدی و محدودیت‌ها")
    t4 = ", ".join(f"{r.cwe} ({r.count})" for r in g.head(4).itertuples())
    para(f"سلطه چند ضعف اصلی: چهار ضعف نخست ({t4}) به‌تنهایی {fa(st['top4_n'])} رکورد از {fa(NV)} رکورد معتبر ({fa(f'{st['top4_pct']:.2f}')}٪) را می‌سازند و پنج ضعف نخست {fa(f'{st['top5_pct']:.2f}')}٪ را.", bullet=True)
    para(f"خانواده تزریق و اجرای کد: مجموع {', '.join(INJECTION_FAMILY)} برابر {fa(st['inj_n'])} رکورد از {fa(NV)} ({fa(f'{st['inj_pct']:.2f}')}٪) است؛ این ضعف‌ها ناشی از عدم خنثی‌سازی ورودی/خروجی پیش از رسیدن به مفسر یا مؤلفه پایین‌دستی‌اند.", bullet=True)
    y23 = df[(df.published_year == 2023) & df.is_valid_cwe].cwe.value_counts()
    para(f"تحول زمانی: در ۲۰۲۳ بیشترین سهم با {y23.index[0]} و {y23.index[1]} (هر کدام {fa(int(y23.iloc[0]))} رکورد، عمدتاً LangChain) بود؛ در ۲۰۲۴ تا ۲۰۲۶ طیف ضعف‌ها با ورود ابزارهای ساخت ایجنت (Langflow، Flowise) و موتورهای استنتاج (Ollama، vLLM، llama.cpp) گسترده‌تر شد و ضعف‌هایی مانند CWE-22، CWE-918، CWE-78، CWE-787 و CWE-125 دیده می‌شوند. به‌دلیل تعداد کم رکورد در هر سال این روند توصیفی است و نه استنباط آماری.", bullet=True)
    para(f"محدودیت‌ها: (الف) {fa(NN)} رکورد NVD-CWE-noinfo ({fa(f'{NN / N * 100:.2f}')}٪) نشان‌دهنده نبود CWE دقیق در زمان برداشت است. (ب) فراوانی بیشتر LangChain و Langflow لزوماً به معنای ناامن‌تر بودن آن‌ها نیست و می‌تواند ناشی از جامعه کاربری بزرگ‌تر و ممیزی بیشتر باشد. (ج) در تحلیل اصلی برای هر CVE فقط CWE اصلی شمرده شده است. (د) {fa(st['n_v4'])} رکورد (CVE-2025-34291) فقط با امتیاز CVSS v4.0 ثبت شده و برای حفظ رکورد در همین دیتاست نگه داشته شد؛ میانگین‌ها با و بدون آن گزارش می‌شوند (کل: {fa(f'{st['mean_all']:.2f}')} در برابر {fa(f'{st['mean_all_v31']:.2f}')}؛ معتبر: {fa(f'{st['mean_valid']:.2f}')} در برابر {fa(f'{st['mean_valid_v31']:.2f}')}). (هـ) CVE-2026-108836 هنگام برداشت فقط امتیاز CNA (VulDB، CVSS v3.1) داشت و تحلیل NVD هنوز ثبت نشده بود. (و) اعداد ستون‌های امتیاز و C/I/A باید پیش از تحویل نهایی با صفحه NVD هر CVE مقایسه شوند.", bullet=True)
    z = doc.settings.element.find(qn("w:zoom"))
    if z is not None: z.set(qn("w:percent"), "100")
    doc.save(path)


# ----------------------------------------------------------------------------
# Notebook (written as plain JSON; outputs are produced by executing the cells)
# ----------------------------------------------------------------------------
NOTEBOOK_CELLS = [
 ("markdown", "# Generative AI / AI Agent vulnerability analysis - RQ2 (most frequent CWE types), 2023-2026\n"
  "Everything below is recomputed from `2_GenAI_Vulnerabilities_Clean_Dataset_Step4.csv`; no number is typed by hand.\n\n"
  "**Method notes**\n1. Denominators: N = all unique CVEs in the clean sample; n_valid = CVEs with a specific CWE; NVD-CWE-noinfo records are reported separately.\n"
  "2. One primary CWE per CVE (1-to-1 mapping). Cumulative share = cumulative raw count / n_valid, rounded once at the end.\n"
  "3. Ranking: count desc, mean CVSS desc, CWE id asc. Ties at 2 CVEs are all listed (Top 5 + all recurring CWEs)."),
 ("code", "import pandas as pd\n"
  "df = pd.read_csv('2_GenAI_Vulnerabilities_Clean_Dataset_Step4.csv', keep_default_na=False)\n"
  "valid = df[df['is_valid_cwe'].astype(str) == 'True']; noinfo = df[df['is_valid_cwe'].astype(str) == 'False']\n"
  "print(f'Total verified records (N): {len(df)}')\n"
  "print(f'Valid specific CWE (n_valid): {len(valid)} ({len(valid)/len(df)*100:.2f}%)')\n"
  "print(f'NVD-CWE-noinfo (n_noinfo): {len(noinfo)} ({len(noinfo)/len(df)*100:.2f}%)')\n"
  "print('Records by year:', df.groupby('published_year').size().to_dict())"),
 ("code", "dist = pd.read_csv('3_RQ2_Top_Valid_CWE_Distribution.csv', keep_default_na=False)\n"
  "print(dist[['rank','cwe','cwe_category','count','share_of_valid_cwe_pct','share_of_total_cve_pct','cumulative_valid_pct','mean_cvss']].to_string(index=False))\n"
  "recurring = dist[dist['count'] >= 2]\n"
  "print(f\"\\nTop 5 cover {dist['count'].head(5).sum()} of {len(valid)} valid CVEs ({dist['count'].head(5).sum()/len(valid)*100:.2f}%)\")\n"
  "print(f\"All {len(recurring)} recurring CWEs (>=2 CVEs) cover {recurring['count'].sum()} of {len(valid)} ({recurring['count'].sum()/len(valid)*100:.2f}%)\")"),
 ("code", "print(pd.read_csv('4_RQ2_Records_Without_CWE_NoInfo_List.csv', keep_default_na=False).to_string(index=False))\n"
  "print()\nprint(pd.read_csv('6_Step5_Data_Audit_and_Denominators.csv', keep_default_na=False)[['Audit_Metric','Count','Denominator','Percentage']].to_string(index=False))"),
 ("markdown", "Charts: `RQ2_Top_CWE_BarChart_Dark.png`, `RQ2_Top_CWE_BarChart_Academic_Light.png`, `RQ2_CWE_Product_Severity_Breakdown.png`, `RQ2_Yearly_Evolution_2023_2026.png`."),
]


def build_notebook(path, outdir):
    import contextlib, io, os
    cwd = os.getcwd(); os.chdir(outdir)
    cells, ns, ec = [], {}, 0
    try:
        for kind, src in NOTEBOOK_CELLS:
            lines = src.split("\n"); source = [l + "\n" for l in lines[:-1]] + [lines[-1]]
            if kind == "markdown":
                cells.append({"cell_type": "markdown", "metadata": {}, "source": source}); continue
            buf = io.StringIO()
            with contextlib.redirect_stdout(buf): exec(src, ns)
            ec += 1
            cells.append({"cell_type": "code", "execution_count": ec, "metadata": {}, "source": source,
                          "outputs": [{"output_type": "stream", "name": "stdout", "text": buf.getvalue().splitlines(True)}]})
    finally:
        os.chdir(cwd)
    nb = {"cells": cells, "metadata": {"kernelspec": {"display_name": "Python 3", "language": "python", "name": "python3"},
                                       "language_info": {"name": "python"}}, "nbformat": 4, "nbformat_minor": 5}
    json.dump(nb, open(path, "w", encoding="utf-8"), ensure_ascii=False, indent=1)


README = """# RQ2 - Generative AI vulnerability analysis (reproducible build)

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
"""


def build_all(outdir="."):
    os.makedirs(outdir, exist_ok=True)
    P = lambda f: os.path.join(outdir, f)
    cat, df = build_frames()
    g, noinfo_t, summary, audit = build_tables(df)
    yearly = yearly_table(df); st = derived_stats(df, g)
    enc = dict(index=False, encoding="utf-8-sig")
    cat.to_csv(P("1_Product_Metadata_Catalog_Step1.csv"), **enc)
    df.to_csv(P("2_GenAI_Vulnerabilities_Clean_Dataset_Step4.csv"), **enc)
    g.to_csv(P("3_RQ2_Top_Valid_CWE_Distribution.csv"), **enc)
    noinfo_t.to_csv(P("4_RQ2_Records_Without_CWE_NoInfo_List.csv"), **enc)
    summary.to_csv(P("5_RQ2_Missing_and_NoInfo_Summary_Table.csv"), **enc)
    audit.to_csv(P("6_Step5_Data_Audit_and_Denominators.csv"), **enc)
    charts = [P("RQ2_Top_CWE_BarChart_Dark.png"), P("RQ2_Top_CWE_BarChart_Academic_Light.png"),
              P("RQ2_CWE_Product_Severity_Breakdown.png"), P("RQ2_Yearly_Evolution_2023_2026.png")]
    chart_bars(g, st, df, charts[0], charts[1]); chart_breakdown(g, st, df, charts[2]); chart_yearly(g, st, df, charts[3])
    build_xlsx(P("RQ2_Complete_Analysis_Workbook.xlsx"), cat, df, g, noinfo_t, summary, audit, yearly, st)
    build_notebook(P("RQ2_Complete_Analysis_Notebook.ipynb"), outdir)
    build_docx(P("RQ2_Full_Methodology_and_Analysis_Report.docx"), cat, df, g, noinfo_t, audit, yearly, st, charts)
    open(P("README.md"), "w", encoding="utf-8").write(README)
    print(f"Build OK: N={st['N']}, n_valid={st['NV']}, n_noinfo={st['NN']}, distinct CWEs={st['ndist']}")


if __name__ == "__main__":
    build_all(os.path.dirname(os.path.abspath(__file__)))
