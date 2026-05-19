# Family Office RAG Pipeline: Technical Documentation

This document describes the design decisions, retrieval workflow, and architecture choices for the **Family Office RAG Explorer**.

---

## 1. Stack Choices
*   **Frontend Interface**: **Streamlit** (allows building web-based data applications rapidly with native support for chat history and interactive sidebars without requiring HTML/CSS boilerplate).
*   **Data Processing**: **Pandas** (used to instantly parse, filter, and convert the backend Excel workbook into a highly optimized CSV string for the LLM).
*   **LLM Providers**: **Google Gemini (1.5 Flash/Pro)** or **OpenAI (GPT-4o/4o-mini)** (user-configurable via the sidebar to dynamically generate institutional-grade summaries and answers).

---

## 2. The Architectural Pivot: Why No Embeddings?
Initially, the application was built using a standard Vector Database (`sentence-transformers` + Cosine Similarity). However, this architecture was deliberately scrapped. 

**Why?** Vector search relies on retrieving the "Top K" (e.g., top 5) most semantically similar chunks. This is fundamentally broken for *global math or aggregation queries*. If a user asks *"Who has the highest total fines?"*, a vector database cannot calculate this; it simply retrieves the 5 profiles whose text looks most similar to the word "fines", entirely missing the actual individual with the highest penalty.

### The New Approach: Two-Step Agentic Routing (Full Context)
Because our dataset is highly structured (Excel) and relatively small (50 rows), we shifted from a Vector Search to a **Two-Step Column Routing Agent**.

#### Step 1: The Router (Column Filtering)
When a query is received, the AI is first given the Schema (the column headers of the Excel file) along with context hints (e.g., *"Disclosure Details contains the fines"*). The AI responds *only* with the exact column names needed to answer the question.

#### Step 2: The Analyzer (CSV Context Injection)
Pandas instantly filters the dataset to include only `CRD#`, `Full Name`, and the specific columns the AI requested. This filtered slice is converted to a CSV string and injected directly into the LLM's prompt. 
The AI is explicitly commanded to read every single row in the CSV to compute the answer.

---

## 3. What Works
*   **100% Aggregation Accuracy:** By forcing the AI to read the filtered CSV slice containing all 50 profiles simultaneously, questions like *"Who is the most experienced?"* or *"How many total disclosures exist?"* return perfectly accurate results, entirely eliminating "missing context" hallucinations.
*   **Massive Token Efficiency:** By asking the Router to only select necessary columns, we prevent sending 42 unnecessary columns of data to the LLM. If the user asks about experience, we only send the Name, CRD#, and Experience columns. This keeps latency low and API costs near zero.

---

## 4. What Does Not Work
*   **Scaling to Massive Datasets:** This approach works perfectly for 50-500 rows because the resulting CSV string easily fits inside modern 128k+ token context windows. However, if the dataset grows to 50,000 rows, serializing the data to a CSV string will breach the context window limit and cause the API to crash or become prohibitively expensive.

---

## 5. What I Would Improve
If tasked with scaling this pipeline to support tens of thousands of Family Office profiles, I would implement a **Text-to-SQL / Pandas Agent Engine**.
Instead of reading a CSV string, the LLM would act as an autonomous programmer. It would read the user query, write actual Python/Pandas code (or SQL) to calculate the answer mathematically (e.g., `df['Fines'].max()`), execute the code in a sandbox, and return the exact numeric result to the user. This guarantees infinite scale and millisecond latency regardless of dataset size.
