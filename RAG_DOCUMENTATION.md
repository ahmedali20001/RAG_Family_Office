# Family Office RAG Pipeline: Technical Documentation

This document describes the design decisions, retrieval workflow, and architecture choices for the **Family Office RAG Explorer**.

---

## 1. Stack Choices
*   **Frontend Interface**: **Streamlit** (allows building web-based data applications rapidly with native support for chat layouts, forms, and custom CSS without requiring HTML/CSS boilerplate).
*   **Embedding Model**: **`sentence-transformers/all-MiniLM-L6-v2`** (a lightweight 384-dimensional sentence transformer model running locally. It provides fast embedding generations with zero API costs, making the app self-sufficient).
*   **Vector Database**: **NumPy Array with Cosine Similarity** (since our dataset consists of 50 profiles (~50 documents), an in-memory NumPy matrix is incredibly fast (sub-millisecond lookups), zero-configuration, and has no infrastructure overhead or dependencies like Docker or SQLite).
*   **LLM Providers**: **Google Gemini (1.5 Flash/Pro)** or **OpenAI (GPT-4o/4o-mini)** (user-configurable via the sidebar to dynamically generate institutional-grade summaries and answers based on retrieved context).

---

## 2. Chunking Strategy
Instead of splitting text files arbitrarily by character length (which ruins biographical data by slicing mid-sentence or separation of disclosure metrics), we use **Entity-Level Semantic Chunking**:
*   Each adviser is processed as a **single discrete document** (1 adviser profile = 1 chunk).
*   A textual document profile is constructed by combining:
    - Demographic Details (Name, CRD#)
    - Licensing scopes (IA, BD)
    - Geographic coverage (Registrations and Branch Locations)
    - Full Employment History (dates and firms)
    - Regulatory Disclosures (each parsed disclosure contains initiating authority, allegations, fines, sanctions, and comments).
*   This ensures that the retriever always brings the **complete biography and disclosure history** of an individual, rather than partial pieces.

---

## 3. Embedding & Retrieval Approach
1.  **Index Time**: The script walks the `raw_json/` directory, compiles the synthesized profile text block for each of the 50 advisers, generates a 384-dimensional vector representation for each text block, and holds them in a NumPy matrix.
2.  **Query Time**: 
    - The user enters a natural language query.
    - The query is embedded using the same `all-MiniLM-L6-v2` model.
    - We calculate the cosine similarity between the query vector and all document vectors.
    - The top $K$ profiles with the highest similarity scores are retrieved.
3.  **Refinement / Generation**: The $K$ retrieved profile texts are injected into an LLM prompt as context, instructing the model to synthesize a response utilizing only the retrieved facts and cite the source name and CRD number.

---

## 4. Evaluation: What Works & What Does Not

### What Works:
*   **Entity Integrity**: Because chunking is kept at the adviser level, the LLM never hallucinates associations between Adviser A's disclosures and Adviser B.
*   **Disclosure Deep-Dive**: The RAG pipeline can semantic-search deep text within broker comments and allegations (e.g. searching "Bolton Securities reporting error" or "Florida lawful registration").
*   **Offline Vector Engine**: Zero dependencies on hosted vector databases.
*   **Dual LLM Config**: Supports both Google Gemini and OpenAI.

### What Does Not:
*   **Complex Cross-Entity Math**: LLMs are notoriously poor at aggregation. If asked, *"Who has the highest number of exams?"*, the semantic search will retrieve the most matching profiles, but the LLM might struggle to evaluate the exact counts across all 50 profiles since it only sees the top $K$ profiles rather than the full table.
*   **Scale Limitation**: Since the NumPy index is kept in memory, it works exceptionally well for hundreds of files, but would need to transition to a dedicated vector database (like Chroma, Pinecone, or pgvector) once scaling to 10,000+ advisers.

---

## 5. Next Steps for Improvement
1.  **Hybrid Retrieval (BM25 + Semantic Search)**: Combine keyword search (essential for exact CRD/Name matching) with semantic search (for conceptual queries like "disciplinary actions regarding registration errors").
2.  **Parent-Child Chunking**: Break disclosures down into child chunks linked to the parent adviser profile, allowing fine-grained disclosure retrieval for advisers with massive regulatory histories.
3.  **Agentic Tool Use**: Enable the model to run Python scripts or SQL queries over the Excel table to answer math/statistical queries (e.g., "average years of experience").
