# AIRA — Final Assignment #1 Submission

## Authoritative source files
1. Minor_Courses_for_BTech_Students(3).xlsx
2. Semester_Spread_Structures_Sept_2026(3).xlsx

These are the actual academic sources used by the project.

## Final architecture
Excel -> loader -> full-field normalization -> question router/entity resolution -> structured tables -> record-level chunks -> SentenceTransformer embeddings -> FAISS -> Top-40 retrieval -> hybrid reranking -> grounded evidence -> deterministic verification -> structured prompt -> Gemini explanation -> verification gate -> final answer.

## Deployment
Upload `app.py`, `requirements.txt`, `data/` and the two Excel files to the Streamlit/GitHub repository. Add:
GEMINI_API_KEY = "..."
to Streamlit Secrets. Never hard-code the key into the repository.

## Assignment alignment
Phase 1: RAG + grounding + functional prototype
Phase 2: synthetic students + edge cases + follow-up/conflicts
Phase 3: four prompt layers
Phase 4: 30-case evaluation framework
Phase 5: Basic LLM -> Prompting -> RAG -> RAG + Structured Data
Phase 6: Streamlit prototype + evidence/uncertainty
Phase 7: 10-page report + PPT + architecture/methodology/results/limitations

## Important
Do not claim 100% accuracy. Report measured results only after executing the benchmark. Do not claim PDF/DOCX/website ingestion because the actual supplied academic sources are Excel workbooks.

## Final accuracy/coverage design update
- Students can enter **course names or course codes**. Course names are a first-class input.
- If a course name maps to multiple distinct course codes, the advisor asks a follow-up instead of guessing.
- If the same course code appears across academic years with harmless wording/punctuation differences, it is treated as the same course.
- Direct factual questions are answered from deterministic records extracted from the two supplied Excel workbooks; the LLM is a grounded language layer and is never the source of truth.
- Broad questions (minor lists, semester lists, structure/category lists, credit requirements) are answered from the corresponding workbook tables or ask for the missing structure/minor/year when necessary.
- Missing, ambiguous, or conflicting information produces an explicit follow-up/insufficient/conflict response rather than an invented answer.
- The app exposes source/evidence records for verification.

## Comprehensive question coverage
The final version routes questions across course records, minor records, semester-spread records, structure/category records, credit requirements and synthetic-student eligibility. Course names are accepted as the normal input; course codes are optional. The system can retrieve individual facts, combined facts, lists, filters and counts supported by the two workbooks. If a requested scope is ambiguous (for example, multiple structures or duplicate course titles), it asks for the missing discriminator rather than guessing.

## Evidence boundary
The two supplied Excel workbooks are the authoritative academic sources. Synthetic student profiles are test data only. The LLM cannot create or override academic facts that are not supported by retrieved/verified source evidence.
