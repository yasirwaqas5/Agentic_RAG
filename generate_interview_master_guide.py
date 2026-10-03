from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH


def add_bullets(doc, items):
    for item in items:
        doc.add_paragraph(item, style='List Bullet')


def add_table(doc, rows, bold_header=True):
    table = doc.add_table(rows=len(rows), cols=len(rows[0]))
    for r_idx, row in enumerate(rows):
        for c_idx, value in enumerate(row):
            cell = table.cell(r_idx, c_idx)
            cell.text = str(value)
            if r_idx == 0 and bold_header:
                for p in cell.paragraphs:
                    for run in p.runs:
                        run.bold = True
    return table


def add_code_block(doc, text):
    p = doc.add_paragraph()
    run = p.add_run(text)
    run.font.name = 'Consolas'
    run.font.size = Pt(9)


def add_section(doc, title):
    doc.add_heading(title, level=1)


def build_document():
    doc = Document()
    doc.styles['Normal'].font.name = 'Calibri'
    doc.styles['Normal'].font.size = Pt(11)

    # Title page
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run('INTERVIEW MASTER GUIDE')
    r.bold = True
    r.font.size = Pt(26)

    doc.add_paragraph('')
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run('production-agentic-rag')
    r.bold = True
    r.font.size = Pt(18)

    doc.add_paragraph('')
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.add_run('Deep technical interview preparation grounded in this repository\'s actual implementation')
    doc.add_page_break()

    # TOC
    add_section(doc, 'TABLE OF CONTENTS')
    items = [
        '1. Project Overview',
        '2. Architecture',
        '3. Complete Request Lifecycle',
        '4. File-by-File Codebase Explanation',
        '5. Agentic RAG vs Basic RAG',
        '6. Planner',
        '7. Retrieval',
        '8. Guardrails',
        '9. Tools',
        '10. Graph / Orchestration',
        '11. LLM + Embedding Providers',
        '12. Ingestion',
        '13. Evaluation',
        '14. Observability',
        '15. Testing',
        '16. API',
        '17. Design Decisions',
        '18. Trade-Offs',
        '19. Failure Modes',
        '20. Security',
        '21. Productionization',
        '22. Interview Questions',
        '23. Interviewer Attack Mode',
        '24. Trick Questions',
        '25. Key Concepts I Must Know',
        '26. What I Should Say About the Project',
        '27. Project Walkthrough Script',
        '28. 1-Day Interview Preparation',
        '29. Final Cheat Sheet',
    ]
    for i in items:
        doc.add_paragraph(i)
    doc.add_page_break()

    # Section 1
    add_section(doc, '1. PROJECT OVERVIEW')
    doc.add_paragraph('Project name: production-agentic-rag')
    doc.add_paragraph('Problem being solved: the repo closes the gap between toy RAG demos and production-grade knowledge systems. Most naive implementations are “embed → top-k → prompt,” while this project adds planning, multi-hop retrieval, grounding validation, guardrails, eval, and observability.')
    doc.add_paragraph('Why the project exists: the README explicitly states that production systems need planning, multi-hop retrieval, grounding checks, evaluation, and observability. This repo turns those concerns into explicit modules with offline defaults so it can be tested without API keys.')
    doc.add_paragraph('Target users / use case: the project is aimed at a domain-specific QA assistant over corpora like product documentation, internal handbooks, and support knowledge. The bundled corpus and eval dataset focus on a support/operations handbook use case.')
    doc.add_paragraph('Core capabilities: hybrid retrieval, multi-hop agent loop, citation tracking, provider abstraction, guardrails, API access, evaluation harness, observability tracing, and optional LangGraph orchestration.')
    doc.add_paragraph('What makes it different from ordinary RAG: it is not just a single retrieval pass and generation step. The code explicitly decomposes complex questions, executes multiple retrieval hops, deduplicates evidence, evaluates grounding, and performs a retry loop when faithfulness is weak.')
    doc.add_paragraph('Technology stack: Python 3.10+, pure-Python core, optional OpenAI/Azure OpenAI, FastAPI, LangGraph, pytest, Docker, GitHub Actions. The core runs without runtime dependencies because MockLLM and HashingEmbedder are used by default.')
    doc.add_paragraph('Architecture style: modular orchestration pipeline. RAGPipeline composes providers, ingestion, retrieval, agent logic, and evaluation; the LangGraph engine is an alternative execution route.')
    doc.add_paragraph('Major components: providers, ingestion, retrieval, agents, evaluation, guardrails, observability, API, graph, and tests.')
    doc.add_paragraph('30-second interview pitch: “This project is an agentic RAG system that reads a document corpus, plans sub-questions, retrieves evidence using hybrid search, checks grounding, and synthesizes a cited answer. It is designed to be more production-aware than a toy demo because it includes guardrails, evaluation, observability, provider abstraction, and an offline mode for CI.”')
    doc.add_paragraph('2-minute interview explanation: “The project is built around RAGPipeline. It loads docs, chunks them, indexes them, and uses hybrid retrieval combining BM25 and vector search via reciprocal rank fusion. A planner decomposes the question into sub-questions, and AgenticRAG executes per-hop retrieval, deduplicates citations, calls the LLM, and measures faithfulness. If the answer is not grounded enough, it reformulates the query and retries. The system also checks for prompt injection, redacts PII, and emits trace summaries with tokens and cost accounting. The design is provider-agnostic and can run with a mock LLM or real OpenAI/Azure models.”')
    doc.add_paragraph('5–10 minute technical explanation: “At the orchestration layer, RAGPipeline sets up a HybridRetriever, builds an LLM provider, wraps a RetrievalTool, and chooses either a built-in AgenticRAG or an optional LangGraphRAG. On query execution, check_input validates the request. The planner splits the question into sub-questions using a heuristic regex-based strategy. For each sub-question, the tool retrieves results, deduplicates by chunk ID, then builds a context list. The LLM answers using only that context and grounding_score compares answer tokens against context tokens. If the score falls below the faithfulness threshold, the system reformulates the question with appended context and retries. That is why this is not a vanilla RAG demo: it has an explicit reasoning loop, evidence selection, and measurements for grounding quality.”')

    # Section 2
    add_section(doc, '2. ARCHITECTURE')
    doc.add_paragraph('High-level system architecture:')
    add_code_block(doc, (
        'User\n'
        '  |\n'
        '  v\n'
        'API / CLI / direct call\n'
        '  |\n'
        '  v\n'
        'RAGPipeline\n'
        '  |\n'
        '  +--> check_input -> InputBlocked guardrail\n'
        '  |\n'
        '  +--> AgenticRAG / LangGraphRAG\n'
        '         |\n'
        '         +--> plan()\n'
        '         +--> RetrievalTool -> HybridRetriever\n'
        '         |      +--> BM25\n'
        '         |      +--> VectorStore\n'
        '         |      +--> RRF fusion + rerank\n'
        '         +--> LLMProvider.complete()\n'
        '         +--> grounding_score\n'
        '         +--> reformulate retry\n'
        '  |\n'
        '  v\n'
        'answer + citations + trace\n'
    ))
    doc.add_paragraph('What each component does: ingestors convert raw documents into chunks, retrieval finds relevant evidence, the planner decomposes work, the agent loops over evidence and synthesis, guardrails protect inputs and outputs, and observability records spans and costs. The system keeps these responsibilities explicit and testable.')
    doc.add_paragraph('Complete end-to-end request flow: query -> check_input -> Tracer -> engine.run -> plan -> retrieval -> LLM synthesis -> grounding_score -> optional reformulate -> final QueryResult with answer, citations, faithfulness, iterations, trace.')
    doc.add_paragraph('Ingestion pipeline: load_directory -> Document -> chunking -> index -> retrieval. Retrieval pipeline: BM25 search + vector search -> reciprocal rank fusion -> rerank -> top-k. Planner -> agent -> tool flow: plan_question -> retrieval tool -> dedup -> context build -> LLM answer. Guardrail flow: input checks before query, output PII redaction and grounding check after generation. Evaluation architecture: run_eval over labeled cases -> metrics. Failure/recovery flow: fail-closed input block, faithfulness retry, engine fallback if LangGraph missing.')

    # Section 3
    add_section(doc, '3. COMPLETE REQUEST LIFECYCLE')
    doc.add_paragraph('Example: a user asks, “What is the return window and how long do refunds take?”')
    steps = [
        'User asks a question.',
        'RAGPipeline.query(question) creates a Tracer and invokes check_input(question).',
        'If the prompt is blocked, QueryResult is returned with blocked=True and the request stops.',
        'If allowed, AgenticRAG.run(question, tracer) starts.',
        'plan(query) decomposes the question into sub-questions using regex heuristics.',
        'For each sub-question, RetrievalTool calls HybridRetriever.retrieve().',
        'HybridRetriever runs BM25 and VectorStore search, fuses rankings with reciprocal rank fusion, then reranks candidates.',
        'The agent deduplicates retrieved chunks and builds a context list from chunk.text values.',
        'The LLM is invoked with a prompt that contains only the retrieved context and the original question.',
        'grounding_score(answer, contexts) measures token overlap between the answer and context.',
        'If the score is below the faithfulness threshold, the agent reformulates the query with context expansion and retries up to max_iterations.',
        'The best-scoring iteration is returned, along with citations and trace summary.',
        'Final answer is sanitized for PII and passed back as QueryResult.'
    ]
    add_bullets(doc, steps)
    doc.add_paragraph('Files/functions involved: src/production_agentic_rag/pipeline.py -> RAGPipeline.query; src/production_agentic_rag/agents/rag_agent.py -> AgenticRAG.run; src/production_agentic_rag/agents/planner.py -> plan; src/production_agentic_rag/retrieval/hybrid.py -> HybridRetriever.retrieve; src/production_agentic_rag/guardrails.py -> grounding_score, redact_pii; src/production_agentic_rag/observability.py -> Tracer.summary.')

    # Section 4
    add_section(doc, '4. FILE-BY-FILE CODEBASE EXPLANATION')
    rows = [
        ['File', 'Purpose / role'],
        ['api/main.py', 'FastAPI app with /health and /ask; lazily creates a shared RAGPipeline and loads the corpus.'],
        ['src/production_agentic_rag/config.py', 'Settings dataclass loads environment variables like provider names, top_k, rerank, chunk sizes, engine, and faithfulness threshold.'],
        ['src/production_agentic_rag/pipeline.py', 'Top-level pipeline orchestrator: wires retriever, LLM, tool, agent, and optional LangGraph engine, applies guardrails, returns QueryResult.'],
        ['src/production_agentic_rag/graph.py', 'Optional LangGraph StateGraph with plan -> retrieve -> generate -> reformulate edges and conditional routing.'],
        ['src/production_agentic_rag/guardrails.py', 'Input injection detection, length cap, PII redaction, grounding_score.'],
        ['src/production_agentic_rag/observability.py', 'Tracer spans, token/cost accounting, and summary.'],
        ['src/production_agentic_rag/agents/planner.py', 'Heuristic planner that decomposes the question into sub-questions.'],
        ['src/production_agentic_rag/agents/rag_agent.py', 'AgenticRAG loop with plan, retrieve, synthesize, and self-correction.'],
        ['src/production_agentic_rag/agents/tools.py', 'RetrievalTool exposes the retriever as a callable tool.'],
        ['src/production_agentic_rag/providers/llm.py', 'MockLLM and OpenAILLM behind the LLMProvider protocol.'],
        ['src/production_agentic_rag/providers/embeddings.py', 'HashingEmbedder and OpenAIEmbedder behind the Embedder protocol.'],
        ['src/production_agentic_rag/retrieval/bm25.py', 'BM25 lexical search with stemming and IDF.'],
        ['src/production_agentic_rag/retrieval/vectorstore.py', 'Dense vector store with cosine similarity and cache.'],
        ['src/production_agentic_rag/retrieval/hybrid.py', 'HybridRetriever uses RRF + rerank.'],
        ['src/production_agentic_rag/ingestion/loaders.py', 'Loads files/directories as documents.'],
        ['src/production_agentic_rag/ingestion/chunking.py', 'Chunk dataclass and fixed-size/markdown-section chunkers.'],
        ['src/production_agentic_rag/text.py', 'Tokenizer, stemmer, stopword filtering, and cosine helper.'],
        ['src/production_agentic_rag/evaluation/metrics.py', 'Retrieval metrics, faithfulness, and answer correctness.'],
        ['src/production_agentic_rag/evaluation/harness.py', 'Dataset loader and end-to-end eval runner.'],
        ['src/production_agentic_rag/integrations.py', 'Optional secure gateway and observability export adapters.'],
        ['data/eval/qa.jsonl', 'Evaluation data with questions, expected answers, relevant_ids.'],
    ]
    add_table(doc, rows)
    doc.add_paragraph('Key point: the repo is intentionally modular. Interfaces are simple and clear, which makes testing and substitution straightforward.')

    # Section 5
    add_section(doc, '5. AGENTIC RAG VS BASIC RAG')
    rows = [
        ['Dimension', 'Basic RAG', 'This project'],
        ['Query handling', 'Single retrieval pass', 'Decomposes query into sub-questions; may retry with reformulation'],
        ['Decision making', 'Usually none', 'Explicit plan() step and iterative faithfulness gate'],
        ['Tools', 'Typically none', 'Retrieval exposed as a tool'],
        ['State', 'Minimal', 'Tracks iterations, contexts, citations, answer'],
        ['Evaluation', 'Often ad hoc', 'Built-in dataset harness and metrics'],
        ['Guardrails', 'Rare', 'Prompt-injection, PII, grounding score'],
        ['Observability', 'Sparse', 'Tracer spans and token/cost accounting'],
    ]
    add_table(doc, rows)
    doc.add_paragraph('IMPLEMENTED: query planning, retrieval tool calling, multi-hop loop, guardrails, tracing, evaluation, provider abstraction. PARTIALLY IMPLEMENTED: real LLM planner, cross-encoder reranking, richer tool schemas. NOT IMPLEMENTED: arbitrary autonomous tool execution, persistent per-user memory, long-lived session state, full enterprise auth. RECOMMENDED EXTENSION: stronger planner validation, real semantic reranker, enterprise authorization, persistent vector DB.')

    # Section 6
    add_section(doc, '6. PLANNER')
    doc.add_paragraph('The planner exists to decompose complex questions into multiple retrieval tasks. In planner.py, plan(question) uses regex splits on question marks, conjunctions, semicolons, and “then” to create sub-questions. If the split produces nothing useful, it falls back to the original question.')
    doc.add_paragraph('Why it exists: multi-hop questions often need evidence from multiple places. Without a planner, the system would dump a broad query into a retriever and hope the top-k covers everything.')
    doc.add_paragraph('What it receives: a raw string question. What it produces: list[str] of sub-questions. Decision logic: deterministic heuristic, not LLM-based. The repo explicitly supports LLM-based planning as a future swap because the interface is simple: same return contract, list[str].')
    doc.add_paragraph('Interaction with the graph: in graph.py, plan_node calls plan() and stores sub_questions in state; retrieve_node iterates over each sub-question and collects results; generate_node uses combined context for synthesis; reformulate_node appends context expansion and loops back to plan. This is a clear node-based orchestration.')
    doc.add_paragraph('Failure handling: if the planner yields poor decomposition, the system still continues with the original question as a fallback. The downstream faithfulness gate acts as an external quality check.')
    doc.add_paragraph('Interview answers: “Why do you need a planner if the LLM can directly answer?” — because planning narrows retrieval and isolates facts. “Why not just call the retriever every time?” — because one large query often produces mixed evidence. “How do you prevent loops?” — max_iterations and faithfulness threshold. “Would you use an LLM planner in production?” — maybe, but only with plan validation and strong guardrails.')

    # Section 7
    add_section(doc, '7. RETRIEVAL')
    doc.add_paragraph('The retrieval subsystem is the evidence engine. It combines lexical BM25 with vector search and then fuses the rankings.')
    doc.add_paragraph('Document loading: load_directory reads .md and .txt files, wraps them into Document(doc_id, text). Chunking: fixed_size_chunks uses size=120 and overlap=24 by default; markdown_section_chunks splits on headings. Metadata: each Chunk includes id, text, metadata with doc_id, section, start, and strategy. Embeddings: HashingEmbedder produces deterministic sparse vectors; OpenAIEmbedder produces real dense vectors from OpenAI/Azure APIs.')
    doc.add_paragraph('Vector search: VectorStore.search measures cosine similarity between query embedding and stored chunk embeddings. BM25: BM25.search scores documents by IDF and term frequency with stemming-aware tokenization. Hybrid retrieval: HybridRetriever.retrieve performs lexical search and dense search, then uses reciprocal rank fusion. Rerank: _rerank adds lexical overlap and phrase bonus before trimming to top-k.')
    doc.add_paragraph('Why the decisions exist: chunk overlap preserves context boundaries; section chunking respects document structure; BM25 is good for keyword precision; vector search handles semantic paraphrases; RRF balances two candidate rankings; rerank improves the final shortlist. Retrieval problems such as poor chunking, semantic mismatch, duplicate chunks, and context overflow are addressed with chunk strategy, hybrid matching, deduplication, and low top_k.')
    doc.add_paragraph('Reranking is [IMPLEMENTED]. It happens inside HybridRetriever._rerank with overlap and phrase bonuses. A production extension would switch to a cross-encoder reranker or an LLM judge when more signal is needed.')

    # Section 8
    add_section(doc, '8. GUARDRAILS')
    doc.add_paragraph('Guardrails are lightweight but important. check_input enforces max length and pattern-based prompt-injection detection. It raises InputBlocked for suspicious content like “ignore all previous instructions”, “reveal the system prompt”, or overly long input. PII is removed via redact_pii in final responses. grounding_score computes token-overlap between answer and context as a crude faithfulness measure.')
    rows = [
        ['Type', 'Current implementation', 'What it protects against'],
        ['Input guardrail', 'Length cap and prompt-injection regex', 'Oversized requests and obvious injection attempts'],
        ['Output guardrail', 'PII redaction', 'SSNs, card numbers, email addresses'],
        ['Grounding guardrail', 'grounding_score()', 'Answers unsupported by retrieved context'],
        ['Tool guardrail', 'Optional gateway integration', 'Auth, rate limiting, audit, schema validation when installed'],
    ]
    add_table(doc, rows)
    doc.add_paragraph('Current Protection vs Production-Grade Protection:')
    rows2 = [
        ['Area', 'Current project', 'Production-grade extension'],
        ['Prompt injection', 'Regex-based blocking', 'Policy-based filtering, document provenance, red-team tests'],
        ['Tool safety', 'Optional secure gateway', 'RBAC, least privilege, execution sandboxing'],
        ['Hallucination', 'Grounding score', 'Citations verification, abstention logic, judge-based evaluation'],
        ['Sensitive data', 'PII redaction', 'Classification, encryption, per-tenant isolation'],
        ['Malicious docs', 'Not implemented', 'Scanning, sanitization, provenance control'],
    ]
    add_table(doc, rows2)
    doc.add_paragraph('Interview answers: “How do you prevent prompt injection?” – fail-closed regex checks at the input boundary. “Can a retrieved document manipulate the agent?” – yes, the current project does not fully defend against indirect prompt injection in the corpus. “How do you prevent unauthorized tool calls?” – optional gateway security plane. “How do you validate tool arguments?” – via schema in the optional gateway; default Python method is not strongly schema validated. “How do you stop infinite loops?” – max_iterations and threshold. “How do you handle malicious documents?” – still a recommended extension, not fully implemented.')

    # Section 9
    add_section(doc, '9. TOOLS')
    doc.add_paragraph('The project treats retrieval as a tool. RetrievalTool in agents/tools.py wraps HybridRetriever and exposes search_knowledge_base as its callable interface. The tool takes a query string and returns a list[Scored] with chunk + score metadata.')
    doc.add_paragraph('Purpose: retrieve relevant chunks when the agent is missing evidence. Input schema: a single string query. Output: ranked candidate list. Invocation mechanism: self.tool(sub_question). Validation: minimal in default implementation; the optional gateway layer enforces a schema and auth layer when installed. Error handling: not extensive by default; bad calls can bubble up. Permissions: absent in the base code; only in optional SecuredRetrievalTool via MCP gateway. Interaction with the agent: the tool result updates local state, gets deduplicated, and is then used for synthesis.')
    doc.add_paragraph('Tool calling makes the system more agentic because evidence gathering is no longer a fixed pre-step. The agent decides when to call the tool, what to retrieve, and how to use the results to refine or re-answer.')

    # Section 10
    add_section(doc, '10. GRAPH / ORCHESTRATION')
    doc.add_paragraph('The optional orchestration engine is in graph.py. When LangGraph is installed, build_rag_graph creates a StateGraph with plan, retrieve, generate, and reformulate nodes. It uses START -> plan -> retrieve -> generate -> conditional edge -> END or reformulate.')
    add_code_block(doc, (
        'START\n'
        '  |\n'
        '  v\n'
        'plan\n'
        '  |\n'
        '  v\n'
        'retrieve\n'
        '  |\n'
        '  v\n'
        'generate\n'
        '  |\n'
        '  +--> faithfulness >= threshold or iterations >= max_iterations => END\n'
        '  |\n'
        '  +--> otherwise => reformulate -> plan\n'
    ))
    doc.add_paragraph('State: RAGState includes question, query, sub_questions, contexts, citations, answer, faithfulness, iterations. Routing is conditional. If grounding is strong enough, execution ends; otherwise it loops. This is a simple but meaningful state machine.')
    doc.add_paragraph('Failure behavior: no relevant documents -> low grounding; planner failure -> fall back to original query; LLM failure -> provider exception; tool failure -> default failure propagation; guardrail rejection -> blocked request; max iterations -> exit with best result found.')

    # Section 11
    add_section(doc, '11. LLM + EMBEDDING PROVIDERS')
    doc.add_paragraph('The provider logic is abstracted via protocols. LLMProvider requires complete(system: str, user: str) -> LLMResult. Embedder requires embed(text: str) -> dict[int, float]. That allows MockLLM, OpenAILLM, HashingEmbedder, and OpenAIEmbedder to be swapped behind the same interface.')
    doc.add_paragraph('Why not hard-code one provider? Because the project wants offline determinism in CI and real model access in production. The environment variables in Settings control the choice: RAG_LLM_PROVIDER, RAG_EMBED_PROVIDER, RAG_LLM_MODEL, RAG_EMBED_MODEL. This is a classic separation-of-concerns design: the orchestration logic is decoupled from vendor APIs.')
    doc.add_paragraph('Trade-offs: the abstraction is intentionally simple and low-level; it does not implement retry logic, rate limiting, fallback chains, or cost-aware routing. That is a conscious choice for readability and a sensible reference architecture. The project demonstrates provider abstraction without overengineering it.')

    # Section 12
    add_section(doc, '12. INGESTION')
    doc.add_paragraph('Ingestion begins with Document objects from loaders. load_file reads a file as UTF-8 text; load_directory recursively loads matching .md and .txt files. Each document is assigned a doc_id from the file stem and then chunked.')
    doc.add_paragraph('Chunking is critical because it defines the evidence unit. fixed_size_chunks uses a sliding window and overlap; markdown_section_chunks splits on headings. The chunk metadata retains structure so the system can later cite the document section and reason about provenance.')
    doc.add_paragraph('The lifecycle is: raw documents -> parse -> chunk -> metadata -> embed -> index. This is the backbone of retrieval quality. Poor chunking is a common cause of missing context or noisy retrieval, which is why the repo emphasizes overlap and section-aware strategies.')

    # Section 13
    add_section(doc, '13. EVALUATION')
    doc.add_paragraph('The evaluation layer is in evaluation/harness.py and evaluation/metrics.py. The dataset is a JSONL file with question, answer, relevant_ids. run_eval loops over each case, invokes pipeline.query, compares actual citations with relevant_ids, and computes retrieval and answer quality metrics.')
    doc.add_paragraph('Retrieval metrics: hit_rate, MRR, recall_at_k, nDCG_at_k. Generation metrics: faithfulness (overlap between answer tokens and context tokens) and answer_correctness (token-F1 vs reference answer). This is a pragmatic set of metrics for a small RAG system.')
    doc.add_paragraph('Unit tests validate low-level behaviors. Integration tests validate composition. RAG evaluation validates retrieval + answer quality over labelled cases. LLM evaluation uses human or model-based judges and is not implemented here. End-to-end evaluation is exactly what the project does with run_eval.')

    # Section 14
    add_section(doc, '14. OBSERVABILITY')
    doc.add_paragraph('Observability is handled by Tracer in observability.py. It records nested spans, computes durations, and tracks input_tokens, output_tokens, and cost_usd. The summary includes the number of spans and token usage. This is a lightweight tracing mechanism suitable for local debugging and simple production monitoring.')
    doc.add_paragraph('Metrics supported in the project: span count, duration by span, token counts, approximate cost. Production recommendations would include request latency, retrieval latency, LLM latency, retrieval hit rate, empty-retrieval rate, tool error rate, guardrail rejection rate, and agent loop count. Those are not fully implemented in current code, but they are the obvious next step.')

    # Section 15
    add_section(doc, '15. TESTING')
    doc.add_paragraph('The project has focused tests covering the real behavior of the system:')
    add_bullets(doc, [
        'test_agent.py checks planner decomposition, multi-hop citations, and grounded answers.',
        'test_guardrails.py checks injection blocking, PII redaction, and grounding score.',
        'test_integrations.py checks observability export fallback and optional security integration.',
        'test_langgraph.py checks optional LangGraph engine behavior and fallback.',
        'test_pipeline.py checks query grounding and end-to-end evaluation quality.',
        'test_retrieval.py checks chunk overlap, BM25 ranking, and hybrid retrieval.',
        'test_text.py checks tokenization, stemming, and cosine similarity.',
    ])
    doc.add_paragraph('Overall strategy: deterministic offline tests that exercise the real retrieval and generation flow without external services. This is exactly the right strategy for a reference implementation and strong interview prep material.')

    # Section 16
    add_section(doc, '16. API')
    doc.add_paragraph('The API is a minimal FastAPI app in api/main.py. It has /health and /ask. AskRequest is a Pydantic model with a question: str. AskResponse includes answer, citations, faithfulness, iterations, blocked. get_pipeline lazily creates the shared RAGPipeline and loads the corpus once. The service is intentionally thin and delegates almost everything to the core pipeline.')
    doc.add_paragraph('Production considerations: no authentication, no authorization, no rate limiting, no request queue, no async workers. This is a valid prototype API but not a production service boundary.')

    # Section 17
    add_section(doc, '17. DESIGN DECISIONS')
    rows = [
        ['Decision', 'Problem addressed', 'Why this approach'],
        ['Agentic RAG', 'Single-hop retrieval is often insufficient', 'Adds looped planning and retrieval'],
        ['Planner', 'Complex queries need decomposition', 'Split multi-fact requests into smaller retrieval tasks'],
        ['Graph orchestration', 'Need explicit control flow', 'StateGraph expresses conditional retry logic'],
        ['Provider abstraction', 'Need offline + cloud flexibility', 'Swap mock or real models without changing business logic'],
        ['Guardrails', 'LLM systems need protection', 'Fail-closed behavior at input and output boundaries'],
        ['Evaluation harness', 'RAG quality must be measurable', 'Labeled dataset + retrieval/answer metrics'],
        ['Separate ingestion', 'Retrieval quality depends on chunk structure', 'Chunk boundaries and metadata matter'],
        ['Separate retrieval', 'Different concerns need different metrics', 'BM25 and dense search can be optimized independently'],
        ['Observability', 'LLM debugging is hard without traces', 'Span-based tracing and token accounting'],
        ['Tools', 'Agent needs evidence fetch boundary', 'Retrieval is explicit and traceable'],
    ]
    add_table(doc, rows)

    # Section 18
    add_section(doc, '18. TRADE-OFFS')
    doc.add_paragraph('Complexity vs capability: the project is more complex than a basic RAG demo, but it is more resilient and more testable. Latency vs reasoning: each extra retrieval hop adds latency and token cost. Cost vs quality: better evidence selection and grounding can improve quality but increases compute. Flexibility vs determinism: the planner and reranker are deterministic heuristics by default; they are easier to reason about but less flexible than an LLM planner. Agent autonomy vs control: the system is agentic enough for retrieval loops but intentionally bounded. More guardrails vs usability: strict checks improve safety but can reduce helpfulness in edge cases. More retrieval vs context size: top_k limits context length, which is a pragmatic trade-off.')
    doc.add_paragraph('Where this architecture may be overkill: simple FAQ bots, tiny support knowledge bases, or deterministic internal policies. Basic RAG is enough when the question distribution is narrow and one retrieval step is sufficient. This project adds value when the questions are multi-hop, the corpus is messy, and the team cares about measured quality and safety.')

    # Section 19
    add_section(doc, '19. FAILURE MODES')
    rows = [
        ['Failure', 'Detection', 'Cause', 'Recovery', 'User impact'],
        ['Empty retrieval', 'No hits in lexical or dense search', 'No relevant chunk matched', 'Low grounding may trigger reformulation or abstention', 'Limited answer or no answer'],
        ['Bad retrieval', 'Low grounding_score', 'Chunk mismatch or poor query', 'Retry with reformulated query', 'Partially grounded answer'],
        ['LLM timeout/failure', 'Exception from provider', 'Network or model issue', '[NOT IMPLEMENTED] retry/backoff', 'Request may fail'],
        ['Tool failure', 'Exception from retrieval tool', 'Bad retriever state or call failure', 'Minimal handling in base code', 'Request may fail'],
        ['Invalid tool args', 'Type or schema mismatch', 'Bad input or malicious payload', 'Optional gateway validates schema', 'Rejected call'],
        ['Planner failure', 'Fallback to original question', 'Heuristic decomposition misses structure', 'Continue with original question', 'Suboptimal retrieval'],
        ['Guardrail rejection', 'InputBlocked / PII redaction', 'Unsafe or sensitive input', 'Block or sanitize', 'Request blocked or answer sanitized'],
        ['Context overflow', 'Large context list', 'Too many chunks retrieved', 'top_k + rerank', 'Truncated or less useful answer'],
        ['Hallucination', 'Grounding score low', 'Missing evidence or weak retrieval', 'Abstain + retry', 'Risk of unsupported answer'],
        ['Prompt injection', 'Regex detection', 'Malicious user prompt', 'Block request', 'Request denied'],
        ['Infinite loop', 'max_iterations reached', 'Repeated low grounding', 'Hard stop', 'Bounded runtime'],
        ['Malformed docs', 'Chunker/parser issue', 'Bad markdown or encoding', 'Load raw text; no deep sanitizer', 'Noisy retrieval'],
    ]
    add_table(doc, rows)

    # Section 20
    add_section(doc, '20. SECURITY')
    doc.add_paragraph('The current project has basic security protections but not full enterprise hardening. It blocks obvious injection patterns, enforces max length, redacts PII from answer text, and provides an optional secured retrieval gateway. However, it does not have a robust auth layer, per-user permissions, or malicious-document scanning. A retrieved document can still influence the answer if it contains manipulative or instruction-like text, because the current implementation does not strongly distinguish corpus content from user instruction.')
    doc.add_paragraph('For enterprise production, add authenticated API access, per-tenant authZ, RBAC for tools, input and output policy enforcement, secrets management, provenance tracking, content sanitization, and stronger prompt-injection detection. The optional gateway is the start of a proper defense-in-depth design.')

    # Section 21
    add_section(doc, '21. PRODUCTIONIZATION')
    doc.add_paragraph('If this system had to serve 10,000 users, I would change several things. CURRENT SYSTEM: in-memory local index, synchronous execution, single-process pipeline, simple tool wrappers, no authN/authZ, no async queueing, no retries or circuit breakers. PRODUCTION EVOLUTION: persistent vector DB, async workers, cache layer, rate limiting, auth and authorization, timeouts and retries, circuit breakers, request tracing to vendor systems, secret management, CI/CD checks, and stronger evaluation gates.')
    doc.add_paragraph('The project is a reference implementation and a sensible architecture base, not a full production deployment. It demonstrates the right control plane and evaluation mindset, which is what matters in interviews.')

    # Section 22
    add_section(doc, '22. INTERVIEW QUESTIONS')
    qna = [
        ['LEVEL 1 — Basic understanding', 'What is this project?', 'It is a provider-agnostic agentic RAG system that ingests a corpus, retrieves relevant evidence, and synthesizes grounded answers with citations.'],
        ['LEVEL 2 — Architecture', 'How is the project organized?', 'The project separates providers, ingestion, retrieval, agents, guardrails, observability, evaluation, and API.'],
        ['LEVEL 3 — RAG', 'Why is retrieval hybrid?', 'Because lexical algorithms and vector similarity catch different signal types.'],
        ['LEVEL 4 — Agentic AI', 'What makes this agentic?', 'It plans sub-questions, calls a retrieval tool, tracks state, and retries with reformulation.'],
        ['LEVEL 5 — LLMs', 'Why have a mock LLM?', 'To make the system deterministic and CI-friendly without external APIs.'],
        ['LEVEL 6 — Retrieval', 'What is RRF?', 'Reciprocal Rank Fusion combines multiple rankings into one fusion score.'],
        ['LEVEL 7 — Guardrails', 'What guardrails exist?', 'Prompt-injection blocking, length checks, PII redaction, and grounding score.'],
        ['LEVEL 8 — Evaluation', 'How do you know it works?', 'Run the dataset harness and check retrieval and answer metrics.'],
        ['LEVEL 9 — Production', 'What would you change for scale?', 'Persistent vector DB, auth, queueing, caching, retries, and stronger monitoring.'],
        ['LEVEL 10 — Hard cross-question', 'Why is this not just RAG with an extra LLM?', 'Because it introduces planning, tool use, state, and grounded retry loops.'],
    ]
    for entry in qna:
        doc.add_paragraph(entry[0])
        doc.add_paragraph('Q: ' + entry[1])
        doc.add_paragraph('A: ' + entry[2])

    # Section 23
    add_section(doc, '23. INTERVIEWER ATTACK MODE')
    attacks = [
        ('Interviewer: “Why is this agentic RAG? Isn’t it just RAG with an extra LLM?”', 'Candidate: “The project adds explicit planning, tool-level retrieval, state tracking, iteration, and grounding-based self-correction. That is a workflow, not just a larger prompt.”'),
        ('Interviewer: “Then why is the planner necessary?”', 'Candidate: “Because multi-fact questions need targeted evidence gathering; a single broad query often mixes multiple intents.”'),
        ('Interviewer: “What happens if retrieval returns garbage?”', 'Candidate: “The grounding score drops, the agent reformulates the query, and the loop exits when max_iterations or threshold is reached.”'),
        ('Interviewer: “What stops dangerous tool calls?”', 'Candidate: “The default implementation doesn’t allow arbitrary tool execution. The optional gateway adds auth, schema validation, and audit layers.”'),
        ('Interviewer: “Does LangGraph make it agentic?”', 'Candidate: “No. LangGraph is orchestration, not the agentic behavior itself. The behavior comes from state, planning, and retrieval loops.”'),
        ('Interviewer: “How do you know your RAG is working?”', 'Candidate: “We measure retrieval quality and groundedness with a labeled evaluation harness.”'),
        ('Interviewer: “Why not just stuff all docs into the prompt?”', 'Candidate: “Because context windows are bounded and retrieval is necessary for precision, cost control, and modularity.”'),
        ('Interviewer: “Why do we need embeddings if BM25 works?”', 'Candidate: “BM25 is lexical; embeddings capture meaning and semantic paraphrases. Hybrid retrieval combines both.”'),
    ]
    for q, a in attacks:
        doc.add_paragraph(q)
        doc.add_paragraph(a)

    # Section 24
    add_section(doc, '24. TRICK QUESTIONS')
    trick = [
        'Why isn’t every RAG system agentic? Because agentic behavior requires an explicit control loop and evidence-driven decision-making.',
        'Does using LangGraph automatically make something agentic? No. It is just orchestration. The system still needs planning and tool decisions.',
        'Does tool calling automatically make something autonomous? No. Tool use becomes autonomous only when the system has guardrails, loops, and bounded decision-making.',
        'Why not simply use a larger context window? Because the cost and noise increase without solving retrieval quality.',
        'Why not put all documents into the prompt? Because the same limits and cost constraints apply; retrieval is still needed.',
        'Why do we need embeddings? For semantic similarity beyond keyword overlap.',
        'Why can good embeddings still produce bad retrieval? Because chunking, metadata, and query drift matter too.',
        'Why can high recall still produce bad answers? The model may still answer with weakly supported evidence.',
        'Why are guardrails not enough for security? They are a boundary check, not a full defense-in-depth system.',
        'Why is evaluation accuracy alone insufficient? Because latency, safety, and cost matter as much as raw correctness.',
    ]
    add_bullets(doc, trick)

    # Section 25
    add_section(doc, '25. KEY CONCEPTS I MUST KNOW')
    add_bullets(doc, [
        'Python concepts: dataclasses, protocol interfaces, type hints, environment-based config, and module separation.',
        'Pydantic: used in FastAPI request/response models.',
        'LLMs: provider interface plus mock and OpenAI/Azure adapters.',
        'Embeddings: hash-based deterministic embedding plus OpenAI embedder.',
        'Vector search: cosine similarity with in-memory cache.',
        'Chunking: fixed-size overlap and markdown section strategies.',
        'RAG: retrieval-before-generation grounded by evidence.',
        'Agentic RAG: planning, tool use, state, and iterative refinement.',
        'Graph: optional StateGraph orchestration with retries.',
        'Guardrails: injection detection, length limits, PII redaction, grounding score.',
        'Observability: span-based tracing and token/cost accounting.',
        'Evaluation: labeled QA harness plus retrieval and answer metrics.',
        'Docker: packaging support for deployment.',
    ])

    # Section 26
    add_section(doc, '26. WHAT I SHOULD SAY ABOUT THE PROJECT')
    doc.add_paragraph('“I worked with a production-aware agentic RAG system built around a central RAGPipeline. It loads a corpus, chunks and indexes it, then retrieves evidence using hybrid search combining BM25 and vector similarity with reciprocal rank fusion. The user query is decomposed by a planner, each sub-question is passed through a retrieval tool, and the system synthesizes a grounded answer while tracking citations and faithfulness. It includes guardrails for prompt injection and sensitive data, along with a built-in evaluation harness and trace-based observability. The core runs offline with a mock LLM and hashing embeddings, making the system easy to test and debug before switching to OpenAI or Azure.”')

    # Section 27
    add_section(doc, '27. PROJECT WALKTHROUGH SCRIPT')
    doc.add_paragraph('“Let me start with the repository structure. The core code lives under src/production_agentic_rag and is organized by providers, ingestion, retrieval, agents, guardrails, evaluation, and observability. I’ll begin with config.py and pipeline.py, then move to planner.py and rag_agent.py to explain the reasoning loop, then show retrieval/hybrid.py and chunking.py, then guardrails.py, evaluation/harness.py, and finally api/main.py.”')
    doc.add_paragraph('Suggested file walk-through order: config.py -> pipeline.py -> agents/planner.py -> agents/rag_agent.py -> agents/tools.py -> retrieval/hybrid.py -> retrieval/bm25.py -> retrieval/vectorstore.py -> ingestion/loaders.py -> ingestion/chunking.py -> providers/llm.py -> providers/embeddings.py -> guardrails.py -> graph.py -> observability.py -> evaluation/harness.py -> evaluation/metrics.py -> api/main.py -> tests/test_pipeline.py.')

    # Section 28
    add_section(doc, '28. 1-DAY INTERVIEW PREPARATION')
    doc.add_paragraph('Hour 1: project overview, architecture, and request lifecycle. Hour 2: retrieval and ingestion. Hour 3: planner and agent loop. Hour 4: guardrails/security. Hour 5: evaluation and observability. Hour 6: mock interview and red-team review.')
    doc.add_paragraph('MUST KNOW: RAGPipeline.query, AgenticRAG.run, HybridRetriever.retrieve, grounding_score, planner.plan, Settings, and the difference between built-in agent and LangGraph engine.')
    doc.add_paragraph('SHOULD KNOW: BM25, RRF, chunking strategies, API contract, evaluation metrics, and failure recovery. NICE TO KNOW: secure gateway, model routing, and production deployment evolution.')

    # Section 29
    add_section(doc, '29. FINAL CHEAT SHEET')
    doc.add_paragraph('Architecture in one line: User -> API/CLI -> RAGPipeline -> input guardrail -> planner -> retrieval -> LLM -> grounding check -> optional reformulate -> final answer with citations and trace.')
    doc.add_paragraph('Important files: pipeline.py, graph.py, agents/planner.py, agents/rag_agent.py, agents/tools.py, retrieval/hybrid.py, guardrails.py, observability.py, evaluation/harness.py, api/main.py.')
    doc.add_paragraph('Key terminology: hybrid retrieval, reciprocal rank fusion, grounding, faithfulness threshold, max_iterations, prompt injection, PII redaction, citations, eval harness, tracer spans.')
    doc.add_paragraph('30-second explanation: “This is a production-aware, provider-agnostic agentic RAG system with hybrid retrieval, explicit planning, grounding checks, guardrails, and evaluation.”')
    doc.add_paragraph('2-minute explanation: “It loads documents, chunks them, indexes them, plans sub-queries, retrieves evidence, synthesizes answers, scores grounding, and retries with reformulated queries if needed.”')
    doc.add_paragraph('5-minute explanation: “The architecture is centered on RAGPipeline. It uses BM25 + vector retrieval with RRF, a planner, a retrieval tool, a faithfulness gate, a LangGraph variant, and an evaluation harness to measure quality.”')

    out = 'INTERVIEW_MASTER_GUIDE_production_agentic_rag.docx'
    doc.save(out)
    return out


if __name__ == '__main__':
    output = build_document()
    print(f'Created {output}')
