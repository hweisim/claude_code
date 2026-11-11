# Langfuse: how it works

A walkthrough of the Langfuse source code, the open-source LLM observability, evaluation and prompt management platform. It covers how the system is put together, how data moves through it, how it is stored, and the design choices behind it.

- **Source analysed:** `github.com/langfuse/langfuse`, `main` @ `47a7034` (2026-11-10), monorepo version `4.78.0`
- **License:** MIT, except `ee/`, `web/src/ee/`, `worker/src/ee/` and `packages/shared/src/server/ee/`, which use a separate enterprise license. The copyright holder is now **ClickHouse, Inc.**
- **Method:** reading the code only. Nothing was run. File paths below are relative to the Langfuse repo root.

---

## 1. Summary

Langfuse is a **TypeScript monorepo** (pnpm + Turborepo) that ships as two Node containers, `web` and `worker`, plus a newer Rust `ai-gateway`. It depends on four stateful backends:

| Backend | Role |
|---|---|
| **PostgreSQL** (Prisma, 83 models) | Control plane: users, orgs, projects, API keys, prompts, datasets, eval configs, model prices, integrations |
| **ClickHouse** | Data plane: traces, observations, scores, and the newer wide `events_*` tables |
| **Redis** (BullMQ + caches) | About 30 job queues, API-key and prompt caches, rate limits, cursors |
| **S3-compatible blob store** (MinIO locally) | Write-ahead log of every raw ingestion event, plus media, exports and overflow storage |

The core design is **"S3 first, then queue, then merge, then ClickHouse."** The API never writes telemetry straight to the database. It stores each raw event in S3, enqueues a small pointer, and returns. The worker later reads *all* S3 files for that entity, merges them into a single row, and batch-inserts it into a ClickHouse `ReplacingMergeTree`. This gives cheap retries and replays, idempotent upserts from many partial updates, and isolation from ClickHouse insert pressure.

---

## 2. Repository layout

```text
web/                 Next.js app: React UI, tRPC (~140 routers), public REST API, OTLP endpoint
worker/              Express + BullMQ consumers: ingestion, evals, exports, integrations, deletions
packages/shared/     Domain models, Prisma schema, ClickHouse migrations, queue contracts, repositories
packages/native/     Rust napi-rs addon loaded by worker (ClickHouse Native encoding, OTel media, UMAP/HDBSCAN)
ai-gateway/          Standalone Rust proxy for OpenAI/Anthropic APIs that emits generations via OTLP
ee/                  Enterprise license check + gated features
fern/                API definitions -> generated SDK clients / OpenAPI
```

Dependency rules are strict and enforced: `web → shared, ee`; `worker → shared, native`; `shared` imports nothing from the others. Queue names and payload schemas live in one place, `packages/shared/src/server/queues.ts`.

---

## 3. Deployment topology

From `docker-compose.yml`:

```
           ┌─────────────── SDKs (Python/JS), OTel exporters, AI gateway ───────────────┐
           │                                                                            │
           ▼                                                                            ▼
   langfuse-web :3000  ──────────────►  Redis (BullMQ)  ◄──────────────  langfuse-worker :3030
   (Next.js, UI+API)                         │                         (queue consumers)
        │   │    │                           │                               │   │   │
        │   │    └──────── S3 / MinIO ◄──────┼───────────────────────────────┘   │   │
        │   └───────────── PostgreSQL ◄──────┴───────────────────────────────────┘   │
        └───────────────── ClickHouse (reads) ◄──────────────── (batched writes) ────┘

   langfuse-ai-gateway :3031 (Rust)  — optional LLM proxy; forwards captured generations as OTLP
```

- `web` and `worker` are the **same codebase and configuration** (`<<: *langfuse-worker-env`). `web` adds NextAuth and UI settings.
- Both scale horizontally. All coordination happens through Redis.
- `LANGFUSE_INIT_*` variables bootstrap an org, project, keys and user on first boot (headless setup).

---

## 4. Write path: ingestion

### 4.1 Two front doors

| Endpoint | Format | Notes |
|---|---|---|
| `POST /api/public/ingestion` | Langfuse batch JSON (`trace-create`, `span-update`, `generation-create`, `score-create`, …) | Legacy SDK protocol, 4.5 MB body limit, now carries deprecation headers |
| `POST /api/public/otel/v1/traces` | OTLP/HTTP, protobuf or JSON, optionally gzip | **Primary path for v3+ SDKs.** Langfuse SDKs are now thin OpenTelemetry wrappers |

### 4.2 Request handling (`web/src/pages/api/public/ingestion.ts` → `packages/shared/src/server/ingestion/processEventBatch.ts`)

1. **Auth.** Basic auth with `publicKey:secretKey`. The secret is checked against `fastHashedSecretKey`, a SHA-256 of the key plus `SALT`, which is fast enough to verify on every request. Results are cached in Redis. Each event then authorizes its own action (`project` versus `scores`-only key scopes).
2. **Guards.** Rate limits (`rate-limiter-flexible` on Redis, per org/plan/resource, per-minute buckets, some costed in bytes), an ingestion-suspended flag for usage thresholds, and Zod validation of each event.
3. **Group** events by entity (`eventBodyId`).
4. **Upload each event to S3** at `<prefix>/<projectId>/<entityType>/<eventBodyId>/<eventId>.json`. OTel batches go to `otel/<projectId>/yyyy/mm/dd/hh/mm/<id>.json`. If S3 returns `SlowDown`, the project is flagged in Redis and routed to a *secondary* queue.
5. **Sample.** `isTraceIdInSample` drops events deterministically by trace id when sampling is configured.
6. **Enqueue** one BullMQ `IngestionJob` per entity. The job is a pointer (`projectId`, `type`, `eventBodyId`, `fileKey`, `skipS3List`), not the payload. Queues are **sharded**: `sha256(projectId-eventBodyId) mod N` picks the Redis queue shard.
7. **Delay.** Jobs normally wait up to 5 s (OTel and scores run immediately). Between **23:45 and 00:15 UTC** they wait 15 s (`LANGFUSE_INGESTION_QUEUE_DELAY_MS`). The reason: ClickHouse sort keys include `toDate(timestamp)`, so an out-of-order update that crosses midnight would create a duplicate row that `ReplacingMergeTree` can never collapse.
8. Return **207 Multi-Status** with per-event success or error.

### 4.3 Worker merge (`worker/src/queues/ingestionQueue.ts`, `worker/src/services/IngestionService/index.ts`)

```
job ─► list S3 prefix for this entity (or direct download if skipS3List)
    ─► download + parse every partial event for the entity
    ─► sort by timestamp
    ─► mergeRecords(): lodash.mergeWith, later non-null values win,
       immutable keys (id, project_id, start_time…) never overwrite
    ─► enrich: model match → usage tokenization → cost calc,
       prompt lookup (name+version → prompt_id), tool-call extraction,
       metadata flattening to Map(String,String)
    ─► ClickhouseWriter.addToQueue(table, record)
```

**Why it re-reads all of S3:** SDKs send `create`, then several `update`s, for the same span, and these can land on different workers. Rebuilding the row from every fragment makes each job idempotent and order-independent. The newest full row always wins in ClickHouse.

### 4.4 `ClickhouseWriter` (`worker/src/services/ClickhouseWriter/`)

An in-memory buffer per table, flushed every **1000 ms** or every **1000 rows** (both configurable), with retry and back-off and a graceful drain on shutdown. Tables written: `traces`, `observations`, `scores`, `dataset_run_items`, `observations_batch_staging`, `events_full`, `blob_storage_file_log`, `topic_*`. The Rust addon in `packages/native` can encode inserts in **ClickHouse Native** format instead of JSONEachRow for throughput.

### 4.5 OTel mapping (`packages/shared/src/server/otel/`)

`OtelIngestionProcessor.ts` (~2000 lines) converts arbitrary OTel spans into Langfuse observations. It reads several semantic-convention dialects:

| Ecosystem | Attributes recognised |
|---|---|
| Langfuse native | `langfuse.trace.*`, `langfuse.observation.{type,input,output,model.name,usage_details,cost_details,prompt.name,…}`, `langfuse.session.id`, `langfuse.user.id`, `langfuse.experiment.*` |
| OTel GenAI semconv | `gen_ai.prompt.N.*`, `gen_ai.completion.N.*`, `gen_ai.input.messages`, `gen_ai.output.messages`, `gen_ai.tool.call.*`, span events `gen_ai.user.message` / `gen_ai.choice` |
| Vercel AI SDK | `ai.prompt.messages`, `ai.response.text`, `ai.toolCall.*`, `ai.telemetry.metadata.*` |
| OpenInference / Arize | `llm.input_messages.*`, `llm.output_messages.*` |
| Others | `mlflow.spanInputs`, `traceloop.entity.*`, `pydantic_ai.all_messages`, Logfire |

`ObservationTypeMapper.ts` then classifies each span as one of `GENERATION | AGENT | TOOL | CHAIN | RETRIEVER | EMBEDDING | EVALUATOR | GUARDRAIL | SPAN | EVENT`. Inline base64 media is pulled out to S3 by `OtelMediaProcessor` (and a Rust fast path).

---

## 5. Storage model

### 5.1 ClickHouse, legacy v3 tables (`packages/shared/clickhouse/migrations/canonical/0001-0003`)

All tables use `ReplacingMergeTree(event_ts, is_deleted)`. The highest `event_ts` wins, and `is_deleted=1` acts as a tombstone.

| Table | ORDER BY | Partition | Notes |
|---|---|---|---|
| `traces` | `(project_id, toDate(timestamp), id)` | month | `input`/`output` ZSTD(3), bloom filters on `id` and metadata keys/values |
| `observations` | `(project_id, type, toDate(start_time), id)` | month | Usage and cost as `Map(LowCardinality(String), UInt64/Decimal64(12))`, so any token type (`input`, `output`, `cache_read`, `reasoning`…) fits without schema changes |
| `scores` | `(project_id, toDate(timestamp), name, id)` | month | Numeric, categorical and boolean scores via `value` / `string_value` / `data_type` |

Reads deduplicate at query time with `FINAL` (about 68 occurrences in repositories) or `LIMIT 1 BY id, project_id`. This read-time deduplication cost is the main thing the v4 redesign tries to remove.

### 5.2 ClickHouse, v4 "wide events" (`0038-0041`, migration 0056 latest)

Based on `.agents/ARCHITECTURE_PRINCIPLES.md` ("observations are the primary analytical unit; prefer wide immutable events; denormalise to avoid joins"), v4 adds:

- **`events_full`**: one very wide row per span. It copies trace-level fields (`trace_name`, `user_id`, `session_id`, `tags`), experiment and dataset context, prompt, model, usage, cost, tools and OTel resource info onto **every** observation.
  - `ORDER BY (project_id, toStartOfMinute(start_time), xxHash32(trace_id), span_id, start_time)`
  - `SAMPLE BY xxHash32(trace_id)` for approximate analytics
  - `MATERIALIZED` cost roll-ups (`calculated_input_cost` and others) computed from the cost map
  - full-text and ngram indexes on name and metadata; 64 MiB granules for large rows
- **`events_core`**: filled by a materialized view from `events_full`, with `input`/`output`/metadata values **truncated to 200 chars**. List views, dashboards and filters scan this compact table. The full payload is fetched only on the detail page.
- **`observations_batch_staging`**: a 3-minute-partitioned landing table with a 48 h TTL. The `eventPropagation` worker walks closed partitions in order, keeping its cursor in Redis, joins in trace attributes, and writes `events_full`. This is how late-arriving trace-level updates such as `userId` set at the end of a request reach every span without UPDATE statements.

### 5.3 PostgreSQL control plane (`packages/shared/prisma/schema.prisma`)

Main model groups:

- **Tenancy and RBAC:** `Organization → Project`, `OrganizationMembership`, `ProjectMembership`, `RoleAssignment`, `ApiKey` (project or org scoped), `SsoConfig`, `VerifiedDomain`, `AuditLog`, SCIM
- **Prompts:** `Prompt` (name, version, labels such as `production`/`latest`, text or chat, config JSON), `PromptDependency` (prompt composition), `PromptProtectedLabels`
- **Datasets and experiments:** `Dataset`, `DatasetItem` (versioned), `DatasetRuns`; run items live in ClickHouse
- **Evals:** `EvalTemplate`, `Evaluator` / `EvaluatorVersion`, `EvaluationRule`, `JobConfiguration`, `JobExecution`, `ScoreConfig`, `AnnotationQueue*`
- **Pricing:** `Model` (`match_pattern` regex, tokenizer, `start_date`), `Price`, `PricingTier`
- **Integrations and automation:** PostHog, Mixpanel, BlobStorage export, Slack, PagerDuty, webhooks (`Trigger` → `Action` → `Automation`), `Monitor`, `BatchExport`, `BatchAction`
- **Newer:** `Skill*`, `InAppAgent*`, `Gateway*`, `Topic*` (embedding-based clustering of traces)

---

## 6. Key subsystems

### 6.1 Model matching and cost (`ingestion/modelMatch.ts`)

The provided model name is matched in Postgres with a regex (`WHERE $model ~ match_pattern`). Project-specific models take priority over global ones, and the latest `start_date` wins. The worker then:

1. uses the SDK-provided `usage_details` if present; otherwise it tokenizes `input`/`output` with the model's tokenizer (tiktoken or Claude);
2. uses provided `cost_details` if present; otherwise it computes `usage × price` per usage key, picking a **pricing tier** by conditions (for example, a long-context tier above N input tokens);
3. stores both `provided_*` and the computed values, so prices can be recalculated later.

### 6.2 Prompt management (`shared/src/server/services/PromptService`)

`GET /api/public/v2/prompts/:name?label=production` reads from a Redis read-through cache (`LANGFUSE_CACHE_PROMPT_TTL_SECONDS`, default 3600). Invalidation uses an **epoch key** per project: a write bumps the epoch, and old keys are ignored and expire on their own. SDKs add their own client cache with stale-while-revalidate, so prompt fetches are effectively free on the hot path. Generations link back through `prompt_name`/`prompt_version` → `prompt_id`, which enables per-prompt-version metrics.

### 6.3 Evaluations (`worker/src/features/evaluation/`)

```
trace/observation/dataset-run-item ingested
   └─► createEvalQueue: createEvalJobs() loads ACTIVE EvaluationRules for project
         ├─ filter (trace filters, environment allow-list, deterministic sampling)
         ├─ delay (let the trace "settle")
         └─► EvalExecutionQueue (sharded, secondary queue for noisy projects)
               └─► extractVariablesFromTracingData()  — JSONPath-style mapping of trace fields → template vars
               └─► runLLMAsJudgeEvaluation()          — calls user's LLM connection (LlmApiKeys, encrypted with ENCRYPTION_KEY)
                    with structured-output schema {score, reasoning}
               └─► score written back through the normal ingestion path (source = EVAL)
```

There are also **code-based evaluators** (`codeBased/`), run in a sandbox or sent to an external HTTP runner signed with `LANGFUSE_CODE_EVAL_EXTERNAL_SIGNING_SECRET`, plus whole-trace transcript evals and observation-level evals. Each evaluator run is itself traced as an internal Langfuse trace.

### 6.4 Queues (`packages/shared/src/server/queues.ts`, `worker/src/queues/`)

About 30 BullMQ queues. The important ones: `ingestion-queue` / `secondary-ingestion-queue`, `otel-ingestion-queue` (+ secondary), `create-eval-queue`, `evaluation-execution-queue` (+ secondary), `event-propagation-queue`, `trace-delete`, `project-delete`, `data-retention-queue`, `batch-export-queue`, `batch-action-queue`, `webhook-queue`, `entity-change-queue`, `posthog/mixpanel/blobstorage-integration-*`, `cloud-usage-metering-queue`, `dead-letter-retry-queue`, `in-app-agent-run-queue`.

Patterns used throughout:

- **Sharded queues** (SHA-256 key → shard) to spread load across Redis Cluster slots.
- **Primary and secondary queues**: noisy or throttled projects are moved off the main lane so they cannot starve others.
- **Fan-out queues** (`*-queue` → `*-processing-queue`): a scheduler job enqueues one job per project.
- **DLQ plus retry runner**, job failure logs, and queue-depth metrics.

### 6.5 Web application

- **Next.js** with both pages and app router; **tRPC** for the UI (`web/src/server/api/root.ts`, about 140 routers); hand-written REST handlers for the public API, with the spec in `fern/` used to generate SDKs.
- **NextAuth** (credentials, Google, GitHub, Okta, Azure AD, Keycloak, generic OIDC…), plus enterprise SSO and SCIM.
- UI filters compile through a **query AST** (`packages/shared/src/server/query-ast`) to ClickHouse SQL. That layer handles time-window enforcement, `FINAL`/`LIMIT 1 BY`, and query tagging (`clickHouseRequestTags`) for cost attribution.
- Feature areas (`web/src/features/*`, about 100 folders): traces, sessions, users, prompts, playground, datasets, experiments, evals, annotation queues, dashboards/widgets, models, monitors, automations, topics, in-app agent, MCP server, and more.

### 6.6 Rust components

- **`packages/native`** (napi-rs, runs inside the worker): ClickHouse Native protocol column encoding, OTel input and media parsing, and **UMAP + HDBSCAN** for the Topics feature, run in a child process so a long clustering job can be killed.
- **`ai-gateway`** (standalone): `/openai/v1/{responses,chat/completions,models}` and `/anthropic/v1/{messages,count_tokens,models}`. It passes provider bytes through unchanged (JSON or SSE), captures a bounded copy of each request and response, and uploads every completed call as a Langfuse generation over OTLP, batched per project. It authenticates with gateway keys that `web` resolves to stored provider connections using signed JWTs.

---

## 7. Design decisions and trade-offs

| Decision | Benefit | Cost |
|---|---|---|
| S3 as write-ahead log before the queue | Durable, replayable (an admin replay endpoint parses S3 keys), small Redis payloads, survives ClickHouse outages | An S3 PUT for every event, and LIST+GET per job (the reason `skipS3List` exists for OTel) |
| Merge-on-read from S3 fragments | Order-independent and idempotent create/update semantics | O(fragments) work per update, and hot entities get re-merged many times |
| `ReplacingMergeTree` + `FINAL` | Simple upserts in an append-only store | Deduplication at read time; duplicates across `toDate()` boundaries (hence the midnight delay) |
| v4 wide `events_full` / `events_core` | No joins, compact list scans, sampling, full-text search | Write amplification, a propagation job with ordering cursors, and a long dual-write migration (`v4-migration`, `v4LegacyApiUsageQueue`) |
| Maps for usage, cost and metadata | Any token or cost type and arbitrary metadata without migrations | Bloom-filter indexes on map keys and values only help so much at high cardinality |
| Postgres for config, ClickHouse for telemetry | Transactions and relations where needed, columnar scans where needed | Two migration systems; cross-store joins happen in application code |
| OTel as the main protocol | Works with any framework that emits OTel | About 2000 lines of attribute-dialect mapping to maintain |

---

## 8. Building something similar: minimal version

To reproduce the core loop on a small budget:

1. **API:** accept OTLP/HTTP plus a small JSON batch API. Auth by project key pair, with a SHA-256 hash of the secret stored and cached.
2. **WAL:** write each raw event to object storage keyed by `project/entityType/entityId/eventId.json`.
3. **Queue:** enqueue `{project, entityType, entityId}` with a few seconds of delay to coalesce updates (any Redis or SQS queue).
4. **Worker:** list and download fragments, merge with "later non-null wins, immutable keys locked", run model-regex price lookup and tokenization, then batch insert.
5. **Store:** one wide ClickHouse table, `ReplacingMergeTree(event_ts, is_deleted)`, `ORDER BY (project_id, toStartOfMinute(start_time), trace_id, span_id)`, with a truncated-I/O materialized view for list pages.
6. **Control plane:** Postgres for projects, keys, prompts (name/version/labels plus a Redis cache) and model prices.
7. **Evals:** a rule table plus a delayed job that templates trace fields into an LLM-judge prompt with structured output and writes a score back through ingestion.

That covers about 80% of Langfuse's core value. The rest of the codebase is hardening for scale (sharding, secondary queues, S3 throttling, propagation) and product surface area.

---

## Appendix: important files

| Concern | Path |
|---|---|
| Legacy ingestion endpoint | `web/src/pages/api/public/ingestion.ts` |
| OTLP endpoint | `web/src/pages/api/public/otel/v1/traces/index.ts`, `web/src/server/otel/processOtelIngestion.ts` |
| Batch → S3 → queue | `packages/shared/src/server/ingestion/processEventBatch.ts` |
| S3 key layout | `packages/shared/src/server/ingestion/eventBucketPath.ts` |
| OTel mapping | `packages/shared/src/server/otel/OtelIngestionProcessor.ts`, `ObservationTypeMapper.ts`, `attributes.ts` |
| Worker ingestion consumer | `worker/src/queues/ingestionQueue.ts` |
| Merge + enrich | `worker/src/services/IngestionService/index.ts` |
| Batched ClickHouse writer | `worker/src/services/ClickhouseWriter/index.ts` |
| ClickHouse schema | `packages/shared/clickhouse/migrations/canonical/*.sql` |
| Postgres schema | `packages/shared/prisma/schema.prisma` |
| Queue contracts | `packages/shared/src/server/queues.ts`, `packages/shared/src/server/redis/*` |
| API key auth | `web/src/features/public-api/server/apiAuth.ts` |
| Rate limits | `web/src/features/public-api/server/RateLimitService.ts` |
| Model/price matching | `packages/shared/src/server/ingestion/modelMatch.ts` |
| Prompt cache | `packages/shared/src/server/services/PromptService/index.ts` |
| Evals | `worker/src/features/evaluation/evalService.ts` |
| v4 propagation | `worker/src/features/eventPropagation/handleEventPropagationJob.ts` |
| Architecture principles | `.agents/ARCHITECTURE_PRINCIPLES.md` |
