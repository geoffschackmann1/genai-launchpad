# GenAI Launchpad

Event-driven AI workflow platform: FastAPI ingests events, Celery processes them asynchronously through chains of AI nodes, results stored to PostgreSQL.

## Common Commands

```bash
make start        # start full Docker stack (API + worker + Redis + Postgres)
make stop         # stop stack
make logs         # tail logs
make test         # run pytest
make lint         # ruff linter
make migrate      # run Alembic migrations
make send-event   # send the placeholder test event
```

## Architecture

```
POST /events/<workflow>/
  → FastAPI: validate schema → persist Event → queue Celery task → 202
  → Worker: fetch event → WorkflowRegistry lookup → run Workflow → persist TaskContext
  → Workflow: Node chain (sequential) + BaseRouter (branching) + ConcurrentNode (parallel)
```

## Key Files

| File | Purpose |
|------|---------|
| `app/api/router.py` | Mount all endpoints here |
| `app/workflows/workflow_registry.py` | Register all workflows here |
| `app/core/nodes/agent.py` | AgentNode base — LLM integration |
| `app/core/nodes/router.py` | BaseRouter + RouterNode |
| `app/core/nodes/base.py` | Node base — save_output / get_output |
| `app/core/workflow.py` | Workflow execution engine |
| `app/services/prompt_loader.py` | PromptManager.get_prompt() |
| `docker/.env` | All secrets and env vars go here |

## Adding a Workflow — Checklist

- [ ] `app/schemas/<name>_schema.py` — event schema
- [ ] `app/workflows/<name>_workflow_nodes/` — node files
- [ ] `app/workflows/<name>_workflow.py` — WorkflowSchema wiring
- [ ] `app/workflows/workflow_registry.py` — add to enum
- [ ] `app/api/<name>_endpoint.py` — POST endpoint
- [ ] `app/api/router.py` — mount endpoint
- [ ] `app/prompts/<name>.j2` — prompt templates
- [ ] `requests/events/<name>_event.json` — test payload
- [ ] `tests/unit/workflows/test_<name>.py` — unit tests

## Existing Workflows

| Workflow | Endpoint | Description |
|----------|----------|-------------|
| `PLACEHOLDER` | `POST /events/` | Minimal pass-through example |
| `SUPPORT_TICKET` | `POST /events/support/` | Classifies tickets → routes to responder or escalation |

## Environment Variables

Copy `docker/.env.example` → `docker/.env` and fill in:
- `ANTHROPIC_API_KEY` — required for support ticket workflow
- `DATABASE_*` — Postgres connection (defaults work with Docker stack)
- `PROJECT_NAME` — used for Redis hostname (default: `launchpad`)

## State Between Nodes

```python
self.save_output(MyOutput(...))      # saves to task_context.nodes["MyClassName"]
self.get_output(OtherNode)           # retrieves other node's output by class
task_context.stop_workflow()         # halt execution after current node
```
