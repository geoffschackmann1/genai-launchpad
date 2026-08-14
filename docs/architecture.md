# Architecture Overview

## System Design

GenAI Launchpad is an event-driven AI workflow platform. The design separates event ingestion (fast, synchronous) from event processing (slow, asynchronous) to ensure the API stays responsive even when LLM calls take several seconds.

```
Client
  │
  ▼
FastAPI (app/api/)
  │  Validates event schema
  │  Persists Event to Postgres
  │  Queues Celery task
  │  Returns 202 Accepted
  │
  ▼
Redis (task broker)
  │
  ▼
Celery Worker (app/worker/)
  │  Fetches Event from Postgres
  │  Looks up Workflow from registry
  │  Runs Workflow
  │  Stores TaskContext back to Postgres
  │
  ▼
Workflow (app/workflows/)
  │  Chains Nodes in sequence
  │  Handles routing via BaseRouter
  │  Returns final TaskContext
```

## Components

### FastAPI (`app/api/`)
One endpoint file per workflow. Each endpoint validates the incoming payload against its Pydantic schema, writes an `Event` row, and immediately returns `202 Accepted` with the Celery task ID.

### Celery Worker (`app/worker/`)
Single task: `process_incoming_event(event_id)`. Fetches the event, instantiates the right workflow class via `WorkflowRegistry`, runs it, and writes results back.

### Workflow Engine (`app/core/workflow.py`)
Executes a `WorkflowSchema` — a directed graph of `NodeConfig` entries. Traverses from `start` node, following `connections` after each node completes. Supports:
- **Linear chains**: `NodeA → NodeB → NodeC`
- **Routing**: a `BaseRouter` node dynamically picks the next node
- **Parallel execution**: `ConcurrentNode` runs `concurrent_nodes` via `asyncio.gather`

### Node Types (`app/core/nodes/`)
| Class | Use |
|-------|-----|
| `Node` | Pure Python logic — no LLM |
| `AgentNode` | LLM call via Pydantic-AI; structured output via Pydantic model |
| `BaseRouter` | Reads task context; returns next Node instance |
| `ConcurrentNode` | Spawns multiple nodes in parallel |

### TaskContext (`app/core/task.py`)
Shared state passed through every node. Key fields:
- `event` — the validated input event
- `nodes` — dict of node outputs, keyed by class name
- `should_stop` — set to `True` to halt the workflow early

### Database (`app/database/`)
Single `events` table with flexible JSON columns:
- `data` — the raw incoming event
- `task_context` — the final state after workflow completes
- `workflow_type` — the `WorkflowRegistry` enum name used for routing

### Prompt Templates (`app/prompts/`)
Jinja2 templates with YAML frontmatter. Loaded via `PromptManager.get_prompt("template_name", **vars)`. Two patterns:
- **Static system prompt**: no variables, baked in at `get_agent_config()` time
- **Dynamic user turn**: per-request variables injected at `process()` time

## Registered Workflows

| Name | Endpoint | Nodes |
|------|----------|-------|
| `PLACEHOLDER` | `POST /events/` | `InitialNode` |
| `SUPPORT_TICKET` | `POST /events/support/` | `ClassifierNode → UrgencyRouter → EscalationNode \| ResponderNode` |

## Data Flow Example (Support Ticket)

1. Client POSTs ticket JSON to `/events/support/`
2. FastAPI validates against `SupportTicketEventSchema`, stores row, queues task
3. Worker runs `SupportTicketWorkflow`
4. `ClassifierNode` (claude-3-5-haiku) classifies: category, urgency, sentiment, summary
5. `UrgencyRouter` checks urgency:
   - `high` / `critical` → `EscalationNode` (deterministic, no LLM)
   - `low` / `medium` → `ResponderNode` (claude-3-5-sonnet, drafts reply)
6. Final `TaskContext` with all node outputs stored to `events.task_context`
