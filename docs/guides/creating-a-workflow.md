# Creating a New Workflow

Follow these steps every time you add a new workflow. The support ticket workflow (`app/workflows/support_ticket_workflow.py`) is the reference implementation.

---

## Step 1 — Define the Event Schema

Create `app/schemas/<name>_schema.py`:

```python
from typing import Optional
from pydantic import BaseModel, Field

class MyEventSchema(BaseModel):
    id: str
    content: str
    optional_field: Optional[str] = None
```

This schema validates every incoming event. FastAPI returns `422 Unprocessable Entity` automatically if required fields are missing.

---

## Step 2 — Create the Node Directory

```
app/workflows/<name>_workflow_nodes/
    __init__.py        ← empty
    first_node.py
    router_node.py     ← if you need branching
    second_node.py
```

---

## Step 3 — Write Each Node

### Plain Node (business logic, no LLM)
```python
from pydantic import BaseModel
from core.nodes.base import Node
from core.task import TaskContext

class MyOutput(BaseModel):
    result: str

class MyNode(Node):
    async def process(self, task_context: TaskContext) -> TaskContext:
        event = task_context.event
        self.save_output(MyOutput(result=event.content.upper()))
        return task_context
```

### AgentNode (LLM call)
```python
from pydantic import BaseModel
from core.nodes.agent import AgentNode, AgentConfig, ModelProvider
from core.task import TaskContext
from services.prompt_loader import PromptManager

class MyOutput(BaseModel):
    summary: str
    score: int

class MyAgentNode(AgentNode):
    def get_agent_config(self) -> AgentConfig:
        return AgentConfig(
            model_provider=ModelProvider.ANTHROPIC,
            model_name="claude-3-5-haiku-latest",
            output_type=MyOutput,
            system_prompt=PromptManager.get_prompt("my_system_prompt"),
        )

    async def process(self, task_context: TaskContext) -> TaskContext:
        result = await self.agent.run(task_context.event.content)
        self.save_output(result.output)
        return task_context
```

### Router Node (branching)
```python
from typing import Optional
from core.nodes.base import Node
from core.nodes.router import BaseRouter, RouterNode
from core.task import TaskContext
from workflows.my_workflow_nodes.first_node import MyNode, MyOutput
from workflows.my_workflow_nodes.path_a import PathANode
from workflows.my_workflow_nodes.path_b import PathBNode

class ScoreRoute(RouterNode):
    def determine_next_node(self, task_context: TaskContext) -> Optional[Node]:
        output: MyOutput = self.get_output(MyNode)
        if output and output.score > 50:
            return PathANode()
        return None

class MyRouter(BaseRouter):
    def __init__(self, task_context=None):
        super().__init__(task_context=task_context)
        self.routes = [ScoreRoute()]
        self.fallback = PathBNode()
```

---

## Step 4 — Define the Workflow

Create `app/workflows/<name>_workflow.py`:

```python
from core.schema import WorkflowSchema, NodeConfig
from core.workflow import Workflow
from schemas.my_schema import MyEventSchema
from workflows.my_workflow_nodes.first_node import MyNode
from workflows.my_workflow_nodes.router_node import MyRouter
from workflows.my_workflow_nodes.path_a import PathANode
from workflows.my_workflow_nodes.path_b import PathBNode

class MyWorkflow(Workflow):
    workflow_schema = WorkflowSchema(
        description="Does something useful.",
        event_schema=MyEventSchema,
        start=MyNode,
        nodes=[
            NodeConfig(node=MyNode, connections=[MyRouter]),
            NodeConfig(node=MyRouter, connections=[PathANode, PathBNode], is_router=True),
            NodeConfig(node=PathANode, connections=[]),
            NodeConfig(node=PathBNode, connections=[]),
        ],
    )
```

---

## Step 5 — Register the Workflow

`app/workflows/workflow_registry.py`:
```python
from workflows.my_workflow import MyWorkflow

class WorkflowRegistry(Enum):
    PLACEHOLDER = PlaceholderWorkflow
    SUPPORT_TICKET = SupportTicketWorkflow
    MY_WORKFLOW = MyWorkflow          # ← add this
```

---

## Step 6 — Add the Endpoint

Create `app/api/<name>_endpoint.py` (copy `support_endpoint.py` and swap the schema + registry name):

```python
from schemas.my_schema import MyEventSchema
from workflows.workflow_registry import WorkflowRegistry

@router.post("/")
def handle_my_event(data: MyEventSchema, session: Session = Depends(db_session)):
    ...
    event = Event(data=raw_event, workflow_type=WorkflowRegistry.MY_WORKFLOW.name)
    ...
```

Mount it in `app/api/router.py`:
```python
from api import my_endpoint
router.include_router(my_endpoint.router, prefix="/events/my-workflow", tags=["my-workflow"])
```

---

## Step 7 — Add Prompt Templates

For each `AgentNode`, create `app/prompts/<template_name>.j2`:

```
---
description: What this prompt does
author: your-name
---

You are a helpful assistant. {{ optional_variable }}
```

Load it:
```python
PromptManager.get_prompt("template_name", optional_variable="value")
```

---

## Step 8 — Create a Test Event

`requests/events/<name>_event.json`:
```json
{
  "id": "test-001",
  "content": "Hello world"
}
```

Send it:
```bash
make send-event   # or write a custom curl in the Makefile
```

---

## Step 9 — Write Tests

`tests/unit/test_<name>.py` — test deterministic nodes (no LLM calls needed):
```python
@pytest.mark.asyncio
async def test_my_node_output(my_payload):
    ctx = TaskContext(event=MyEventSchema(**my_payload))
    node = MyNode(task_context=ctx)
    await node.process(ctx)
    output = ctx.nodes["MyNode"]
    assert output.result == "EXPECTED VALUE"
```
