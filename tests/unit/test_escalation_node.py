import pytest

from core.task import TaskContext
from schemas.support_schema import SupportTicketEventSchema
from workflows.support_ticket_workflow_nodes.classifier_node import (
    ClassificationOutput,
    ClassifierNode,
    TicketCategory,
    TicketSentiment,
    TicketUrgency,
)
from workflows.support_ticket_workflow_nodes.escalation_node import (
    EscalationNode,
    EscalationOutput,
)


def make_task_context(payload: dict, urgency: TicketUrgency) -> TaskContext:
    event = SupportTicketEventSchema(**payload)
    ctx = TaskContext(event=event)
    ctx.nodes[ClassifierNode.__name__] = ClassificationOutput(
        category=TicketCategory.ACCOUNT,
        urgency=urgency,
        sentiment=TicketSentiment.FRUSTRATED,
        summary="Customer cannot access account.",
        reasoning="Password reset not working, paying customer blocked.",
    )
    return ctx


@pytest.mark.asyncio
async def test_critical_escalates_to_tier1(support_ticket_payload):
    ctx = make_task_context(support_ticket_payload, TicketUrgency.CRITICAL)
    node = EscalationNode(task_context=ctx)
    await node.process(ctx)

    output: EscalationOutput = ctx.nodes[EscalationNode.__name__]
    assert output.escalated is True
    assert output.escalation_tier == "tier-1-oncall"
    assert output.priority_label == "P0 - Immediate Response"


@pytest.mark.asyncio
async def test_high_escalates_to_tier2(support_ticket_payload):
    ctx = make_task_context(support_ticket_payload, TicketUrgency.HIGH)
    node = EscalationNode(task_context=ctx)
    await node.process(ctx)

    output: EscalationOutput = ctx.nodes[EscalationNode.__name__]
    assert output.escalated is True
    assert output.escalation_tier == "tier-2-senior"
    assert output.priority_label == "P1 - Urgent"


@pytest.mark.asyncio
async def test_escalation_reason_contains_summary(support_ticket_payload):
    ctx = make_task_context(support_ticket_payload, TicketUrgency.HIGH)
    node = EscalationNode(task_context=ctx)
    await node.process(ctx)

    output: EscalationOutput = ctx.nodes[EscalationNode.__name__]
    assert "Customer cannot access account." in output.escalation_reason


@pytest.mark.asyncio
async def test_escalated_at_is_set(support_ticket_payload):
    ctx = make_task_context(support_ticket_payload, TicketUrgency.CRITICAL)
    node = EscalationNode(task_context=ctx)
    await node.process(ctx)

    output: EscalationOutput = ctx.nodes[EscalationNode.__name__]
    assert output.escalated_at is not None
    assert "T" in output.escalated_at  # ISO format check
