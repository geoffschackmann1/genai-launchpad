from datetime import datetime, timezone

from pydantic import BaseModel, Field

from core.nodes.base import Node
from core.task import TaskContext
from workflows.support_ticket_workflow_nodes.classifier_node import ClassifierNode, TicketUrgency


class EscalationOutput(BaseModel):
    escalated: bool = Field(default=True)
    escalation_tier: str = Field(..., description="Support tier to escalate to")
    escalation_reason: str = Field(..., description="Why this ticket was escalated")
    escalated_at: str = Field(..., description="ISO timestamp of escalation")
    priority_label: str = Field(..., description="Human-readable priority label for the ticket queue")


class EscalationNode(Node):
    async def process(self, task_context: TaskContext) -> TaskContext:
        event = task_context.event
        classification = self.get_output(ClassifierNode)

        urgency = classification.urgency
        tier = "tier-1-oncall" if urgency == TicketUrgency.CRITICAL else "tier-2-senior"
        priority_label = "P0 - Immediate Response" if urgency == TicketUrgency.CRITICAL else "P1 - Urgent"

        output = EscalationOutput(
            escalated=True,
            escalation_tier=tier,
            escalation_reason=(
                f"Ticket '{event.subject}' classified as {urgency.value} urgency "
                f"({classification.category.value}) with {classification.sentiment.value} "
                f"customer sentiment. Summary: {classification.summary}"
            ),
            escalated_at=datetime.now(timezone.utc).isoformat(),
            priority_label=priority_label,
        )
        self.save_output(output)
        return task_context
