from typing import Optional

from core.nodes.base import Node
from core.nodes.router import BaseRouter, RouterNode
from core.task import TaskContext
from workflows.support_ticket_workflow_nodes.classifier_node import (
    ClassifierNode,
    TicketUrgency,
)
from workflows.support_ticket_workflow_nodes.escalation_node import EscalationNode
from workflows.support_ticket_workflow_nodes.responder_node import ResponderNode


class HighUrgencyRoute(RouterNode):
    """Routes critical and high urgency tickets to EscalationNode."""

    def determine_next_node(self, task_context: TaskContext) -> Optional[Node]:
        classification = self.get_output(ClassifierNode)
        if classification is None:
            return None
        if classification.urgency in (TicketUrgency.CRITICAL, TicketUrgency.HIGH):
            return EscalationNode()
        return None


class UrgencyRouter(BaseRouter):
    def __init__(self, task_context: TaskContext = None):
        super().__init__(task_context=task_context)
        self.routes = [HighUrgencyRoute()]
        self.fallback = ResponderNode()
