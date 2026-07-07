from core.schema import NodeConfig, WorkflowSchema
from core.workflow import Workflow
from schemas.support_schema import SupportTicketEventSchema
from workflows.support_ticket_workflow_nodes.classifier_node import ClassifierNode
from workflows.support_ticket_workflow_nodes.escalation_node import EscalationNode
from workflows.support_ticket_workflow_nodes.responder_node import ResponderNode
from workflows.support_ticket_workflow_nodes.urgency_router import UrgencyRouter


class SupportTicketWorkflow(Workflow):
    workflow_schema = WorkflowSchema(
        description="Classifies incoming support tickets and routes them to auto-responder or escalation.",
        event_schema=SupportTicketEventSchema,
        start=ClassifierNode,
        nodes=[
            NodeConfig(
                node=ClassifierNode,
                connections=[UrgencyRouter],
                description="AI classifies ticket by category, urgency, and sentiment",
            ),
            NodeConfig(
                node=UrgencyRouter,
                connections=[EscalationNode, ResponderNode],
                is_router=True,
                description="Routes critical/high urgency to escalation, others to responder",
            ),
            NodeConfig(
                node=EscalationNode,
                connections=[],
                description="Marks ticket as escalated with tier and reason",
            ),
            NodeConfig(
                node=ResponderNode,
                connections=[],
                description="Drafts a professional support response",
            ),
        ],
    )
