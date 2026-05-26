from enum import Enum

from workflows.placeholder_workflow import PlaceholderWorkflow
from workflows.support_ticket_workflow import SupportTicketWorkflow


class WorkflowRegistry(Enum):
    PLACEHOLDER = PlaceholderWorkflow
    SUPPORT_TICKET = SupportTicketWorkflow
