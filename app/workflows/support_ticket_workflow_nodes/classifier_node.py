from enum import Enum

from pydantic import BaseModel, Field

from core.nodes.agent import AgentConfig, AgentNode, ModelProvider
from core.task import TaskContext
from services.prompt_loader import PromptManager


class TicketCategory(str, Enum):
    BILLING = "billing"
    TECHNICAL = "technical"
    ACCOUNT = "account"
    FEATURE_REQUEST = "feature_request"
    OTHER = "other"


class TicketUrgency(str, Enum):
    LOW = "low"
    MEDIUM = "medium"
    HIGH = "high"
    CRITICAL = "critical"


class TicketSentiment(str, Enum):
    FRUSTRATED = "frustrated"
    NEUTRAL = "neutral"
    POSITIVE = "positive"


class ClassificationOutput(BaseModel):
    category: TicketCategory = Field(..., description="Category of the support ticket")
    urgency: TicketUrgency = Field(..., description="Urgency level of the ticket")
    sentiment: TicketSentiment = Field(..., description="Customer sentiment")
    summary: str = Field(..., description="One-sentence summary of the ticket issue")
    reasoning: str = Field(..., description="Brief reasoning for the classification")


class ClassifierNode(AgentNode):
    class OutputType(ClassificationOutput):
        pass

    def get_agent_config(self) -> AgentConfig:
        return AgentConfig(
            model_provider=ModelProvider.ANTHROPIC,
            model_name="claude-3-5-haiku-latest",
            output_type=ClassificationOutput,
            system_prompt=PromptManager.get_prompt("support_classifier"),
        )

    async def process(self, task_context: TaskContext) -> TaskContext:
        event = task_context.event
        user_message = (
            f"Customer: {event.customer_name} ({event.customer_email})\n"
            f"Product: {event.product or 'Not specified'}\n"
            f"Subject: {event.subject}\n\n"
            f"{event.body}"
        )
        result = await self.agent.run(user_message)
        self.save_output(result.output)
        return task_context
