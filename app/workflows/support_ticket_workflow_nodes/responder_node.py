from pydantic import BaseModel, Field

from core.nodes.agent import AgentConfig, AgentNode, ModelProvider
from core.task import TaskContext
from services.prompt_loader import PromptManager
from workflows.support_ticket_workflow_nodes.classifier_node import ClassifierNode


class ResponderOutput(BaseModel):
    response_subject: str = Field(..., description="Email subject for the response")
    response_body: str = Field(..., description="Full customer-facing response email body")
    internal_notes: str = Field(..., description="Internal notes for the support team")


class ResponderNode(AgentNode):
    class OutputType(ResponderOutput):
        pass

    def get_agent_config(self) -> AgentConfig:
        return AgentConfig(
            model_provider=ModelProvider.ANTHROPIC,
            model_name="claude-3-5-sonnet-latest",
            output_type=ResponderOutput,
            system_prompt=PromptManager.get_prompt("support_responder"),
        )

    async def process(self, task_context: TaskContext) -> TaskContext:
        event = task_context.event
        classification = self.get_output(ClassifierNode)

        user_message = PromptManager.get_prompt(
            "support_responder_user",
            customer_name=event.customer_name,
            category=classification.category.value,
            urgency=classification.urgency.value,
            sentiment=classification.sentiment.value,
            summary=classification.summary,
            ticket_subject=event.subject,
            ticket_body=event.body,
        )

        result = await self.agent.run(user_message)
        self.save_output(result.output)
        return task_context
