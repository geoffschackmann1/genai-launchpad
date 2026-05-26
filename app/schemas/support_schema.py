from typing import Optional

from pydantic import BaseModel, Field


class SupportTicketEventSchema(BaseModel):
    ticket_id: str = Field(..., description="Unique identifier for the support ticket")
    subject: str = Field(..., description="Subject line of the support ticket")
    body: str = Field(..., description="Full body text of the support ticket")
    customer_name: str = Field(..., description="Name of the customer submitting the ticket")
    customer_email: str = Field(..., description="Email address of the customer")
    product: Optional[str] = Field(None, description="Product or service the ticket relates to")
