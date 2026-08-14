import pytest
from pydantic import ValidationError

from schemas.support_schema import SupportTicketEventSchema


def test_valid_ticket(support_ticket_payload):
    ticket = SupportTicketEventSchema(**support_ticket_payload)
    assert ticket.ticket_id == "TKT-TEST-001"
    assert ticket.customer_email == "test@example.com"
    assert ticket.product == "Core Platform"


def test_product_is_optional(support_ticket_payload):
    support_ticket_payload["product"] = None
    ticket = SupportTicketEventSchema(**support_ticket_payload)
    assert ticket.product is None


def test_missing_required_field(support_ticket_payload):
    del support_ticket_payload["subject"]
    with pytest.raises(ValidationError):
        SupportTicketEventSchema(**support_ticket_payload)


def test_missing_customer_email(support_ticket_payload):
    del support_ticket_payload["customer_email"]
    with pytest.raises(ValidationError):
        SupportTicketEventSchema(**support_ticket_payload)
