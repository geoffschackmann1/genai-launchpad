import sys
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).parent.parent / "app"))


@pytest.fixture
def support_ticket_payload():
    return {
        "ticket_id": "TKT-TEST-001",
        "subject": "Cannot log in",
        "body": "I cannot access my account.",
        "customer_name": "Test User",
        "customer_email": "test@example.com",
        "product": "Core Platform",
    }


@pytest.fixture
def low_urgency_payload():
    return {
        "ticket_id": "TKT-TEST-002",
        "subject": "How do I export data?",
        "body": "Hi, how do I export my data to CSV? Thanks!",
        "customer_name": "Mark Chen",
        "customer_email": "mark@example.com",
        "product": None,
    }
