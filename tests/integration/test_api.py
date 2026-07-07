"""
Integration tests — require the full Docker stack to be running.

Run with:
    make start
    make test
"""

import pytest
import requests

BASE_URL = "http://localhost:8080"


@pytest.mark.integration
def test_support_endpoint_accepts_valid_ticket(support_ticket_payload):
    response = requests.post(
        f"{BASE_URL}/events/support/",
        json=support_ticket_payload,
        timeout=10,
    )
    assert response.status_code == 202
    assert "process_incoming_event started" in response.json()["message"]


@pytest.mark.integration
def test_support_endpoint_rejects_invalid_ticket():
    response = requests.post(
        f"{BASE_URL}/events/support/",
        json={"ticket_id": "TKT-001"},  # missing required fields
        timeout=10,
    )
    assert response.status_code == 422


@pytest.mark.integration
def test_placeholder_endpoint_still_works():
    response = requests.post(
        f"{BASE_URL}/events/",
        json={"id": "test-123", "type": "PLACEHOLDER"},
        timeout=10,
    )
    assert response.status_code == 202
