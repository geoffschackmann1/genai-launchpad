from fastapi import APIRouter

from api import endpoint
from api import support_endpoint

"""
API Router Module

This module sets up the API router and includes all defined endpoints.
It uses FastAPI's APIRouter to group related endpoints and provide a prefix.
"""

router = APIRouter()

router.include_router(endpoint.router, prefix="/events", tags=["events"])
router.include_router(support_endpoint.router, prefix="/events/support", tags=["support"])
