import json
from http import HTTPStatus

from fastapi import APIRouter, Depends
from sqlalchemy.orm import Session
from starlette.responses import Response

from database.event import Event
from database.repository import GenericRepository
from database.session import db_session
from schemas.support_schema import SupportTicketEventSchema
from worker.config import celery_app
from workflows.workflow_registry import WorkflowRegistry

router = APIRouter()


@router.post("/")
def handle_support_event(
    data: SupportTicketEventSchema,
    session: Session = Depends(db_session),
) -> Response:
    repository = GenericRepository(session=session, model=Event)
    raw_event = data.model_dump(mode="json")
    event = Event(data=raw_event, workflow_type=WorkflowRegistry.SUPPORT_TICKET.name)
    repository.create(obj=event)

    task_id = celery_app.send_task(
        "process_incoming_event",
        args=[str(event.id)],
    )

    return Response(
        content=json.dumps({"message": f"process_incoming_event started `{task_id}`"}),
        status_code=HTTPStatus.ACCEPTED,
    )
