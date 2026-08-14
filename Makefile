.PHONY: start stop logs test lint migrate send-event send-support-event

start:
	cd docker && ./start.sh

stop:
	cd docker && ./stop.sh

logs:
	cd docker && ./logs.sh

test:
	cd app && python -m pytest ../tests/ -v

lint:
	ruff check app/

migrate:
	cd app && ./migrate.sh

send-event:
	cd requests && python send_event.py

send-support-event:
	curl -s -X POST http://localhost:8080/events/support/ \
		-H "Content-Type: application/json" \
		-d @requests/events/support_ticket_event.json | python3 -m json.tool

send-support-low:
	curl -s -X POST http://localhost:8080/events/support/ \
		-H "Content-Type: application/json" \
		-d @requests/events/support_ticket_low_urgency.json | python3 -m json.tool
