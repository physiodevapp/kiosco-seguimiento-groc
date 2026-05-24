.PHONY: push-personal push-prod pull deploy open logs

push-personal:
	cp .clasp.personal.json .clasp.json
	clasp push

push-prod:
	cp .clasp.prod.json .clasp.json
	clasp push
	cp .clasp.personal.json .clasp.json

pull:
	cp .clasp.personal.json .clasp.json
	clasp pull

deploy:
	cp .clasp.prod.json .clasp.json
	clasp push
	clasp deploy --deploymentId $(ID) --description "$(DESC)"
	cp .clasp.personal.json .clasp.json

open:
	cp .clasp.personal.json .clasp.json
	clasp open

logs:
	cp .clasp.personal.json .clasp.json
	clasp logs
