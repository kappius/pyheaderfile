
.PHONY: setup test clean build

setup:
	@pip install -r requirements_local.txt
	@pip install -r requirements.txt

test:
	@coverage run --branch `which nosetests` --with-yanc --logging-clear-handlers -s tests/*
	@coverage report -m

clean:
	@find . -name "*.pyc" -exec rm -rf {} \;

build:
	@python setup.py sdist upload -r pypi
