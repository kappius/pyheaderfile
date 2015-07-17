
.PHONY: setup test clean build

setup:
	@pip install -r requirements_local.txt
	@pip install -r requirements.txt

test:
	@py.test --doctest-modules pyheaderfile

clean:
	@find . -name "*.pyc" -exec rm -rf {} \;

build:
	@python setup.py sdist upload -r pypi

