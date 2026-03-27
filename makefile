# Run the streamlit app
run:
	python -m streamlit run app.python

# Install dependencies
install:
	pip install -r requirements.txt

# Git add + commit + push
#push:
#	git add . 
#	git commit -m "Update"
#	git push

# Format code
format:
	black .

# Clean temp files
clean:
	find . -type f -name "*.pyc" -delete
	find . -type d -name "__pycache__" -delete
	
test:
	pytest

lint:
	flake8 .