#!/bin/bash

python3 md2docx.py test.docx && soffice --headless --convert-to pdf test.docx
