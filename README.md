# Introduction

`md2docx` is a tool developed for internal Syndis use for generating docx files from security notes written in markdown. For details on syntax, see `template.md`.

# How to use

## Clone the repository

`git clone git@github.com:syndisgudni/md2docx.git`

## Setup virtual environment (recommended)

`python -m venv .`

`source .venv/bin/activate`

## Install dependencies

`pip install -r requirements.txt`

## Convert the test file

`python md2docx.py --file test.md output.docx`

## Copy template, insert your findings and do the thing!

`python md2docx.py --file copy_of_template.md 'SYN-XXX-YY-ZZ - Security Note'`

# Docker run

## Build docker image

`docker build . -t md2docx`

## run docker with output test.docx

`docker run -v $(pwd):/data -it md2docx test.docx`

