

# Docker run

## Build docker image
`docker build . -t md2docx`
## run docker with output test.docx
`docker run -v $(pwd):/data -it md2docx test.docx`

