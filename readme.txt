docker build --no-cache -t docx_generator .
docker run -p 11312:11312 docx_generator 