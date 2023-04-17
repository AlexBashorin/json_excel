FROM golang:alpine
RUN GOOS=linux CGO_ENABLED=0
RUN apk add git
RUN mkdir /app
ADD . /app
WORKDIR /app
RUN go build -o main .
CMD ["/app/main"]
