FROM denoland/deno:alpine

WORKDIR /app

COPY deno.json deno.lock ./
COPY main.ts ./
COPY index.html ./
COPY css ./css
COPY js ./js
COPY vendor ./vendor
COPY pia-verträge ./pia-verträge

RUN deno cache main.ts

EXPOSE 8000

CMD ["run", "--allow-net", "--allow-read", "main.ts"]
