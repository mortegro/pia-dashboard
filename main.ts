import { serveDir } from "jsr:@std/http/file-server";

const staticRoot = new URL(".", import.meta.url).pathname;

Deno.serve((req) =>
  serveDir(req, {
    fsRoot: staticRoot,
    quiet: true,
  })
);
