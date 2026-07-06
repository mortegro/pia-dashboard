import { serveDir } from "jsr:@std/http/file-server";

const staticRoot = import.meta.dirname;
const tarifDir = `${staticRoot}/pia-verträge`;

const TARIF_FILENAME_RE = /^PIA-Leistungen-Preise-(\d{8})-(\d{8})\.csv$/;

function toIsoDate(yyyymmdd: string): string {
  return `${yyyymmdd.slice(0, 4)}-${yyyymmdd.slice(4, 6)}-${yyyymmdd.slice(6, 8)}`;
}

async function listTarifPeriods() {
  const periods = [];
  for await (const entry of Deno.readDir(tarifDir)) {
    if (!entry.isFile) continue;
    const m = entry.name.match(TARIF_FILENAME_RE);
    if (!m) continue;
    const start = toIsoDate(m[1]);
    const end = toIsoDate(m[2]);
    const csvText = await Deno.readTextFile(`${tarifDir}/${entry.name}`);
    periods.push({
      start,
      end,
      filename: entry.name,
      label: `PIA-Leistungen-Preise ${start} – ${end}`,
      csvText,
    });
  }
  periods.sort((a, b) => a.start.localeCompare(b.start));
  return periods;
}

Deno.serve(async (req) => {
  const url = new URL(req.url);
  if (url.pathname === "/api/tarif-periods") {
    return Response.json(await listTarifPeriods());
  }
  return serveDir(req, { fsRoot: staticRoot, quiet: true });
});
