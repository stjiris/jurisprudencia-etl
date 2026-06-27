import { spawn } from "child_process";
import { client } from "./client";

// On-demand full-update consumer. Runs nightly (19:00) from cron and executes a
// `node dist --full` run only if the app scheduled one for today.
//
// NOTE: the index name and document shape are shared with the Next.js app
// (nextjs-jurisprudencia/src/core/etl-trigger.ts). Keep them in sync.
const ETL_TRIGGER_INDEX = "etl-trigger.0.0";
const REPORT_INDEX = "jurisprudencia-indexer-report.2.0";

// Hard cap on a single run. Protects against the DGSI stalls seen in history
// (e.g. the 2026-06-13 full run that dragged on for ~2.6 days).
const MAX_RUNTIME_MS = 12 * 60 * 60 * 1000;

async function findDueRequest() {
    const r = await client.search({
        index: ETL_TRIGGER_INDEX,
        size: 1,
        sort: [{ scheduledFor: "asc" }],
        query: {
            bool: {
                filter: [
                    { term: { status: "scheduled" } },
                    { range: { scheduledFor: { lte: "now" } } }
                ]
            }
        }
    }).catch(() => null);
    if (!r || r.hits.hits.length === 0) return null;
    return r.hits.hits[0];
}

// Stale-lock aware: only treats a "running" request as blocking if it started
// within MAX_RUNTIME (a crashed run leaves a stale "running" doc otherwise).
async function anotherRunInProgress(excludeId: string): Promise<boolean> {
    const since = new Date(Date.now() - MAX_RUNTIME_MS).toISOString();
    const r = await client.search({
        index: ETL_TRIGGER_INDEX,
        size: 1,
        query: {
            bool: {
                filter: [
                    { term: { status: "running" } },
                    { range: { startedAt: { gte: since } } }
                ],
                must_not: [{ ids: { values: [excludeId] } }]
            }
        }
    }).catch(() => null);
    return !!r && r.hits.hits.length > 0;
}

async function setStatus(id: string, doc: Record<string, unknown>) {
    await client.update({ index: ETL_TRIGGER_INDEX, id, doc, refresh: "true" });
}

function runFullUpdate(): Promise<{ ok: boolean; error?: string }> {
    return new Promise((resolve) => {
        // Equivalent to the cron's `node dist --full`; reuses the existing entry
        // point so the run writes its own jurisprudencia-indexer-report.2.0 doc.
        const child = spawn(process.execPath, ["dist", "--full"], {
            cwd: process.cwd(),
            stdio: "inherit"
        });

        const watchdog = setTimeout(() => {
            child.kill("SIGKILL");
            resolve({ ok: false, error: "stalled (watchdog)" });
        }, MAX_RUNTIME_MS);

        child.on("exit", (code) => {
            clearTimeout(watchdog);
            resolve(code === 0 ? { ok: true } : { ok: false, error: `process exited with code ${code}` });
        });
        child.on("error", (e) => {
            clearTimeout(watchdog);
            resolve({ ok: false, error: e.message });
        });
    });
}

async function latestReportSince(since: string) {
    const r = await client.search({
        index: REPORT_INDEX,
        size: 1,
        sort: [{ dateEnd: "desc" }],
        query: {
            bool: {
                filter: [
                    { term: { soft: false } },
                    { range: { dateEnd: { gte: since } } }
                ]
            }
        }
    }).catch(() => null);
    if (!r || r.hits.hits.length === 0) return null;
    return r.hits.hits[0];
}

async function main() {
    const hit = await findDueRequest();
    if (!hit) {
        console.log("[consume-trigger] No due request. Nothing to do.");
        return;
    }
    const id = hit._id!;

    if (await anotherRunInProgress(id)) {
        console.log(`[consume-trigger] Another run already in progress, skipping ${id}.`);
        await setStatus(id, { status: "skipped", error: "another run in progress" });
        return;
    }

    const startedAt = new Date().toISOString();
    console.log(`[consume-trigger] Starting full update for ${id} at ${startedAt}`);
    await setStatus(id, { status: "running", startedAt });

    const result = await runFullUpdate();
    const endedAt = new Date().toISOString();

    const report = await latestReportSince(startedAt);
    const counts = (report?._source as any) || {};

    await setStatus(id, {
        status: result.ok ? "success" : "failed",
        endedAt,
        error: result.error || null,
        reportId: report?._id || null,
        created: counts.created ?? null,
        updated: counts.updated ?? null,
        deleted: counts.deleted ?? null,
        skiped: counts.skiped ?? null
    });

    console.log(`[consume-trigger] Finished ${id}: ${result.ok ? "success" : "failed"} at ${endedAt}`);
}

main().catch(e => { console.error(e); process.exit(1); });
