import path from "path";
import { fileURLToPath } from "url";
import fs from "fs-extra";
import https from "https";
import archiver from "archiver";
import esbuild from "esbuild";

// ---- paths
const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const ROOT = path.resolve(__dirname, "..");
const BUILD = path.join(ROOT, "static-resources-build");

// ---- utils
async function ensureDir(p) { await fs.ensureDir(p); }

function download(url, dest) {
    return new Promise(async (resolve, reject) => {
        await ensureDir(path.dirname(dest));
        const file = fs.createWriteStream(dest);
        const onError = (err) => { try { file.close(); } catch { } reject(err); };
        https.get(url, (res) => {
            // follow a single redirect if present
            if (res.statusCode >= 300 && res.statusCode < 400 && res.headers.location) {
                https.get(res.headers.location, (res2) => pipe(res2));
            } else {
                pipe(res);
            }
        }).on("error", onError);

        function pipe(res) {
            if (res.statusCode !== 200) {
                onError(new Error(`GET ${url} -> ${res.statusCode}`));
                return;
            }
            res.pipe(file);
            file.on("finish", () => file.close(resolve));
            file.on("error", onError);
        }
    });
}

async function zipDir(inputDir, outZipPath) {
    await ensureDir(path.dirname(outZipPath));
    await fs.access(inputDir); // throws if missing
    const output = fs.createWriteStream(outZipPath);
    const archive = archiver("zip", { zlib: { level: 9 } });
    return new Promise((resolve, reject) => {
        output.on("close", resolve);
        archive.on("error", reject);
        archive.pipe(output);
        archive.directory(inputDir, false);
        archive.finalize();
    });
}

async function buildCopilot() {
    // Bundle your wrapper to a single IIFE that attaches to window for Locker
    const inFile = path.join(ROOT, "build-src/copilotstudio-global.js");
    const outDir = path.join(BUILD, "copilotStudioClient", "dist");
    const outFile = path.join(outDir, "copilotStudioClient.js");

    await ensureDir(outDir);
    await esbuild.build({
        entryPoints: [inFile],
        outfile: outFile,
        format: "iife",
        platform: "browser",
        target: "es2019",
        bundle: true,
        minify: true,
        legalComments: "none",
        // Ensure globalThis + window are available for UMD-ish libs under Locker
        banner: { js: "var globalThis = globalThis || window;" }
    });

    // Zip as Salesforce static resource
    await zipDir(outDir, path.join(BUILD, "copilotStudioClient.resource"));

    // Metadata
    await fs.writeFile(path.join(BUILD, "copilotStudioClient.resource-meta.xml"), STATIC_META, "utf8");
    console.log("🟢 [build] Copilot Studio client ready");
}

async function buildMSAL() {
    const outDir = path.join(BUILD, "msalbrowser", "dist");
    const outFile = path.join(outDir, "msal-browser.min.js");
    await ensureDir(outDir);

    // Pin MSAL version for deterministic builds
    const MSAL_URL = "https://alcdn.msauth.net/browser/2.37.0/js/msal-browser.min.js";
    await download(MSAL_URL, outFile);

    await zipDir(outDir, path.join(BUILD, "msalbrowser.resource"));
    await fs.writeFile(path.join(BUILD, "msalbrowser.resource-meta.xml"), STATIC_META, "utf8");
    console.log("🟢 [build] MSAL Browser ready");
}

async function buildAdaptive() {
    // Copy from node_modules (already minified in the package dist)
    const src = path.join(ROOT, "node_modules", "adaptivecards", "dist", "adaptivecards.js");
    const outDir = path.join(BUILD, "adaptiveCards", "dist");
    const outFile = path.join(outDir, "adaptivecards.js");

    await ensureDir(outDir);
    await fs.copyFile(src, outFile);

    await zipDir(outDir, path.join(BUILD, "adaptiveCards.resource"));
    await fs.writeFile(path.join(BUILD, "adaptiveCards.resource-meta.xml"), STATIC_META, "utf8");
    console.log("🟢 [build] AdaptiveCards ready");
}

const STATIC_META = `<?xml version="1.0" encoding="UTF-8"?>
<StaticResource xmlns="http://soap.sforce.com/2006/04/metadata">
  <cacheControl>Public</cacheControl>
  <contentType>application/zip</contentType>
</StaticResource>`;

async function main() {
    await ensureDir(BUILD);
    const which = (process.argv[2] || "").toLowerCase();

    if (!which) {
        await buildCopilot();
        await buildMSAL();
        await buildAdaptive();
        console.log("✅ [build] All static resources complete");
        return;
    }

    if (which === "copilot") return buildCopilot();
    if (which === "msal") return buildMSAL();
    if (which === "adaptive") return buildAdaptive();

    console.error(`⚠️ [build] Unknown target "${which}"`);
    process.exit(2);
}

main().catch((e) => {
    console.error("❌ [build] FAILED:", e);
    process.exit(1);
});