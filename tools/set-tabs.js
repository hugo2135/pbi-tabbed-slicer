#!/usr/bin/env node
/*
 * set-tabs.js - change the number of "tab field" data roles.
 *
 * Usage:
 *   node tools/set-tabs.js 12
 *   npm run set-tabs -- 12
 *   npm run set-tabs 12        (PowerShell may swallow the "--")
 *
 * Updates, in one shot:
 *   1. src/dataModel.ts   MAX_TABS
 *   2. capabilities.json  dataRoles            (tab1..tabN)
 *   3. capabilities.json  categories.select    (tab1..tabN)
 *   4. capabilities.json  objects.tabs         (name1..nameN, bg1..bgN)
 *
 * Default tab colours need no update: src/settings.ts spreads hues evenly
 * across MAX_TABS, so N tabs always get N distinct colours.
 *
 * This file is intentionally ASCII-only. Non-ASCII output has been observed
 * to vanish in some Windows consoles, which made failures look like silence.
 *
 * After running this, do `npm run package`, then in Power BI Desktop remove
 * the old custom visual and import the new .pbiviz again.
 */

"use strict";

const fs = require("fs");
const path = require("path");

const root = path.join(__dirname, "..");
const capPath = path.join(root, "capabilities.json");
const modelPath = path.join(root, "src", "dataModel.ts");
const logPath = path.join(__dirname, "set-tabs.log");

const lines = [];
let logError = null;

function flushLog() {
    try {
        fs.writeFileSync(logPath, lines.join("\n") + "\n", "utf8");
    } catch (err) {
        logError = err;
    }
}

function say(msg) {
    lines.push(String(msg));
    console.log(msg);
    flushLog();
}

function fail(msg) {
    lines.push(String(msg));
    console.error(msg);
    flushLog();
}

function main() {
    say("set-tabs: adjust the number of tab field wells");
    say("  node      : " + process.version + " on " + process.platform);
    say("  project   : " + root);
    say("  argv      : " + JSON.stringify(process.argv.slice(2)));
    say("  log file  : " + logPath);

    for (const f of [capPath, modelPath]) {
        if (!fs.existsSync(f)) {
            fail("");
            fail("FAILED: file not found: " + f);
            fail("  Run this from the project root; tools/ and src/ must exist.");
            return 1;
        }
    }

    const modelBefore = fs.readFileSync(modelPath, "utf8");
    const currentMatch = modelBefore.match(/MAX_TABS\s*=\s*(\d+)/);
    const current = currentMatch ? currentMatch[1] : null;
    say("  current   : MAX_TABS = " + (current === null ? "(not found)" : current));

    const n = parseInt(process.argv[2], 10);
    if (!Number.isInteger(n) || n < 1 || n > 50) {
        fail("");
        fail("FAILED: no valid number given (need an integer 1-50)");
        fail("  node tools/set-tabs.js 20");
        fail("  npm run set-tabs -- 20");
        fail("  npm run set-tabs 20        (if PowerShell eats the --)");
        return 1;
    }

    // ---- capabilities.json ------------------------------------------------
    const cap = JSON.parse(fs.readFileSync(capPath, "utf8"));

    const tabRole = i => ({
        displayName: "頁籤 " + i + " 欄位",
        description: "放入第 " + i
            + " 個頁籤的欄位；依序拖入多個欄位即形成"
            + "該頁籤的階層（第一個欄位為最上層，不限層數）。",
        name: "tab" + i,
        kind: "Grouping"
    });

    // keep non-tab roles (e.g. the optional measure) after the tab roles
    const otherRoles = cap.dataRoles.filter(r => !/^tab\d+$/.test(r.name));
    cap.dataRoles = [];
    for (let i = 1; i <= n; i++) {
        cap.dataRoles.push(tabRole(i));
    }
    cap.dataRoles.push.apply(cap.dataRoles, otherRoles);

    const categories = cap.dataViewMappings[0].categorical.categories;
    categories.select = [];
    for (let i = 1; i <= n; i++) {
        categories.select.push({ for: { in: "tab" + i } });
    }

    // keep general tab settings, rebuild nameN / bgN
    const tabProps = cap.objects.tabs.properties;
    const rebuilt = {};
    for (const key of Object.keys(tabProps)) {
        if (!/^(name|bg)\d+$/.test(key)) {
            rebuilt[key] = tabProps[key];
        }
    }
    for (let i = 1; i <= n; i++) {
        rebuilt["name" + i] = tabProps["name" + i] || {
            displayName: "頁籤 " + i + " 名稱",
            type: { text: true }
        };
    }
    for (let i = 1; i <= n; i++) {
        rebuilt["bg" + i] = tabProps["bg" + i] || {
            displayName: "頁籤 " + i + " 背景色",
            type: { fill: { solid: { color: true } } }
        };
    }
    cap.objects.tabs.properties = rebuilt;

    fs.writeFileSync(capPath, JSON.stringify(cap, null, 2) + "\n", "utf8");

    // ---- src/dataModel.ts -------------------------------------------------
    const MAX_TABS_RE = /export const MAX_TABS = \d+;/;
    // Test whether the pattern matches - do NOT test whether the text changed.
    // Asking for the value it already has is not an error.
    if (!MAX_TABS_RE.test(modelBefore)) {
        fail("");
        fail("FAILED: could not find `export const MAX_TABS = <number>;` in " + modelPath);
        return 1;
    }
    fs.writeFileSync(
        modelPath,
        modelBefore.replace(MAX_TABS_RE, "export const MAX_TABS = " + n + ";"),
        "utf8"
    );

    // ---- verify by reading everything back --------------------------------
    const vCap = JSON.parse(fs.readFileSync(capPath, "utf8"));
    const vModel = fs.readFileSync(modelPath, "utf8");
    const vMax = (vModel.match(/MAX_TABS\s*=\s*(\d+)/) || [])[1];
    const propNames = Object.keys(vCap.objects.tabs.properties);

    const checks = [
        ["src/dataModel.ts MAX_TABS", vMax],
        ["capabilities dataRoles", vCap.dataRoles.filter(r => /^tab\d+$/.test(r.name)).length],
        ["capabilities categories.select", vCap.dataViewMappings[0].categorical.categories.select.length],
        ["capabilities objects.tabs nameN", propNames.filter(k => /^name\d+$/.test(k)).length],
        ["capabilities objects.tabs bgN", propNames.filter(k => /^bg\d+$/.test(k)).length]
    ];

    say("");
    say("Set tab field wells to " + n + ":");
    let ok = true;
    for (const [label, got] of checks) {
        const pass = String(got) === String(n);
        ok = ok && pass;
        say("  [" + (pass ? "OK" : "!!") + "] " + label + " = " + got + (pass ? "" : " (expected " + n + ")"));
    }

    if (!ok) {
        fail("");
        fail("FAILED: something did not get written. Is a file read-only or locked?");
        return 1;
    }

    say("");
    say("Next:  npm run package");
    say("Then in Power BI Desktop: Visualizations pane -> remove the old custom");
    say("visual -> import the new .pbiviz from dist/ again.");
    say("(This script alone changes nothing inside Power BI.)");
    say("");
    say("RESULT: OK, tabs=" + n);
    return 0;
}

let code = 0;
try {
    code = main();
} catch (err) {
    code = 1;
    fail("");
    fail("FAILED with an unexpected error:");
    fail(String((err && err.stack) || err));
}

if (logError) {
    console.error("(could not write the log file: " + String(logError.message || logError) + ")");
}

process.exit(code);
