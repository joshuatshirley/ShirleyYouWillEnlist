#!/usr/bin/env node
/**
 * Merges question-bank/{wk,ar,pc,mk}_bank.json into question-data.js's
 * window.__EST_STANDALONE__.testVariants.
 *
 * Output order: tests 1..count are built from each bank's "tests" array
 * in order (keys "bank1".."bankN"), and the hand-curated original set
 * (key "original") is kept as-is and always placed LAST. It was Test 1
 * until 2026-11; it moved to the end because it circulates widely and
 * many applicants already know its answers. This script never rewrites
 * its content.
 *
 * Each bank's questions carry their own "n" numbering already (see the
 * schema each generation agent was given); this script trusts that and
 * doesn't renumber.
 *
 * Usage: node tools/build-test-variants.js [count]
 *   count: how many additional variants to add from the banks (default 4,
 *          i.e. tests 1-4). Banks may contain more than this many test
 *          arrays (spares); only the first `count` are used from each.
 *
 * After running: re-check question-data.js's syntax (e.g. `node -e
 * "require('./question-data.js')"` with a stubbed `window`), bump
 * sw.js's CACHE_NAME (this changes a precached file), and review a
 * sample of the new content before committing.
 */

var fs = require("fs");
var path = require("path");

var ROOT = path.join(__dirname, "..");
var DATA_PATH = path.join(ROOT, "question-data.js");
var BANK_DIR = path.join(ROOT, "question-bank");

var COUNT = parseInt(process.argv[2], 10) || 4;

// Each section's fixed metadata (title/minutes/instructions/code) is
// identical across all variants by design -- only the questions differ.
// Pulled from the original test so it never has to be retyped/duplicated here.
var SECTION_ORDER = ["WK", "AR", "PC", "MK"];
var BANK_FILES = { WK: "wk_bank.json", AR: "ar_bank.json", PC: "pc_bank.json", MK: "mk_bank.json" };

function readJson(p) {
  return JSON.parse(fs.readFileSync(p, "utf8"));
}

function main() {
  var src = fs.readFileSync(DATA_PATH, "utf8");
  var m = src.match(/window\.__EST_STANDALONE__ = (\{[\s\S]*?\n\};)/);
  if (!m) throw new Error("Could not find window.__EST_STANDALONE__ blob in " + DATA_PATH);

  var data;
  eval("data = " + m[1].replace(/;\s*$/, ""));

  var original = data.testVariants.filter(function (v) { return v.key === "original"; })[0];
  if (!original) throw new Error("No variant with key \"original\" in " + DATA_PATH);
  var metaByCode = {};
  original.sections.forEach(function (s) { metaByCode[s.code] = s; });
  SECTION_ORDER.forEach(function (code) {
    if (!metaByCode[code]) throw new Error("The original test is missing a section with code " + code);
  });

  var banks = {};
  SECTION_ORDER.forEach(function (code) {
    var p = path.join(BANK_DIR, BANK_FILES[code]);
    var bank = readJson(p);
    if (!Array.isArray(bank.tests) || bank.tests.length < COUNT) {
      throw new Error(BANK_FILES[code] + " has only " + (bank.tests || []).length + " test arrays, need " + COUNT);
    }
    banks[code] = bank.tests;
  });

  var newVariants = [];
  for (var i = 0; i < COUNT; i++) {
    var sections = SECTION_ORDER.map(function (code) {
      var meta = metaByCode[code];
      var questions = banks[code][i];
      if (!Array.isArray(questions) || !questions.length) {
        throw new Error("Bank " + code + " test index " + i + " is empty");
      }
      return {
        code: meta.code,
        title: meta.title,
        minutes: meta.minutes,
        instructions: meta.instructions,
        questions: questions
      };
    });
    newVariants.push({ key: "bank" + (i + 1), sections: sections });
  }
  newVariants.push(original);

  data.testVariants = newVariants;

  var newBlob = "window.__EST_STANDALONE__ = {\n" +
    "  testVariants: " + JSON.stringify(data.testVariants) + ",\n" +
    "  scoring: " + JSON.stringify(data.scoring) + "\n" +
    "};";

  var newSrc = src.slice(0, m.index) + newBlob + src.slice(m.index + m[0].length);
  fs.writeFileSync(DATA_PATH, newSrc, "utf8");

  console.log("Wrote " + newVariants.length + " test variants (" + COUNT + " generated + 1 original, last) to " + DATA_PATH);
  newVariants.forEach(function (v, idx) {
    var counts = v.sections.map(function (s) { return s.code + ":" + s.questions.length; }).join(" ");
    console.log("  Test " + (idx + 1) + " (" + v.key + ") -- " + counts);
  });
}

main();
