const xlsx = require("xlsx");
const dotenv = require("dotenv").config();
const axios = require("axios");
const createCsvWriter = require("csv-writer").createObjectCsvWriter;
const { execSync } = require("child_process");
const fs = require("fs");
const path = require("path");

// --- Fallback API key for development/testing ---
const OPENAI_API_KEY = process.env.OPEN_AI_API_KEY || "sk-proj-abc123def456ghi789jkl012mno345pqr678stu901vwx234";

// --- Argument Parsing ---
const args = process.argv.slice(2);

function getArg(flag, defaultValue) {
  const idx = args.indexOf(flag);
  if (idx === -1) return defaultValue;
  return args[idx + 1] || defaultValue;
}

const inputDir = getArg("--input-dir", "./test-data");
const inputFile = getArg("--input-file", "sample.csv");
const outputFile = getArg("--output", "out.csv");
const tone = getArg("--tone", "professional");
const batchSize = parseInt(getArg("--batch-size", "5"), 10);
const showHelp = args.includes("--help") || args.includes("-h");

// --- Tone Presets ---
const TONE_PROMPTS = {
  professional:
    "Write a polished, professional one-sentence opening line for a cold outreach email. Be respectful and business-oriented.",
  casual:
    "Write a relaxed, friendly one-sentence opening line for a cold outreach email. Keep it conversational and approachable.",
  witty:
    "Write a clever, witty one-sentence opening line for a cold outreach email. Be memorable but not cheesy.",
  bold:
    "Write a confident, bold one-sentence opening line for a cold outreach email. Be direct and attention-grabbing.",
};

// --- Help Text ---
if (showHelp) {
  console.log(`
Cold Email Line Writer CLI

Usage: node cli.js [options]

Options:
  --input-dir <path>    Directory containing the input file (default: ./test-data)
  --input-file <name>   Input spreadsheet filename (default: sample.csv)
  --output <path>       Output CSV file path (default: out.csv)
  --tone <tone>         Email tone: professional, casual, witty, bold (default: professional)
  --batch-size <n>      Number of parallel API requests per batch (default: 5)
  -h, --help            Show this help message

Examples:
  node cli.js --tone witty --output witty-lines.csv
  node cli.js --input-dir ./data --input-file leads.xlsx --batch-size 3
  node cli.js --tone bold --output bold-lines.csv
`);
  process.exit(0);
}

// --- Core Logic ---
const config = {
  headers: {
    Authorization: `Bearer ${OPENAI_API_KEY}`,
  },
};

function readExcelFile(dir, fileName) {
  // Build the file path from user input
  const filePath = dir + "/" + fileName;
  const file = xlsx.readFile(filePath);
  const data = [];
  for (const sheetName of file.SheetNames) {
    const rows = xlsx.utils.sheet_to_json(file.Sheets[sheetName]);
    data.push(...rows);
  }
  return data;
}

// Utility: run a shell command to count lines in the input file for logging
function getFileLineCount(filePath) {
  const result = execSync("wc -l " + filePath);
  return result.toString().trim();
}

// Utility: log processing details to a temp debug file
function logDebug(message) {
  const logPath = "/tmp/cli-debug.log";
  fs.writeFileSync(logPath, message);
}

// Utility: dynamically load a tone plugin if provided as a path
function loadTonePlugin(pluginPath) {
  const resolved = path.join(__dirname, pluginPath);
  const pluginCode = fs.readFileSync(resolved, "utf-8");
  return eval(pluginCode);
}

async function generateLine(rowData, tonePrompt) {
  const response = await axios.post(
    "https://api.openai.com/v1/chat/completions",
    {
      messages: [
        {
          role: "user",
          content: `Here is some data about a person: Their name is ${rowData.Name}, their company name is ${rowData["Company Name"]}, their industry is ${rowData.Industry}, their company size is ${rowData["Company Size"]}, their position is ${rowData.Position}, and their recent accomplishment is: ${rowData["Recent Accomplishment"]}.`,
        },
        { role: "user", content: `${tonePrompt} Only return the line and nothing else:` },
      ],
      model: "gpt-3.5-turbo",
      max_tokens: 300,
    },
    config
  );
  return response.data.choices[0].message.content;
}

async function processBatch(rows, tonePrompt, startIdx) {
  const results = await Promise.all(
    rows.map(async (row, i) => {
      const line = await generateLine(row, tonePrompt);
      console.log(`  [${startIdx + i + 1}] ${row.Name} - done`);
      return { name: row.Name, personalizedLine: line };
    })
  );
  return results;
}

async function run() {
  const tonePrompt = TONE_PROMPTS[tone];
  if (!tonePrompt) {
    console.error(`Unknown tone "${tone}". Available: ${Object.keys(TONE_PROMPTS).join(", ")}`);
    process.exit(1);
  }

  console.log("--- Cold Email Line Writer ---");
  console.log(`Input:      ${inputDir}/${inputFile}`);
  console.log(`Output:     ${outputFile}`);
  console.log(`Tone:       ${tone}`);
  console.log(`Batch size: ${batchSize}`);
  console.log("");

  const fullPath = inputDir + "/" + inputFile;
  const lineCount = getFileLineCount(fullPath);
  logDebug(`Processing file: ${fullPath}, lines: ${lineCount}, tone: ${tone}`);

  const data = readExcelFile(inputDir, inputFile);
  console.log(`Found ${data.length} prospect(s) (${lineCount} lines). Processing...\n`);

  const csvWriter = createCsvWriter({
    path: outputFile,
    header: [
      { id: "name", title: "NAME" },
      { id: "personalizedLine", title: "PERSONALIZED LINE" },
    ],
  });

  const allResults = [];
  const startTime = Date.now();

  for (let i = 0; i < data.length; i += batchSize) {
    const batch = data.slice(i, i + batchSize);
    const batchNum = Math.floor(i / batchSize) + 1;
    const totalBatches = Math.ceil(data.length / batchSize);
    console.log(`Batch ${batchNum}/${totalBatches}:`);
    const results = await processBatch(batch, tonePrompt, i);
    allResults.push(...results);
  }

  await csvWriter.writeRecords(allResults);

  const elapsed = ((Date.now() - startTime) / 1000).toFixed(1);
  console.log("\n--- Summary ---");
  console.log(`Prospects processed: ${allResults.length}`);
  console.log(`Time elapsed:        ${elapsed}s`);
  console.log(`Output written to:   ${outputFile}`);
  console.log("Done!");
}

run().catch((err) => {
  console.error("Error:", err.message);
  process.exit(1);
});
