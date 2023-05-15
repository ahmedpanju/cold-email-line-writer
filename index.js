const xlsx = require("xlsx");
const dotenv = require("dotenv").config();
const axios = require("axios");
const createCsvWriter = require("csv-writer").createObjectCsvWriter;

const csvWriter = createCsvWriter({
  path: "out.csv",
  header: [
    { id: "name", title: "NAME" },
    { id: "personalizedLine", title: "PERSONALIZED LINE" },
  ],
});

const config = {
  headers: {
    Authorization: `Bearer ${process.env.OPEN_AI_API_KEY}`,
  },
};

const readExcelFile = async (path, filePath) => {
  try {
    const file = xlsx.readFile(`${path}/${filePath}`);
    let data = [];
    const sheets = file.SheetNames;
    for (let i = 0; i < sheets.length; i++) {
      const temp = xlsx.utils.sheet_to_json(file.Sheets[file.SheetNames[i]]);
      temp.forEach((res) => {
        data.push(res);
      });
    }
    return data;
  } catch (err) {
    console.log(err);
  }
};

const createFinalCSV = async (PATH_DOWNLOADED_FILE, sourceFileName) => {
  const spreadsheetData = await readExcelFile(
    PATH_DOWNLOADED_FILE,
    sourceFileName
  );

  const finalDataArray = [];

  await Promise.all(
    spreadsheetData.map(async (rowData) => {
      const customLineResponse = await axios.post(
        `https://api.openai.com/v1/chat/completions`,
        {
          messages: [
            {
              role: "user",
              content: `Here is some data about a person: Their name is ${rowData.Name}, their company name is ${rowData["Company Name"]}, their industry is ${rowData.Industry}, their company size is ${rowData["Company Size"]}, their position is ${rowData.Position}, and their recent accomplishment is: ${rowData["Recent Accomplishment"]}.`,
            },
            {
              role: "user",
              content:
                "Your job is to return a 1 sentence, personalized line that I can include in a cold outreach email to this person. Only return the line and nothing else:",
            },
          ],
          model: "gpt-3.5-turbo",
          max_tokens: 300,
        },
        config
      );

      finalDataArray.push({
        name: rowData.Name,
        personalizedLine: customLineResponse.data.choices[0].message.content,
      });
    })
  );

  csvWriter.writeRecords(finalDataArray).then(() => {
    console.log("...Done");
  });
};

createFinalCSV("./test-data", "sample.csv");
