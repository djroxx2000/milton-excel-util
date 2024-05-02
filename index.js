const xlsxFileInput = document.getElementById("xlsx-upload");
const csvFileInput = document.getElementById("csv-upload");
const eventLog = document.querySelector(".update-event");
const keyOrder = [
  "utm_source",
  "utm_campaign",
  "utm_medium",
  "utm_term",
  "utm_content",
  "gclid",
  "fbclid",
  "checkout_token",
  "cart_token",
  "PG Transaction Id",
  "Breeze Order Id",
  "CustomerGSTIN",
  "Company Name",
];
const iframeStyles = `
<style>
  tr:nth-child(even) {background-color: #f2f2f2;}                

  table,
  td {
    font-family: Georgia, serif;
    border-bottom: 1px solid #ddd;
    border-collapse: collapse;
    text-align: center;
    padding: 15px;
    margin: 10px;
  }

  tbody tr:first-child {
    font-weight: bold;
    background-color: #777;
  }

  tr:hover {background-color: #9999;}
</style>
`;

const reader = new FileReader();
const sheetToCsvConfig = {
  raw: false,
  defval: null,
  RS: "\r\n",
  FS: "\t\t",
};
const sheetToJsonConfig = { raw: false };
let uploadedFile;

/**
 * @description Converts csv string to an array of arrays
 * @param {string} csvData
 * @param {string} lineSeparator
 * @param {string} dataSeparator
 */
function convertCsvToAOA(csvData, lineSeparator, dataSeparator) {
  try {
    const rows = csvData.split(lineSeparator);
    return rows.map((row) => {
      return row.split(dataSeparator).map((data) => data.trim());
    });
  } catch (err) {
    eventLog.textContent =
      "Unable to convert CSV Data to required format. Please check file delimiters. Full Error: " +
      err;
  }
}

/**
 * Safely parses each row in csv when data itself might contain lineSeparator. Requires header row to be present where line separator will not be used otherwise fails.
 * @param {string} csvData
 * @param {string} lineSeparator
 * @param {string} dataSeparator
 * @param {string} groupLimit when data blocks are grouped together since the data itself contains line or data separator characters
 * @returns
 */
function safeConvertCsvToAOA(
  csvData,
  lineSeparator,
  dataSeparator,
  groupLimit
) {
  try {
    const firstNewLine = csvData.indexOf(lineSeparator);
    const firstDataBlock = csvData.indexOf(dataSeparator);
    if (firstNewLine === -1 || firstDataBlock === -1) {
      eventLog.textContent = `Unable to safely convert CSV Data to required format. Please check file contains a valid header row and that line delimiter is ${lineSeparator}`;
      return;
    }
    console.log(csvData.length);
    // const headerRow = csvData.slice(0, firstNewLine).split(dataSeparator);
    const aoa = [[]];
    let isGroupedData = false;
    let rowNum = 0;
    let dataBuffer = "";
    for (let i = 0; i < csvData.length; i++) {
      const curCharacter = csvData[i];
      switch (csvData[i]) {
        case lineSeparator:
          if (isGroupedData) {
            dataBuffer += curCharacter;
          } else {
            aoa[rowNum].push(dataBuffer);
            dataBuffer = "";
            aoa.push([]);
            rowNum++;
          }
          break;
        case dataSeparator:
          if (isGroupedData) {
            dataBuffer += curCharacter;
          } else {
            aoa[rowNum].push(dataBuffer);
            dataBuffer = "";
          }
          break;
        case groupLimit:
          if (csvData[i - 1] !== "\\") {
            isGroupedData = !isGroupedData;
          }
          dataBuffer += curCharacter;
          break;
        default:
          dataBuffer += curCharacter;
      }
    }
    console.log("AOA:", aoa);
    return aoa;
  } catch (err) {
    eventLog.textContent =
      "Unable to safely convert CSV Data to required format. Please check file contains a valid header row and that line delimiter is \\n. Full Error: " +
      err;
  }
}

/**
 *
 * @param {blob} fileData csv or xlsx file data received in file input
 * @returns {XLSX.WorkBook} XLSX workbook
 */
function createWorkbook(fileData) {
  try {
    if (uploadedFile.name.endsWith("xlsx")) {
      return XLSX.read(fileData, { type: "binary" }, { dateNF: "dd/mm/yyyy" });
    } else {
      const sheetAOA = safeConvertCsvToAOA(fileData.toString(), "\n", ",", '"');
      const sheet = XLSX.utils.aoa_to_sheet(sheetAOA);
      const workbook = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(workbook, sheet, "Orders");
      return workbook;
    }
  } catch (err) {
    eventLog.textContent =
      "Unable to convert given file data to required format. Please check file format. Full Error: " +
      err;
  }
}

reader.addEventListener("error", (err) => {
  eventLog.textContent =
    "Unable to parse file data to apply changes. Full Error: " + err;
});

reader.addEventListener("loadend", (e) => {
  console.log("DEBUGGING::LOADEND", e);
  const fileData = e.target.result;
  const workbook = createWorkbook(fileData);
  const orderSheet = workbook.Sheets["Orders"];
  const sheetData = XLSX.utils.sheet_to_json(orderSheet, sheetToJsonConfig);
  const sheetCsv = XLSX.utils.sheet_to_csv(orderSheet, sheetToCsvConfig);
  const sheetAOA = convertCsvToAOA(sheetCsv, "\r\n", "\t\t");
  // console.log("DEBUGGING::SHEETDATA", sheetData, sheetCsv, sheetAOA);
  sheetAOA[0].push(...keyOrder);
  for (let idx = 0; idx < sheetData.length; idx++) {
    const row = sheetData[idx];
    let attributes = row["Note Attributes"] !== undefined ? row["Note Attributes"] : row["Additional Details"];
    // console.log("ATTRIBUTES", attributes, Object.keys(row), typeof attributes);
    if (attributes !== null && attributes !== undefined) {
      if(attributes[0] === '"') {
        attributes = attributes.slice(1, attributes.length - 1);
      }
      const details = attributes.split("\n");
      for (const detail of details) {
        const [key, value] = detail.split(":").map((val) => val.trim());
        row[key] = value;
      }
    } else {
      console.log("IDX:", idx, sheetData[idx]);
      eventLog.textContent = "Unable to find column 'Note Attributes' or 'Additional Details' in given file.";
      return;
    }
    keyOrder.forEach((key) => {
      sheetAOA[idx + 1].push(row[key] ? row[key] : "NA");
    });
  }
  // console.log("DEBUGGING::AVAILABLE", sheetAOA);
  const updatedXLSX = XLSX.utils.aoa_to_sheet(sheetAOA);
  const updatedWorkbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(updatedWorkbook, updatedXLSX, "Orders");
  // const table = XLSX.utils.sheet_to_html(updatedXLSX);
  // const iframe = document.getElementById("table-iframe");
  // const iframeDoc = iframe.contentWindow.document;
  // iframeDoc.open();
  // iframeDoc.write(table);
  // iframeDoc.write(iframeStyles);
  // iframeDoc.close();
  XLSX.writeFile(
    updatedWorkbook,
    uploadedFile.name.split(".")[0] + "-updated.xlsx"
  );
  eventLog.textContent = "Completed parsing and updating file. Ready for next file..."
});

/**
 *
 * @param {Event} _event Event object for file input change event
 * @param {string} fileType csv/xlsx
 */
function handleSelected(_event, fileType) {
  eventLog.textContent = `Parsing and updating ${fileType} file...`;
  if (fileType === "xlsx" && xlsxFileInput.files[0]) {
    uploadedFile = xlsxFileInput.files[0];
    reader.readAsBinaryString(uploadedFile);
  } else if (fileType === "csv" && csvFileInput.files[0]) {
    uploadedFile = csvFileInput.files[0];
    reader.readAsText(uploadedFile);
  } else {
    eventLog.textContent = `Invalid file format received...`;
  }
}

xlsxFileInput.addEventListener("change", (e) => handleSelected(e, "xlsx"));
csvFileInput.addEventListener("change", (e) => handleSelected(e, "csv"));
