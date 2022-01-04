const Excel = require("exceljs");

// Create a new instance of a Workbook class
const workbook = new Excel.Workbook();

var worksheet, worksheet_educative, worksheet_leetcode;

var linkStyle_red, linkStyle_green;
workbook.xlsx
  .readFile("export.xlsx")
  .then(() => {
    console.log("ssssqwert");
    worksheet = workbook.addWorksheet("Analysis_final");
    worksheet_educative = workbook.getWorksheet("Educative1");
    worksheet_leetcode = workbook.getWorksheet("G_F_M_A_A");

    worksheet.state = "visible";
    console.log("sssswqs");

    linkStyle_red = {
      underline: true,
      color: { argb: "FFF0000" },
    };

    linkStyle_blue = {
      underline: true,
      color: { argb: "0000FF" },
    };

    worksheet.columns = [
      { header: "Available at Educative?", key: "available", width: "20" },
      { header: "Title", key: "title", width: 40 },
      { header: "Acceptance Rate", key: "rate", width: 32 },
      { header: "Difficulty Level", key: "level", width: 15 },
      { header: "Frequence of occurrence", key: "freq", width: 15 },
    ];
    compareData();
    /*worksheet_leetcode.eachRow((leetcode_row, rowNumber) => {
      test();
    });*/
  })
  .catch((err) => {
    console.log("error : " + err);
  });

const writeToExcelSheet = async (row_toinsert, found) => {
  //console.log("in test");
  let available_str = "";
  if (found) available_str = "yes";
  else available_str = "no";

  let row = worksheet.addRow({
    available: available_str,
    title: {
      text: row_toinsert.getCell(1).text,
      hyperlink: row_toinsert.getCell(1).hyperlink,
    },
    //title: question.question_title,
    rate: row_toinsert.getCell(2).value,
    level: row_toinsert.getCell(3).value,
    freq: row_toinsert.getCell(4).value,
  });
  //row.getCell(1).font = found ? linkStyle_red : linkStyle_green;

  row.getCell(1).font = linkStyle_blue;

  if (found)
    row.fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "FFF0000" },
    };

  // save under export.xlsx
  await workbook.xlsx.writeFile("export.xlsx");
};

const compareData = () => {
  var educative_rows = [];

  worksheet_educative.eachRow(function (row, rowNumber) {
    console.log("Row " + rowNumber + " = " + JSON.stringify(row.values));
    educative_rows.push(row.getCell(1).text.toString().toLowerCase());
  });
  let count = 0;
  worksheet_leetcode.eachRow((leetcode_row, rowNumber) => {
    let leetcode_cellvalue = leetcode_row.getCell(1).text.toLowerCase();
    //console.log("Finding: " + leetcode_cellvalue);

    if (
      educative_rows.find((rowValue) => {
        return rowValue.includes(leetcode_cellvalue);
      })
    ) {
      console.log(count++ + "Found");
      // question exists in both the sheets
      writeToExcelSheet(leetcode_row, true);
    } else {
      console.log("not found");
      // problem does not exist on educative platform
      //new_row.getCell(1).font = linkStyle_green;
      writeToExcelSheet(leetcode_row, false);
    }
  });
  // });
};

//compareData();
