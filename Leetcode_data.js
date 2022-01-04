// This code open chrome new instance with an existing chrome user profile located at /Users/Ishrat/Desktop/ChromeProfile
// Chrome profile must be created and passed as argument to the webdriver to use login information for educative.io
// There is no need to open chrome in debugger mode, hence, we wont need anything of the following sort
// //options.addArguments('debuggerAddress=127.0.0.1:9222');

//var endLessons = require("./lessonsData1");

const TIME_OUT = 20000;
var PAGE_NO = 0;
var TOTAL_PAGES = 7; // Microsoft, the total number of pages on leetcode are 14
const { Options } = require("selenium-webdriver/chrome");

const Excel = require("exceljs");
// Create a new instance of a Workbook class
const workbook = new Excel.Workbook();

var worksheet;

var linkStyle;
workbook.xlsx.readFile("export.xlsx").then(() => {
  worksheet = workbook.getWorksheet("G_M_F_A_A");

  linkStyle = {
    underline: true,
    color: { argb: "FF0000FF" },
  };

  worksheet.columns = [
    { header: "Title", key: "title", width: 40 },
    { header: "Acceptance Rate", key: "rate", width: 32 },
    { header: "Difficulty Level", key: "level", width: 15 },
    { header: "Frequence of occurrence", key: "freq", width: 15 },
  ];
});

const options = new Options();
options.addArguments("--user-data-dir=/Users/Ishrat/Desktop/ChromeProfile2");

const webdriver = require("selenium-webdriver"),
  By = webdriver.By,
  until = webdriver.until;

const driver = new webdriver.Builder()
  .setChromeOptions(options)
  .forBrowser("chrome")
  .build();

const quit_chrome = () => {
  //console.log("Chrome closed.");
  driver.quit();
};

function quit() {
  //console.log("Quitting chrome...");
  setTimeout(quit_chrome, TIME_OUT);
}

function ExtractQuestionData(questions_rows) {
  let questions = [];

  questions_rows.map(async (question_row, index) => {
    // find title
    let question_title_anchor = await question_row.findElement(
      By.xpath("./div[2]/div/div/div/div/a")
      //By.xpath("//a")
    );
    let title = await question_title_anchor.getText();
    let url = await question_title_anchor.getAttribute("href");
    console.log("Question Title is : " + title);

    // find acceptance
    let acceptance_span = await question_row.findElement(
      By.xpath("./div[4]/span")
    );
    let acceptance = await acceptance_span.getText();
    //console.log("Question Title is : " + acceptance);

    // difficulty level
    let difficulty_level_span = await question_row.findElement(
      By.xpath("./div[5]/span")
    );
    let difficulty = await difficulty_level_span.getText();
    //console.log("Difficulty level is: " + difficulty);

    // frequency
    let frequence_div = await question_row.findElement(
      By.xpath("./div[6]/div/div/div[2]")
    );
    let frequency = await frequence_div.getCssValue("width");
    console.log("Frequency is: " + frequency);

    let question = {
      question_title: title,
      question_url: url,

      acceptance_rate: acceptance,
      difficulty_level: difficulty,
      frequency_rate: frequency,
    };

    questions.push(question);
    writeToExcel(index + (PAGE_NO - 1) * 50, question);
  });
  console.log(questions.length);
}

const writeToExcel = async (question_no, question) => {
  console.log(
    question_no +
      " " +
      "Problem row is : " +
      question.question_title +
      " " +
      question.acceptance_rate +
      " " +
      question.difficulty_level +
      " " +
      question.frequency_rate
  );

  let row = worksheet.addRow({
    title: {
      text: question.question_title,
      hyperlink: question.question_url,
    },
    //title: question.question_title,
    rate: question.acceptance_rate,
    level: question.difficulty_level,
    freq: question.frequency_rate.split("px")[0],
  });
  row.getCell(1).font = linkStyle;

  // save under export.xlsx
  await workbook.xlsx.writeFile("export.xlsx");
};

const fetchQuestions = () => {
  driver
    .wait(
      until.elementsLocated(
        /*By.xpath(
          '//div[@class="odd:bg-overlay-3 dark:odd:bg-dark-overlay-1 even:bg-overlay-1 dark:even:bg-dark-overlay-3"]'
        )*/
        By.xpath('//div[@role="row"]')
      ),
      TIME_OUT,
      "Timed out after 30 seconds",
      20000
    )
    .then((questions_rows) => {
      console.log(questions_rows.length);
      ExtractQuestionData(questions_rows);

      setTimeout(getCompanyPage, 10000);
    })
    .catch((err) => {
      console.log("Questions could not be fetched: " + err);
      quit();
    });
};

const getCompanyPage = async () => {
  PAGE_NO++;
  if (PAGE_NO <= TOTAL_PAGES)
    var BASE_URL =
      "https://leetcode.com/problemset/all/?companySlugs=microsoft%2Cgoogle%2Cfacebook%2Camazon%2Capple&page=" +
      PAGE_NO;
  await driver.get(BASE_URL);

  setTimeout(fetchQuestions, 5000);
};

getCompanyPage();
