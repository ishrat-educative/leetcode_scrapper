const TIME_OUT = 20000;
var PAGE_NO = 0;
const { Options } = require("selenium-webdriver/chrome");

const Excel = require("exceljs");
// Create a new instance of a Workbook class
const workbook = new Excel.Workbook();

var worksheet;

var linkStyle;
workbook.xlsx.readFile("export.xlsx").then(() => {
  worksheet = workbook.addWorksheet("Educative1");

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

getExceptionLessonTitles = () => {
  // Get lesson titles that have problem statements yet DIY is not appended to them

  let linkStyle = {
    underline: true,
    color: { argb: "FF0000FF" },
  };
  for (let i = 21; i < 23; i++) {
    driver
      .findElements(
        By.xpath(
          "/html/body/div[1]/div[2]/div[2]/div/div[2]/div[2]/div/ul/div[" +
            i +
            "]/div/ul/li"
        )
      )
      .then((lessons) => {
        lessons.forEach(async (lesson) => {
          let anchor = await lesson.findElement(
            By.xpath("./div/div/div[1]/div[2]/span/a")
          );
          let text = await anchor.getText();
          console.log("Lesson is : " + text);
          let row = worksheet.addRow({
            title: text,
          });
          row.getCell(1).font = linkStyle;
          workbook.xlsx.writeFile("export.xlsx");
          console.log("written to file");
        });
      })
      .catch();
  }

  //quit();
};

getlessonTitle = () => {
  driver
    .wait(
      until.elementsLocated(
        By.xpath(
          "//div[@class='styles__ArticleTitle-sc-1ttnunj-6 kmIYWG']//span[@class='overflow-ellipsis overflow-hidden whitespace-nowrap max-w-sm text-base']//a"
        )
      ),
      TIME_OUT,
      "Timed out after 30 seconds",
      1000
    )
    .then((lessonlinks) => {
      let count = 0;
      lessonlinks.forEach(async (lessonlink, index) => {
        lesson_title = await lessonlink.getText();
        //console.log("DIY length " + lesson_title.indexOf("DIY"));
        //console.log(index + ". Lesson title:" + lesson_title);

        if (lesson_title.indexOf("DIY") == 0) {
          //console.log(index + ". Lesson title:" + lesson_title);
          console.log(
            count++ + " DIY SUBSTRING is " + lesson_title.split("DIY: ")[1]
          );
          let row = worksheet.addRow({
            title: lesson_title.split("DIY: ")[1],
          });
          row.getCell(1).font = linkStyle;
          workbook.xlsx.writeFile("export.xlsx");
        }
      });
    })
    .catch((err) => {
      console.log("Lessons could not be loaded");
    });
  //quit();
};

const loadModule = async () => {
  await driver.get(
    "https://www.educative.io/collectioneditor/10370001/6112523134173184"
  );

  setTimeout(getExceptionLessonTitles, 5000, driver, By, worksheet, quit);

  setTimeout(getlessonTitle, 15000, driver, By, worksheet, quit);
};

loadModule();
