const puppeteer = require("puppeteer");
const ProgressBar = require("progress");
const params = process.argv.slice(2);
const inputFileName = params[0];
const outputFileName = params[1];
const XLSX = require("xlsx");
const Excel = require("exceljs");
const readline = require("readline");
const { log } = require("console");
const URL = "https://mnemonicdictionary.com/word/";

main();

// const rl = readline.createInterface({
//   input: process.stdin,
//   output: process.stdout
// });

// rl.question("Wanna find comments from Word or words? ", function(url) {

//   if(url.toLowerCase() =="Word"){
//     URL = "https://mazii.net/#!/search?type=k&query="
//     console.log("start searching Word comment")
//   } else {
//     URL = "https://mazii.net/search?type=w&query="
//     console.log("start searching word comment")
//   }
//   rl.close();
//   main();
// });

function main() {
    console.log("======= START ======");
    const wb = XLSX.readFile(`${inputFileName}.xlsx`);
    const sheet_name_list = wb.SheetNames;
    const data = XLSX.utils.sheet_to_json(wb.Sheets[sheet_name_list[0]]);

    doCrawl(data);
}

async function doCrawl(data) {
    let tickLength = Math.ceil(data.length / 10);
    let Word = [];
    let count = 0;
    let bar = new ProgressBar("  Crawling [:bar] :percent :etas", {
        complete: "=",
        incomplete: " ",
        width: 20,
        total: tickLength,
    });

    while (data.length > 0) {
        Word = data.splice(0, 10).map((k) => k["Kanji"]);
        let num = count * 10;
        let eachRs = await getCommentsOfWord(URL, Word);
        await appendToExcel(eachRs, num);
        bar.tick();
        count++;
    }
    console.log("======= COMPLETED ======");
}

async function getCommentsOfWord(URL, listWord) {
    try {
        const browser = await puppeteer.launch();
        const pdfs = listWord.map(async (k, i) => {
            const page = await browser.newPage();
            await page.goto(`${URL}${k}`, {
                waitUntil: "networkidle2",
                timeout: 120000,
            });
            let content = await page.evaluate(async () => {
                let comment = document.querySelector(
                    ".slick-active p:not(.md-attribution)"
                );
                if (!comment) return "Không có comment";
                return comment.innerText.trim();
            });
            await page.close();
            return {
                Word: k,
                comment: content,
            };
        });

        return Promise.all(pdfs).then((result) => {
            browser.close();
            return result;
        });
    } catch (err) {
        console.log("err", err);
    }
}

function appendToExcel(each10Rs, num) {
    const workbook = new Excel.Workbook();
    return workbook.xlsx
        .readFile(`${outputFileName}.xlsx`)
        .then(function (data) {
            var worksheet = workbook.getWorksheet(1);
            each10Rs.forEach((rs, index) => {
                let row = worksheet.getRow(num + index + 1);
                row.getCell(1).value = rs.Word;
                row.getCell(2).value = rs.comment;
                row.commit();
            });
            return workbook.xlsx.writeFile(`${outputFileName}.xlsx`);
        });
}
