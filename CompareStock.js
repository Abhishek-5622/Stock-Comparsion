// node CompareStock.js "ICICI BANK LTD." "HDFC BANK LTD" "GOODYEAR INDIA LTD." "SAT INDUSTRIES LTD." "STATE BANK OF INDIA"
// AIM : Comparsion of Stock Details of different companies. 

//Require modules
let puppeteer = require("puppeteer");
let fs = require('fs');
let path = require('path');
let xlsx = require("xlsx");
var nodemailer = require('nodemailer');

//Link of the page 
let link = "https://www.bseindia.com/";

//Get name of the stock which we want to compare 
let StockName = process.argv.slice(2);

//Get email and password from seperate file(secrete.js)
let {password , email} = require("../activity/secrete.js");


(async function () {
    try {
        //launch browser
        let browserInstance = await puppeteer.launch({
            headless: false,
            defaultViewport: null,
            args: ["--start-maximized"]
        });
        //Create new tab
        let newTab = await browserInstance.newPage();
        await newTab.goto(link);
        
        //Create empty array
        let newarr=[];

        //Function call that gives us details of stocks.
        let stockA = await getDetailsOfStock(newTab,newarr,StockName[0]);

        for(let i=1;i<StockName.length;i++)
        {
            stockA = await getDetailsOfStock(newTab,stockA,StockName[i]);
        }

        //Sort array on the basic of Change in stock value
        stockA.sort((a,b) => b.Change - a.Change);
        
       //Function call to make json file
        makefile();
        //Path of json file
        let filePath = path.join(__dirname, "Stock" + ".json");
        
        //Write stock details in json file.
        fs.writeFileSync(filePath, JSON.stringify(stockA));

        // file stock details in excel 
        excelWriter("Stock.xlsx",stockA,"First");

        //Function call to send mail
        SendMailByJS();

        await browserInstance.close(); 
    //handle error
    } catch (err) {
        console.log(err);
    }
})();

//Function that store stock details in an array
async function getDetailsOfStock(newTab,newarr,StockName)
{
        //Wait for selector to display
        await newTab.waitForSelector("#SearchQuotediv", { visible: true });

        //Click 
        await newTab.click("#SearchQuotediv");

        //Type stockName in search box
        await newTab.type("#SearchQuotediv", StockName, { delay: 500 });

        //Press Enter
        await newTab.keyboard.press("Enter", { delay: 200 });
        await newTab.keyboard.press("Enter");

        //Wait for selector
        await newTab.waitForSelector(".stockreach_title.ng-binding", { visible: true });

        //Browser console function
        function consoleFn(newarr) {

            //Extract details of stock
            let Name=document.querySelector(".stockreach_title.ng-binding").innerText;
            let UpdatedValue =document.querySelector("#idcrval").innerText;
            let Change =document.querySelector(".sensexbluetext.ng-binding").innerText;
            let ChangeValue = Change.split("(");
            let ChangePercent = Change.split(" ");
            let PreviousClose = document.querySelectorAll(".textvalue.ng-binding");
            let Code = document.querySelector('div[style="color: #333;font-size:13px; padding-top:20px;"]').innerText;
            let DateArr =document.querySelector(".col-lg-12.ng-binding").innerText;
            let Date = DateArr.split("|");

            //push details in array
            newarr.push(
                {
                    NameOfStock :Name,
                    Date:Date[0],
                    Code:Code,
                    UpdatedValue:UpdatedValue,
                    Change:ChangeValue[0],
                    ChangePercent : ChangePercent[1],
                    PreviousClose:PreviousClose[0].innerText,
                    Open : PreviousClose[1].innerText,
                    High : PreviousClose[2].innerText,
                    Low : PreviousClose[3].innerText,
                    UpperPriceBand : PreviousClose[7].innerText,
                    LowerPriceBand :PreviousClose[8].innerText
                });
            return newarr;
        }
    
    return await newTab.evaluate(consoleFn,newarr);
 
}

//Function that use to make json file
function makefile() {
    let pathoffile = path.join(__dirname, "Stock" + ".json")
    var createStream = fs.createWriteStream(pathoffile);
    createStream.end();
}

//Function that use to write in excel file
function excelWriter(filePath, content, sheetName) {
    let newWB = xlsx.utils.book_new();
    let newWS = xlsx.utils.json_to_sheet(content);
    xlsx.utils.book_append_sheet(newWB, newWS, sheetName);
    //File => create , replace
    xlsx.writeFile(newWB, filePath);
}

//Function that use to send mail.
function  SendMailByJS()
{
    var mail = nodemailer.createTransport({
        service: 'gmail',
        auth: {
          //Enter email and password
          user: email,
          pass: password
        }
      });
       
      var mailOptions = {
         from: email,
         to: 'abhishekmehta765@gmail.com',
         subject: 'Stock Record',
         //Attachment 
         attachments: [{
          filename: 'Stock.xlsx',
          path: 'C:/Users/hp/Desktop/Hackethon/activity/Stock.xlsx' 
         }]
      }
       
      mail.sendMail(mailOptions, function(error, info){
            if (error) {
              console.log(error);
            } else {
              console.log('Email sent: ' + info.response);
            }
      });
}
