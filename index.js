const axios = require("axios");
const xlsx = require("node-xlsx");
let qs = require("qs");
const HTMLParser = require("node-html-parser");
const fs = require("fs");
const monment = require("moment");
monment.locale("zh-cn");

//url
const url = "http://202.61.88.188/xmgk/Person/rList.aspx";

//所有人的网址
let allHumanData = [
  {
    name: "sheet",
    data: [["姓名", "地址"]],
  },
];

let allDetails = [
  {
    name: "sheet",
    data: [["名字", "在建项目数量", "链接"]],
  },
];

// //读取excel
// const workSheetsFromBuffer = xlsx.parse("./assets/names.xlsx");
// //索引
// let index = workSheetsFromBuffer[0].data.length;
// let currentIndex = 0;

// const outPutHuman = () => {
//   if (currentIndex < index) {
//     //config
//     let config = {
//       url: url,
//       method: "POST",
//       headers: {
//         "Content-Type": "application/x-www-form-urlencoded",
//         "Accept-Encoding": "gzip, deflate",
//         Accept:
//           "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.9",
//         "Accept-Language": "zh-CN,zh;q=0.9,en;q=0.8,en-GB;q=0.7,en-US;q=0.6",
//         "User-Agent":
//           "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/88.0.4324.182 Safari/537.36 Edg/88.0.705.74",
//         Host: "202.61.88.188",
//         Origin: "http://202.61.88.188",
//         Referer: "http://202.61.88.188/xmgk/Person/rList.aspx",
//       },
//       data: qs.stringify({
//         __VIEWSTATE:
//           "/wEPDwUKLTk0OTA3NDI0OGRktfRWnDBOmG3vQzZKy/XGkQBYTf6y/E1Q3cPcXzK5DBs=",
//         __VIEWSTATEGENERATOR: "BF7D64B3",
//         __EVENTVALIDATION:
//           "/wEdAAIuDpyLtyFgJS1TWechhO4W+d3fBy3LDDz0lsNt4u+CuEELtJQu+cdOas7KA4kYTvVsitsWzJFprZmaicfjvdlH",
//         mc: workSheetsFromBuffer[0].data[currentIndex][0],
//         qymc: "中科标禾工程项目管理有限公司",
//         ctl00$MainContent$Button1: "搜索",
//       }),
//     };
//     axios(config)
//       .then((response) => {
//         console.log("当前时间为:", monment().format("YYYY MMMM Do"));
//         console.log("状态码:", response.status);
//         console.log("输出为:");
//         let result = HTMLParser.parse(response.data);
//         //抓取信息
//         console.log(
//           workSheetsFromBuffer[0].data[currentIndex][0] +
//             ": http://202.61.88.188/xmgk/Person/" +
//             result.querySelectorAll("a")[12].getAttribute("href")
//         );
//         console.log(workSheetsFromBuffer[0].data[currentIndex][0] + "抓取成功");

//         allHumanData[0].data.push([
//           workSheetsFromBuffer[0].data[currentIndex][0],

//           "http://202.61.88.188/xmgk/Person/" +
//             result.querySelectorAll("a")[12].getAttribute("href"),
//         ]);

//         //增加索引
//         currentIndex += 1;
//       })
//       .catch((error) => {
//         console.log("错误为:", error);
//       });
//   } else {
//     console.log("完成抓取");
//     //buffer
//     let buffer = xlsx.build(allHumanData);

//     fs.writeFile("data.xlsx", buffer, (err) => {
//       if (err) {
//         return console.error(err);
//       }
//     });
//     console.log("完成抓取稍等开始解析");
//     clearInterval(intervel);
//   }
// };

// //开始运行
// let intervel = setInterval(outPutHuman, 3000);

//TODO toExcel

// let excelData = [
//   {
//     name: "sheet",
//     data: [["店名", "地址", "区域", "最低消费", "评价留言", "评分"]],
//   },
// ];
let allData = fs.readFileSync("data.xlsx");
let formatData = xlsx.parse(allData);
//最终数据buffer
let detailBuffer;
formatData[0].data.forEach((value) => {
  allDetails[0].data.push([value[0], "链接", value[1]]);
});

detailBuffer = xlsx.build(allDetails);

console.log(allDetails);
fs.writeFile("final.xlsx", detailBuffer, (err) => {
  if (err) {
    console.log("遇到错误:");
    return console.error(err);
  }
});

// console.log("全部完毕可以查看!!");

// formatData.data.searchResult.forEach((element) => {
//   excelData[0].data.push([
//     element.title,
//     element.address,
//     element.areaname,
//     element.lowestprice,
//     element.comments,
//     element.avgscore,
//   ]);
// });

// let buffer = xlsx.build(excelData);

// fs.writeFile("人员数据"+monment().format("YYYY-MM-Do")+".xlsx", buffer, (err) => {
//   if (err) {
//     console.log("写入失败:" + err);
//     return;
//   }

//   console.log("写入完成");
// });
