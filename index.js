const request = require("request");
const xlsx = require("node-xlsx");
const HTMLParser = require("node-html-parser");
const fs = require("fs");
const monment = require("moment");
const { SSL_OP_EPHEMERAL_RSA } = require("constants");
const { resolve } = require("path");
monment.locale("zh-cn");

//定义休眠
const sleep = (time) => {
  return new Promise((resolve) => setTimeout(resolve, time));
};

//查询url
const url = "http://202.61.88.188/xmgk/Person/rList.aspx";

let allHumanData;

//读取excel
const workSheetsFromBuffer = xlsx.parse("./assets/names.xlsx");
workSheetsFromBuffer[0].data.forEach((element) => {
  //定义参数
  const options = {
    url: url,
    headers: {
      Connection: "keep-alive",
      "Accept-Encoding": "gzip, deflate",
      Accept:
        "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.9",
      "User-Agent":
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/88.0.4324.182 Safari/537.36 Edg/88.0.705.74",
      Host: "202.61.88.188",
      Origin: "http://202.61.88.188",
      Referer: "http://202.61.88.188/xmgk/Person/rList.aspx",
    },
    form: {
      __VIEWSTATE:
        "/wEPDwUKLTk0OTA3NDI0OGRktfRWnDBOmG3vQzZKy/XGkQBYTf6y/E1Q3cPcXzK5DBs=",
      __VIEWSTATEGENERATOR: "BF7D64B3",
      __EVENTVALIDATION:
        "/wEdAAIuDpyLtyFgJS1TWechhO4W+d3fBy3LDDz0lsNt4u+CuEELtJQu+cdOas7KA4kYTvVsitsWzJFprZmaicfjvdlH",
      mc: element[0],
      rybc: null,
      zsbh: null,
      sfzh: null,
      qymc: "中科标禾工程项目管理有限公司",
      ctl00$MainContent$Button1: "搜索",
    },
  };
  //防止服务器被抓爆
  sleep(5000);

  //获得数据
  request.post(options, (error, response, body) => {
    console.log("当前时间为:", monment().format("YYYY MMMM Do"));
    console.log("错误为:", error);
    console.log("状态码:", response && response.statusCode);
    console.log("输出到为:");
    let result = HTMLParser.parse(body);
    console.log(
      element[0] + "  " + result.querySelectorAll("a")[12].getAttribute("href")
    );
    //   fs.writeFile("data.json", body, (err) => {
    //     if (err) {
    //       return console.error(err);
    //     }
    //     console.log("成功");
    //   });
  });
});

//TODO toExcel

// let excelData = [
//   {
//     name: "sheet",
//     data: [["店名", "地址", "区域", "最低消费", "评价留言", "评分"]],
//   },
// ];
// let allData = fs.readFileSync("data.json");
// let formatData = JSON.parse(allData);

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
