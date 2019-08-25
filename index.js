const express = require("express");
// 引入所需要的第三方包
const superagent = require("superagent");
const xlsx = require("xlsx");
var iconv = require("iconv-lite");

const app = express();

const jquery = require("jquery");
const jsdom = require("jsdom");
let es = null;

let getHotNews = res => {
  let hotNews = { title: {}, data: [] };
  let dom = new jsdom.JSDOM(res);
  let $ = jquery(dom.window);
  // 找到目标数据所在的页面元素，获取数据
  $("div#content table tr").each((idx, ele) => {
    if (idx == 0) {
      hotNews.title = {
        date: $(ele).find("td").eq(0).text().trim(),
        status: $(ele).find("td").eq(1).text().trim(),
        code: $(ele).find("td").eq(2).text().trim(),
        winddiretion: $(ele).find("td").eq(3).text().trim()
      };
    } else {
      let item = {
        date: $(ele).find("td").eq(0).text().replace(/[\r\n]/g, "").replace(/\ +/g, ""),
        status: $(ele).find("td").eq(1).text().replace(/[\r\n]/g, "").replace(/\ +/g, ""),
        code: $(ele).find("td").eq(2).text().replace(/[\r\n]/g, "").replace(/\ +/g, ""),
        winddiretion: $(ele).find("td").eq(3).text().replace(/[\r\n]/g, "").replace(/\ +/g, "")
      };
      hotNews.data.push(item);
    }
  });
  return hotNews;
};

app.get("/", async (req, res, next) => {
  let month = req.query.month;
  superagent
    .get("http://www.tianqihoubao.com/lishi/gdqingyuan/month/" + month + ".html")
    .set("content-type", "text/html; charset=utf-8")
    .responseType("blob")
    .end((err, r) => {
      if (err) {
        // 报错拦截
        console.log(`热点新闻抓取失败 - ${err}`);
      } else {
        //将获取的二进制数据转中文文本
        var resBody = iconv.decode(Buffer.from(r.body), "gb2312");
        //数据提取格式化
        let formatBody = getHotNews(resBody);
        
        //通过工具将json转表对象
        let target = xlsx.utils.json_to_sheet(formatBody.data); 
        //定义表头
        target.A1.v = formatBody.title.date;
        target.B1.v = formatBody.title.status;
        target.C1.v = formatBody.title.code;
        target.D1.v = formatBody.title.winddiretion;
        let keys = Object.keys(target).sort();
        let ref = target["!ref"]; //定义表范围
        let filename = `./清远天气数据.xls`;
        try{
          es = xlsx.readFile('./清远天气数据.xls', {cellStyles: true})
        }catch(error) {
          
        }
        if(es == null) {
          let workbook = {
            //定义操作文档
            SheetNames: [month], //定义表明
            Sheets: {
              [month]: Object.assign({}, target, { "!ref": ref }) //表对象
            }
          };
          xlsx.writeFile(workbook, filename); //将数据写入文件
          res.send(formatBody);
          console.log('创建成功')
        }else {
          if(es.SheetNames.indexOf(month) > -1) {
            console.log("本地已有改月份数据，无需再创建")
            res.send(formatBody);
            return
          }
          es.SheetNames.push(month);
          es.Sheets[month] = Object.assign({}, target, { "!ref": ref }) //表对象
          es.SheetNames.sort();
          Object.keys(es.Sheets).sort();
          xlsx.writeFile(es, filename); //将数据写入文件
          res.send(formatBody);
          console.log('创建成功')
        }
      }
    });
});



let server = app.listen(3000, function() {
  let host = server.address().address;
  let port = server.address().port;
  console.log(`Your App is running at http://${host}:${port}`);
});
