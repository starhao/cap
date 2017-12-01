var express = require('express');
var router = express.Router();
const Moment = require('moment');
const Process = require('child_process');
const Excel = require('exceljs');

/* GET home page. */
router.get('/', function (req, res, next) {
  res.render('index', {title: 'Express'});
});

/* 生成文件的post请求*/
router.post('/', (req, res) => {
  var params = req.body;
  if (params.path && params.username && params.template) {
    try {
      //获取当前时间
      var endDate = Moment().format('Y-M-D HH:mm:ss');
      //4天前的时间
      var startDate = Moment().subtract(4, 'days').format('Y-M-D 00:00:00');

      var command = `git log --author=${params.username} --pretty=format:'%ad,%s' --date=iso --since='${startDate}' --before='${endDate}'`;

      var path = params.path
      Process.exec(command, {cwd: path}, async (error, stdout, stderr) => {
        if (error) {
          console.error(`exec error :${error}`);
          res.render('error', {message: '出错了啊喂', error: {status: '', stack: '可能是用户名或者路径输错了什么的'}})
        }
        //截取每个提交
        var resultArr = stdout.split('\n');

        //将每个提交的时间和提交信息封装成json格式
        var commitInfo = [];

        resultArr.forEach(value => {
          commitInfo.push({time: value.split(',')[0], info: value.split(',')[1]})
        });

        //execl表格的10个时间段
        var timeArr = [];
        for (var i = 4; i >= 0; i--) {
          var morning = {
            start: Moment().subtract(i, 'days').format('Y-M-D 00:00:00'),
            end: Moment().subtract(i, 'days').format('Y-M-D 12:00:00'),
            name: Moment().subtract(i, 'days').format('Y-M-D') + '上午'
          };
          var afternoon = {
            start: Moment().subtract(i, 'days').format('Y-M-D 12:00:00'),
            end: Moment().subtract(i, 'days').format('Y-M-D 23:59:59'),
            name: Moment().subtract(i, 'days').format('Y-M-D') + '下午'
          };
          timeArr.push(morning, afternoon)
        }

        //新的提交信息, 按照10个时间段区分开
        var newCommitInfo = [];

        timeArr.forEach(value => {
          var infoArr = [];
          commitInfo.forEach(curInfo => {
            //在这个时间段内
            if (curInfo.time < value.end && curInfo.time > value.start)
              infoArr.push(curInfo.info)
          });
          newCommitInfo.push({'time': value.name, info: infoArr});
        });
        //写入xlsx表格
        var distPath = await writeToXlsx(newCommitInfo, params.template);
        res.render('success', {title: `生成成功,新文件的路径是:${distPath}`});
      });
    } catch (err) {
      console.log(err);
      res.render('error', {message: '出错了啊喂', error: {status: '', stack: 'xxxxx'}})
    }
  } else {
    res.render('error', {message: '参数未填写完整', error: {status: '', stack: 'xxxxx'}})
  }
});

async function writeToXlsx(commitInfo, template) {

  var workbook = new Excel.Workbook();

  return new Promise((resolve => {
    workbook.xlsx.readFile(`${template}`).then(() => {
      var worksheet = workbook.getWorksheet(1);
      //第5~14行是要填写信息的行
      for (var i = 5; i <= 14; i++) {
        var rowData = worksheet.getRow(i);
        rowData.getCell(4).value = commitInfo[i - 5].info.join(';\n');
      }
      rowData.commit();
      //目标路径
      var distPath = `${template.split('.')[0]} (${Moment().format('Y-M-D')}).xlsx`;
      //写入新的文件
      workbook.xlsx.writeFile(distPath);
      resolve(distPath);
    });
  }))


}

module.exports = router;
