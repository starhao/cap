var express = require('express');
var router = express.Router();
const moment = require('moment');
const process = require('child_process');
const Excel = require('exceljs')

/* GET home page. */
router.get('/', function (req, res, next) {

  res.render('index', {title: 'Express'});
});

router.get('/writeToXlsx', async (req, res, next) => {
  //获取当前时间
  var endDate = moment().format('Y-M-D HH:mm:ss');
  //4天前的时间
  var startDate = moment().subtract(4, 'days').format('Y-M-D 00:00:00');

  var command = `git log --author=chris --pretty=format:'%ad,%s' --date=iso --since='${startDate}' --before='${endDate}'`;

  await process.exec(command, {cwd: `/Users/wuchenghao/Projects/angular2/platform-web`}, async (error, stdout, stderr) => {
    if (error) {
      console.error(`exec error :${error}`);
      console.log('出错了啊喂~');
      return;
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
        start: moment().subtract(i, 'days').format('Y-M-D 00:00:00'),
        end: moment().subtract(i, 'days').format('Y-M-D 12:00:00'),
        name: moment().subtract(i, 'days').format('Y-M-D') + '上午'
      };
      var afternoon = {
        start: moment().subtract(i, 'days').format('Y-M-D 12:00:00'),
        end: moment().subtract(i, 'days').format('Y-M-D 23:59:59'),
        name: moment().subtract(i, 'days').format('Y-M-D') + '下午'
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
    await writeToXlsx(newCommitInfo);

  });

  res.render('writeToXlsx', {title: '生成成功'});
})

function writeToXlsx(commitInfo) {
  var workbook = new Excel.Workbook();

  workbook.xlsx.readFile('/Users/wuchenghao/Resources/knrt/word/吴成浩工作周报.xlsx').then(() => {

    var worksheet = workbook.getWorksheet(1);

    for (var i = 5; i <= 14; i++) {
      var rowData = worksheet.getRow(i);
      rowData.getCell(4).value = commitInfo[i - 5].info.join(';\n');
      console.log(rowData.getCell(4).value);
    }

    rowData.commit();
    return workbook.xlsx.writeFile('/Users/wuchenghao/Resources/knrt/word/test.xlsx')
  });
}

module.exports = router;
