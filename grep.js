'use strict';
const fs = require('fs');
const path = require('path');
const XLSX = require("xlsx");
const moment = require("moment");
const program = require('commander');
const clc = require("cli-color");
const sprintf = require('sprintf-js').sprintf;

const DateHolidays = require('date-holidays');
const holidays = new DateHolidays('JP');

program
  .version('0.0.1')
  .option('-q, --quittingTime [value]', 'Check quittingTime')
  .option('-w, --workTime [value]', 'Check workTime')
  .option('-o, --overTime [value]', 'Check overTime')
  .option('-t, --totalWorkTime [value]', 'Check totalWorkTime')
  .option('--holiday', 'Check holiday')
  .parse(process.argv);

const EXCEL_DIR = './excel';

/*
const arrMin = arr => Math.min(...arr);
const arrMax = arr => Math.max(...arr);
const arrSum = arr => arr.reduce((a,b) => a + b, 0);
const arrAvg = arr => arr.reduce((a,b) => a + b, 0) / arr.length;
*/

const min2timestr = min => {
	if(!min) return 0;
	return sprintf("%02d:%02d",parseInt(min / 60),parseInt(min % 60));
}

class xlsxAdapter{
	constructor(fileName){
		this.workbook = XLSX.readFile(fileName);
		this.worksheet = this.workbook.Sheets[this.workbook.SheetNames[0]];
	}
	async cell(cell){
		return this.worksheet[cell].v;
	}
}

class jobkanParser{
	constructor(fileName){
		this.xlsx = new xlsxAdapter(fileName);
	}
	async init(){
		const monthStr = await this.xlsx.cell("A1");
		this.year = monthStr.match(/^(\d{4})/)[0];

		this.staffName = await this.xlsx.cell("A4");
		this.staffCode = await this.xlsx.cell("C4");
		this.group = await this.xlsx.cell("E4");
		this.staffType = await this.xlsx.cell("H4");
		this.totalWorkTime = await this.parseTimeCell("A8");
		this.totalOverTime = await this.parseTimeCell("B8");
		this.totalNightTime = await this.parseTimeCell("C8");
	}
	async parseTimeCell(cell){
		const value = await this.xlsx.cell(cell);
		return this.timestr2min(value);
	}
	timestr2min(timestr){
		const times = timestr.split(/:/,2);
		return parseInt(times[0]) * 60 + parseInt(times[1]);
	}

	async workSchedule(){
		const workSchedule = [];
		for(let i = 11; i <= 41; i++){
			const dateStr = await this.xlsx.cell("A" + i);
			const date = dateStr.match(/^(\d{2})\/(\d{2})/);
			if(dateStr.match(/^[^\d]/)) break;

			workSchedule.push({
				date: moment([this.year,date[1] - 1,date[2]]),
				attendanceScheduleTime: await this.xlsx.cell(`D${i}`),
				quittingScheduleTime: await this.xlsx.cell(`E${i}`),
				attendanceTime: await this.xlsx.cell(`F${i}`),
				quittingTime: await this.xlsx.cell(`G${i}`),
				workTime:  await this.xlsx.cell(`H${i}`),
				overTime:  await this.xlsx.cell(`J${i}`),
				nightTime: await this.xlsx.cell(`K${i}`),
				breakTime: await this.xlsx.cell(`L${i}`),
				paidTime:  await this.xlsx.cell(`M${i}`),
			});
		}
		return workSchedule;
	}
}

const main = async () =>{
	const result = fs.readdirSync(EXCEL_DIR);

	for(let i = 0;i < result.length;i++){
		if(result[i].match(/^~\$/)) continue;
		const fileName = path.join(EXCEL_DIR, result[i]);

		const jobkan = new jobkanParser(fileName);
		await jobkan.init();
		const workSchedule = await jobkan.workSchedule();

		if(program.totalWorkTime && jobkan.totalWorkTime > jobkan.timestr2min(program.totalWorkTime)){
			console.log(clc.blue(`${jobkan.staffName}(${jobkan.staffCode}:${jobkan.staffType}) 労働時間合計が${min2timestr(jobkan.totalWorkTime)}`));
		}

		for(let time of workSchedule){
			if(time.attendanceTime > '00:00' && time.quittingTime === '00:00'){
				console.log(clc.red(`${jobkan.staffName}(${jobkan.staffCode}:${jobkan.staffType}) ${time.date.format('YYYY年MM月DD日')}の退勤時間が未入力`));
			}

			if(program.holiday && time.workTime > '00:00'){
				const hd = holidays.isHoliday(time.date.toDate());
				if(hd){
					console.log(clc.blue(`${jobkan.staffName}(${jobkan.staffCode}:${jobkan.staffType}) ${time.date.format('YYYY年MM月DD日')}(${hd.name})に出勤、実労働時間:${time.workTime}`))
				} else if((time.date.weekday() == 0 || time.date.weekday() == 6)){
					console.log(clc.blue(`${jobkan.staffName}(${jobkan.staffCode}:${jobkan.staffType}) ${time.date.format('YYYY年MM月DD日')}(${time.date.weekday() == 0 ? '日曜日' : '土曜日'})に出勤、実労働時間:${time.workTime}`))
				}
			}

			let hasError = false;
			for(let key of ['quittingTime','workTime','overTime']){
				if (!program[key]) continue;
				if (time[key] > program[key]) hasError = true;
			}
			if(hasError) {
				console.log(clc.blue(`${jobkan.staffName}(${jobkan.staffCode}:${jobkan.staffType}) ${time.date.format('YYYY年MM月DD日')}が条件に該当`));
			}
		}
	}
};

main();


