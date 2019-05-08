'use strict';
const fs = require('fs');
const path = require('path');
const XLSX = require("xlsx");
const moment = require("moment");
const program = require('commander');

program
  .version('0.0.1')
  .option('-q, --quittingTime [value]', 'Check quittingTime')
  .option('-w, --workTime [value]', 'Check workTime')
  .option('-o, --overTime [value]', 'Check overTime')
  .option('-t, --totalWorkTime [value]', 'Check totalWorkTime')
  .parse(process.argv);

const EXCEL_DIR = './excel';

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
		this.totalWorkTime = await this.xlsx.cell("A8");
		this.totalOverTime = await this.xlsx.cell("B8");
		this.totalNightTime = await this.xlsx.cell("C8");
	}
	async parseTimeCell(cell){
		const value = await this.xlsx.cell(cell);
		return parseInt(value.split(/:/,2).join(''));
	}
	async workSchedule(){
		const workSchedule = [];
		for(let i = 11; i <= 41; i++){
			const dateStr = await this.xlsx.cell("A" + i);
			const date = dateStr.match(/^(\d{2})\/(\d{2})/);
			if(dateStr.match(/^[^\d]/)) break;

			workSchedule.push({
				date: moment([this.year,date[1] - 1,date[2]]).format("YYYY/MM/DD"),
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
		if(program.totalWorkTime && parseInt(jobkan.totalWorkTime.split(/:/,2).join('')) > parseInt(program.totalWorkTime.split(/:/,2).join(''))){
			console.log(`${result[i]} ${jobkan.staffName}(${jobkan.staffType}) <- totalWorkTime:${jobkan.totalWorkTime}`);
		}

		for(let time of workSchedule){
			if(time.attendanceTime > '00:00' && time.quittingTime === '00:00'){
				console.warn(`${result[i]} ${jobkan.staffName}(${jobkan.staffType}) ${time.date} <- quittingTime:00:00`);
			}

			let hasError = false;

			for(let key of ['quittingTime','workTime','overTime']){
				if (!program[key]) continue;
				if (time[key] > program[key]) hasError = true;
			}
			if(hasError) {
				console.log(`${result[i]} ${jobkan.staffName}(${jobkan.staffType}) ${time.date}`);
			}
		}
	}
};

main();


