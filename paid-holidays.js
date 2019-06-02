'use strict';
const fs = require('fs');
const path = require('path');
const moment = require("moment");
const program = require('commander');
const clc = require("cli-color");
const sprintf = require('sprintf-js').sprintf;
const csvParse = require('csv-parse/lib/sync')

program
	.version('0.0.1')
  .option('-o, --output [filename]', 'Output CSV filename')
  .parse(process.argv);

const EXCEL_DIR = './excel';

const main = async () => {

	const result = fs.readdirSync(EXCEL_DIR);

	const dates = {};
	const staffs = {};

	// 仮データの作成と最終日判定
	for(let i = 0;i < result.length;i++){
		if(!result[i].match(/^custom_[\dA-Za-z]+\.csv$/)) continue;
		const fileName = path.join(EXCEL_DIR, result[i]);
		let data = csvParse(fs.readFileSync(fileName),{from: 2});
		for(let row of data){
			const dateTime = moment(new Date(row[0].split('/')));
			if(!dateTime.isValid()) continue;
			dates[dateTime.unix()] = true;

			const staffCode = parseInt(row[1]);
			if(!staffs[staffCode]){
				staffs[staffCode] = {
					code: staffCode,
					name: row[2],
					transactions: []
				}
			}
		}
	}
	const lastSummaryTime = parseInt((Object.keys(dates).sort().reverse())[0]);

	for(let i = 0;i < result.length;i++){
		if(!result[i].match(/^custom_.+\.csv$/)) continue;
		const fileName = path.join(EXCEL_DIR, result[i]);
		let data = csvParse(fs.readFileSync(fileName),{from: 2});
		for(let row of data){
			const dateTime = moment(new Date(row[0].split('/')));
			if(!dateTime.isValid()) continue;

			const staffCode = parseInt(row[1]);
			if(dateTime.unix() === lastSummaryTime){
				staffs[staffCode].existPaidHolidays = parseFloat(row[9])
			}
			if(row[5] == "0" && row[6] == "0" && row[7] == "0:00" && row[8] == "0") continue;

			staffs[staffCode].transactions.push({
				dateUnix: dateTime.unix(),
				position: row[3],
				belongs: row[4],
				lostPaidDays: parseFloat(row[5]),
				usedPaidDays: parseFloat(row[6]),
				addPaidDays: parseFloat(row[8])
			});
		}
	}
	// データの出力
	for(let staffCode in staffs){
		const staff = staffs[staffCode];
		staff.addPaidDays = [];
		staff.lastYear = {
			totalLostPaidDays: 0,
			totalUsedPaidDays: 0
		};
		staff.thisYear = {
			totalLostPaidDays: 0,
			totalUsedPaidDays: 0
		};

		for(let transaction of staff.transactions){
			if(transaction.addPaidDays >= 10) staff.addPaidDays.push(transaction.dateUnix);
		}

		// 最終有給付与日
		if(staff.addPaidDays.length > 0){
			staff.lastPaidDay = (staff.addPaidDays.sort().reverse())[0];
		}

		if(staff.lastPaidDay){
			if(staff.lastPaidDay < moment.unix(lastSummaryTime).add(-12,'months').add(1,'days').unix()){
				console.error(clc.red(`${staff.name}(${staffCode}) 最終有給付与日が${moment.unix(staff.lastPaidDay).format('YYYY年MM月DD日')}で以降付与されていません`));
				staff.lastYear.from = moment.unix(staff.lastPaidDay).add(-12,'months').unix();
				staff.lastYear.to = moment.unix(staff.lastPaidDay).add(-1,'days').unix();
			} else {
				// 最終付与日から最終集計日まで
				staff.thisYear.from = moment.unix(staff.lastPaidDay).unix();
				staff.thisYear.to = moment.unix(lastSummaryTime).unix();

				// 最終付与日の1年前から最終付与日前日まで
				staff.lastYear.from = moment.unix(staff.lastPaidDay).add(-12,'months').unix();
				staff.lastYear.to = moment.unix(staff.lastPaidDay).add(-1,'days').unix();
			}
		}
		for(let transaction of staff.transactions){
			if(staff.thisYear.from && staff.thisYear.from <= transaction.dateUnix && staff.thisYear.to >= transaction.dateUnix) {
				staff.thisYear.totalLostPaidDays += transaction.lostPaidDays;
				staff.thisYear.totalUsedPaidDays += transaction.usedPaidDays;
			}
			if(staff.lastYear.from && staff.lastYear.from <= transaction.dateUnix && staff.lastYear.to >= transaction.dateUnix) {
				staff.lastYear.totalLostPaidDays += transaction.lostPaidDays;
				staff.lastYear.totalUsedPaidDays += transaction.usedPaidDays;
			}
		}
	}
	const fileName = program.output || path.join(EXCEL_DIR, `summary-${moment.unix(lastSummaryTime).format('YYYYMMDD')}.csv`);
	if(fs.existsSync(fileName)) fs.unlinkSync(fileName);
	console.log(fileName);

	// utf8 bom
	fs.appendFileSync(fileName,'\uFEFF');
	fs.appendFileSync(
		fileName,
		`スタッフコード,姓 名,最終有給付与日,有給残数,`
	);
	fs.appendFileSync(
		fileName,
		`有給消滅日数（集計期間内）,有給消化日数（前期間）,有給消化日数（現在期間）\r\n`
	);

	for(let staffCode in staffs){
		const staff = staffs[staffCode];
		fs.appendFileSync(
			fileName,
			`${staffCode},${staff.name},${staff.lastPaidDay ? moment.unix(staff.lastPaidDay).format('YYYY/MM/DD') : ""},${staff.existPaidHolidays || 0},`
		);
		fs.appendFileSync(
			fileName,
			`${staff.lastYear.totalLostPaidDays},${staff.lastYear.totalUsedPaidDays},${staff.thisYear.totalUsedPaidDays}\r\n`
		);
	}
};

main();


