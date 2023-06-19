import { Component, OnInit } from '@angular/core';
import { read, writeFile, utils, WorkSheet, WorkBook, write } from "xlsx";
import { DatePipe } from '@angular/common';
import parse from 'date-fns/parse';
import differenceInMinutes from 'date-fns/differenceInMinutes';
import formatDuration from 'date-fns/formatDuration';

@Component({
  selector: 'app-root',
  templateUrl: './app.component.html',
  styleUrls: ['./app.component.scss'],
})
export class AppComponent implements OnInit {
  title = 'timesheet';
  loaded = false;

  today: Date;
  clockedIn: string | undefined | null;
  clockedOut: string | undefined | null;
  weekTotal = '';
  total = '';

  fileName = '';
  workbook: WorkBook | undefined;
  worksheet: any[] = [];
  todayRow: number = 0;

  constructor(private datePipe: DatePipe) {
    this.today = new Date();
  }

  ngOnInit(): void {
    // step 1: read spreadsheet
    // step 2: convert to array of objects
    // step 3: check today's date
    // step 4: fill in today's clock times if applicable
    // step 5: on clock functions, add new cells to the sheet

  }

  async onFileSelected(event: any) {
    const file: File = event.target?.files[0];

    if (file) {
      if(!file.type.includes('spreadsheet')) throw new Error('file is not a spreadsheet');
      this.fileName = file.name;
      const contents = await file.arrayBuffer()
      this.loaded = true;
      this.workbook = read(contents);
      this.parseSpreadsheet();
    }
    else throw new Error('file not found');
  }
  
  parseSpreadsheet() {
    const months = this.workbook?.SheetNames;
    const currentMonth = this.datePipe.transform(this.today, 'MMMM') || '';
    let sheet;
    const newDay = {
      day: this.datePipe.transform(this.today, 'shortDate'),
      timeIn: '',
      timeOut: '',
      totalMinutes: '',
      total: '',
    }
    if (!months?.includes(currentMonth)) {
      //make new sheet
      sheet = utils.json_to_sheet([newDay]);
      utils.book_append_sheet(this.workbook!, sheet, currentMonth);
    } else {
      //find current month sheet
      sheet = this.workbook?.Sheets[currentMonth];
    }
    this.worksheet = utils.sheet_to_json(sheet!);
    console.log('current month:', this.worksheet);
    const day = this.datePipe.transform(this.today, 'shortDate');
    if(!this.worksheet.find(row => row.day === day)){
      //make a new row
      this.worksheet.push(newDay);
    }
    this.todayRow = this.worksheet.findIndex((row) => row.day === day);
    this.clockedIn = this.worksheet[this.todayRow].timeIn || '';
    this.clockedOut = this.worksheet[this.todayRow].timeOut || '';
    console.log('current day:', this.worksheet[this.todayRow]);
  }
  saveSheet() {
    if (this.workbook && this.worksheet) {
      const currentMonth = this.datePipe.transform(this.today, 'MMMM') || '';
      this.workbook.Sheets[currentMonth] = utils.json_to_sheet(this.worksheet);
      writeFile(this.workbook, 'timesheet.xlsx');
      console.log('saved spreadsheet', this.worksheet);
    }
  }

  clockIn() {
    if (this.clockedIn || this.clockedOut) {
      return;
    }
    this.clockedIn = this.datePipe.transform(new Date(), 'shortTime');
    this.worksheet[this.todayRow].timeIn = this.clockedIn;
    this.saveSheet();
  }
  clockOut() {
    if (!this.clockedIn || this.clockedOut) {
      return;
    }
    this.clockedOut = this.datePipe.transform(new Date(), 'shortTime');
    this.worksheet[this.todayRow].timeOut = this.clockedOut;

    this.calculateTotal();
  }
  calculateTotal() {
    const inDate = parse(this.clockedIn!, 'h:mm a', new Date());
    const outDate = parse(this.clockedOut!, 'h:mm a', new Date());
    const totalMin = differenceInMinutes(outDate, inDate);
    this.total = formatDuration({
      hours: Math.floor(totalMin / 60),
      minutes: totalMin % 60,
    });

    this.worksheet[this.todayRow].totalMinutes = totalMin;
    this.worksheet[this.todayRow].total = this.total;
    this.saveSheet();
  }
  calculateWeek() {
    let startDay = this.worksheet.findIndex(row => row.day === 'weekTotal');
    startDay = startDay === -1 ? 0 : startDay;
    const week = this.worksheet.slice(startDay);
    const totalMin = week.reduce((prev, curr) => prev + curr.totalMinutes, 0);
    this.weekTotal = formatDuration({
      hours: Math.floor(totalMin / 60),
      minutes: totalMin % 60,
    });
    this.worksheet.push({day: 'weekTotal', total: this.weekTotal});
    this.saveSheet();
  }
}
