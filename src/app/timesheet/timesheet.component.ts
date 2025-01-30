import { Component } from '@angular/core';
import { CommonModule } from '@angular/common';
import * as XLSX from 'xlsx';
import saveAs from 'file-saver';

@Component({
  selector: 'app-timesheet',
  standalone: true,
  imports: [CommonModule],
  templateUrl: './timesheet.component.html',
  styleUrls: ['./timesheet.component.css'],
})
export class TimesheetComponent {
  dailyData: any[] = [];
  uploadedFilesCount = 0;
  uploadedFileNames: string[] = [];

  onFileUpload(event: any): void {
    const files = event.target.files;
    this.uploadedFilesCount = files.length;
    this.uploadedFileNames = [];
    this.dailyData = [];

    if (files.length > 0) {
      Array.from(files).forEach((file: any) => {
        this.uploadedFileNames.push(file.name);
        const reader = new FileReader();
        reader.onload = (e: any) => {
          const data = new Uint8Array(e.target.result);
          const workbook = XLSX.read(data, { type: 'array' });
          const sheetName = workbook.SheetNames[0];
          const worksheet = workbook.Sheets[sheetName];
          const parsedData = XLSX.utils.sheet_to_json(worksheet);
          this.dailyData.push(...parsedData);
        };
        reader.readAsArrayBuffer(file);
      });
    }
  }

  generateWeeklySummary(): void {
    if (!this.dailyData.length) return;

    const currentUserEmail = this.dailyData[0]['Owner Mailid'];
    const userTasks = this.dailyData.filter((task) => task['Owner Mailid'] === currentUserEmail);

    const weeklySummary = this.prepareWeeklySummary(userTasks);
    this.exportToExcelWithColors(weeklySummary, `Weekly_Timesheet_${currentUserEmail}`);
  }

  generateGroupLeaderFile(): void {
    if (!this.dailyData.length) return;

    const userSheets = this.groupBy(this.dailyData, 'Owner Mailid');
    const workbook = XLSX.utils.book_new();

    Object.keys(userSheets).forEach((userEmail) => {
      const userTasks = userSheets[userEmail];
      const weeklySummary = this.prepareWeeklySummary(userTasks);

      const worksheet = this.createColoredWorksheet(weeklySummary);
      XLSX.utils.book_append_sheet(workbook, worksheet, userEmail.split('@')[0]);
    });

    const excelBuffer = XLSX.write(workbook, { bookType: 'xlsx', type: 'array' });
    const blob = new Blob([excelBuffer], { type: 'application/octet-stream' });
    saveAs(blob, `Group_Weekly_Timesheet.xlsx`);
  }

  private prepareWeeklySummary(tasks: any[]): any[] {
    const groupedData = this.groupBy(tasks, 'Task/Issue ID');
    const weeklySummary: any[] = [];

    for (const taskId in groupedData) {
      const taskGroup = groupedData[taskId];

      const totalLogHours = this.calculateTotalHours(
        taskGroup.map((task) => task['Daily Log'] || '00:00')
      );

      const combinedComments = taskGroup
        .map((task) => task['Notes']?.trim() || '')
        .filter((comment) => comment && comment !== '-')
        .map((comment) => `• ${comment}`)
        .join('\n');

      const ownerMail = taskGroup[0]['Owner Mailid'] || '';
      const taskOwnerName = this.formatOwnerName(ownerMail);

      weeklySummary.push({
        'Task ID': taskId,
        'Task Name': taskGroup[0]['Task/General/Issue'] || '',
        'Project Name': 'I_TTF_DEV_ALL ROUNDERS',
        'Task List Name': taskGroup[0]['Task List/Module'] || '',
        'Custom Status': 'Completed',
        'Task Owner': taskOwnerName,
        'Total Log Hours': totalLogHours,
        'Task Comment': combinedComments || 'No Comments',
      });
    }

    return weeklySummary;
  }

  private groupBy(array: any[], key: string): { [key: string]: any[] } {
    return array.reduce((result, currentValue) => {
      (result[currentValue[key]] = result[currentValue[key]] || []).push(currentValue);
      return result;
    }, {});
  }

  private calculateTotalHours(timeStrings: string[]): string {
    let totalMinutes = 0;

    timeStrings.forEach((time) => {
      const [hours, minutes] = time.split(':').map((val) => parseInt(val, 10) || 0);
      totalMinutes += hours * 60 + minutes;
    });

    const totalHours = Math.floor(totalMinutes / 60);
    const remainingMinutes = totalMinutes % 60;

    return `${this.padWithZero(totalHours)}:${this.padWithZero(remainingMinutes)}`;
  }

  private formatOwnerName(email: string): string {
    const namePart = email.split('@')[0];
    return namePart.replace(/\./g, ' ');
  }

  private padWithZero(value: number): string {
    return value < 10 ? `0${value}` : `${value}`;
  }

  private createColoredWorksheet(data: any[]): any {
    const worksheet = XLSX.utils.json_to_sheet(data);

    const range = XLSX.utils.decode_range(worksheet['!ref']!);
    const headerStyle = {
      font: { bold: true, color: { rgb: 'FFFFFF' } }, // Bold white text
      fill: { fgColor: { rgb: 'E4C2E3' } }, // Light purple background
    };
    const rowColors = ['FFFFFF', 'C9DFF2']; // Alternating row colors: white and light blue

    // Apply header styles to the first row
    for (let col = range.s.c; col <= range.e.c; col++) {
      const cellAddress = XLSX.utils.encode_cell({ r: 0, c: col });
      if (worksheet[cellAddress]) {
        worksheet[cellAddress].s = headerStyle;
      }
    }

    // Apply alternating row colors
    for (let row = 1; row <= range.e.r; row++) {
      const rowStyle = { fill: { fgColor: { rgb: rowColors[row % 2] } } }; // Alternating color
      for (let col = range.s.c; col <= range.e.c; col++) {
        const cellAddress = XLSX.utils.encode_cell({ r: row, c: col });
        if (worksheet[cellAddress]) {
          worksheet[cellAddress].s = rowStyle;
        }
      }
    }

    return worksheet;
  }

  private exportToExcelWithColors(data: any[], fileName: string): void {
    const worksheet = this.createColoredWorksheet(data);

    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, 'Weekly Summary');

    const excelBuffer = XLSX.write(workbook, { bookType: 'xlsx', type: 'array' });
    const blob = new Blob([excelBuffer], { type: 'application/octet-stream' });
    saveAs(blob, `${fileName}.xlsx`);
  }
}
