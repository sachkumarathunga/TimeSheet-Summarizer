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
  // Weekly Timesheet Data
  uploadedFiles: any[] = [];
  uploadedFilesCount = 0;
  uploadedFileNames: string[] = [];

  // Team Leader Timesheet Data
  leaderUploadedFiles: any[] = [];
  leaderUploadedFilesCount = 0;
  leaderUploadedFileNames: string[] = [];

  onWeeklyFileUpload(event: any): void {
    this.processFileUpload(event, 'weekly');
  }

  onLeaderFileUpload(event: any): void {
    this.processFileUpload(event, 'leader');
  }

  private processFileUpload(event: any, type: 'weekly' | 'leader'): void {
    const files = event.target.files;
    if (!files.length) return;

    if (type === 'weekly') {
      this.uploadedFilesCount = files.length;
      this.uploadedFileNames = [];
      this.uploadedFiles = [];
    } else {
      this.leaderUploadedFilesCount = files.length;
      this.leaderUploadedFileNames = [];
      this.leaderUploadedFiles = [];
    }

    Array.from(files).forEach((file: any) => {
      const reader = new FileReader();
      reader.onload = (e: any) => {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array' });
        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];
        const parsedData = XLSX.utils.sheet_to_json(worksheet);

        if (type === 'weekly') {
          this.uploadedFileNames.push(file.name);
          this.uploadedFiles.push(...parsedData);
        } else {
          this.leaderUploadedFileNames.push(file.name);
          this.leaderUploadedFiles.push(...parsedData);
        }
      };
      reader.readAsArrayBuffer(file);
    });
  }

  generateWeeklySummary(): void {
    if (!this.uploadedFiles.length) return;

    const currentUserEmail = this.uploadedFiles[0]['Owner Mailid'];
    const userTasks = this.uploadedFiles.filter((task) => task['Owner Mailid'] === currentUserEmail);

    const weeklySummary = this.prepareSummary(userTasks);
    this.exportToExcelWithSuccess(weeklySummary, `Weekly_Timesheet_${currentUserEmail}`);
  }

  generateTeamLeaderTimesheet(): void {
    if (!this.leaderUploadedFiles.length) return;

    // Group by "Task Owner" first
    const groupedByOwner = this.groupBy(this.leaderUploadedFiles, 'User');
    let teamTimesheet: any[] = [];

    Object.keys(groupedByOwner).forEach((owner) => {
      const tasks = groupedByOwner[owner];

      // Now, Group by "Task ID" to prevent duplicates
      const groupedByTaskID = this.groupBy(tasks, 'Task/Issue ID');

      Object.keys(groupedByTaskID).forEach((taskId) => {
        const taskGroup = groupedByTaskID[taskId];

        // Merge Daily Log hours for each task
        const totalLogHours = this.calculateTotalHours(
          taskGroup.map((task) => task['Daily Log'] || '00:00')
        );

        // Combine task comments
        const combinedComments = taskGroup
          .map((task) => task['Notes']?.trim() || '')
          .filter((comment) => comment && comment !== '-')
          .map((comment) => `• ${comment}`)
          .join('\n');

        // Push unique task per owner
        teamTimesheet.push({
          'Task ID': taskId,
          'Task Name': taskGroup[0]['Task/General/Issue'] || '',
          'Project Name': taskGroup[0]['Project Name'] || '',
          'Task List Name': taskGroup[0]['Task List/Module'] || '',
          'Custom Status': 'Completed',
          'Task Owner': taskGroup[0]['User'] || '',
          'Total Log Hours': totalLogHours,
          'Task Comment': combinedComments || 'No Comments',
        });
      });
    });

    this.exportToExcelWithSuccess(teamTimesheet, `Team_Leader_Timesheet`);
  }

  private prepareSummary(tasks: any[]): any[] {
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

      weeklySummary.push({
        'Task ID': taskId,
        'Task Name': taskGroup[0]['Task/General/Issue'] || '',
        'Project Name': taskGroup[0]['Project Name'] || '',
        'Task List Name': taskGroup[0]['Task List/Module'] || '',
        'Custom Status': 'Completed',
        'Task Owner': taskGroup[0]['User'] || '',
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

  private calculateTotalHours(timeStrings: string | string[]): string {
    let totalMinutes = 0;
    
    // Ensure timeStrings is always an array
    const timeArray = Array.isArray(timeStrings) ? timeStrings : [timeStrings];

    timeArray.forEach((time) => {
      if (typeof time !== 'string') return; // Ignore invalid data
      const [hours, minutes] = time.split(':').map((val) => parseInt(val, 10) || 0);
      totalMinutes += hours * 60 + minutes;
    });

    return `${Math.floor(totalMinutes / 60)}:${totalMinutes % 60}`;
  }

  private exportToExcelWithSuccess(data: any[], fileName: string): void {
    const worksheet = XLSX.utils.json_to_sheet(data);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, 'Summary');
    const excelBuffer = XLSX.write(workbook, { bookType: 'xlsx', type: 'array' });
    const blob = new Blob([excelBuffer], { type: 'application/octet-stream' });
    saveAs(blob, `${fileName}.xlsx`);

    setTimeout(() => {
      alert(`✅ ${fileName}.xlsx has been successfully downloaded!`);
      window.location.reload(); // Auto refresh after download
    }, 500);
  }
}
