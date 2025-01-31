import { Component } from '@angular/core';
import { TimesheetComponent } from './timesheet/timesheet.component';

@Component({
  selector: 'app-root',
  standalone: true,
  imports: [TimesheetComponent], // Import TimesheetComponent
  template: `<app-timesheet></app-timesheet>`,
})
export class AppComponent {
  title(title: any) {
    throw new Error('Method not implemented.');
  }
}
