import { Component } from '@angular/core';
import * as XLSX from 'xlsx';
import * as FileSaver from 'file-saver';
import * as whatsapp from 'whatsapp-chat-parser';

@Component({
  selector: 'app-root',
  templateUrl: './app.component.html',
  styleUrl: './app.component.scss'
})
export class AppComponent {
  title = 'whatsapp-to-excel';
  file: File | undefined = undefined;

  onFileChange(event: any) {
    this.file = event.target.files[0];
  }

  importAndExport() {
    const reader = new FileReader();
    if (this.file) {
      reader.readAsText(this.file);
      reader.onload = () => {
        const text = reader.result as string;
        const messages = whatsapp.parseString(text);
        console.log(messages);
        const workbook = XLSX.utils.book_new();
        const worksheet = this.createWorksheet(messages);
        XLSX.utils.book_append_sheet(workbook, worksheet, 'All Chats');
        const names = [...new Set(messages.map(message => message.author))].filter(name => name) as string[];
        names.forEach(name => {
          const nameWorksheet = this.createWorksheet(messages.filter(message => message.author === name));
          XLSX.utils.book_append_sheet(workbook, nameWorksheet, name);
        })
        const excelBuffer = XLSX.write(workbook, { bookType: 'xlsx', type: 'array' });
        const blob = new Blob([excelBuffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
        FileSaver.saveAs(blob, 'exported_data.xlsx');
      };
    }
  }

  createWorksheet(messages: whatsapp.Message[]) {
    const worksheet = XLSX.utils.json_to_sheet(messages);
    const authorMaxWidth = messages.map(messages => messages.author ? messages.author.length : 0).reduce((w, r) => Math.max(w, r), 10);
    worksheet["!cols"] = [{ wch: 14 }, { wch: authorMaxWidth }, { wch: 50 }];
    this.formatColumn(worksheet, 0, 'm/d/yyyy h:mm');
    return worksheet;
  }

  formatColumn(worksheet: XLSX.WorkSheet, col: number, fmt: string) {
    if (worksheet['!ref']) {
      const range = XLSX.utils.decode_range(worksheet['!ref'])
      for (let row = range.s.r + 1; row <= range.e.r; ++row) {
        const ref = XLSX.utils.encode_cell({ r: row, c: col });
        if (worksheet[ref] && worksheet[ref].t === 'n') {
          worksheet[ref].z = fmt
        }
      }
    }
  }
}
