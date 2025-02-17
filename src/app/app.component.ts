import { Component } from '@angular/core';
import * as XLSX from 'xlsx';
import * as FileSaver from 'file-saver';
import * as whatsapp from 'whatsapp-chat-parser';
import { Message } from './message';

@Component({
  selector: 'app-root',
  templateUrl: './app.component.html',
  styleUrl: './app.component.scss'
})
export class AppComponent {
  cardsWords = ['cards', 'uth', 'stud', 'ms', '3cp', '1cp', 'ocp', 'tcp', 'bj', 'blackjack'];
  audreyWords = ['start', 'bdp', 'bd', 'sj', 'bl', 'lbl', 'chd', 'msd', 'cmd', 'lmd', 'csd', 'dl', 'dlg', 'gi5', 'gl', 'rl', 'lol', 'pyl', 'sdjw', 'rts', 'sxr', 'rr', 'pp', 'tgt', 'tbs', 'twd', 'tjg', 'wsdr', 'wsds', 'wre', 'cws', 'w4c', 'w4cb', 'w4cf', 'w4cp', 'w4cw'];
  freeplayWords = ['fsp', 'freeplay'];
  w2gWords = ['w2g', 'taxes', 'tax'];
  multipleInOutWords = ['\\+'];
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
        const whatsAppMessages = whatsapp.parseString(text).map(message => ({ ...message, message: message.message.replace(/[\r\n]+/g, " ") }));
        const messages: Message[] = whatsAppMessages.map(msg => {
          const message = msg?.message;
          const inMatch = message?.match(/(?<=in.*)\d+\.*\d*/i);
          const outMatch = message?.match(/(?<=out.*)\d+\.*\d*/i);
          const additionalInMatch = message?.match(/(?<=in.*)(?<!.*out.*)(?<=\+\w*)\d+\.*\d*/ig);
          const additionalOutMatch = message?.match(/(?<=out.*)(?<=\+\w*)\d+\.*\d*/ig);
          let moneyIn: number | '' = inMatch ? +inMatch[0] : '';
          if (additionalInMatch) {
            additionalInMatch.forEach(inMatch => {
              moneyIn = +moneyIn + +inMatch;
            });
          }
          let out: number | '' = outMatch ? +outMatch[0] : '';
          if (additionalOutMatch) {
            additionalOutMatch.forEach(outMatch => {
              out = +out + +outMatch;
            });
          }
          const net = moneyIn || out ? +out - +moneyIn : '';
          const cards = this.matchesAnyWord(this.cardsWords, message) ? 'Yes' : '';
          const audrey = this.matchesAnyWord(this.audreyWords, message) ? 'Yes' : '';
          const freeplay = this.matchesAnyWord(this.freeplayWords, message) ? 'Yes' : '';
          const w2g = this.matchesAnyWord(this.w2gWords, message) ? 'Yes' : '';
          const multipleInOut = /\+\w*\d+/.test(message) ? 'Yes' : '';
          return { ...msg, in: moneyIn, out, net, cards, audrey, freeplay, w2g, multipleInOut };
        });
        console.log(messages);
        const workbook = XLSX.utils.book_new();
        this.createFilteredAndUnfilteredWorksheet(workbook, 'All Chats', messages);
        const names = [...new Set(messages.map(message => message.author))].filter(name => name) as string[];
        names.forEach(name => {
          this.createFilteredAndUnfilteredWorksheet(workbook, name, messages.filter(message => message.author === name))
        })
        const excelBuffer = XLSX.write(workbook, { bookType: 'xlsx', type: 'array' });
        const blob = new Blob([excelBuffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
        FileSaver.saveAs(blob, 'exported_data.xlsx');
      };
    }
  }

  createFilteredAndUnfilteredWorksheet(workbook: XLSX.WorkBook, sheetName: string, messages: whatsapp.Message[]) {
    const worksheet = this.createWorksheet(messages);
    XLSX.utils.book_append_sheet(workbook, worksheet, sheetName);
    const worksheetFiltered = this.createWorksheet(messages.filter(message => /in.*\d+.*out.*\d+/i.test(message.message)));
    XLSX.utils.book_append_sheet(workbook, worksheetFiltered, `${sheetName} Filtered`);
    const worksheetFiltered2 = this.createWorksheet(messages.filter(message => /\d+./.test(message.message) && !/in.*\d+.*out.*\d+/i.test(message.message)));
    XLSX.utils.book_append_sheet(workbook, worksheetFiltered2, `${sheetName} Filtered 2`);
  }

  createWorksheet(messages: whatsapp.Message[]) {
    const worksheet = XLSX.utils.json_to_sheet(messages);
    const authorMaxWidth = messages.map(messages => messages.author ? messages.author.length : 0).reduce((w, r) => Math.max(w, r), 10);
    worksheet["!cols"] = [{ wch: 14 }, { wch: authorMaxWidth }, { wch: 50 }];
    worksheet["!cols"][10] = { wch: 13.38 };
    this.formatColumn(worksheet, 0, 'm/d/yyyy h:mm');
    worksheet['!autofilter'] = { ref: "A1:K1" };
    return worksheet;
  }

  matchesAnyWord(matchableWords: string[], stringToTest: string) {
    return new RegExp('\\b' + matchableWords.join('\\b|\\b') + '\\b', 'i').test(stringToTest);
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
