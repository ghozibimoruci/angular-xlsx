import { CommonModule } from '@angular/common';
import { Component, ViewChild, ElementRef } from '@angular/core';
import { RouterOutlet } from '@angular/router';
import * as XLSX from 'xlsx';

@Component({
  selector: 'app-root',
  imports: [RouterOutlet, CommonModule],
  templateUrl: './app.component.html',
  styleUrl: './app.component.css'
})
export class AppComponent {
  title = 'myapp';
  isDragging = false;
  @ViewChild('fileInput') fileInput!: ElementRef<HTMLInputElement>;

  // Allowed file extensions for XLSX
  allowedExtensions = ['xls', 'xlsx', 'csv', 'ods'];

  fileList: FileList | null = null;
  theFile: File | null = null;
  jsonData: string = '';

  onDragOver(event: DragEvent) {
    event.preventDefault();
    event.stopPropagation();
    this.isDragging = true;
  }

  onDragLeave(event: DragEvent) {
    event.preventDefault();
    event.stopPropagation();
    this.isDragging = false;
  }

  onDrop(event: DragEvent) {
    event.preventDefault();
    event.stopPropagation();
    this.isDragging = false;

    if (event.dataTransfer?.files.length) {
      this.fileList = this.filterValidFiles(event.dataTransfer.files);
      this.theFile = this.fileList[0];
      this.handleExcelFile(this.theFile);
    }
  }

  onFileChange(event: Event) {
    const input = event.target as HTMLInputElement;
    if (input.files?.length) {
      this.fileList = this.filterValidFiles(input.files);
      this.theFile = this.fileList[0];
      this.handleExcelFile(this.theFile);
    }
  }

  filterValidFiles(files: FileList): FileList {
    const validFilesArray = Array.from(files).filter(file => {
      const fileExtension = file.name.split('.').pop()?.toLowerCase();
      return fileExtension && this.allowedExtensions.includes(fileExtension);
    });

    // Convert array back to FileList
    const dataTransfer = new DataTransfer();
    validFilesArray.forEach(file => dataTransfer.items.add(file));

    return dataTransfer.files;
  }

  clearFileList(){
    this.fileList = null;
    this.theFile = null;
  }

  openFileSelector() {
    this.fileInput.nativeElement.click();
  }

  async handleExcelFile(file: File) {
    const reader = new FileReader();
    reader.readAsBinaryString(file);

    reader.onload = (e) => {
      const binaryStr = e.target?.result as string;
      const workbook = XLSX.read(binaryStr, { type: 'binary' });

      // Get the first sheet
      const sheetName = workbook.SheetNames[0];
      const sheet = workbook.Sheets[sheetName];

      // Convert sheet to JSON
      const jsonData = XLSX.utils.sheet_to_json(sheet, { header: 1 });
      
      // Convert to beautified JSON
      this.jsonData = JSON.stringify(jsonData, null, 2);
    };

    reader.onerror = (error) => {
      console.error('Error reading file:', error);
    };
  }
}
