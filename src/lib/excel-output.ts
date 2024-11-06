import * as XLSX from 'xlsx';
import { Comment } from './parser';

export class ExcelOutput {
  private static readonly HEADERS = [
    'Number',
    'Created By',
    'Created On',
    'Status',
    'Category',
    'Reference',
    'Content',
    'Response',
    'Comment Addressed',
    'Comment Backchecked'
  ];

  static createHeaderRow(): XLSX.WorkSheet {
    const ws = XLSX.utils.aoa_to_sheet([this.HEADERS]);
    
    // Set column widths
    const wscols = this.HEADERS.map(() => ({ wch: 15 }));
    ws['!cols'] = wscols;

    return ws;
  }

  static createContentRow(comment: Comment, index: number): any[] {
    return [
      comment.number,           // A: Number
      comment.createdBy,        // B: Created By
      comment.createdOn,        // C: Created On
      comment.status,           // D: Status
      comment.category,         // E: Category
      comment.reference,        // F: Reference
      comment.content,          // G: Content
      '',                       // H: Response
      'False',                  // I: Comment Addressed
      'False'                   // J: Comment Backchecked
    ];
  }

  static createSpreadsheet(
    filepath: string,
    sheetName: string,
    comments: Comment[]
  ): boolean {
    try {
      // Create new workbook
      const workbook = XLSX.utils.book_new();
      
      // Create worksheet with headers
      const worksheet = this.createHeaderRow();
      
      // Add content rows
      const contentRows = comments.map((comment, idx) => 
        this.createContentRow(comment, idx + 2)
      );
      
      // Add content rows to worksheet starting from row 2
      XLSX.utils.sheet_add_aoa(worksheet, contentRows, { origin: 'A2' });
      
      // Add worksheet to workbook
      XLSX.utils.book_append_sheet(workbook, worksheet, sheetName);
      
      // Write to file
      XLSX.writeFile(workbook, filepath);
      
      return true;
    } catch (error) {
      console.error('Error creating spreadsheet:', error);
      return false;
    }
  }

  // Helper method to create a workbook buffer (useful for API responses)
  static createWorkbookBuffer(comments: Comment[]): Buffer {
    // Create new workbook
    const workbook = XLSX.utils.book_new();
    
    // Create worksheet with headers
    const worksheet = this.createHeaderRow();
    
    // Add content rows
    const contentRows = comments.map((comment, idx) => 
      this.createContentRow(comment, idx + 2)
    );
    
    // Add content rows to worksheet starting from row 2
    XLSX.utils.sheet_add_aoa(worksheet, contentRows, { origin: 'A2' });
    
    // Add worksheet to workbook
    XLSX.utils.book_append_sheet(workbook, worksheet, 'Formatted Comments');
    
    // Write to buffer
    const buffer = XLSX.write(workbook, { type: 'buffer', bookType: 'xlsx' });
    
    return buffer;
  }
} 