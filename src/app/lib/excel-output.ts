import * as XLSX from 'xlsx';
import { Comment } from './parser';

export class ExcelOutput {
  // Constants for column headers
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

  // Helper to create cell references (e.g., "A1", "B2")
  private static createCellReference(header: string, index: number): string {
    return `${header}${index}`;
  }

  // Create a cell with number value
  private static createNumberCell(value: number, ref: string): XLSX.CellObject {
    return {
      t: 'n', // number type
      v: value,
      r: ref
    };
  }

  // Create a cell with text value
  private static createTextCell(value: string, ref: string): XLSX.CellObject {
    return {
      t: 's', // string type
      v: value,
      r: ref
    };
  }

  // Create header row
  private static createHeaderRow(): XLSX.WorkSheet {
    const ws = XLSX.utils.aoa_to_sheet([this.HEADERS]);
    
    // Set column widths (similar to OpenXML formatting)
    const wscols = this.HEADERS.map(() => ({ wch: 15 }));
    ws['!cols'] = wscols;

    return ws;
  }

  // Create content row from a comment
  private static createContentRow(comment: Comment): (string | number)[] {
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

  // Main function to create spreadsheet (equivalent to F# createSpreadsheet)
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
      const contentRows = comments.map(comment => 
        this.createContentRow(comment)
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

  // Helper method for API responses (creates buffer instead of file)
  static createWorkbookBuffer(comments: Comment[]): Buffer {
    // Create new workbook
    const workbook = XLSX.utils.book_new();
    
    // Create worksheet with headers
    const worksheet = this.createHeaderRow();
    
    // Add content rows
    const contentRows = comments.map(comment => 
      this.createContentRow(comment)
    );
    
    // Add content rows to worksheet starting from row 2
    XLSX.utils.sheet_add_aoa(worksheet, contentRows, { origin: 'A2' });
    
    // Add worksheet to workbook
    XLSX.utils.book_append_sheet(workbook, worksheet, 'Formatted Comments');
    
    // Write to buffer
    return XLSX.write(workbook, { type: 'buffer', bookType: 'xlsx' }) as Buffer;
  }
} 