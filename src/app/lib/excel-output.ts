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

  private static createWorksheet(comments: Comment[]): XLSX.WorkSheet {
    // Create the data array (without headers since they'll be added by json_to_sheet)
    const data = comments.map(comment => ({
      'Number': comment.number,
      'Created By': comment.createdBy,
      'Created On': comment.createdOn,
      'Status': comment.status,
      'Category': comment.category,
      'Reference': comment.reference,
      'Content': comment.content.replace(/\r\n/g, '\n'),
      'Response': '',
      'Comment Addressed': 'False',
      'Comment Backchecked': 'False'
    }));

    // Create worksheet with headers
    const ws = XLSX.utils.json_to_sheet(data, {
      header: this.HEADERS,
      skipHeader: false
    });

    // Set column widths
    ws['!cols'] = [
      { wch: 10 },  // Number
      { wch: 15 },  // Created By
      { wch: 15 },  // Created On
      { wch: 12 },  // Status
      { wch: 15 },  // Category
      { wch: 15 },  // Reference
      { wch: 60 },  // Content
      { wch: 40 },  // Response
      { wch: 15 },  // Comment Addressed
      { wch: 15 }   // Comment Backchecked
    ];

    // Get worksheet range
    const range = XLSX.utils.decode_range(ws['!ref'] || 'A1');

    // Apply styles to all cells
    for (let R = range.s.r; R <= range.e.r; ++R) {
      for (let C = range.s.c; C <= range.e.c; ++C) {
        const cell = XLSX.utils.encode_cell({ r: R, c: C });
        if (!ws[cell]) continue;

        // Convert cell to proper format if needed
        if (typeof ws[cell] === 'number' || typeof ws[cell] === 'string' || typeof ws[cell] === 'boolean') {
          const value = ws[cell];
          ws[cell] = {
            v: value,
            t: typeof value === 'number' ? 'n' : typeof value === 'boolean' ? 'b' : 's'
          };
        }

        // Apply styles based on whether it's a header or content
        if (!ws[cell].s) ws[cell].s = {};
        Object.assign(ws[cell].s, R === 0 ? {
          fill: { fgColor: { rgb: "D3D3D3" }, patternType: 'solid' },
          font: { bold: true, sz: 12 },
          alignment: { horizontal: 'center', vertical: 'center', wrapText: true },
          border: {
            top: { style: 'thin' },
            bottom: { style: 'thin' },
            left: { style: 'thin' },
            right: { style: 'thin' }
          }
        } : {
          font: { sz: 11 },
          alignment: { vertical: 'center', wrapText: true },
          border: {
            top: { style: 'thin' },
            bottom: { style: 'thin' },
            left: { style: 'thin' },
            right: { style: 'thin' }
          }
        });
      }
    }

    return ws;
  }

  static createWorkbookBuffer(comments: Comment[]): Buffer {
    const wb = XLSX.utils.book_new();
    const ws = this.createWorksheet(comments);
    XLSX.utils.book_append_sheet(wb, ws, 'Formatted Comments');
    
    // Set workbook properties
    wb.Workbook = {
      Views: [{ RTL: false }],
      Sheets: [{
        Hidden: 0,
      }]
    };

    return XLSX.write(wb, { 
      type: 'buffer',
      bookType: 'xlsx',
      cellStyles: true,
      compression: true,
      Props: {
        Author: "FDOT Comment Reformatter"
      },
      bookSST: false
    }) as Buffer;
  }

  static createSpreadsheet(filepath: string, sheetName: string, comments: Comment[]): boolean {
    try {
      const wb = XLSX.utils.book_new();
      const ws = this.createWorksheet(comments);
      XLSX.utils.book_append_sheet(wb, ws, sheetName);
      
      // Set workbook properties
      wb.Workbook = {
        Views: [{ RTL: false }],
        Sheets: [{
          Hidden: 0,
          Selected: true
        }]
      };

      XLSX.writeFile(wb, filepath, { cellStyles: true });
      return true;
    } catch (error) {
      console.error('Error creating spreadsheet:', error);
      return false;
    }
  }
} 