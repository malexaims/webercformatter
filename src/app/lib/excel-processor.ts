import * as XLSX from 'xlsx';
import { Parser } from './parser';
import { ExcelOutput } from './excel-output';

export class ExcelProcessor {
  /**
   * Process an Excel file and create a formatted output
   * @param inputBuffer The input file as an ArrayBuffer
   * @returns Buffer containing the processed Excel file
   */
  static processBuffer(inputBuffer: ArrayBuffer): Buffer {
    try {
      // Read the workbook from buffer
      const workbook = XLSX.read(inputBuffer);
      
      // Get the first worksheet
      const firstSheetName = workbook.SheetNames[0];
      const worksheet = workbook.Sheets[firstSheetName];
      
      // Convert to array of arrays (equivalent to F#'s nested list)
      const rows = XLSX.utils.sheet_to_json(worksheet, { 
        header: 1,
        raw: false, // Get formatted text values
        defval: ''  // Default empty cells to empty string
      }) as string[][];
      
      // Extract and process comments
      const comments = Parser.extractCommentRows(rows);
      
      // Create output workbook buffer
      return ExcelOutput.createWorkbookBuffer(comments);
      
    } catch (error) {
      console.error('Error processing Excel file:', error);
      throw error;
    }
  }

  /**
   * Process Excel files using file paths (for Node.js environments)
   * @param inputPath Path to input Excel file
   * @returns boolean indicating success
   */
  static processFile(inputPath: string, outputPath: string): boolean {
    try {
      // Read the input file
      const workbook = XLSX.readFile(inputPath);
      
      // Get the first worksheet
      const firstSheetName = workbook.SheetNames[0];
      const worksheet = workbook.Sheets[firstSheetName];
      
      // Convert to array of arrays
      const rows = XLSX.utils.sheet_to_json(worksheet, { 
        header: 1,
        raw: false,
        defval: ''
      }) as string[][];
      
      // Extract and process comments
      const comments = Parser.extractCommentRows(rows);
      
      // Create output workbook buffer
      return ExcelOutput.createSpreadsheet(outputPath, 'Formatted Comments', comments);
      
    } catch (error) {
      console.error('Error processing file:', error);
      return false;
    }
  }
} 