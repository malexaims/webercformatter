import { NextResponse } from 'next/server';
import * as XLSX from 'xlsx';
import { Parser } from '@/app/lib/parser';
import { ExcelOutput } from '@/app/lib/excel-output';

export async function POST(request: Request) {
  try {
    const formData = await request.formData();
    const file = formData.get('file') as File;
    
    if (!file) {
      return NextResponse.json(
        { error: 'No file provided' },
        { status: 400 }
      );
    }

    // Read the file
    const buffer = await file.arrayBuffer();
    const workbook = XLSX.read(buffer);
    
    // Get the first worksheet
    const worksheet = workbook.Sheets[workbook.SheetNames[0]];
    
    // Convert to array of arrays
    const rows = XLSX.utils.sheet_to_json(worksheet, { header: 1 }) as string[][];
    
    // Extract comments
    const comments = Parser.extractCommentRows(rows);
    const processedBuffer = ExcelOutput.createWorkbookBuffer(comments);

    // Return the processed file
    return new NextResponse(processedBuffer, {
      headers: {
        'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        'Content-Disposition': `attachment; filename="reformatted_comments.xlsx"`,
      },
    });
  } catch (error) {
    console.error('Error processing file:', error);
    return NextResponse.json(
      { error: 'Failed to process file' },
      { status: 500 }
    );
  }
} 