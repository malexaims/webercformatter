// Define the Comment type as an interface (equivalent to F# record)
export interface Comment {
  number: number;
  createdBy: string;
  createdOn: string;
  status: string;
  category: string;
  reference: string;
  content: string;
}

export class Parser {
  static extractCommentRows(rows: string[][]): Comment[] {
    // Create indexed rows (equivalent to List.zip with filtering)
    const iRows = rows
      .map((row, index) => [index, row] as const)
      .filter(([, row]) => row.length > 0);

    // Extract reviewer comments (groups of 4 rows starting with "No")
    const reviewersComments = iRows
      .filter(([, row]) => row[0] === "No")
      .map(([index]) => rows.slice(index, index + 4));

    // Helper functions to extract specific fields
    // These are direct translations of the F# point-free functions
    const number = (c: string[][]) => c[0][0];
    const createdBy = (c: string[][]) => c[2][1];
    const createdOn = (c: string[][]) => c[2][2];
    const status = (c: string[][]) => c[0][1];
    const category = (c: string[][]) => c[0][4];
    const reference = (c: string[][]) => c[0][3];
    const content = (c: string[][]) => c[3][1];

    // Map comment groups to Comment objects and sort by number
    // This is equivalent to the F# list comprehension and List.sortBy
    return reviewersComments
      .map(c => ({
        number: parseInt(number(c)),
        createdBy: createdBy(c),
        createdOn: createdOn(c),
        status: status(c),
        category: category(c),
        reference: reference(c),
        content: content(c)
      }))
      .sort((a, b) => a.number - b.number);
  }
} 