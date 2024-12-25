import * as XLSX from "xlsx";

/**
 * Exports data to an Excel file (.xlsx).
 *
 * @param {any[]} data - The data to export. Each object in the array represents a row.
 * @param {string} fileName - The name of the generated Excel file (without extension).
 *
 * @example
 * const data = [
 *   { name: "John", age: 30 },
 *   { name: "Jane", age: 25 }
 * ];
 * exportToExcel(data, "Users");
 * // This will generate a file named "Users.xlsx"
 */
export function exportToExcel(data: any[], fileName: string): void {
  const worksheet = XLSX.utils.json_to_sheet(data);
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, "Sheet1");
  XLSX.writeFile(workbook, `${fileName}.xlsx`);
}

/**
 * Imports data from an Excel file and maps rows to objects using the header row.
 *
 * @template T - The shape of the object each row will be mapped to.
 * @param {File} file - The Excel file to import.
 * @returns {Promise<T[]>} A promise that resolves to an array of objects representing the Excel data.
 *
 * @example
 * interface User {
 *   name: string;
 *   age: number;
 * }
 * const file = new File([...], "example.xlsx");
 * importFromExcel<User>(file).then(data => {
 *   console.log(data);
 * });
 */
export function importFromExcel<T = Record<string, any>>(
  file: File
): Promise<T[]> {
  return new Promise((resolve, reject) => {
    if (!file) {
      reject("No file provided");
      return;
    }

    const reader = new FileReader();
    reader.onload = (event: any) => {
      try {
        const data = new Uint8Array(event.target.result);
        const workbook = XLSX.read(data, { type: "array" });
        const firstSheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheetName];

        // Convert worksheet to JSON
        const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
        const headers: string[] = jsonData[0] as string[];
        const dataRows: any[][] = (jsonData as any[][]).slice(1);

        // Map rows to objects using headers
        const formattedData = dataRows.map((row: any[]) => {
          return row.reduce(
            (acc: Record<string, any>, cell: any, index: number) => {
              acc[headers[index]] = cell;
              return acc;
            },
            {}
          ) as T;
        });

        resolve(formattedData);
      } catch (error) {
        reject(`Error processing the file: ${(error as Error).message}`);
      }
    };

    reader.onerror = (error) => {
      reject(`File read error: ${error.target?.error?.message}`);
    };

    reader.readAsArrayBuffer(file);
  });
}
