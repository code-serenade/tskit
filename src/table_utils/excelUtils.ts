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
 * Imports data from an Excel file (.xlsx) selected via an <input type="file"> element.
 *
 * @returns {Promise<any[]>} A promise that resolves to an array of objects,
 *                           where each object represents a row of data from the Excel file.
 *
 * @throws {Error} If no file is selected or the file fails to load.
 *
 * @example
 * // HTML:
 * // <input type="file" id="fileInput" />
 *
 * importFromExcel().then(data => {
 *   console.log(data);
 * }).catch(error => {
 *   console.error(error);
 * });
 *
 * @remarks
 * This function relies on a hidden <input type="file"> element in the DOM to get the file.
 * Ensure the input element is present and accessible in the document.
 */
export function importFromExcel(): Promise<any[]> {
  return new Promise((resolve, reject) => {
    const fileInput = document.querySelector(
      'input[type="file"]'
    ) as HTMLInputElement;
    const file = fileInput?.files?.[0];

    if (!file) {
      reject("No file selected");
      return;
    }

    const reader = new FileReader();
    reader.onload = (event: any) => {
      const data = new Uint8Array(event.target.result);
      const workbook = XLSX.read(data, { type: "array" });
      const firstSheetName = workbook.SheetNames[0];
      const worksheet = workbook.Sheets[firstSheetName];

      const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
      const headers: string[] = jsonData[0] as string[];
      const dataRows: any[][] = (jsonData as any[][]).slice(1);

      const formattedData = dataRows.map((row: any[]) => {
        return row.reduce((acc: any, cell: any, index: number) => {
          acc[headers[index]] = cell;
          return acc;
        }, {});
      });

      resolve(formattedData);
    };

    reader.onerror = (error) => {
      reject(error);
    };

    reader.readAsArrayBuffer(file);
  });
}
