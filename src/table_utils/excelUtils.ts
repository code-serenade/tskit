import * as XLSX from "xlsx";

export function exportToExcel(data: any[], fileName: string): void {
  const worksheet = XLSX.utils.json_to_sheet(data);
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, "Sheet1");
  XLSX.writeFile(workbook, `${fileName}.xlsx`);
}

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

      // 读取数据时指定从第二行开始，第一行作为表头
      const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
      const headers: string[] = jsonData[0] as string[];
      const dataRows: any[][] = (jsonData as any[][]).slice(1);

      // 格式化数据
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
