export declare function parseExcelToJson(excel: Uint8Array, fileName: string): Data;
declare type KVData = {
    [key: string]: string;
};
declare type SheetData = {
    [key: string]: string | KVData | KVData[];
};
declare type Data = {
    [tabName: string]: SheetData;
};
export {};
//# sourceMappingURL=excel.d.ts.map