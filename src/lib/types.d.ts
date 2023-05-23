
export enum ExcelReturnType {
    ARRAY_OF_OBJECT = 'ARRAYOFOBJECT',
    ARRAY_OF_ARRAY = 'ARRAYOFARRAY',
}


interface SheetDataWithAOOReturnType {
    returnType: ReturnType.ARRAY_OF_OBJECT;
    sheet: string;
    data: unknown[];
    merges: number[][];
    totalCols: number;
    totalRows: number;
}

interface SheetDataWithAOAReturnType { 
    returnType: ReturnType.ARRAY_OF_ARRAY;
    sheet: string;
    data: unknown[][];
    merges: number[][];
    totalCols: number;
    totalRows: number;
}

export type ExcelData = (SheetDataWithAOOReturnType | SheetDataWithAOAReturnType)[];