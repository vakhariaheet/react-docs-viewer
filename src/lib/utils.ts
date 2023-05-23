import { read, utils } from 'xlsx';
import {
	ExcelReturnType,
	ExcelData
} from './types.d';
export const readXlsxFile = async (
	file: string | File,
	returnType: ExcelReturnType = ExcelReturnType.ARRAY_OF_OBJECT,
): Promise<ExcelData> => {
	try {
		const buffer = await readFile(file);

		const wb = read(buffer, { type: 'array' });
		const json: ExcelData =
			[];

		for (const sheet of wb.SheetNames) {
			const ws = wb.Sheets[sheet];
			let totalCols = 0;
			let totalRows = 0;

			if (returnType === ExcelReturnType.ARRAY_OF_OBJECT) {
				const data = utils.sheet_to_json(ws);
				if (data.length > totalRows) {
					totalRows = data.length;
				}
				data.forEach((row: any) => {
					if (Object.keys(row).length > totalCols) {
						totalCols = Object.keys(row).length;
					}
				});

				json.push({
					sheet,
					returnType,
					data,
					merges:
						wb.Sheets[sheet]['!merges']?.map((item: any) => [
							item.s.r,
							item.s.c,
							item.e.r,
							item.e.c,
						]) || [],
					totalCols,
					totalRows,
				});
			} else {
				const data = utils.sheet_to_json(ws, { header: 1 }) as any;
				if (data.length > totalRows) {
					totalRows = data.length;
				}
				data.forEach((row: any) => {
					if (row.length > totalCols) {
						totalCols = row.length;
					}
				});

				json.push({
					sheet,
					returnType,
					data,
					merges:
						wb.Sheets[sheet]['!merges']?.map((item: any) => [
							item.s.r,
							item.s.c,
							item.e.r,
							item.e.c,
						]) || [],
					totalCols,
					totalRows,
				});
			}
		}
		return json;
	} catch (e: any) {
		throw new Error(e);
	}
};

const readFile = async (file: File | string): Promise<ArrayBuffer> =>
	new Promise((resolve, reject) => {
        if (typeof file === 'string') {
            
			fetch(file)
				.then((resp) => resp.arrayBuffer())
				.then((buffer) => {
					resolve(buffer);
                }).catch((e) => {
                    reject(e);

                });
		}

		if (file instanceof File) {
			const reader = new FileReader();
			reader.onload = (e) => {
				const data = e.target?.result;
				if (!data) reject(new Error('Failed to read file'));
				if (data instanceof ArrayBuffer) resolve(data);
				else {
					reject(new Error('Failed to read file'));
				}
			};
			reader.readAsArrayBuffer(file);
		}
	});
