import React, {
	ButtonHTMLAttributes,
	DetailedHTMLProps,
	HTMLProps,
	TdHTMLAttributes,
	ThHTMLAttributes,
	useEffect,
	useState,
} from 'react';
import {
	ExcelData,
	ExcelReturnType
} from '../../types.d';
import styles from './Excel.module.css';
import { readXlsxFile } from '../../utils';

export interface ExcelProps {
	file: string | File;
	loadingComponent?: React.ReactNode;
	errorComponent?: React.ReactNode;
	children?: (
		data: ExcelData,
	) => React.ReactNode;
	elements?: Partial<{
		sheetBtn: DetailedHTMLProps<
			ButtonHTMLAttributes<HTMLButtonElement>,
			HTMLButtonElement
		>;
		sheetBtnText: (name: string, index: number, allSheet: string[]) => string;
		sheetBtnContainer: HTMLProps<HTMLDivElement>;
		tableContainer: HTMLProps<HTMLDivElement>;
		table: HTMLProps<HTMLTableElement>;
		row: HTMLProps<HTMLTableRowElement>;
		headerCell: DetailedHTMLProps<
			ThHTMLAttributes<HTMLTableCellElement>,
			HTMLTableCellElement
		>;
		cell: DetailedHTMLProps<
			TdHTMLAttributes<HTMLTableCellElement>,
			HTMLTableCellElement
		>;
		thead: HTMLProps<HTMLTableSectionElement>;
		tbody: HTMLProps<HTMLTableSectionElement>;
		mainContainer: HTMLProps<HTMLDivElement>;
	}>;
	disabledCopy?: boolean;
}

const Excel: React.FC<ExcelProps> = ({ file, children, elements,disabledCopy }) => {
	const [data, setData] = useState<ExcelData>([]);
	const [loading, setLoading] = useState(false);
	const [error, setError] = useState<string>('');
	const [currentSheetIndex, setCurrentSheetIndex] = useState(0);
	const renderCells = (row: any, index: number) => {
		const cells = [];
		for (let i = 0; i < row.length; i++) {
			if (row[i] === undefined) {
				cells.push(
					<td
						key={i}
						{...elements?.cell}
						className={`${styles.cell} ${styles.cell}`}
					></td>,
				);
				continue;
			}
			const mergeAddr = data[currentSheetIndex].merges.find(
				(cell) => cell[1] === i && cell[0] === index,
			);
			const rowSpan = mergeAddr ? mergeAddr[2] - mergeAddr[0] + 1 : 1;
			const colSpan = mergeAddr ? mergeAddr[3] - mergeAddr[1] + 1 : 1;
			cells.push(
				<td
					key={i}
					{...elements?.cell}
					colSpan={colSpan}
					rowSpan={rowSpan}
					className={`${styles.cell} ${styles.cell}`}
					onCopy={
						disabledCopy ? (e) => e.preventDefault() : undefined
					}
				>
					{row[i]}
				</td>,
			);
		}
		return cells;
	};
	useEffect(() => {
		(async () => {
			try {
				setLoading(true);
				const json = (await readXlsxFile(file, ExcelReturnType.ARRAY_OF_ARRAY)).map(
					(table) => {
						const { totalCols } = table;

						const rows = [];

						for (const row of table.data) {
							if (!(row instanceof Array)) continue;
							if (row.length === 0) rows.push(new Array(totalCols).fill(''));
							else {
								const rowIndex = table.data.indexOf(row);
								let colsToAdd = totalCols - row.length;

								table.merges.forEach((cell) => {
									if (rowIndex >= cell[1] && rowIndex <= cell[2]) {
										const totolMergeCell = cell[3] - cell[1] + 1;

										if (totolMergeCell > colsToAdd) {
											colsToAdd = 0;
											return;
										}
										colsToAdd = colsToAdd - totolMergeCell;
									}
								});
								rows.push([...row, ...new Array(colsToAdd).fill('')]);
							}
						}
						return {
							...table,
							data: rows,
						};
					},
				);
				setData(json);
				setLoading(false);
			} catch (err) {
				setLoading(false);
				if (err instanceof Error) setError(err.message);
				else setError('Something went wrong');
			}
		})();
	}, [file]);
	if (loading) return <div>Loading...</div>;
	if (error) return <div>{error} </div>;
	if (children) return <>{children(data)}</>;

	return (
		<div {...elements?.mainContainer}>
			{/* Header */}
			<div
				{...elements?.sheetBtnContainer}
				className={`${styles.btnContainer} ${elements?.sheetBtnContainer?.className}`}
			>
				{data.map(({ sheet }, index) => (
					<button
						key={index}
						{...elements?.sheetBtn}
						onClick={(e) => {
							setCurrentSheetIndex(index);
							elements?.sheetBtn?.onClick && elements?.sheetBtn?.onClick(e);
						}}
						disabled={
							index === currentSheetIndex ||
							(elements?.sheetBtn?.disabled
								? elements?.sheetBtn?.disabled
								: false)
						}
						onCopy={
							disabledCopy ? (e) => e.preventDefault() : undefined
						}
						className={`${styles.btn} ${elements?.sheetBtn?.className}`}
					>
						{elements?.sheetBtnText
							? elements?.sheetBtnText(
									sheet,
									index,
									data.map(({ sheet }) => sheet),
							  )
							: sheet}
					</button>
				))}
			</div>

			<div
				{...elements?.tableContainer}
				className={`${styles.tableContainer} ${elements?.tableContainer}`}
			>
				{data[currentSheetIndex] && (
					<table
						{...elements?.table}
						className={`${styles.table} ${elements?.table?.className}`}
					>
						<thead {...elements?.thead}>
							<tr
								{...elements?.row}
								className={`${styles.row} ${elements?.row?.className}`}
							>
								{(data[currentSheetIndex].data[0] as any) &&
									(data[currentSheetIndex].data[0] as string[]).map(
										(item: number | string, index: number) => {
											if (item === undefined)
												return (
													<th
														{...elements?.headerCell}
														key={index}
														className={`${styles.headerCell} ${elements?.cell?.className}`}
													></th>
												);
											if (typeof item === 'number') item = item.toString();
											const mergeAddr = data[currentSheetIndex].merges.find(
												(cell) => cell[1] === index && cell[0] === 0,
											);
											const rowSpan = mergeAddr
												? mergeAddr[2] - mergeAddr[0] + 1
												: 1;
											const colSpan = mergeAddr
												? mergeAddr[3] - mergeAddr[1] + 1
												: 1;
											return (
												<th
													key={index}
													{...elements?.headerCell}
													colSpan={colSpan}
													rowSpan={rowSpan}
													className={`${styles.headerCell} ${elements?.cell?.className}`}
													onCopy={
														disabledCopy ? (e) => e.preventDefault() : undefined
													}
												>
													{item.startsWith('__EMPTY') ? '' : item}
												</th>
											);
										},
									)}
							</tr>
						</thead>
						<tbody {...elements?.tbody}>
							{data[currentSheetIndex].data.map((row: any, index: number) => {
								if (index === 0) return null;
								return (
									<tr
										{...elements?.row}
										key={index}
										className={`${styles.row} ${elements?.row?.className}`}
									>
										{renderCells(row, index)}
									</tr>
								);
							})}
						</tbody>
					</table>
				)}
			</div>
		</div>
	);
};
export default Excel;
