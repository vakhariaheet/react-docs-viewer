import { useEffect, useState } from 'react';
import './App.css';
import { ExcelReturnType, Excel, readXlsxFile } from './lib/index';

function App() {
	useEffect(() => {
		const fetchData = async () => {
			const result = await readXlsxFile(
				'https://papeer.s3.ap-south-1.amazonaws.com/Untitled.xlsx',
				ExcelReturnType.ARRAY_OF_ARRAY,
			);

			console.log(result);
		};
		fetchData();
	}, []);
	return (
		<Excel
			file={'https://papeer.s3.ap-south-1.amazonaws.com/Untitled.xlsx'}
			elements={{
				sheetBtn: {
					disabled: false,
				},
			}}
			disabledCopy
		/>
	);
}

export default App;
