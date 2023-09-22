import { IWorkbook } from ".."

const generateSharedStrings = (workbook: IWorkbook) => {

	// console.log(workbook.strings)

	const uniqueCount = workbook.strings.reduce((acc:string[], cur) => acc.includes(cur) ? acc : [...acc, cur], []).length;



	return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
	<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="${workbook.strings.length}" uniqueCount="${uniqueCount}">
		${workbook.strings.map(string => `
		<si><t>${string}</t></si>
		`).join('\n')}
	</sst>`
}


export { generateSharedStrings }