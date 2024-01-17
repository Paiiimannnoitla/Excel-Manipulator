const Excel = require('exceljs')
const fs = require('fs')

const filelist = fs.readdirSync('./xlsx')
const xlsxArr = []
const main = async()=>{
	// Detect file extension
	for(var i=0;i<filelist.length;i++){
		const f = filelist[i]
		const str = f.substring(f.length-5)
		if(str == '.xlsx'){
			xlsxArr[xlsxArr.length] = './xlsx/' + f
		}
	}
	// Modification function
	for(var i=0;i<xlsxArr.length;i++){
		const f = xlsxArr[i]
		const workbook = new Excel.Workbook
		console.log(f)
		const wb = await workbook.xlsx.readFile(f)
		if(wb){
			const worksheet = workbook.worksheets[0]
			for(var j=7;j<35;j++){
				const row = worksheet.getRow(j)
				const height = row.height
				if(height){
					row.height = 95
				}else{
					row.height = height + 80
				}
				await workbook.xlsx.writeFile(f)
			}
		}			
	}
}


const init = ()=>{
	main()
}
init()