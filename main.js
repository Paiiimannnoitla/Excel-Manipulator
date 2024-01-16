const Excel = require('exceljs')
const fs = require('fs')

const filelist = fs.readdirSync('./xlsx')
const xlsxArr = []
// Detect file extension
for(var i=0;i<filelist.length;i++){
	const f = filelist[i]
	const str = f.substring(f.length-5)
	if(str == '.xlsx'){
		xlsxArr[xlsxArr.length] = './xlsx/' + f
	}
}




const workbook = new Excel.Workbook()