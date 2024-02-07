const Excel = require('exceljs')
const fs = require('fs')

const filelist = fs.readdirSync('./xlsx')
const xlsxArr = []
/*
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
*/
// Small procurement filter
const main = async()=>{
	const file = './a.xlsx'
	const workbook = new Excel.Workbook
	const wb = await workbook.xlsx.readFile(file)
	const contentArr = ['郵資','印花稅','借款','所得稅','能量','天然氣','匯率','燃料','旅費','計程車','電力','汽電','容量',
		'電費','電能','匯調']
	const nameArr = ['電力','業務處週轉金','法院','經濟部','斯其大','郵政','水庫','電信','優必闊',
		'財政部','秀豐','嘉樂寶','福昇','麥寮']
	const delArr = []
	
	const nameMatch = (name)=>{
		for(var i=0;i<nameArr.length;i++){
			const isMatch = name.includes(nameArr[i])
			if(isMatch){
				return true
			}
		}
	}
	const contentMatch = (content)=>{
		for(var i=0;i<contentArr.length;i++){
			const isMatch = content.includes(contentArr[i])
			if(isMatch){
				return true
			}
		}
	}
	if(wb){
		// Match process
		const worksheet = workbook.worksheets[0]
		for(var i=2;i<999999;i++){
			const row = worksheet.getRow(i)
			const name = row.getCell(3).value
			if(!name){
				delArr[delArr.length] = i
				break
			}
			const nameStatus = nameMatch(name)
			if(nameStatus){
				delArr[delArr.length] = i
			}else{
				const content = row.getCell(4).value
				const contentStatus = contentMatch(content)
				if(contentStatus){
					delArr[delArr.length] = i
				}
			}			
		}
		// Delete process
		for(var i=0;i<delArr.length;i++){
			const c = delArr.length-i-1
			worksheet.spliceRows(delArr[c],1)
		}
		await workbook.xlsx.writeFile(file)
	}
}
const autoFiller = async()=>{
	const key = fs.readFileSync('./key.txt','utf8')
	const keyArr = key.split('\r\n')
	
	const value = fs.readFileSync('./value.txt','utf8')
	const valArr = value.split('\r\n')
	
	const file = './1130103-轉直供筆記-更新中.xlsx'
	const workbook = new Excel.Workbook
	const wb = await workbook.xlsx.readFile(file)
	
	if(wb){
		const worksheet = workbook.worksheets[0]
		for(var i=660;i<1000;i++){
			const row = worksheet.getRow(i)
			const contract = row.getCell(4).value
			const id = keyArr.indexOf(contract)
			if(id+1){
				console.log(i)
				const v = valArr[id]
				row.getCell(11).value = v
			}
			
		}
		const fin = await workbook.xlsx.writeFile(file) 
		if(fin){
			console.log('end')
		}
	}
}
const func = 2
const init = ()=>{
	if(func == 1){
		main()
	}else if(func == 2){
		autoFiller()
	}
}
init()