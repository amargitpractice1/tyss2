import {getLastRow,readLastRowAndColumn,readDataFromExcel} from '../ExcelUtility/excel.js'


describe("read multiple values",async ()=>{
    let filepath='test/excelfile/abcd.xlsx'
    it("reading the attributes",async ()=>{
      let lastrow=   await getLastRow(filepath,"Sheet2");

      for(let i=0;i<=lastrow;i++)
      {
         let data = await readDataFromExcel(filepath,"Sheet2",i,1)
         console.log(data)
      }
         
    })

    it("reading values",async ()=>{
        let lastrow= await getLastRow(filepath,"Sheet2");
        for(let i=0;i<=lastrow;i++)
        {
        let attribute = await readDataFromExcel(filepath,"Sheet2",i,1);
        let value = await readDataFromExcel(filepath,"Sheet2",i,2)
        console.log(attribute+"------->"+value)
        }
    })
})