import excel  from 'exceljs';

export async function readDataFromExcel (filepath,sheetname,rownum,cellno)
                                    {
                                        const workbook=new excel.Workbook();
                                        await workbook.xlsx.readFile(filepath);
                                        const worksheet=workbook.getWorksheet(sheetname)
                                        let data= worksheet.getRow(rownum).getCell(cellno).toString();
                                      
                                        return data


                                    }


                                   

export async function readLastRowAndColumn(filePath,sheetname) {
                                      try {
                                        // Load the Excel workbook
                                        const workbook = new excel.Workbook();
                                        await workbook.xlsx.readFile(filePath);
                                    
                                        // Access the active worksheet
                                        const worksheet = workbook.getWorksheet(sheetname); // Assuming the first worksheet
                                    
                                        // Find the last row with data (considering non-empty cells)
                                        let lastRowIndex = 0;
                                        for (let i = 1; i <= worksheet.rowCount; i++) { // Start from row 1 (index 0)
                                          const row = worksheet.getRow(i);
                                          if (row.values.some(cellValue => !!cellValue)) { // Check for non-empty value
                                            lastRowIndex = i;
                                          }
                                        }
                                    
                                        // Find the last column with data
                                        let lastColumnIndex = 0;
                                        for (let col = 1; col <= worksheet.columnCount; col++) { // Start from column 1 (index 0)
                                          const columnValues = worksheet.getColumnValues(col);
                                          if (columnValues.some(cellValue => !!cellValue)) { // Check for non-empty value
                                            lastColumnIndex = col;
                                          }
                                        }
                                    
                                        return { lastRowIndex, lastColumnIndex };
                                    
                                      } catch (error) {
                                        console.error('Error reading Excel file:', error);
                                        return {}; // Return empty object on error
                                      }
                                    }

 export async function readLastrow(filePath,sheetname)
 {
                                        const workbook=new excel.Workbook();
                                        await workbook.xlsx.readFile(filePath);
                                        const worksheet=workbook.getWorksheet(sheetname)
                                        let range =worksheet.getU
 }  
 
 export async function getLastRow(filepath,sheetname)
                                   {
                                    const workbook=new excel.Workbook();
                                    await workbook.xlsx.readFile(filepath);
                                    const worksheet=workbook.getWorksheet(sheetname)
                                    const rows = worksheet.getColumn(1);
                                    const rowsCount = rows['_worksheet']['_rows'].length;
                                    return rowsCount
                                   }
                                    
                      
                                                                    