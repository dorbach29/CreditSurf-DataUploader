const XLSX = require('xlsx');
const chk = require('chalk');


//FETCHING EXCEL WORKBOOK + GETTING FIRST sheet; 
let workbook; 
let sheet; 
try{
    workbook = XLSX.readFile(process.argv[2]);
    sheet = workbook.Sheets[workbook.SheetNames[0]];
    console.log(chk.green("Read File Successfully"));
    if(!sheet) throw "Sheet not found";

} catch (error) {
    console.log(chk.red.bold("ERROR READING THE FILE:"));
    throw error;
}







let row = 1;
//Getting cell range
let a1_range = sheet['!ref'];
//Decoding range
let range = XLSX.utils.decode_range(a1_range);
let endCol = range.e.c;
while(true){

    //CHECK IF END//
    const primary_address = {c:0, r: row};
    const primary_cell = sheet[XLSX.utils.encode_cell(primary_address)].v;
    if(primary_cell === "~END~") break;
    //ALL OPERATIONS BELLOW//


    for(let col = 0; col <= endCol; col++){

        //Getting Cell
        const cell_adress = {c:col, r:row};
        const cell = sheet[XLSX.utils.encode_cell(cell_adress)].v;
        
        //Getting Header
        const h_adress = {c:col, r:0};
        const header = sheet[XLSX.utils.encode_cell(h_adress)].v;
    }
    
    
    row++;
}

