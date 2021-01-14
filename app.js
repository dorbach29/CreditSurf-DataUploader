const XLSX = require('xlsx');
const chk = require('chalk');
const {MongoClient} = require('mongodb');

//Parser handles parsing of the sheet and uploading to mongo 
const parser = require('./functions.js');


//Function to get the sheet given a path
function getSheet(path){
    let workbook; 
    let sheet; 
    try{
        workbook = XLSX.readFile(path);
        sheet = workbook.Sheets[workbook.SheetNames[0]];
        console.log("Read File Successfully");
        if(!sheet) throw "Sheet not found";
    
    } catch (error) {
        console.log(chk.red.bold("app.js: Error reading xlsx file"));
        throw error;
    }
    return {workbook : workbook, sheet : sheet};
}

//Initiates connections to mongodb
//Upserts data to MongoDB Database given the MondoDbClient, Database, Collection, and ExcelSheet
async function main(client, dbName, collectionName, excelSheet, parser){
    try {
         
        //Getting the excell sheet with the data
        const {workbook, sheet} =  getSheet(excelSheet);

        await client.connect(); 
        console.log("app.js: Connected To DB");

        //Getting the mongoDB collection
        const database = client.db(dbName);
        const collection = database.collection(collectionName);

      

        await parser.importData(sheet, collection);
    }
    catch (error){
        console.log( error);
    } finally {
        await client.close();
    }
}

///MONGO OPTIONS
const client = MongoClient('mongodb://127.0.0.1:27017/', { useUnifiedTopology: true });
const dbName = "creditsurf";
const collectionName = "cards";


///RUNNING PORGRAM
if(process.argv[2]){
    let excelSheet = process.argv[2];
    main(client, dbName, collectionName, excelSheet, parser);
} else {
    console.log(chk.redBright("Please provide a path to a valid excelSheet: node app.js <path>"));
}
 












