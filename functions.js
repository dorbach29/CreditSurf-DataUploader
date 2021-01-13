const XLSX = require('xlsx');
const chk = require('chalk');

function getString(data){
    return data; 
};

function getInt(data){
    return parseInt(data, 10);
};
function getArray(data){
    return data.split(';');
};


headerFunctions = {
CardName : getString,
CardType :getString,
CreditNetwork : getString,
Bank : getString,
GuideLink:getString,
CardLink:getString,
TimeDomestic:getInt,
TimeInternational:getInt,      
CoverageLimit: getArray,
PriceLimit:getInt,
CoverageType:getString,
CarsExcluded:getArray,
CountriesExcluded:getArray,
Steps:getArray,
ClaimSteps:getArray,
ClaimDocuments:getArray,
}


//Have a better _id function for the future
function generateID(object){
    return object.CardName.replace(/\s/g, "");  
}

//Builds Mongo Object given row of the excel sheet
//To generate more complex objects a possibility is to then parse through the objects and generate objects
//NEEDS REFACTORING FOR BETTER ERROR HANDLING
function buildObject(sheet, endCol, row){

    //Initializing Object
    const object = {};
    
    //Looping through columns
    for(let col = 0; col <= endCol; col++){

        //Getting Cell
        const cell_adress = {c:col, r:row};
        const cell = sheet[XLSX.utils.encode_cell(cell_adress)];

        
        //Getting Header
        const h_adress = {c:col, r:0};
        const header = sheet[XLSX.utils.encode_cell(h_adress)];

        //Getting correct method for the datatype
        const getValue = headerFunctions[header.v];

        //Getting the value
        if(!getValue){
            throw `functions.js: HeaderField <${header.v}> lacks equivalent method`;
        } else {
            const value = getValue(cell.v);
            object[header.v] = value; 
        }
    }

    object["_id"] = generateID(object);

    return object;
}


//Row by row creates mongo objects and uploads them to data bases
async function importData(sheet, collection){
    let row = 1;

    //Getting the amount of collumns the data has
    let a1_range = sheet['!ref'];   //Encoded range of data
    let range = XLSX.utils.decode_range(a1_range);  //Decoded range of data
    let endCol = range.e.c;         //Last collumn that includes data as an int

    //Initializing where we will store the mongo-objects
    const objectArray = [];

    while(true){

        //Checking if the end is reached
        const primary_address = {c:0, r: row};
        const primary_cell = sheet[XLSX.utils.encode_cell(primary_address)].v;
        if(primary_cell === "~END~") break;


        //Try to build mongo-object for that row, and push onto the objectArray
        try { 
            const object = buildObject(sheet, endCol, row);
            objectArray.push(object);
        } catch (err) {
            console.log(chk.red(`Object at row ${row} was not built`))
            console.log(err);
        }
        
        row++;
    }

    //Only proceeds if their are documents to be inserted
    if(objectArray.length === 0){
        console.log(chk.blue("functions.js: No documents to be inserted/updated"));
        return;
    };

    //Pushing object array to mongodb
    try {
        //Inserting array of objects
        let {result} = await collection.insertMany(objectArray,  {"orderd" : false});
        console.log(chk.green(result.insertedCount));

        //Optionally check for all rows not inserted below

    } catch (err){
        console.log(chk.red("functions.js: Error inserting documents"))
        throw err;
    }
    return;
}

const parser = {
    importData : importData,
    
    happy : ()=>{
        console.log(chk.green.bold("YAY!"));
    }

}

module.exports = parser;

