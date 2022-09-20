//Accessing Process Button
const processFileBtn = document.getElementById("processFileBtn");
// Accessing file input
const excelFile = document.getElementById('excelFile');


// Adding event Listener to handle click evenr
processFileBtn.addEventListener("click", () => {
    processExcelFile();
});




// VARABLES
// Variables for excel file 
var excelBook;
var excelSheets;
var excelSheetNo = 0;

// varibale to strore headings of the selected sheet
var headings = [];




// Getting data in required sheet from excel file
function getSheetData(){
    // console.log(excelSheetNo);
    // console.log(excelSheets[excelSheetNo]);
    var sheetData =  XLSX.utils.sheet_to_json(excelBook.Sheets[excelSheets[excelSheetNo]], { header: 1 });
    return sheetData;
}



function processExcelFile() {
    //console.log("Processing");

    // Hiding File input area
    document.getElementById("fileInputArea").style.display = "none";

    try {
        if (checkExcelOrNot()) {
            // If input is excel file
            //console.log("Correct File...");

            readExcelFile();
        } else {
            // If file is not an excel file
            //console.log("Wrong File...");
            const errorArea = document.getElementById("errorArea");
            errorArea.style.display = 'block';
            errorArea.innerHTML = "<p>Only Excel files are allowed</p>";
            setTimeout(() => {
                location.reload();
            }, 2000);
        }
    } catch (e) {
        // If file input in empty
        const errorArea = document.getElementById("errorArea");
        errorArea.style.display = 'block';
        errorArea.innerHTML = "<p>No file Choosen</p>";
        setTimeout(() => {
            location.reload();
        }, 2000);
    }

}



// file type checker
function checkExcelOrNot() {

    if (!['application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        'application/vnd.ms-excel'
    ].includes(excelFile.files[0].type)) {

        //excelFile.value = '';

        return false;
    }
    return true;
}


// reading excel file
function readExcelFile(){

    // Creating reader for excel file
    var reader = new FileReader();

    // Reading First Excel file from the input
    reader.readAsArrayBuffer(excelFile.files[0]);

    // ````````Anonymous function````````
    // after reading full file invoking an anonymous function for fetching data from excel file
    reader.onload = (event)=>{

        // Storing reader reasult into a variable in array format
        var excelResult = new Uint8Array(reader.result);

        // Excel result book similar to excel file
        excelBook = XLSX.read(excelResult, { type: 'array' });

        // Getting excel sheets array from excel reasult book
        excelSheets = excelBook.SheetNames;

        // Show sheets dropdown
        showSheetsDropdown(excelSheets);

    }
}


// Display dropdown for selection os sheets
function showSheetsDropdown(sheets){

    // Hiding File input area
    document.getElementById("fileInputArea").style.display = "none";

    // Creating dropdown for selection
    const sheetInputArea = document.getElementById("sheetInputArea");
    var  temp = `<p>Select Sheet:</p><select id="selectSheet" class="form-select" aria-label="Default select example">`;
    for(var i=0; i<sheets.length; i++){
        temp += `<option value="${i}">${i+1}:${sheets[i]}</option>`;
    }
    temp += `</select>`; 
    // assigning dropdown to select dropdown area
    sheetInputArea.innerHTML = temp;

    // EVENT LISTENER FOR SELECT SHEET DROPDOWN
    // Accessing dropdown selection
    const selectSheet = document.getElementById("selectSheet");

    //calling show Subjects
    showSubjects();

    // Adding event Listener to select sheet dropdown
    selectSheet.addEventListener("change", () => {
        excelSheetNo = selectSheet.value;
        showSubjects();
    });
}



// Display inputs to senter passmarks for each subject
function showSubjects(){
    // getting data from required sheet in excel file
    var sheetData = getSheetData();
    
    // getting headings 
    h
    for(var i=0; i<sheetData[0].length; i++){

    }
}



/*

        //
        // Based on sheet name reading the data present in sheet into json array format
        //
        //console.log(excelSheets.length);
        //Show dropdown sheets and input marks for each subject

*/