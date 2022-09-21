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

// array to strore headings of the selected sheet
var headings = [];

// array to strore pass marks for different subjects
var passMarks = [];

// array to strore students data
var students = [];



// Getting data in required sheet from excel file
function getSheetData() {
    // console.log(excelSheetNo);
    // console.log(excelSheets[excelSheetNo]);
    var sheetData = XLSX.utils.sheet_to_json(excelBook.Sheets[excelSheets[excelSheetNo]], { header: 1 });
    return sheetData;
}



function processExcelFile() {
    //console.log("Processing");

    // Hiding File input area
    document.getElementById("fileInputArea").style.display = "none";

    //showing backbtn on navigationbar
    document.getElementById("navBarBackBtn").style.display = "block";

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
function readExcelFile() {

    // Creating reader for excel file
    var reader = new FileReader();

    // Reading First Excel file from the input
    reader.readAsArrayBuffer(excelFile.files[0]);

    // ````````Anonymous function````````
    // after reading full file invoking an anonymous function for fetching data from excel file
    reader.onload = (event) => {

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
function showSheetsDropdown(sheets) {

    // Hiding File input area
    document.getElementById("fileInputArea").style.display = "none";

    // Creating dropdown for selection
    const sheetInputArea = document.getElementById("sheetInputArea");
    var temp = `<p>Select Sheet:</p><select id="selectSheet" class="form-select" aria-label="Default select example">`;
    for (var i = 0; i < sheets.length; i++) {
        temp += `<option value="${i}">${i + 1}:${sheets[i]}</option>`;
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
function showSubjects() {
    // getting data from required sheet in excel file
    var sheetData = getSheetData();

    // getting headings  and creating input fields for different subjects
    // creating new headings for different subjects
    headings = [];
    var temp = ``;
    for (var i = 0; i < sheetData[0].length; i++) {
        if (i < 3)
            headings.push(sheetData[0][i]);
        else {
            if (i % 2 == 1) {
                headings.push(sheetData[0][i]);
                temp += `
                        <div class="input-group mb-3">
                            <span class="input-group-text" id="inputGroup-sizing-default">MARKS</span>
                            <input type="number" class="form-control inputMarksForSubject" aria-label="Sizing example input" aria-describedby="inputGroup-sizing-default" min="0" max="100" placeholder="${sheetData[0][i]}">
                        </div>` ;
            }
        }
    }

    temp += `<button class="btn btn-dark" id="claculateBtn">Calculate</button>`;
    // Appending inputs to subject area
    document.getElementById('subjectsArea').innerHTML = temp;

    //Calculate
    const claculateBtn = document.getElementById('claculateBtn');
    claculateBtn.addEventListener("click", () => {
        getPassMarks();
        getStudentsObjects();
        prepareHTMLTable(students);
        showSearchArea();
    });
}


// Getting pass marks from subjects input area
function getPassMarks() {
    // creating new array for passing marks
    passMarks = [];
    var temp = document.getElementsByClassName("inputMarksForSubject");
    for (var i = 0; i < temp.length; i++) {
        //console.log(Number(temp[i].value));
        passMarks.push(Number(temp[i].value));
    }
    //console.log(passMarks);

}


// Students class
class Student {
    studentId;
    studentName;
    studentEmail;
    studentMarks;
    objectiveStatus;
    subjectiveStatus;
    finalStatus;
    finalMarks;

    constructor(studentId, studentName, studentEmail, studentMarks, objectiveStatus, subjectiveStatus, finalStatus, finalMarks) {
        this.studentId = studentId;
        this.studentName = studentName;
        this.studentEmail = studentEmail;
        this.studentMarks = studentMarks;
        this.objectiveStatus = objectiveStatus;
        this.subjectiveStatus = subjectiveStatus;
        this.finalStatus = finalStatus;
        this.finalMarks = finalMarks;
    }
}

// Subjects class
class Subject {
    subjectName;
    subjectMarks;
    subjectPassMark;
    subjectAttempts;
    subjectStatus;

    constructor(subjectName, subjectMarks, subjectPassMark, subjectAttempts, subjectStatus) {
        this.subjectName = subjectName;
        this.subjectMarks = subjectMarks;
        this.subjectPassMark = subjectPassMark;
        this.subjectAttempts = subjectAttempts;
        this.subjectStatus = subjectStatus;
    }
}



function getStudentsObjects() {

    // creating new students array
    students = [];

    // getting data from required sheet in excel file
    var sheetData = getSheetData();
    //console.log(sheetData);

    // getting noOfRows and noOfColumns in excel file sheet
    const noOfRows = sheetData.length;
    const noOfCols = sheetData[0].length;

    // Iterating from second row since first row contains headings
    // For each students...
    for (var i = 1; i < noOfRows; i++) {
        // Creating required variables which will be used while creating objects
        var studentId, studentName, studentEmail;
        var subjectName, subjectMarks, subjectAttempts, subjectStatus;
        var studentMarks = [];

        // temp variable is indexing variable for headings array
        var temp1 = 3;
        // temp variable is indexing variable for headings array
        var temp2 = 0;

        for (var j = 0; j < noOfCols; j++) {
            if (j == 0)        // col:1 - studens id 
                studentId = sheetData[i][j];
            else if (j == 1)   // col:2  - student name
                studentName = sheetData[i][j];
            else if (j == 2)   // col:3  - student email
                studentEmail = sheetData[i][j];
            else            // col:3 to ...  - subjects
            {
                subjectName = headings[temp1++];     // subject name
                subjectMarks = sheetData[i][j];     // subject marks
                j++;
                subjectAttempts = sheetData[i][j];  // subject attemps

                // function to validate students subject status
                subjectStatus = validate(subjectMarks, subjectAttempts, passMarks[temp2]);

                // creating new subject object for the student
                var obj = new Subject(subjectName, subjectMarks, passMarks[temp2++], subjectAttempts, subjectStatus);

                // pushing the new subject into student marks
                studentMarks.push(obj);
            }
        }

        // required varibale for students status
        var objectiveStatus = "PASS", subjectiveStatus = "PASS";

        // Iterating through  students sbjets array
        for (var k = 0; k < studentMarks.length; k++) {

            // getting subject object
            var subject = studentMarks[k];

            // if its a last column then its subjective type
            if (k == studentMarks.length - 1) {
                if (subject.subjectStatus == 'FAIL') {
                    subjectiveStatus = "FAIL";
                    break;
                }

                if (subject.subjectStatus == 'PENDING') {
                    subjectiveStatus = "PENDING";
                }

                if (subject.subjectStatus == 'NA') {
                    subjectiveStatus = "NOT APPEARED";
                }

            }
            else    // if its not a last column then its objective type
            {
                if (subject.subjectStatus == 'FAIL') {
                    objectiveStatus = "FAIL";
                    break;
                }

                if (subject.subjectStatus == 'PENDING') {
                    objectiveStatus = "PENDING";
                }

                if (subject.subjectStatus == 'NA') {
                    objectiveStatus = "NOT APPEARED";
                }

            }
        }


        // Conditions for finalstatus and sinal marks
        // required varibale for students status
        var finalStatus, finalMarks;

        // Condtion for student to FAIL
        if (objectiveStatus == 'FAIL' || subjectiveStatus == 'FAIL') {
            finalStatus = 'FAILED';
            finalMarks = '--';
        }
        // Condtion for student to NOT APPEARED
        else if (objectiveStatus == 'NOT APPEARED' || subjectiveStatus == 'NOT APPEARED') {
            finalStatus = 'NOT APPEARED';
            finalMarks = '--';
        }
        // Condtion for student to PENDING
        else if (objectiveStatus == 'PENDING' || subjectiveStatus == 'PENDING') {
            finalStatus = 'PENDING';
            finalMarks = '--';
        }
        // Conditions for student to PASS
        else {

            finalStatus = 'PASS';
            var objMarks = 0, subMarks = 0;

            // Iterating through  students sbjets array
            for (var k = 0; k < studentMarks.length; k++) {

                // getting subject object
                var subject = studentMarks[k];

                // then its subjective
                if (k == studentMarks.length - 1) {
                    subMarks = subject.subjectMarks;
                }
                else // then its objective
                {
                    objMarks += subject.subjectMarks;
                }
            }

            finalMarks = ((objMarks / (studentMarks.length - 1)) + (subMarks)) / 2;
            finalMarks = Math.round(finalMarks);
        }




        // creating new studemt object
        var obj = new Student(studentId, studentName, studentEmail, studentMarks, objectiveStatus, subjectiveStatus, finalStatus, finalMarks);

        // pushing the new student object to the list of students
        students.push(obj);
    }
    //console.log(students);
}


// Marks validation table
function validate(marks, attempts, criteria) {
    if (marks >= criteria && attempts <= 3) {
        return "PASS";
    }
    else if (marks < criteria && attempts < 3) {
        return "PENDING";
    }
    else if (marks < criteria && attempts == 3)
        return "FAIL";
    else
        return "NA";
}


// preparing html tbale
function prepareHTMLTable() {

    // Hiding sheetarea -  dropdown and  subjects input area - passmarks inputs
    document.getElementById("subjectsArea").style.display = "none";
    document.getElementById("sheetInputArea").style.display = "none";

    
    //creating HTML table based on the result of students data
    //Table start
    var tableStart = `<table class="table" id="htmlResultTable">`;



    //table Headings -> id, name, mail, subjects...
    // Openinig row
    var tableHeadings = `<tr>`;
    for (var i = 0; i < headings.length; i++) {
        if(i<3)
            tableHeadings = tableHeadings + `<th>${headings[i]}</th>`;
        else
            tableHeadings = tableHeadings + `<th>${headings[i]} <br> ${passMarks[i-3]}</th>`;
    }
    
    // Objective status
    tableHeadings = tableHeadings + `<th>Objective Status</th>`;
    // Subjective status
    tableHeadings = tableHeadings + `<th>Subjective Status</th>`;
    // Final status
    tableHeadings = tableHeadings + `<th>Final Status</th>`;
    // Finak Marks
    tableHeadings = tableHeadings + `<th>Final Marks</th>`;
    // Closing row
    tableHeadings = tableHeadings + `</tr>`;



    //console.log(students);
    // table rows 
    // All students data
    var tableRows = ``;
    for(var i=0; i<students.length; i++){
        // creating row tag based on status of the student
        var tempRow = ``;
        if(students[i].finalStatus === 'PASS')
            tempRow = tempRow + `<tr class="passRow">`;
        else if(students[i].finalStatus === 'FAILED')
            tempRow = tempRow + `<tr class="failRow">`;
        else if(students[i].finalStatus === 'PENDING')
            tempRow = tempRow + `<tr class="pendingRow">`;
        else
            tempRow = tempRow + `<tr class="notAppearedRow">`;


        // creating col tags for student data
        // 1.Student ID
        tempRow = tempRow + `<td>${students[i].studentId}</td>`;
        // 2.Student Name
        tempRow = tempRow + `<td>${students[i].studentName}</td>`;
        // 3.Student Email
        tempRow = tempRow + `<td>${students[i].studentEmail}</td>`;

        // 4.Student marks and attempts for each Subject (Adding dynamically through looping subjects array from student object)
        for(var j=0; j<students[i].studentMarks.length; j++){
            tempRow = tempRow + `<td>
                                    ${students[i].studentMarks[j].subjectMarks}
                                    -
                                    ${students[i].studentMarks[j].subjectAttempts}
                                </td>`;
        }

        // 5.Student Objective Status
        tempRow = tempRow + `<td>${students[i].objectiveStatus}</td>`;

        // 6.Student Subjective Status
        tempRow = tempRow + `<td>${students[i].subjectiveStatus}</td>`;

        // 7.Student Final Status
        tempRow = tempRow + `<td>${students[i].finalStatus}</td>`;

        // 7.Student Final Marks
        tempRow = tempRow + `<td>${students[i].finalMarks}</td>`;

        // closing row tag
        tempRow = tempRow + '</tr>';


        // Adding each temporary student row to table rows string 
        tableRows = tableRows + tempRow;
    }


    // table end
    var tableEnd = `</table>`;

    // download Button 
    var downloadButton = '<button class="btn btn-dark" id="downloadExcel">Get Excel</button>';

    // Adding table to HTML page
    document.getElementById("htmlTableArea").innerHTML = tableStart + tableHeadings + tableRows + tableEnd + downloadButton;
    document.getElementById("htmlTableArea").style.overflow = 'scroll';


    // Downloading excelfile using event listner
    document.getElementById("downloadExcel").addEventListener("click", function(){
        downloadExcelFile();
    });

}



// Show serch Area
function showSearchArea(){
    // showing serch input area
    document.getElementById("searchInputhArea").style.display = 'flex';

    // variable for type of search
    var serchField = 1;

    // Accessing serch selection
    const searchSelection = document.getElementById("searchSelection");

    // Adding event handlers to select serch type
    searchSelection.addEventListener("change", () => {
        serchField = searchSelection.value;
        //console.log(serchField);
    });

    // Accessing serch value
    const searchInput = document.getElementById("searchInput");

    // serch data
    var searchData = '';

    // Adding event handler to serch value
    searchInput.addEventListener("input", () => {
        searchData = searchInput.value;
        sortStudentData(serchField, searchData);
    });

    // Accessing serch btn
    const serchBtn = document.getElementById("serchBtn");

    // adding event handler to serch btn
    serchBtn.addEventListener("click", () => {
        sortStudentData(serchField, searchData.toLowerCase());
    });

    
}

// Sorting students data to show in serch result
function sortStudentData(field, data){
    if(data === '')
        prepareHTMLSerchTable([]);

    var tempStudents = [];
    for(var i=0; i<students.length; i++){
        if(field == 1){
            var tempData = students[i].studentId;
            if(tempData == data)
                tempStudents.push(students[i]);
        }
        else if(field == 2){
            var tempData = students[i].studentName.toLowerCase();
            if(tempData.includes(data))
                tempStudents.push(students[i]);
        }
        else{
            var tempData = students[i].studentEmail.toLowerCase();
            if(tempData.includes(data))
                tempStudents.push(students[i]);
        }
    }

    //console.log(tempStudents);
    prepareHTMLSerchTable(tempStudents);
}


// preparing html tbale for serch result
function prepareHTMLSerchTable(serchStudents) {
    //console.log(serchStudents);
    //creating HTML table based on the result of students data
    //Table start
    var tableDivHeading = '<p>Serch Result:</p>';
    var tableStart = `<table class="table" id="htmlSearchTable">`;



    //table Headings -> id, name, mail, subjects...
    // Openinig row
    var tableHeadings = `<tr>`;
    for (var i = 0; i < headings.length; i++) {
        if(i<3)
            tableHeadings = tableHeadings + `<th>${headings[i]}</th>`;
        else
            tableHeadings = tableHeadings + `<th>${headings[i]} <br> ${passMarks[i-3]}</th>`;
    }
    
    // Objective status
    tableHeadings = tableHeadings + `<th>Objective Status</th>`;
    // Subjective status
    tableHeadings = tableHeadings + `<th>Subjective Status</th>`;
    // Final status
    tableHeadings = tableHeadings + `<th>Final Status</th>`;
    // Finak Marks
    tableHeadings = tableHeadings + `<th>Final Marks</th>`;
    // Closing row
    tableHeadings = tableHeadings + `</tr>`;



    //console.log(students);
    // table rows 
    // All students data
    var tableRows = ``;
    for(var i=0; i<serchStudents.length; i++){
        // creating row tag based on status of the student
        var tempRow = ``;
        if(serchStudents[i].finalStatus === 'PASS')
            tempRow = tempRow + `<tr class="passRow">`;
        else if(serchStudents[i].finalStatus === 'FAILED')
            tempRow = tempRow + `<tr class="failRow">`;
        else if(serchStudents[i].finalStatus === 'PENDING')
            tempRow = tempRow + `<tr class="pendingRow">`;
        else
            tempRow = tempRow + `<tr class="notAppearedRow">`;


        // creating col tags for student data
        // 1.Student ID
        tempRow = tempRow + `<td>${serchStudents[i].studentId}</td>`;
        // 2.Student Name
        tempRow = tempRow + `<td>${serchStudents[i].studentName}</td>`;
        // 3.Student Email
        tempRow = tempRow + `<td>${serchStudents[i].studentEmail}</td>`;

        // 4.Student marks and attempts for each Subject (Adding dynamically through looping subjects array from student object)
        for(var j=0; j<serchStudents[i].studentMarks.length; j++){
            tempRow = tempRow + `<td>
                                    ${serchStudents[i].studentMarks[j].subjectMarks}
                                    -
                                    ${serchStudents[i].studentMarks[j].subjectAttempts}
                                </td>`;
        }

        // 5.Student Objective Status
        tempRow = tempRow + `<td>${serchStudents[i].objectiveStatus}</td>`;

        // 6.Student Subjective Status
        tempRow = tempRow + `<td>${serchStudents[i].subjectiveStatus}</td>`;

        // 7.Student Final Status
        tempRow = tempRow + `<td>${serchStudents[i].finalStatus}</td>`;

        // 7.Student Final Marks
        tempRow = tempRow + `<td>${serchStudents[i].finalMarks}</td>`;

        // closing row tag
        tempRow = tempRow + '</tr>';


        // Adding each temporary student row to table rows string 
        tableRows = tableRows + tempRow;
    }


    // table end
    var tableEnd = `</table>`;

    if(serchStudents.length <= 0 || serchStudents.length == students.length) {
        // Adding table to HTML page
        document.getElementById("searchResultArea").style.display = "none";
    }
    else{
        // Adding table to HTML page
        document.getElementById("searchResultArea").style.display = "block";
        document.getElementById("searchResultArea").innerHTML = tableDivHeading + tableStart + tableHeadings + tableRows + tableEnd;
        document.getElementById("searchResultArea").style.overflow = 'scroll';
    }
    
}


// Downloading excelfile
function downloadExcelFile(){
    
    function html_table_to_excel(type) {
        var data = document.getElementById('htmlResultTable');

        var file = XLSX.utils.table_to_book(data, { sheet: "sheet1" });
        XLSX.write(file, { bookType: type, bookSST: true, type: 'base64' });
        XLSX.writeFile(file, 'Exco Result.' + type);

        setTimeout(() => {
            location.reload();
        }, 1000);
    }

    html_table_to_excel('xlsx');
}

/*

        

*/