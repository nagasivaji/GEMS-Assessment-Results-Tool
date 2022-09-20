// //Accessing Process Button
// const processFileBtn = document.getElementById("processFileBtn");
// // Accessing file input
// const excelFile = document.getElementById('excelFile');



// // Adding event Listener to handle click evenr
// processFileBtn.addEventListener("click", () =>{
//     processExcelFile();
// });



// // Variables



// function processExcelFile(){
//     //console.log("Processing");

//     try {
//         if(checkExcelOrNot()){
//             console.log("Correct File...");
//         }else{
//             console.log("Wrong File...");
//             const errorArea = document.getElementById("errorArea");
//             errorArea.style.display = 'block';
//             errorArea.innerHTML = "<p>Only Excel files are allowed</p>";
//             setTimeout(() => {
//                 location.reload();
//             }, 2000);
//         }
//     } catch (e) {
//         const errorArea = document.getElementById("errorArea");
//         errorArea.style.display = 'block';
//         errorArea.innerHTML = "<p>No file Choosen</p>";
//         setTimeout(() => {
//             location.reload();
//         }, 2000);
//     }
    
// }



// // file type checker
// function checkExcelOrNot() {

//     if (!['application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
//             'application/vnd.ms-excel'
//         ].includes(excelFile.files[0].type)) {

//         //excelFile.value = '';

//         return false;
//     }
//     return true;
// }
