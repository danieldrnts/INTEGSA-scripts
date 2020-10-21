// Grupo Ya Quedó.
// Developers: Daniel Dorantes García, Valeria Gomez

// lIBRARIES
// --- Exceljs:   https://github.com/exceljs/exceljs
import { Workbook } from "exceljs";
import "exceljs/dist/exceljs.min.js";
import "./styles.css";

// Global variables
// Variables GLOBALES para la insercion de marca y submarca.
let nombre_de_marca = "";
let submarcas = [];
let eliminatedRows = [];
// Worksheet instance
const workbook = new Workbook();
workbook.creator = "Daniel";
// let workbook_cat = new Workbook();
// workbook_cat.creator = "workbook_categorias";

let inputWorksheet = document.querySelector("#excelInput");
document.querySelector("#SendButton").addEventListener("click", function () {
  nombre_de_marca = document.getElementById("distribuidor").value;
  // first we added the categories.
  // CallToMethod
  // then we add 'Submarca' and 'Fabricante'
  addFabricanteSubmarca();
});

// Method to add 'Submarca' and 'Fabricante' to all worksheets.
function addFabricanteSubmarca() {
  let index = 0;
  // Call to add Categories
  for (let k = 1; k <= workbook.worksheets.length; k++) {
    handleCategories(k);
  }
  //
  // Iterate over every workbook's worksheet.
  for (let j = 1; j <= workbook.worksheets.length; j++) {
    fill_new_columns(j, index);
    index++;
  }
  // finally, we download the worksheet.
  downloadsheet(workbook);
}

// Function para ingresar las columnas al nuevo excel
function fill_new_columns(page, pos) {
  let marca = ["Marca"];
  let submarca = ["Submarca"];

  for (
    let index = 0;
    index < workbook.getWorksheet(page).rowCount - eliminatedRows[page];
    index++
  ) {
    marca.push(nombre_de_marca);
    submarca.push(submarcas[pos]);
  }
  // Agrego la nueva columna 'newColumnValues'
  workbook.getWorksheet(page).spliceColumns(1, 0, marca, submarca);
}

// obtener el spreedsheet
function loadExcel() {
  const arryBuffer = new Response(this.files[0]).arrayBuffer();
  arryBuffer.then(function (data) {
    // worksheetGlobal = 3;
    workbook.xlsx.load(data).then(function () {
      alert("Work sheet cargado con éxito");
      workbook.worksheets.forEach((element) => {
        submarcas.push(element.name);
      });
    });
  });
}

// Methods to handle categories in spreedsheet.
function handleCategories(pageNumber) {
  // Array de categorias y posiciones
  let categories = [];
  let categoriesFilter = [];

  // get current page of workbook.
  let page = workbook.getWorksheet(pageNumber);
  let counter = 0;

  page.eachRow(function (row, rowNumber) {
    if (row.values.length === 2) {
      if (counter >= 1) {
        if (categories.length) {
          let rowItem = {};
          rowItem.name = row.values[1];
          rowItem.start = rowNumber;
          rowItem.children = 0;
          categories.push(rowItem);
          categories[categories.length - 2].children = counter - 1;
          counter = 0;
        } else {
          let rowItem = {};
          rowItem.name = row.values[1];
          rowItem.start = rowNumber;
          rowItem.children = 0;
          categories.push(rowItem);
          counter = 0;
        }
      }
    }
    counter++;

    if (rowNumber === page.rowCount) {
      categories[categories.length - 1].children = counter;
    }
    // console.log("Row: " + rowNumber + " Value: " + row.values);
  });

  categories.forEach((element) => {
    if (element.children !== 0) {
      categoriesFilter.push(element);
    }
  });

  page.eachRow(function (row, rowNumber) {
    while (row.values.length === 2) {
      page.spliceRows(rowNumber, 1);
    }
  });

  eliminatedRows[pageNumber] = categoriesFilter.length;
  assignCategories(categoriesFilter, pageNumber);

  return true;
}

// Method to assign categories to the new spreedsheet
function assignCategories(categories, pageNumber) {
  let categories_final = ["Categoría"];
  let worksheet_cat = workbook.getWorksheet(pageNumber);

  // building the Array
  for (let index = 0; index < categories.length; index++) {
    for (let i = 0; i < categories[index].children; i++) {
      categories_final.push(categories[index].name);
    }
  }

  worksheet_cat.spliceColumns(1, 0, categories_final);
  return true;
}

// Method to remove categories rows
// function removeCategories(shorksheet_remove_instance, categories) {
//   let flag = true;
//   for (let index = 0; index < 4; index++) {
//     if (flag) {
//       shorksheet_remove_instance.spliceRows(categories[index].start, 1);
//       flag = false;
//     } else {
//       shorksheet_remove_instance.spliceRows(categories[index].start - 1, 1);
//     }
//   }
//   alert("Excel Terminado");
// }

// function loadWorksheetCategories() {
//   const arryBuffer = new Response(this.files[0]).arrayBuffer();
//   arryBuffer.then(function (data) {
//     // worksheetGlobal = 3;
//     workbook_cat.xlsx.load(data).then(function () {
//       alert("Work sheet categorías loaded");
//       // mOVER ESTE METODO AL SEGUNDO PLANO.
//       // handleCategories();
//     });
//   });
// }

// download final worksheet.
function downloadsheet(workbook_instance) {
  workbook_instance.xlsx.writeBuffer().then(function (data) {
    let blob = new Blob([data], {
      type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    });
    let blobURL = URL.createObjectURL(blob);
    let link = document.querySelector("#downloadLinkId");
    link.download = "Excel_modificado.xlsx";
    link.href = blobURL;
    alert("Excel listo");
  });
}

inputWorksheet.addEventListener("change", loadExcel, false);
// inputCategorias.addEventListener("change", loadWorksheetCategories, false);
