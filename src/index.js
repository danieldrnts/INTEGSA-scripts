import { Workbook } from "exceljs";
import "exceljs/dist/exceljs.min.js";
import "./styles.css";
// WORKSHEET PARA AÑADIR MARCA Y SUBMARCAS
const workbook = new Workbook();
workbook.creator = "Daniel";
// WORKSHEET PARA CATEGORIAS
let workbook_cat = new Workbook();
workbook_cat.creator = "workbook_categorias";

let inputWorksheet = document.querySelector("#myexcel");
document.querySelector("#SendButton").addEventListener("click", function () {
  getNames();
});

let inputCategorias = document.querySelector("#excelcategoria");

// Variables GLOBALES para la insercion de marca y submarca.
let nombre_de_marca = "";
let submarcas = [];

function getNames() {
  console.log("fired");
  nombre_de_marca = document.getElementById("distribuidor").value;
  var submarcas_raw = document.getElementById("fabricantes").value;
  submarcas = submarcas_raw.split(",");
  console.log(submarcas);
  handleWorksheets();
}

// Function para ingresar las columnas al nuevo excel/
function fill_new_columns(page, pos) {
  let marca = ["Marca"];
  let submarca = ["Submarca"];

  for (let index = 0; index < workbook.getWorksheet(page).rowCount; index++) {
    marca.push(nombre_de_marca);
    submarca.push(submarcas[pos]);
  }

  // Agrego la nueva columna 'newColumnValues'
  workbook.getWorksheet(page).spliceColumns(1, 0, marca, submarca);
}

// Manje los worksheets una vez asignados.
function handleWorksheets() {
  let index = 0;
  console.log("Work pages: " + workbook.worksheets.length);

  for (let j = 1; j <= workbook.worksheets.length; j++) {
    fill_new_columns(j, index);
    index++;
  }

  // Cuando termina descargo el excel
  downloadsheet();
}

// Function para descargar el xlsx como respuesta on client.
function downloadsheet(workbook_instance) {
  workbook_instance.xlsx.writeBuffer().then(function (data) {
    let blob = new Blob([data], {
      type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    });
    let blobURL = URL.createObjectURL(blob);
    let link = document.querySelector("#downloadLinkId");
    link.download = "Excel_modificado.xlsx";
    link.href = blobURL;
  });
}

// obtener el spreedsheet
function handleFile() {
  const arryBuffer = new Response(this.files[0]).arrayBuffer();
  arryBuffer.then(function (data) {
    // worksheetGlobal = 3;
    workbook.xlsx.load(data).then(function () {
      alert("Work sheet cargado con éxito");
    });
  });
}

// Methods to handle categories in spreedsheet.
function handleCategories() {
  // Array de categorias y posiciones
  let categories = [];
  let categoriesFilter = [];

  // Display the columns values
  let page_1 = workbook_cat.getWorksheet(1);
  let counter = 0;

  page_1.eachRow(function (row, rowNumber) {
    if (row.values.length === 2) {
      if (counter > 1) {
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
      // while (row.values.length === 2) {
      //   page_1.spliceRows(rowNumber, 1);
      // }
    }
    counter++;

    if (rowNumber === page_1.rowCount) {
      categories[categories.length - 1].children = counter;
    }
    // console.log("Row: " + rowNumber + " Value: " + row.values);
  });

  categories.forEach((element) => {
    if (element.children !== 0) {
      categoriesFilter.push(element);
    }
  });

  // downloadsheet(workbook_cat);
  console.log(categories);
  console.log(categoriesFilter);

  assignCategories(categoriesFilter);
}

// Method to assign categories to the new spreedsheet
function assignCategories(categories) {
  let categories_final = ["Categoría"];
  let worksheet_cat = workbook_cat.getWorksheet(1);

  // building the Array
  for (let index = 0; index < categories.length; index++) {
    for (let i = 0; i < categories[index].children; i++) {
      categories_final.push(categories[index].name);
    }
  }

  worksheet_cat.spliceColumns(1, 0, categories_final);
  console.log(categories_final);

  downloadsheet(workbook_cat);
}

// Method to remove categories rows
function removeCategories(shorksheet_remove_instance, categories) {
  let flag = true;
  for (let index = 0; index < 4; index++) {
    if (flag) {
      shorksheet_remove_instance.spliceRows(categories[index].start, 1);
      flag = false;
    } else {
      shorksheet_remove_instance.spliceRows(categories[index].start - 1, 1);
    }
  }
  alert("Excel Terminado");
  // downloadsheet(workbook_cat);
}

function loadWorksheetCategories() {
  const arryBuffer = new Response(this.files[0]).arrayBuffer();
  arryBuffer.then(function (data) {
    // worksheetGlobal = 3;
    workbook_cat.xlsx.load(data).then(function () {
      alert("Work sheet categorías loaded");
      handleCategories();
    });
  });
}

inputWorksheet.addEventListener("change", handleFile, false);
inputCategorias.addEventListener("change", loadWorksheetCategories, false);
