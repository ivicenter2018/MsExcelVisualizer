/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */


import MyFormula from './MyFormula.js';
import MyTable from './MyTable.js';

//#region Global variables
let globalTables = new Map();
let currentIndex = 0;
let currentTableIndex = 0 ;
let maxFormulaContainer = 0 ;
let maxTablaContainer = 0;
let currentTable = "";
let currentPageInterface = 1;
let currentPageTableInterface = 1;
let customColor1="#478C3B";
let tableInfoId =0;
const NM_PAGES = 10 ;

//#endregion Global variables


/*---------------------------------------------------------------------------------------------------------------------------------------
-Main function 
---------------------------------------------------------------------------------------------------------------------------------------*/
Office.onReady((info) => {
    if (info.host === Office.HostType.Excel) {



        document.getElementById("reloadButton").addEventListener("click", function () {
            location.reload();
            
        });
        document.getElementById("loadData").addEventListener("click", function () {
            
            getTables();
            

         });

        document.getElementById("findFormulasSelectedRange").addEventListener("click", function () {
            currentTable = "RangeTable";
            globalTables.set("RangeTable", new MyTable("RangeTable"));
            (async () => {
                await findFormulaBySelectedRange();
                await loadPageData("RangeTable");
                await showFormulaData("RangeTable",0);
            })();
           
           document.getElementById("cancelFindFormulasSelectedRange").hidden=false;
           document.getElementById("tablesSpace").hidden = true;
        });

        document.getElementById("generateWorksheetData").addEventListener("click", function () {
            generateWorksheetInfo();
        });

        document.getElementById("cancelFindFormulasSelectedRange").addEventListener("click", function () {
            
            //Clean the data 
            document.getElementById("formulesSpace").innerHTML="";
            document.getElementById("idFormula").textContent="";
            document.getElementById("tablesSpace").hidden = false;
            document.getElementById("cancelFindFormulasSelectedRange").hidden=true;
            document.getElementById("showSearchError").hidden=true;
         });
         document.getElementById("findFormulaByNameButton").addEventListener("click",async function () {
            
            currentTable = "RangeTable";
            globalTables.set("RangeTable", new MyTable("RangeTable"));
            
            let fName = document.getElementById("findFormulaByName").value;
            
            if(checkInputInterface(1,fName)){
                
                (async () => {
                    
                    await findFormulasByName(fName);
                    await loadPageData("RangeTable");
                    await showFormulaData("RangeTable",0);
                })();
                document.getElementById("cancelFindFormulasByName").hidden=false;
                document.getElementById("tablesSpace").hidden = true;
            }else{
                showInputError()
            }
            
         });
 
         document.getElementById("cancelFindFormulasByName").addEventListener("click", function () {
            
             //Clean the data 
            document.getElementById("formulesSpace").innerHTML="";
            document.getElementById("idFormula").textContent="";
            document.getElementById("tablesSpace").hidden = false;
            document.getElementById("showSearchError").hidden=true;
            let buttonCancelSearch= document.getElementById("cancelFindFormulasByName");
            let searchInput= document.getElementById("findFormulaByName");
            searchInput.value="";
            buttonCancelSearch.hidden=true;
             
          });

          document.getElementById("backPageBtn").addEventListener("click", async function () {

            document.getElementById("currentPageDisplay").innerHTML = await getPageIndexInterface(false,2);
            let newStartIndex = await getPageIndex(false,2);
            currentIndex = newStartIndex;
            console.log(" current Index:" + currentIndex);
            showFormulaData(currentTable,newStartIndex);
         });

         document.getElementById("nextPageBtn").addEventListener("click", async function () {
            document.getElementById("currentPageDisplay").innerHTML = await getPageIndexInterface(true,2);
            let newStartIndex = await getPageIndex(true,2);
            currentIndex = newStartIndex;
            console.log(" current Index:" + currentIndex);
            showFormulaData(currentTable,newStartIndex);
         });

         document.getElementById("backPageTableBtn").addEventListener("click", async function () {
            
            document.getElementById("currentPageTableDisplay").innerHTML = await getPageIndexInterface(false,1);
            let newStartIndex = await getPageIndex(false,1);
            currentTableIndex = newStartIndex;
            console.log(" current Table Index:" + currentTableIndex);
            showTableData(newStartIndex);
         });

         document.getElementById("nextPageTableBtn").addEventListener("click", async function () {
           
            document.getElementById("currentPageTableDisplay").innerHTML = await getPageIndexInterface(true,1);
            let newStartIndex = await getPageIndex(true,1);
            currentTableIndex = newStartIndex;
            console.log(" current Table Index:" + currentTableIndex);
            showTableData(newStartIndex);
         });
        
         document.getElementById("filterPageBtn").addEventListener("click",  async function () {
            let filterPageValue = document.getElementById("pageInputFilter").value;
            filterPageValue = parseInt(filterPageValue);
            if(checkInputInterface(2,filterPageValue)){
                if(filterPageValue > maxFormulaContainer/NM_PAGES)
                    filterPageValue = Math.max(0,maxFormulaContainer/NM_PAGES);    
                if(filterPageValue < 1)
                    filterPageValue = 1;
                let newStartIndex = (filterPageValue-1)*NM_PAGES;
                showFormulaData(currentTable,newStartIndex);
                document.getElementById("currentPageDisplay").innerHTML = filterPageValue;
                currentPageInterface = filterPageValue;
                currentIndex = newStartIndex;
            }else{
                showInputError();
            }
            
            
         });
         document.getElementById("filterTablePageBtn").addEventListener("click",  async function () {
            let filterPageValue = document.getElementById("pageInputTableFilter").value;
            filterPageValue = parseInt(filterPageValue);
            if(checkInputInterface(3,filterPageValue)){
                if(filterPageValue > maxTablaContainer/NM_PAGES)
                    filterPageValue =Math.max(0,maxTablaContainer/NM_PAGES);    
                if(filterPageValue < 1)
                    filterPageValue = 1;
                let newStartIndex = (filterPageValue-1)*NM_PAGES;
                showTableData(newStartIndex);
                document.getElementById("currentPageTableDisplay").innerHTML = filterPageValue;
                currentPageTableInterface = filterPageValue;
                currentTableIndex = newStartIndex;
            }else{
                showInputError();
            }
            
            
         });
         document.getElementById('colorButtonInterface').addEventListener('click',async function() {
            let color = document.getElementById('colorPicker').value;
            customColor1=color;
            showFormulaData(currentTable,currentIndex);
            
        });
    }
});



//#region TASK PANE FUNCTIONS

/*---------------------------------------------------------------------------------------------------------------------------------------
-This function works by getting the tables from the active worksheet.
-Does not receive parameters
-Call to more functions: getFormulasOfTable , showTableData.
---------------------------------------------------------------------------------------------------------------------------------------*/
export async function getTables() {
    try {

        await Excel.run(async (context) => {

            

            //We load the data from the tables of the active Worksheet
            const sheetAux = context.workbook.worksheets.getActiveWorksheet();
            const tablesAux = sheetAux.tables;

            tablesAux.load('items/name');

            await context.sync();
            

            if(tablesAux.items.length==0){

                document.getElementById('ms-warning').className = "ms-warningActive";

            }else{
                
                document.getElementById("loadInfo").hidden=false;
                document.getElementById("findFormulasSelectedRange").hidden=false;
                document.getElementById("searchByName").hidden=false;
                document.getElementById("initialInterface").hidden = true;
                document.getElementById("tablesSpace").innerHTML = "";
                document.getElementById("tablesSpace").hidden = false;
                document.getElementById("generateWorksheetData").hidden = false;

                
                for(const table of tablesAux.items){
                    //Store the table structure
                    let currentTable = new MyTable(table.name);
                    globalTables.set(table.name, currentTable);
    
                    await getFormulasOfTable(table.name);
                    
                    //showTableData(table.name)

                }
                maxTablaContainer = globalTables.size;
                loadPageTableData()
                showTableData(0);
                //Pre-load a AuxTable por the range filter
                let currentTable = new MyTable("RangeTable");
                globalTables.set("RangeTable", currentTable);
                
            }
            document.getElementById("loadInfo").hidden=true;

        });


    } catch (error) {
        console.error(error);
    }
}



/*---------------------------------------------------------------------------------------------------------------------------------------
-Search cell by cell in a tablefor possible formulas.Stores the position of the formula and the body.After that it searches
for the precedents of the formula.It ends by storing the data.
-Receives one parameter, the name of the table works with.
-Call another function: getAddressPrecedents.
---------------------------------------------------------------------------------------------------------------------------------------*/
export async function getFormulasOfTable(nameTable) {
    try {
        return await Excel.run(async (context) => {

            const sheetAux = context.workbook.worksheets.getActiveWorksheet();
            const tableAux = sheetAux.tables.getItem(nameTable);
            //Get the range values of the table
            let rangeTable = tableAux.getDataBodyRange();
            //Get the values and the adress property of each cell
            rangeTable.load("formulas");
            const propertiesToGet = rangeTable.getCellProperties({
                address: true
            });
            await context.sync();


            let n=1;
            for (let i = 0; i < rangeTable.formulas.length; i++) {

                for (let j = 0; j < rangeTable.formulas[i].length; j++) {
                    //When we get a cell that is a formule, we store it properly
                    if (rangeTable.formulas[i][j].toString().startsWith("=")) {
                        
                        const cellAddress = propertiesToGet.value[i][j];
                        // cellAddress returns sheet1!F5 , we get only F5
                        let addressAux = cellAddress.address.slice(cellAddress.address.lastIndexOf("!") + 1);
                        let formuleAux = rangeTable.formulas[i][j].toString();
                        let bodyAux = formuleAux.slice(1, formuleAux.length);


                        let currentFormula = new MyFormula(addressAux, bodyAux);

                        //Set precedents of the formula
                        const rangeAux = sheetAux.getRange(addressAux);
                        const precedentsAux = rangeAux.getDirectPrecedents()
                        precedentsAux.areas.load("address");
                        await context.sync();

                        for (let p = 0; p < precedentsAux.areas.items.length; p++) {

                            const precedents = await getAddressPrecedents(precedentsAux.areas.items[p].address);
                            currentFormula.setPrecedentsCells(precedents);
                            
                        }

                        //Store the formula in his table
                        let tableAux2 = globalTables.get(nameTable);
                        tableAux2.newFormula(addressAux, currentFormula);
                        document.getElementById("loadInfoMs").innerHTML= "Tablas cargadas "+ globalTables.size + ". Procesando " + n + " formulas ...";
                        n=n+1;
                        globalTables.set(nameTable, tableAux2);

                    }


                }

            }
            return 0;


        });


    } catch (error) {
        console.error(error);
    }
}

/*---------------------------------------------------------------------------------------------------------------------------------------
-This functions searches for the precedents of a range filtering the data to be returned in a more undertandable format
-Receives one parameter, a string which return the function getPrecedents.
-Does not call another aux function.
---------------------------------------------------------------------------------------------------------------------------------------*/
export async function getAddressPrecedents(outString) {
    try {
        return await Excel.run(async (context) => {

            const sheetName = context.workbook.worksheets.getActiveWorksheet();
            const addressAux = outString.split(','); // Split the string in diferent precedents
            const precedentsAux = [];

            sheetName.load("name");
            await context.sync();
            addressAux.forEach(a => {

                const elements = a.split('!');
                let newPrecedent;
                if (sheetName.name != elements[0]) {
                    newPrecedent = a;
                } else {
                    newPrecedent = elements[1]; //Store only the range without the Sheet name

                }
                precedentsAux.push(newPrecedent);

            });

            return precedentsAux;
        });
    } catch (error) {
        console.error(error);
    }
}

/*---------------------------------------------------------------------------------------------------------------------------------------
-This function is responsible for obtaining the formulas from a  selected range range .
-It receive parameters: the range over witch to obtain the cells
-Does not call another aux function: getCellsFromRange,generateFormulaGrid, generateVisGraph,
---------------------------------------------------------------------------------------------------------------------------------------*/
export async function findFormulaBySelectedRange() {
    try {
        await Excel.run(async (context) => {

            const sheetAux = context.workbook.worksheets.getActiveWorksheet();
            let rangeTable = context.workbook.getSelectedRange();
            rangeTable.load("formulas");
            const propertiesToGet = rangeTable.getCellProperties({
                address: true
            });
            await context.sync();

            let formulasNotFindedflag = true;

            for (let i = 0; i < rangeTable.formulas.length; i++) {

                for (let j = 0; j < rangeTable.formulas[i].length; j++) {
                    //When we get a cell that is a formule, we store it properly
                    if (rangeTable.formulas[i][j].toString().startsWith("=")) {
                        formulasNotFindedflag=false;
                        const cellAddress = propertiesToGet.value[i][j];
                        // cellAddress returns sheet1!F5 , we get only F5
                        let addressAux = cellAddress.address.slice(cellAddress.address.lastIndexOf("!") + 1);
                        let formuleAux = rangeTable.formulas[i][j].toString();
                        let bodyAux = formuleAux.slice(1, formuleAux.length);



                        let currentFormula = new MyFormula(addressAux, bodyAux);

                        //Set precedents of the formula
                        const rangeAux = sheetAux.getRange(addressAux);
                        const precedentsAux = rangeAux.getDirectPrecedents()
                        precedentsAux.areas.load("address");
                        await context.sync();

                        for (let p = 0; p < precedentsAux.areas.items.length; p++) {

                            const precedents = await getAddressPrecedents(precedentsAux.areas.items[p].address);
                            currentFormula.setPrecedentsCells(precedents);
                        }

                        //Store the formula in his table
                        let tableAux2 = globalTables.get("RangeTable");
                        tableAux2.newFormula(addressAux, currentFormula);
                        globalTables.set("RangeTable", tableAux2);
                       
                        
                        
                    }


                }

            }
            if(formulasNotFindedflag){
                document.getElementById("showSearchError").hidden=false;
            }
            
             
            
        });
    } catch (error) {
        console.error(error);
    }
}


/*---------------------------------------------------------------------------------------------------------------------------------------
-This function is responsible for obtaining the cells from a range .
-It receive parameters: the range over witch to obtain the cells
-Does not call another aux function: 
---------------------------------------------------------------------------------------------------------------------------------------*/
export async function getCellsFromRange(rangeSelected) {
    try {
        return await Excel.run(async function(context) {
         
          var cells = [];
          
          // Obtener el rango especificado
          let cellsRange = context.workbook.getSelectedRange();
          
          cellsRange.load("formulas");
          const propertiesToGet = cellsRange.getCellProperties({
              address: true
              });
          await context.sync();
  
  
          for (let i = 0; i < cellsRange.formulas.length; i++) {
  
              for (let j = 0; j < cellsRange.formulas[i].length; j++) {
                  //When we get a cell that is a formule, we store it properly
                  
  
                      const cellAddress = propertiesToGet.value[i][j];
                      
                      // cellAddress.address returns sheet1!F5 , we get only F5
                      let addressAux = cellAddress.address.slice(cellAddress.address.lastIndexOf("!") + 1);
                      cells.push(addressAux);
                  
              }
          }
          return cells;
       });
      } catch (error) {
        console.error(error);
      }
}

/*---------------------------------------------------------------------------------------------------------------------------------------
-This function is responsible for obtaining the formulas from a searched name .
-It receive parameters: the name of a formula
-Does not call another aux function
---------------------------------------------------------------------------------------------------------------------------------------*/
export async function findFormulasByName(fName) {
    try {
        await Excel.run(async (context) => {
            const sheetAux = context.workbook.worksheets.getActiveWorksheet();

            //Clean the data 
            document.getElementById("formulesSpace").innerHTML="";
            document.getElementById("idFormula").textContent=""; 

            let formulasNotFindedflag = true;  
            for(const [tableName, table] of globalTables){          
                if(tableName !="RangeTable"){
                    let formulasAux = table.getFormulas();
                    for(const [fNameAux, f] of formulasAux){
                        
                        
                        let addressFormula = f.getNombre();
                        let bodyFormula = f.getBody();
 
                        if((bodyFormula.includes(fName.toLowerCase())) ||(bodyFormula.includes(fName.toUpperCase()) )){ 
                            formulasNotFindedflag=false; 
                            //Store the formula in his table
                            let currentFormula = new MyFormula(addressFormula, bodyFormula);

                            //Set precedents of the formula
                            const rangeAux = sheetAux.getRange(fNameAux);
                            const precedentsAux = rangeAux.getDirectPrecedents()
                            precedentsAux.areas.load("address");
                            await context.sync();

                            for (let p = 0; p < precedentsAux.areas.items.length; p++) {

                            const precedents = await getAddressPrecedents(precedentsAux.areas.items[p].address);
                            currentFormula.setPrecedentsCells(precedents);
                            }


                            let tableAux2 = globalTables.get("RangeTable");
                            tableAux2.newFormula(addressFormula, currentFormula);
                            globalTables.set("RangeTable", tableAux2);
                            
                        }
                    };
                }
                
                
            };
            if(formulasNotFindedflag){
                document.getElementById("showSearchError").hidden=false;
            }

             
            
        });
    } catch (error) {
        console.error(error);
    }
}



/*---------------------------------------------------------------------------------------------------------------------------------------
-This function is responsible for obtaining the index to get the page .
-It receive 2 parameters : a boolean for increment (true) or decrement(false)
-Does not call another aux function
---------------------------------------------------------------------------------------------------------------------------------------*/
export async function getPageIndex(increment,type) {
    try {
        return await Excel.run(async (context) => {
            if(type == 1)// Tablas
            {
                if (increment) {
                    currentTableIndex = currentTableIndex + NM_PAGES;
                    if (currentTableIndex >= maxTablaContainer)    
                        currentTableIndex = Math.max(0, maxTablaContainer - NM_PAGES);         
                } else {
                    currentTableIndex = currentTableIndex - NM_PAGES;
                    // Asegurarse de no ser menos de 0
                    if (currentTableIndex < 0) 
                        currentTableIndex = 0; 
                }                  
                return (currentTableIndex);
            }
            else{ //Type = 2 Formulas
                if (increment) {
                    currentIndex = currentIndex + NM_PAGES;
                    if (currentIndex >= maxFormulaContainer)      
                        currentIndex = Math.max(0, maxFormulaContainer - NM_PAGES);         
                } else {
                    currentIndex = currentIndex - NM_PAGES;
                    // Asegurarse de no ser menos de 0
                    if (currentIndex < 0) 
                        currentIndex = 0; 
                }                  
                return (currentIndex);
            }                                    
        });
    } catch (error) {
        console.error(error);
    }
}


/*---------------------------------------------------------------------------------------------------------------------------------------
-This function is responsible for obtaining the index for the page interface.
-its receive parameters : a boolean for increment (true) or decrement(false)
-Does not call another aux function
---------------------------------------------------------------------------------------------------------------------------------------*/
export async function getPageIndexInterface(increment,type) {
    try {
        return await Excel.run(async (context) => {
            if(type == 1){ //Type == 1 Tables
                if (increment) {
                    currentPageTableInterface = Math.min(currentPageTableInterface + 1, Math.ceil(maxTablaContainer/NM_PAGES));               
                } else {
                    currentPageTableInterface = Math.max(currentPageTableInterface - 1, 1);        
                } 
                console.log(currentPageTableInterface) ;                   
                return (currentPageTableInterface); 
            }else{ //Type == 2 Pages
                if (increment) {
                    currentPageInterface = Math.min(currentPageInterface + 1, Math.ceil(maxFormulaContainer/NM_PAGES));               
                } else {
                    currentPageInterface = Math.max(currentPageInterface - 1, 1);        
                } 
                console.log(currentPageInterface) ;                   
                return (currentPageInterface); 
            }
                  
        });
    } catch (error) {
        console.error(error);
    }
}



//#endregion TASK PANE FUNCTIONS





//#region TASK PANE VIEW FUNCTIONS

/*---------------------------------------------------------------------------------------------------------------------------------------
-This function is responsible for generating the HTML elements that represent the tables of the active worksheet in the control panel.
-Receives as a parameter the name of the table with which to generate the data.
-Call another function:  showFormulaData.
---------------------------------------------------------------------------------------------------------------------------------------*/
export async function showTableData(startPageIndex) {
    try {
        await Excel.run(async (context) => {
            let tablesSpaceAux = document.getElementById("tablesSpace");
            tablesSpaceAux.innerHTML = '';
            await context.sync();
            let endPageIndex = Math.min(startPageIndex + NM_PAGES, maxTablaContainer);
            const keys = Array.from(globalTables.keys());
            for (let i = startPageIndex; i < endPageIndex && i < globalTables.size; i++) {
                
                
                let nameTable = keys[i];
                let newElement = document.createElement("button");
                newElement.className = "buttonLink";
                newElement.addEventListener("click", function () { 
                    currentTable = nameTable;
                    loadPageData(nameTable);
                    showFormulaData(nameTable,0);       
                });
                newElement.id = nameTable;
                newElement.innerHTML = `${nameTable}`;
                tablesSpaceAux.appendChild(newElement);
            }
           
            

        });
    } catch (error) {
        console.error(error);
    }
}

/*---------------------------------------------------------------------------------------------------------------------------------------
-This function is responsible for generating the HTML content that displays the formulas and their content of a table in the control panel
-Receives as a parameter the name of the table with which to generate the data.
-Call anotherfunction : GenrateFormulaGrid,generateVisGraph
---------------------------------------------------------------------------------------------------------------------------------------*/
export async function showFormulaData(nameTable,startPageIndex) {
    try {
        await Excel.run(async (context) => {
            //Clean the data 
            document.getElementById("formulesSpace").innerHTML="";
            document.getElementById("idFormula").textContent="";
            document.getElementById("colorPickerSpace").hidden=false;

            let tablaAux = globalTables.get(nameTable);            
            let formulasAux = tablaAux.formulas;
            
            document.getElementById("idFormula").textContent = nameTable;
            const keysArray = Array.from(formulasAux.keys());
            const valuesArray = Array.from(formulasAux.values());
            let endPageIndex = Math.min(startPageIndex + NM_PAGES, maxFormulaContainer);

            for (let i = startPageIndex; i < endPageIndex && i < keysArray.length; i++) {
                const key = keysArray[i];
                const value = valuesArray[i];
              

                generateFormulaGrid(key, value.getBody());
                generateVisGraph(key, nameTable);
            }
            
        });
    } catch (error) {
        console.error(error);
    }
}

/*---------------------------------------------------------------------------------------------------------------------------------------
-This function is responsible for generating the HTML content that displays the formula and their content of a table in the control panel
-Receives as a parameter the key (range of the formula) and the value of it
-Does not call another function  
---------------------------------------------------------------------------------------------------------------------------------------*/
export async function generateFormulaGrid(key, value) {
    let formulasSpaceAux = document.getElementById("formulesSpace");
    
    let formulaContainer = document.createElement("div");
    formulaContainer.id = key;
    formulaContainer.className = "formulaContent";
    formulaContainer.addEventListener("mouseover", function () {
        locateRange(key);
    });


    let formulaKeyContainer = document.createElement("h3");
    formulaKeyContainer.innerHTML = key;

    let formulaBodyContainer = document.createElement("p");
    formulaBodyContainer.innerHTML = value;

    formulaContainer.appendChild(formulaKeyContainer);
    formulaContainer.appendChild(formulaBodyContainer);
    formulasSpaceAux.appendChild(formulaContainer);
}
/*---------------------------------------------------------------------------------------------------------------------------------------
-This function is responsible for generating the HTML content that displays the graphs and their content of a formula in the control panel
-Receives as a parameters the name of the table and the range formula.
-Call another function : networkFunctionRange
---------------------------------------------------------------------------------------------------------------------------------------*/
export async function generateVisGraph(rangeFormula, nameTable) {
    try {
        await Excel.run(async (context) => {
            
            let nodesAux = [];
            let edgesAux = [];
            let tablaAux = globalTables.get(nameTable);
            let formulasAux = tablaAux.getFormulas();
            let currentFormula = formulasAux.get(rangeFormula);

            let firstNode = { id: rangeFormula, label: rangeFormula, color: { background: customColor1 } };
            nodesAux.push(firstNode);

            currentFormula.getPrecedentsCells().forEach((precedent) => {

                if (!nodesAux.includes(precedent)) {
                    let nodeAux, edgeAux;
                    //console.log(precedent.length);
                    for (let i = 0; i < precedent.length; i++) {

                        //console.log(precedent[i]);
                        nodeAux = {
                            id: precedent[i],
                            label: precedent[i]
                        };
                        edgeAux = {
                            from: precedent[i],
                            to: rangeFormula
                        };
                        nodesAux.push(nodeAux);
                        edgesAux.push(edgeAux);
                    }

                } else {
                   // console.log(precedent);
                    let edgeAux = {
                        from: precedent,
                        to: rangeFormula
                    };
                    edgesAux.push(edgeAux);
                }

            });
           //console.log(nodesAux);
            // Datos del grafo
            let nodes = new vis.DataSet(nodesAux);
            let edges = new vis.DataSet(edgesAux);

            //console.log(rangeFormula);
            //Make a contaner for the control panel with the HTML
            let containerFormula = document.getElementById(rangeFormula);
            let containerGraph = document.createElement("div");           
            containerGraph.classList.add("graphContainer");
            
            containerFormula.appendChild(containerGraph);

            //Create the graph using Vis.js library

            const data = { nodes: nodes, edges: edges };
            const options = {};
            const network = new vis.Network(containerGraph, data, options);

            //Add the functionality for locate range clicking on the node in the control panel
            network.on('click', function (params) {
                if (params.nodes.length > 0) {
        
                    let nodeId = params.nodes[0];
                    let node = nodes.get(nodeId); 
                    locateRange(node.id);
                    
                }
            });

        });


    } catch (error) {
        console.error(error);
    }
}


/*---------------------------------------------------------------------------------------------------------------------------------------
-This function is responsible for indicating in the Excel sheet the cell that is being viewed in the control panel.
-It receives as a parameter the range that it must indicate
-Does not call another aux function
---------------------------------------------------------------------------------------------------------------------------------------*/
export async function locateRange(range) {
    try {
        await Excel.run(async (context) => {

            const currentSheet = context.workbook.worksheets.getActiveWorksheet();
            const rangeAux = currentSheet.getRange(range);

            rangeAux.select();
            await context.sync();
        });
    } catch (error) {
        console.error(error);
    }
}

/*---------------------------------------------------------------------------------------------------------------------------------------
-This function is responsible for load the values for the control panel page .
-It receives as a parameter a nameTable 
-Does not call another aux function
---------------------------------------------------------------------------------------------------------------------------------------*/
export async function loadPageData(nameTable){
    
    try {
        await Excel.run(async (context) => {
            
            document.getElementById("pageControl").hidden=false;
            let tablaAux = globalTables.get(nameTable); 
            maxFormulaContainer=tablaAux.getSize();
            document.getElementById("currentPageDisplay").innerHTML = currentPageInterface;
            document.getElementById("totalPagesDisplay").innerHTML = "MAX : " + Math.max(Math.ceil(maxFormulaContainer/NM_PAGES),1);
            
        });
    } catch (error) {
        console.error(error);
    }
}

/*---------------------------------------------------------------------------------------------------------------------------------------
-This function is responsible for load the values for the control panel page for tables.
-It not receives  a parameter 
-Does not call another aux function
---------------------------------------------------------------------------------------------------------------------------------------*/
export async function loadPageTableData(){
    
    try {
        await Excel.run(async (context) => {
            document.getElementById("pageTableControl").hidden=false;
            maxTablaContainer=globalTables.size - 1;
            document.getElementById("currentPageTableDisplay").innerHTML = currentPageTableInterface;
            document.getElementById("totalPagesTableDisplay").innerHTML = "MAX : " + Math.max(Math.ceil(maxTablaContainer/NM_PAGES),1);
            
        });
    } catch (error) {
        console.error(error);
    }
}


/*---------------------------------------------------------------------------------------------------------------------------------------
-This function is responsible of checking the input interface.
-Does not receive parameters : a number for use the correct check
-Does not call another aux function
---------------------------------------------------------------------------------------------------------------------------------------*/
export async function checkInputInterface(typeCheck,value) {
    try {
        return await Excel.run(async (context) => {
            
            switch(typeCheck){
                case 1:
                    if(typeof value != "string"){
                        console.log("Error");
                        return false;
                    }
                    break;
                case 2:
                    if(typeof value == NaN || typeof value != "number" || value > Math.ceil(maxFormulaContainer/NM_PAGES)){
                        console.log("Error");
                        return false;
                    }
                    break;
                case 3:
                    if(typeof value == NaN || typeof value != "number" || value > Math.ceil(maxTablaContainer/NM_PAGES)){
                        console.log("Error");
                        return false;
                    }
                    break;

            }
           return true;
            
        });
    } catch (error) {
        console.error(error);
    }
}
//#endregion



//#region GetInfo
/*---------------------------------------------------------------------------------------------------------------------------------------
-This function works by check if a new worksheet name exist.
-Does  receive parameters: 
-Does not call another aux function
---------------------------------------------------------------------------------------------------------------------------------------*/

export async function doesWorksheetExist(context, sheetName) {
    const sheets = context.workbook.worksheets;
    sheets.load("items/name");
    await context.sync();

    // Buscar si alguna hoja tiene el mismo nombre
    for (const sheet of sheets.items) {
        if (sheet.name === sheetName) {
            return true; // La hoja ya existe
        }
    }

    return false; // No existe
}

/*---------------------------------------------------------------------------------------------------------------------------------------
-This function works by generating a new worksheet with the data obtained.
-Does not receive parameters
-Call another aux function : doesWorksheetExist()
---------------------------------------------------------------------------------------------------------------------------------------*/
export async function generateWorksheetInfo() {
    try {
        await Excel.run(async (context) => {
            
            let worksheetNameAux = "FormulaViewerData_" + tableInfoId;
            tableInfoId += 1;

            // Verifica si la hoja ya existe
            while (await doesWorksheetExist(context, worksheetNameAux)) {
                worksheetNameAux = "FormulaViewerData_" + tableInfoId;
                tableInfoId += 1;
            }

            const worksheet = context.workbook.worksheets.add(worksheetNameAux);
            await context.sync(); // Sincroniza para asegurarse de que la hoja está creada

            let auxRange = 1; 

            for (const [tableName, table] of globalTables) {
                if(tableName != "RangeTable"){
                    let startRow = auxRange + 1; 
                let startRange = `B${startRow}:C${startRow}`;
                //let startRangeChart = `E${startRow}:F${startRow}`;  

                // Crear la tabla en el rango especificado
                let tableData = worksheet.tables.add(startRange, true);  // hasHeaders = true
                // Crear grafico     
                //const chart = sheet.charts.add(Excel.ChartType.pie, dataRange, Excel.ChartSeriesBy.columns);
        

                await context.sync();  // Sincroniza antes de asignar el nombre para asegurarse de que tableData está listo
                
                // Verificar que el nombre de la tabla sea único y válido
                let nombreTabla = tableName + "_Aux" + tableInfoId;
                console.log(nombreTabla);
                // Asignar el nombre único a la tabla
                tableData.name = nombreTabla;
                
                // Asignar los encabezados
                tableData.getHeaderRowRange().values = [["Cell", "Formula"]];

                // Preparar los datos de las fórmulas
                let formulaRows = [];
                for (const [formulaId, formula] of table.formulas) {
                    formulaRows.push([formulaId, formula.body]);
                }
                
                // Agregar los datos de las fórmulas
                tableData.rows.add(null, formulaRows);

                auxRange += formulaRows.length + 2;

                // Auto-ajustar columnas y filas si la versión de la API lo soporta
                if (Office.context.requirements.isSetSupported("ExcelApi", "1.2")) {
                    worksheet.getUsedRange().format.autofitColumns();
                    worksheet.getUsedRange().format.autofitRows();
                }

                worksheet.activate();
                await context.sync();  // Sincroniza los cambios
                }
                
            }
        });
    } catch (error) {
        console.error(error);
    }
}


/*---------------------------------------------------------------------------------------------------------------------------------------
-This function is responsible for show the error input.
-Does not receives  a parameter
-Does not call another aux function
---------------------------------------------------------------------------------------------------------------------------------------*/
export async function showInputError(){
    
    try {
        await Excel.run(async (context) => {
            alert("Error al filtrar");
            
        });
    } catch (error) {
        console.error(error);
    }
}

