/* xlsx.js (C) 2013-present SheetJS -- https://sheetjs.com */
const XLSX = require('xlsx-js-style');
var Mustache = require('mustache');
Mustache.escape = function(text) {return text;};

const EXTENSIONS = "xls|xlsx|xlsm|xlsb|xml|csv|txt|dif|sylk|slk|prn|ods|fods|htm|html".split("|");
const ABCS = ["A","B","C","D","E","F","G","H","I","J","K","L","M","N","O","P","Q","R","S","T","U","V","W","X","Y","Z"];

var grid = x_spreadsheet('#gridctr');

let allTasks = [];


const processWb = function(wb) {
	const HTMLOUT = document.getElementById('htmlout');
	const XPORT = document.getElementById('exportBtn');
	XPORT.disabled = false;
	HTMLOUT.innerHTML = "";
	wb.SheetNames.forEach(function(sheetName) {
		const htmlstr = XLSX.utils.sheet_to_html(wb.Sheets[sheetName],{editable:false});
		HTMLOUT.innerHTML += htmlstr;
	});
};

const readFile = function(files) {
	const f = files[0];
	const reader = new FileReader();
	reader.onload = function(e) {
		let data = e.target.result;
		data = new Uint8Array(data);
        generateTasks(XLSX.read(data, {type: 'array'}));
		//processWb(XLSX.read(data, {type: 'array'}));
	};
	reader.readAsArrayBuffer(f);
};

const generateTasks = function(wb) {
    //reset
    allTasks = [];
    console.log(wb);
    let rangeRow = wb.Sheets[wb.SheetNames[0]]['!ref'].match(/\d+/g);
    let currentRow = Number(rangeRow[0]);
    let maxRow = Number(rangeRow[1]);

    let rangeColumn = wb.Sheets[wb.SheetNames[0]]['!ref'].match(/[a-zA-Z]+/g);
    let currentColumn = ABCS.indexOf(rangeColumn[0]);
    let maxColumn = ABCS.indexOf(rangeColumn[1]);

    console.log({currentRow, maxRow, currentColumn, maxColumn});
    //wb.Sheets[wb.SheetNames[0]];

    let sheet = wb.Sheets[wb.SheetNames[0]];
    console.log(sheet);
    console.log(XLSX.utils.sheet_to_json(sheet));
    let arrayDoc = XLSX.utils.sheet_to_json(sheet);
    let previewData = {empty: true};
    let subIndex = 0;
    arrayDoc.forEach(function(row) {
        if (row.Tarea) {
            if (!previewData.empty) {
                //guardar la tarea anterior
                allTasks.push(fillTask(previewData));
            }
            subIndex = 0;
            previewData = Object.assign({}, previewData, {
                Tarea: Mustache.render(row.Tarea, row),
                tasks: "",
                task: `${row.Index}. ${Mustache.render(row.Tarea, row)}`,
                Index: row.Index,
                empty: false
            });
        }
        subIndex++;
        let fullTask = mergeTask(previewData, row);
        let currentSubtask = {
            ...fullTask,
            Tarea: Mustache.render(fullTask.Tarea || "", fullTask),
            Subtarea: Mustache.render(row.Subtarea || "", fullTask)
        };
        
        let taskTemplate = `${currentSubtask.tasks}${currentSubtask.Index}.${subIndex} ${currentSubtask.Subtarea}

`;
        currentSubtask.tasks = taskTemplate;
        console.log('currentSubtask', currentSubtask);
        console.log('previewData', previewData);
        previewData = currentSubtask;
    });
    allTasks.push(fillTask(previewData));
    
    let ws = XLSX.utils.aoa_to_sheet(allTasks);
    console.log("ws", ws);

    let wbTask = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wbTask, ws, "Tareas");
    grid.loadData(stox(wbTask));
    
    XLSX.writeFile(wbTask, "Tasks.xlsx");
    //processWb(wbTask);
}

const fillTask = function(task) {
    return [task.Espacio, task.task, task.Area || "SGT_IN_IM_MX_AOT", "Ejecutor asignado", task.tasks];
}

const handleReadBtn = async function() {
	const o = await dialog.showOpenDialog({
		title: 'Select a file',
		filters: [{
			name: "Spreadsheets",
			extensions: EXTENSIONS
		}],
		properties: ['openFile']
	});
	if(o.filePaths.length > 0) processWb(XLSX.readFile(o.filePaths[0]));
};

const exportXlsx = async function() {
	var new_wb = xtos(grid.getData());
    /* generate download */
    XLSX.writeFile(new_wb, "SheetJS.xlsx");
};

const mergeTask = function (previewTask, newTask) {
    let answer = {}

    for (key in newTask) {
        if (newTask[key] && newTask[key] !== "")
            answer[key] = newTask[key];
    }
    return {
        ...previewTask,
        ...answer
    }
}

// add event listeners
//const readBtn = document.getElementById('readBtn');
const readIn = document.getElementById('readIn');

//readBtn.addEventListener('click', handleReadBtn, false);
readIn.addEventListener('change', (e) => { readFile(e.target.files); }, false);