async function getCSVData(fileInput) {
    const csv = fileInput.files[0];
    const fileData = await csv.arrayBuffer();
    const data = XLSX.read(fileData, {cellDates: true});
    const worksheet = data.Sheets[data.SheetNames[0]]
    const rawData = XLSX.utils.sheet_to_json(worksheet);
    return rawData;
}
function addTask(targetObj, toIncrease) {
    if (typeof targetObj[toIncrease] !== undefined && targetObj[toIncrease]) {
        targetObj[toIncrease] += 1;
    } else {
        targetObj[toIncrease] = 1;
    }

    return targetObj;
}
function updateTasks(taskObj, task, label = '') {
    if (task['Completed At'] || task['Due Date']) {

        if (task['Due Date']) {
            doneDate = task['Due Date'];
        } else {
            task['Completed At'];
        }

        completeDate = doneDate.toLocaleDateString('en-ZA');
        firstWeekDay = new Date(doneDate
            .setDate(doneDate.getDate()
            - doneDate.getDay() ))
            .toLocaleDateString('en-ZA');

        lastWeekDay = new Date(doneDate
            .setDate(doneDate.getDate()
            - doneDate.getDay() + 6))
            .toLocaleDateString('en-ZA');
        completeWeek = firstWeekDay + ' - ' + lastWeekDay;
        completeMonth = doneDate.getFullYear() + '/' + (doneDate.toLocaleDateString('en-ZA', {month: '2-digit'}));

        if (label) {
            if (!taskObj['All']['byDay'][completeDate]) {
                taskObj['All']['byDay'][completeDate] = {};
            }
            if (!taskObj['All']['byWeek'][completeWeek]) {
                taskObj['All']['byWeek'][completeWeek] = {};
            }
            if (!taskObj['All']['byMonth'][completeMonth]) {
                taskObj['All']['byMonth'][completeMonth] = {};
            }

            taskObj['All']['byDay'][completeDate] = addTask(taskObj['All']['byDay'][completeDate], label);
            taskObj['All']['byWeek'][completeWeek] = addTask(taskObj['All']['byWeek'][completeWeek], label);
            taskObj['All']['byMonth'][completeMonth] = addTask(taskObj['All']['byMonth'][completeMonth], label);
        } else {
            taskObj['All']['byDay'] = addTask(taskObj['All']['byDay'], completeDate);
            taskObj['All']['byWeek'] = addTask(taskObj['All']['byWeek'], completeWeek);
            taskObj['All']['byMonth'] = addTask(taskObj['All']['byMonth'], completeMonth);
        }

        if (task['Projects']) {
            if(!taskObj[task['Projects']]) {
                taskObj[task['Projects']] = {
                    byDay: {},
                    byWeek: {},
                    byMonth: {}
                };
            }

            if (label) {
                if (!taskObj[task['Projects']]['byDay'][completeDate]) {
                    taskObj[task['Projects']]['byDay'][completeDate] = {};
                }
                if (!taskObj[task['Projects']]['byWeek'][completeWeek]) {
                    taskObj[task['Projects']]['byWeek'][completeWeek] = {};
                }
                if (!taskObj[task['Projects']]['byMonth'][completeMonth]) {
                    taskObj[task['Projects']]['byMonth'][completeMonth] = {};
                }

                taskObj[task['Projects']]['byDay'][completeDate] = addTask(taskObj[task['Projects']]['byDay'][completeDate], label);
                taskObj[task['Projects']]['byWeek'][completeWeek] = addTask(taskObj[task['Projects']]['byWeek'][completeWeek], label);
                taskObj[task['Projects']]['byMonth'][completeMonth] = addTask(taskObj[task['Projects']]['byMonth'][completeMonth], label);
            } else {
                taskObj[task['Projects']]['byDay'] = addTask(taskObj[task['Projects']]['byDay'], completeDate);
                taskObj[task['Projects']]['byWeek'] = addTask(taskObj[task['Projects']]['byWeek'], completeWeek);
                taskObj[task['Projects']]['byMonth'] = addTask(taskObj[task['Projects']]['byMonth'], completeMonth);
            }

        }
    }

    return taskObj;
}
async function getComparisonReport(e) {
    const loader = document.getElementById('loader');
    if (loader) {
        loader.innerHTML = 'loading...';
    }

    const firstCSV = e.target.querySelector('#comparison-csv-1');
    const firstLabel = e.target.querySelector('#comparison-label-1').value;
    const secondCSV = e.target.querySelector('#comparison-csv-2');
    const secondLabel = e.target.querySelector('#comparison-label-2').value;

    const firstData = await getCSVData(firstCSV);
    const secondData = await getCSVData(secondCSV);

    let projectTasks = {
        'All': {
            'byDay': {},
            'byWeek': {},
            'byMonth': {}
        }
    };
    let doneDate;
    let completeDate;
    let completeWeek;
    let completeMonth;
    let firstWeekDay;
    let lastWeekDay;

    const workbook = XLSX.utils.book_new();
    let daysheet;
    let weeksheet;
    let monthsheet;
    let sheetname;
    let usedsheets = [];
    let k = 1;

    for (let i = 0; i < firstData.length; i++) {
        projectTasks = updateTasks(projectTasks, firstData[i], firstLabel);
    }

    for (let j = 0; j < secondData.length; j++) {
        projectTasks = updateTasks(projectTasks, secondData[j], secondLabel);
    }

    for (const [key, value] of Object.entries(projectTasks)) {
        if (key.length > 20) {
            sheetname = key.substring(0, 20).replace(/[^a-zA-Z0-9]/g, '');
        } else {
            sheetname = key.replace(/[^a-zA-Z0-9]/g, '');
        }

        if (usedsheets.includes(sheetname)) {
            sheetname += k;
            k++;
        } else {
            usedsheets.push(sheetname);
        }

        daysheet = makeSheet(value['byDay'], 'date');
        XLSX.utils.book_append_sheet(workbook, daysheet, sheetname + " Per Day");

        weeksheet = makeSheet(value['byWeek'], 'week');
        XLSX.utils.book_append_sheet(workbook, weeksheet, sheetname + " Per Week");

        monthsheet = makeSheet(value['byMonth'], 'month');
        XLSX.utils.book_append_sheet(workbook, monthsheet, sheetname + " Per Month");
    }

    if (loader) {
        loader.innerHTML = 'finished loading';
    }

    XLSX.writeFile(workbook, "Tasks.xlsx", { compression: true });
}
async function getAsanaReport(e) {

    const loader = document.getElementById('loader');
    if (loader) {
        loader.innerHTML = 'loading...';
    }

    const csvInput = e.target.querySelector('#csv-file');
    const rawData = await getCSVData(csvInput);
    let projectTasks = {
        'All': {
            'byDay': {},
            'byWeek': {},
            'byMonth': {}
        }
    };
    let doneDate;
    let completeDate;
    let completeWeek;
    let completeMonth;
    let firstWeekDay;
    let lastWeekDay;

    for (let i = 0; i < rawData.length; i++) {
        projectTasks = updateTasks(projectTasks, rawData[i]);
    }

    const workbook = XLSX.utils.book_new();
    let daysheet;
    let weeksheet;
    let monthsheet;
    let sheetname;
    let usedsheets = [];
    let k = 1;

    for (const [key, value] of Object.entries(projectTasks)) {
        if (key.length > 20) {
            sheetname = key.substring(0, 20).replace(/[^a-zA-Z0-9]/g, '');
        } else {
            sheetname = key.replace(/[^a-zA-Z0-9]/g, '');
        }

        if (usedsheets.includes(sheetname)) {
            sheetname += k;
            k++;
        } else {
            usedsheets.push(sheetname);
        }

        daysheet = makeSheet(value['byDay'], 'date');
        XLSX.utils.book_append_sheet(workbook, daysheet, sheetname + " Per Day");

        weeksheet = makeSheet(value['byWeek'], 'week');
        XLSX.utils.book_append_sheet(workbook, weeksheet, sheetname + " Per Week");

        monthsheet = makeSheet(value['byMonth'], 'month');
        XLSX.utils.book_append_sheet(workbook, monthsheet, sheetname + " Per Month");
    }

    if (loader) {
        loader.innerHTML = 'finished loading';
    }

    XLSX.writeFile(workbook, "Tasks.xlsx", { compression: true });
}

function makeSheet(taskObj, type) {
    let dateObj;
    const taskArray = [];
    const ordered = Object.keys(taskObj).sort().reduce(
      (obj, key) => {
        obj[key] = taskObj[key];
        return obj;
      },
      {}
    );

    for (const [key, value] of Object.entries(ordered)) {
        dateObj = {};
        dateObj[type] = key;

        if (typeof value === "object") {
            for (const [subkey, subvalue] of Object.entries(value)) {
                dateObj[subkey] = subvalue;
            }
        } else {
            dateObj['number of tasks'] = value;
        }
        taskArray.push(dateObj);
    }

    const tasksheet = XLSX.utils.json_to_sheet(taskArray);
    return tasksheet;
}

Date.prototype.getWeek = function() {
  var onejan = new Date(this.getFullYear(),0,1);
  var today = new Date(this.getFullYear(),this.getMonth(),this.getDate());
  var dayOfYear = ((today - onejan + 86400000)/86400000);
  return Math.ceil(dayOfYear/7)
};

document.getElementById('asana-csv-form').addEventListener('submit', getAsanaReport);
document.getElementById('asana-comparison-form').addEventListener('submit', getComparisonReport);
