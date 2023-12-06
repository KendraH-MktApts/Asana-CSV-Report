async function getCSVData(e) {
    const loader = document.getElementById('loader');
    if (loader) {
        loader.innerHTML = 'loading...';
    }

    const csv = e.target.querySelector('#csv-file').files[0];
    const fileData = await csv.arrayBuffer();
    const data = XLSX.read(fileData, {cellDates: true});
    const worksheet = data.Sheets[data.SheetNames[0]]
    const rawData = XLSX.utils.sheet_to_json(worksheet);
    const projectTasks = {};
    const totalTasks = {
        'byDay': {},
        'byWeek': {},
        'byMonth': {}
    };
    let doneDate;
    let completeDate;
    let completeWeek;
    let completeMonth;
    let firstWeekDay;
    let lastWeekDay;

    for (let i = 0; i < rawData.length; i++) {
        if (rawData[i]['Completed At'] || rawData[i]['Due Date']) {

            if (rawData[i]['Due Date']) {
                doneDate = rawData[i]['Due Date'];
            } else {
                rawData[i]['Completed At'];
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
            completeMonth = doneDate.getFullYear() + '/' + (doneDate.toLocaleDateString('en-ZA', {month: '2-digit'}))

            if (totalTasks['byDay'][completeDate]) {
                totalTasks['byDay'][completeDate] += 1;
            } else {
                totalTasks['byDay'][completeDate] = 1;
            }

            if (totalTasks['byWeek'][completeWeek]) {
                totalTasks['byWeek'][completeWeek] += 1;
            } else {
                totalTasks['byWeek'][completeWeek] = 1;
            }

            if (totalTasks['byMonth'][completeMonth]) {
                totalTasks['byMonth'][completeMonth] += 1;
            } else {
                totalTasks['byMonth'][completeMonth] = 1;
            }

            if (rawData[i]['Projects']) {
                if(!projectTasks[rawData[i]['Projects']]) {
                    projectTasks[rawData[i]['Projects']] = {
                        byDay: {},
                        byWeek: {},
                        byMonth: {}
                    };
                }

                if (projectTasks[rawData[i]['Projects']]['byDay'][completeDate]) {
                    projectTasks[rawData[i]['Projects']]['byDay'][completeDate] += 1;
                } else {
                    projectTasks[rawData[i]['Projects']]['byDay'][completeDate] = 1;
                }

                if (projectTasks[rawData[i]['Projects']]['byWeek'][completeWeek]) {
                    projectTasks[rawData[i]['Projects']]['byWeek'][completeWeek] += 1;
                } else {
                    projectTasks[rawData[i]['Projects']]['byWeek'][completeWeek] = 1;
                }

                if (projectTasks[rawData[i]['Projects']]['byMonth'][completeMonth]) {
                    projectTasks[rawData[i]['Projects']]['byMonth'][completeMonth] += 1;
                } else {
                    projectTasks[rawData[i]['Projects']]['byMonth'][completeMonth] = 1;
                }

            }

            projectTasks['All'] = totalTasks;
        }

        if (loader) {
            loader.innerHTML = 'finished loading';
        }
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
        dateObj['number of tasks'] = value;
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

document.getElementById('asana-csv-form').addEventListener('submit', getCSVData);
