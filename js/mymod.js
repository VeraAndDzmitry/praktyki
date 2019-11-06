{
    let depends = [
        "js/depend/jszip.js",
        "js/depend/xlsx.js",
        "js/depend/fileSaver.js"
    ];

    depends.forEach(name => {
        let script = document.createElement('script');
        script.src = name;
        document.body.appendChild(script);
    });

    function processExcel(data) {
        let workbook = XLSX.read(data, {
            type: 'binary'
        });
        let firstSheet = workbook.SheetNames[0];
        let excelRows = XLSX.utils.sheet_to_row_object_array(workbook.Sheets[firstSheet]);
        let headers = Object.keys(excelRows[0]);
        return [headers, excelRows];
    };

    function createExcel(headers, data, type, nik = []) {
        console.clear();
        let replacement = {};
        let result = [];
        let head = headers.map(val => nik[val] || val);
        result.push(head);
        let regExp = /^[^\\]+?\\[^\\]*?$/gm;
        data.forEach((el,data_index) => {
            let temp = [];
            headers.forEach(head => {

                let arr = el[head] === undefined ? null: (el[head]+'').match(regExp);
                if(arr) {
                    if (!replacement.hasOwnProperty(head))
                        replacement[head] = {};
                    if (!replacement[head].hasOwnProperty(data_index))
                        replacement[head][data_index] = {};
                    arr.forEach(value => {
                        let newHeader = value.trim().split('\\')[0];
                        let newValue = value.trim().split('\\')[1];

                            if(replacement[head][data_index].hasOwnProperty(newHeader)){
                                (replacement[head])[data_index][newHeader] += '\n' + newValue;
                            } else (replacement[head])[data_index][newHeader] = newValue;
                        }
                    );
                    temp.push([]);
                } else temp.push(el[head]);
            });
            result.push(temp);
        });

        for (let header in replacement) {
            let temp = [];
            temp[0] = [];
            for (let i = 1; i < data.length; i++) {
                if (!replacement[header].hasOwnProperty(i-1)) continue;
                temp[i] = new Array(temp[0].length);
                for (let head in replacement[header][i-1]) {
                    let index = temp[0].indexOf(head);
                    if (!~index){
                        temp[0].push(head);
                        temp[i].push(replacement[header][i-1][head]);
                    } else
                    if (replacement[header][i-1].hasOwnProperty(head))
                        temp[i][index] = replacement[header][i-1][head];
                }
            }
            result[0].push(...temp[0]);
            for (let i = 1; i < data.length ; i++) {
                result[i] = result[i].concat(temp[i],
                    (new Array(temp[0].length-
                        [temp[i]].length)));
            }
        }
        let ws = XLSX.utils.aoa_to_sheet(result);
        let wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, 'report');
        XLSX.writeFile(wb, 'report.' + (type || 'xlsx')) ||
        XLSX.write(wb, {bookType: type, bookSST: true, type: 'base64'});

    };
}