{
    var nickName = {};
    function upload() {
        let fileUpload = document.getElementById("file");
        let regex = /(.xls|.xlsx|.csv)$/;
        let reader = new FileReader();

        reader.onloadstart = (e=>{
            let fileUpload = document.getElementById("file");
            document.getElementById('info').innerText = 'Loadind... wait';
            document.cont.enabled = false;
            console.log(e)
        });

        reader.loaded = (e=> document.cont.enabled = true);
        if (regex.test(fileUpload.files[0].name.toLowerCase())) {
            if (typeof (FileReader) != "undefined") {
                if (reader.readAsBinaryString) {
                    reader.onload = e => {
                        render(...processExcel(e.target.result));
                    };
                    reader.readAsBinaryString(fileUpload.files[0]);
                } else {
                    reader.onload = e => {
                        let data = "";
                        let bytes = new Uint8Array(e.target.result);

                        for (let i = 0; i < bytes.byteLength; i++) {
                            data += String.fromCharCode(bytes[i]);
                        }
                        render(...processExcel(data));
                    };
                    reader.readAsArrayBuffer(fileUpload.files[0]);
                }
            } else {
                alert("This browser does not support HTML5.");
            }
        } else {
            fileUpload.value = '';
            alert("Please upload a valid Excel or Csv file.");
        }
    };


    function render(headers, data, el) {
        let data_out = [];
        let data_in = headers || [];

        let div = () => {
            let div = document.createElement('div');

            div.style.margin = '3px';
            div.style.alignItems = 'center';
            div.style.display = 'flex';
            div.style.flexDirection = 'column';
            div.style.justifyContent = 'flex=center';

            el.appendChild(div);
            return div;
        };

        let select = (data, name, nick, size) => {
            let select = document.createElement('select');

            select.size = size || '10';
            select.name = name || '';
            select.style.minWidth = '200px';
            select.style.maxWidth = '400px';
            data = data || [];
            nick = nick || {};
            select.redraw = (data,nick) => {
                select.innerHTML = '';
                data.forEach((element) => {
                    let option = document.createElement('option');

                    option.value = element;
                    option.text = nick.hasOwnProperty(element)? nick[element]: element;
                    option.title = element;
                    select.appendChild(option);
                });
            };
            select.redraw(data,nick);
            return select;
        };

        let button = (name, parent) => {
            let button = document.createElement('button');

            parent = parent || el;
            button.innerText = name;
            parent.appendChild(button);
            button.style.margin = '2px';
            return button;
        };

        el = el || document.getElementById('option_el');
        el.innerHTML = '';
        el.style.display = 'flex';
        el.style.justifyContent = 'space-between';
        document.getElementById('info').innerText = 'Select headers for report';

        let div_in = div();
        let headers_in = select(data_in, 'select_in',nickName);
        let label_in = document.createElement('label');

        label_in.for = 'select_in';
        label_in.innerText = 'Available headers';
        div_in.appendChild(label_in);
        div_in.appendChild(headers_in);

        let div_control = div();
        let btnAdd = button('ADD', div_control);
        let btnRemoveAll = button('REMOVE ALL', div_control);
        let btnRemove = button('REMOVE', div_control);
        let btnCreateXlsx = button('CREATE XLSX', div_control);
        let btnCreateOds = button('CREATE ODS', div_control);
        let btnnickName = button('Add NICK NAME', div_control);
        let div_out = div();
        let headers_out = select(data_out, 'select_out',nickName);
        let label_out = document.createElement('label');

        label_out.for = 'select_out';
        label_out.innerText = 'Selected headers';
        div_out.appendChild(label_out);
        div_out.appendChild(headers_out);

        headers_in.ondblclick = e => {
            let element = e.target.value;

            data_out.push(element);
            data_in.splice(data_in.indexOf(element), 1);
            headers_in.redraw(data_in,nickName);
            headers_out.redraw(data_out,nickName);
        };

        headers_out.ondblclick = e => {
            let element = e.target.value;

            data_in.push(element);
            data_out.splice(data_out.indexOf(element), 1);
            headers_in.redraw(data_in,nickName);
            headers_out.redraw(data_out,nickName);
        };

        btnAdd.onclick = e => {
            if (!!headers_in.value) {
                let element = headers_in.value;

                data_out.push(element);
                data_in.splice(data_in.indexOf(element), 1);
                headers_in.redraw(data_in,nickName);
                headers_out.redraw(data_out,nickName);
            }
        };

        btnRemove.onclick = e => {
            if (!!headers_out.value) {
                let element = headers_out.value;

                data_in.push(element);
                data_out.splice(data_out.indexOf(element), 1);
                headers_in.redraw(data_in,nickName);
                headers_out.redraw(data_out,nickName);
            }
        };

        btnRemoveAll.onclick = e => {
            if (data_out.length) {
                data_in.splice(0, 0, ...data_out);
                data_out.splice(0, data_out.length);
                headers_in.redraw(data_in,nickName);
                headers_out.redraw(data_out,nickName);
            }
        };

        btnCreateXlsx.onclick = e => {
            if (data_out.length) {
                createExcel(data_out, data, 'xlsx',nickName);
            } else alert('Select headers');
        };

        btnCreateOds.onclick = e => {
            if (data_out.length) {
                createExcel(data_out, data, 'ods',nickName);
            } else alert('Select headers');
        };

        btnnickName.onclick = e =>{
            let textToJSON = {};
            let input = document.createElement('input');

            input.type = 'file';
            input.onchange = e => {
                let reader = new FileReader();

                if (typeof (FileReader) != "undefined") {
                    if (reader.readAsText) {
                        reader.onload = e => {
                            reader.result.split("\n")
                                .map(val => val.split(';'))
                                .filter(val => val.length > 1)
                                .forEach(val => textToJSON[val[1].trim()] = val[0].trim());
                            Object.assign(nickName,textToJSON);
                            headers_in.redraw(data_in,nickName);
                            headers_out.redraw(data_out,nickName);
                            console.log(nickName);
                        };
                        reader.readAsText(input.files[0]);
                    } else {
                        reader.onload = e => {
                            let data = "";
                            let bytes = new Uint8Array(e.target.result);

                            for (let i = 0; i < bytes.byteLength; i++) {
                                data += String.fromCharCode(bytes[i]);
                            }
                            data.split("\n")
                                .map(val => val.split(';'))
                                .filter(val => val.length > 1)
                                .forEach(val => textToJSON[val[1].trim()] = val[0].trim());
                            Object.assign(nickName,textToJSON);
                            headers_in.redraw(data_in,nickName);
                            headers_out.redraw(data_out,nickName);
                        };

                        reader.readAsText(input.files[0]);
                    }
                } else {
                    alert("This browser does not support HTML5.");
                    input.value = '';
                }
            };
            input.click();

        };

    }
}
