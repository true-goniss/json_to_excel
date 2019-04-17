export default class JsonToExcel {

    constructor() {
        this.rows = [];
        this.worksheetName = 'Worksheet';
    }

    addRow = (row) => {
        this.rows.push(row);
    }

    download = (filename) => {
        dataType = 'data:application/vnd.ms-excel;base64,';
        template = '<html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel" xmlns="http://www.w3.org/TR/REC-html40"><head><!--[if gte mso 9]><xml><x:ExcelWorkbook><x:ExcelWorksheets><x:ExcelWorksheet><x:Name>{worksheet}</x:Name><x:WorksheetOptions><x:DisplayGridlines/></x:WorksheetOptions></x:ExcelWorksheet></x:ExcelWorksheets></x:ExcelWorkbook></xml><![endif]--></head><body><table>{table}</table></body></html>';

        let rowsHtml = '';
        Object.values(this.rows).forEach((row) => {
            rowsHtml += '<tr>';
  
            row.cells.forEach((cell, i) => {
                if (('styles' in row) && ('all' in row.styles)) { // если прописан стиль для целой строки
                    let style = row.styles.all;
                    rowsHtml += `<${style.tagname} ${style.attr}>${cell}</${style.tagname}>`;
                }
                else if (('styles' in row) && (i in row.styles)) { // если прописан стиль для отельных ячеек
                    let style = row.styles[i];
                    rowsHtml += `<${style.tagname} ${style.attr}>${cell}</${style.tagname}>`;
                } else { // если не прописано никаких стилей
                    rowsHtml += `<td>${cell}</td>`;
                }
            });
  
            rowsHtml += '</td>'
        });

        const file = dataType + base64(format(template, context(rowsHtml)));
        saveFile(file, filename + '.xls');
    }

    saveFile = (data, filename) => {
        var link = document.createElement('a');
        if (typeof link.download === 'string') {
            link.href = data;
            link.download = filename;

            //Firefox requires the link to be in the body
            document.body.appendChild(link);

            //simulate click
            link.click();

            //remove the link when done
            document.body.removeChild(link);
        } else {
            window.open(data);
        }
    }

    
    context = (rowsHtml) => {
        return {
            worksheet: worksheetName || 'Worksheet',
            table:
            `<meta http-equiv="content-type" content="application/vnd.ms-excel; charset=UTF-8">
                <table id="excelTable" style="width:100%" border="1">
                ${rowsHtml}
                </table>`
        };
    }

    base64 = (s) => {
        return window.btoa(unescape(encodeURIComponent(s)))
    }

    format = (s, c) => {
        return s.replace(/{(\w+)}/g, (m, p) => {
            return c[p];
        });
    }
}
