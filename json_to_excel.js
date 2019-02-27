(async function() {
    'use strict';

    const data = {};
    const names = data.names_text;
    const results = data.results;

    let rows = []; // Все строки в таблице

    let worksheetName = 'Отчёт';

    let header = {cells: [names.position, names.name, names.enter_data, names.users_new_favorite, names.users_new_intention],
                  styles: {
                      //all: {tagname: 'th', attr: ``},
                      0: {tagname: 'th', attr: ``},
                      1: {tagname: 'th', attr: ``},
                      2: {tagname: 'th', attr: `colspan="2"`},
                      3: {tagname: 'th', attr: `colspan="2"`},
                      4: {tagname: 'th', attr: `colspan="2"`},
                  }
                 };

    let header2 = {cells: ['', '', names.week, names.raising, names.week, names.raising, names.week, names.raising],
                   styles: {
                       all: {tagname: 'th', attr: ``},
                   }
                  };

    rows.push(header);
    rows.push(header2);


    Object.values(results).forEach((element, i) => {
        console.log(element);
        let row = {
            cells: [i + 1, element.name,
                    element.values.enter_data.week,
                    element.values.enter_data.raising,
                    element.values.users_new_favorite.week,
                    element.values.users_new_favorite.raising,
                    element.values.users_new_intention.week,
                    element.values.users_new_intention.raising,
                   ]
        };

        rows.push(row);
    });

// Создание .csv файла:  downloadTextAsFile('results.csv', createCSV(columns, rows));

// json to xls:
let tableToExcel = (() => {
    let dataType = 'data:application/vnd.ms-excel;base64,',
        template = '<html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel" xmlns="http://www.w3.org/TR/REC-html40"><head><!--[if gte mso 9]><xml><x:ExcelWorkbook><x:ExcelWorksheets><x:ExcelWorksheet><x:Name>{worksheet}</x:Name><x:WorksheetOptions><x:DisplayGridlines/></x:WorksheetOptions></x:ExcelWorksheet></x:ExcelWorksheets></x:ExcelWorkbook></xml><![endif]--></head><body><table>{table}</table></body></html>',

      base64 = (s) => {
          return window.btoa(unescape(encodeURIComponent(s)))
      },

      format = (s, c) => {
          return s.replace(/{(\w+)}/g, (m, p) => {
              return c[p];
          })
      }

    return (table, name) => {

      table = document.getElementById('excelTable');

      let rowsHtml = '';
      Object.values(rows).forEach((row, i) => {
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

      let ctx = {
        worksheet: worksheetName || 'Worksheet',
        table:
`<meta http-equiv="content-type" content="application/vnd.ms-excel; charset=UTF-8">
<table id="excelTable" style="width:100%" border="1">
${rowsHtml}
</table>`
      }

      console.log(ctx.table);

      window.location.href = dataType + base64(format(template, ctx))
    }
  })()

let jdf = tableToExcel();

})();


// Создать контент для CSV файла. Аргументы: столбцы, строки.
function createCSV(columns, rows){
    let fileContent = '';

    columns.forEach((column, i) => {
        fileContent += column + ';'
    });

    fileContent += '\r\n';

    rows.forEach((row, i) => {
        row.forEach((value, i) => {
            fileContent += value + ';';
        });
    });

    return fileContent;
}

// Загрузить текст как файл
function downloadTextAsFile(filename, text){
  var element = document.createElement('a');
  //element.setAttribute('href', 'data:text/plain;charset=utf-8,' + encodeURIComponent(text));
  element.setAttribute('href', 'data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64, ' + encodeURIComponent(tab_text));
  element.setAttribute('download', filename);

  element.style.display = 'none';
  document.body.appendChild(element);

  element.click();

  document.body.removeChild(element);
}
