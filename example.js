import JsonToExcel from './json_to_excel';

let excel = new JsonToExcel();

excel.worksheetName = 'Таблица';
    
excel.addRow(
  {
    cells: ['cell_1', 'cell_2', 'cell_3'],
    styles: {
        0: {tagname: 'th', attr: ``},
        1: {tagname: 'th', attr: ``},
        2: {tagname: 'th', attr: `colspan="2"`},
    }
  }
);

excel.addRow(
  {
    cells: ['1', '2', '3', '4'],
           styles: {
               all: {tagname: 'th', attr: ``},
           }
  }
);

excel.download('таблица');
