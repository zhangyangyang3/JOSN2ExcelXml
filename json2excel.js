/**
 * json 对象转换成excel文件
 * Created by zhangyangyang on 2016/5/24.
 */

$J2E = {};
$J2E.xmlHead = '<?xml version="1.0"?>';
$J2E.applicationHead = '<?mso-application progid="Excel.Sheet"?>';
$J2E.createWorkBook = function () {
    var workBook = {};
    workBook.head='<Workbook '+
        'xmlns:x="urn:schemas-microsoft-com:office:excel" '+
        'xmlns="urn:schemas-microsoft-com:office:spreadsheet" '+
        'xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet">';
    workBook.tail='</Workbook>';
    workBook.workSheet = [];
    workBook.addWorkSheet = function(sheet) {
        workBook.workSheet.push(sheet);
    }
    workBook.createWorksheet = function (){
        var worksheet = {};
        worksheet.name='Sheet1';
        worksheet.setName=function setSheetName(name) {
            worksheet.name = name;
        }
        worksheet.head='<Worksheet ss:Name="'+ worksheet.name +'">  <ss:Table>';
        worksheet.tail ='  </ss:Table></Worksheet>';
        worksheet.Rows=[];
        worksheet.createRow = function (){
            var row={};
            row.cells = [];
            row.head='<ss:Row>';
            row.tail='</ss:Row>';
            row.addCell = function(val){
                var cell = {
                    head: '<ss:Cell><Data ss:Type="String">',
                    tail:'</Data></ss:Cell>',
                    val: val,
					type:'String',
					setType: function (val) {
						if (typeof val =='number') {
							cell.type='Number';
						} else {
							cell.type='String';
						}
						cell.head = '<ss:Cell><Data ss:Type="'+cell.type+'">';
					}
				};
				cell.setType(val);
				console.log(cell);
				row.cells.push(cell);
            }
            return row;
        };
        worksheet.addRow = function (row) {
            worksheet.Rows.push(row);
        }
		return worksheet;
    }
	return workBook;
}

$J2E.export= function (fileName, workbook) {
    if (fileName==null) {
        fileName = 'default';
    }
    var csv = '';
    csv += $J2E.xmlHead;
    csv += $J2E.applicationHead;
    csv += workbook.head;
    var sheets = workbook.workSheet;
    for (var i = 0 ; i < sheets.length; i ++) {
        csv += sheets[i].head;
        var rows = sheets[i].Rows;
        for (var j = 0; j < rows.length; j ++) {
            csv += rows[j].head;
            var cells = rows[j].cells;
            for (var k = 0; k < cells.length; k ++) {
                csv += cells[k].head;
                csv += cells[k].val;
                csv += cells[k].tail;
            }
            csv += rows[j].tail;
        }
        csv += sheets[i].tail;
    }
    csv += workbook.tail;
    //Initialize file format you want csv or xls
    var uri = 'data:text/csv;charset=utf-8,\ufeff' + encodeURIComponent(csv);
    var link = document.createElement("a");
    link.href = uri;

    link.style = "visibility:hidden";
    link.download = fileName + ".xml";

    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
}






