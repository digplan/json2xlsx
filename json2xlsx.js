module.exports = function(filename, sheetname){

XLSX = require('xlsx');
FS = require('fs');

//var workbook = XLSX.readFile('txt.xlsx');
//var sheets = workbook.Props.SheetNames;
var indata = '';
process.stdin.on('readable', function(){
  indata += process.stdin.read() || '';
})
process.stdin.on('end', processData);

function processData(){
  var o = JSON.parse(indata);
  if(o.push && sheetname){
  	var ob = {};
  	ob[sheetname] = o;
  	o = ob;
  }
  var wb = FS.existsSync(filename) ? XLSX.readFile(filename) : new Workbook();

  for(ws_name in o){
  	wb.SheetNames.push(ws_name);
    var twodarr = o[ws_name];
    if(!twodarr[0].push)
    	twodarr = convertObjArray(twodarr);
    var ws = sheet_from_array_of_arrays(twodarr);
  	wb.Sheets[ws_name] = ws;
  	//console.log(ws_name)
  }
  XLSX.writeFile(wb, filename);
}

function convertObjArray(objarray){
  var arrarr = [Object.keys(objarray[0])];
  for(var n=0; n<objarray.length;n++){
  	var row = [];
  	for(var i in objarray[0])
  	  row.push(objarray[n][i]);
    arrarr.push(row);
  }
  console.log(arrarr);
  return arrarr;
}

function datenum(v, date1904) {
	if(date1904) v+=1462;
	var epoch = Date.parse(v);
	return (epoch - new Date(Date.UTC(1899, 11, 30))) / (24 * 60 * 60 * 1000);
}
 
function sheet_from_array_of_arrays(data, opts) {
	var ws = {};
	var range = {s: {c:10000000, r:10000000}, e: {c:0, r:0 }};
	for(var R = 0; R != data.length; ++R) {
		for(var C = 0; C != data[R].length; ++C) {
			if(range.s.r > R) range.s.r = R;
			if(range.s.c > C) range.s.c = C;
			if(range.e.r < R) range.e.r = R;
			if(range.e.c < C) range.e.c = C;
			var cell = {v: data[R][C] };
			if(cell.v == null) continue;
			var cell_ref = XLSX.utils.encode_cell({c:C,r:R});
			
			if(typeof cell.v === 'number') cell.t = 'n';
			else if(typeof cell.v === 'boolean') cell.t = 'b';
			else if(cell.v instanceof Date) {
				cell.t = 'n'; cell.z = XLSX.SSF._table[14];
				cell.v = datenum(cell.v);
			}
			else cell.t = 's';
			
			ws[cell_ref] = cell;
		}
	}
	if(range.s.c < 10000000) ws['!ref'] = XLSX.utils.encode_range(range);
	return ws;
}

function Workbook() {
	if(!(this instanceof Workbook)) return new Workbook();
	this.SheetNames = [];
	this.Sheets = {};
}

}