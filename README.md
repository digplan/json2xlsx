Json2XLSX
=========

Provide or pipe a JSON object to a new Excel XLSX file

Takes an object with one worksheet (tab) per key, and a 2d array or array of objects
Updates an existing file if already created

````
{
	"worksheet1": [
		[1,2,3],
		[4,5,6]
	],

	"worksheet2": [
		{a: 1, b: 2},
		{a: 3, b: 4}
	]
}
````

````
###Write Excel
require('json2xlsx').write(filename, sheetname, object [optional instead of piping in]);    
object is optionally using object instead of pipe    

or pipe..

echo '{"work1": [["TRUE",2,3], [4,5,6]], "work2": [{"a": 1, "b":2},{"a":3, "b": 4}]}'\
| node -e "require('./json2xlsx.js').write('file.xlsx')"

# update existing
echo '{"work3": [["TRUE",2,3], [4,5,6]], "work4": [{"a": 1, "b":2},{"a":3, "b": 4}]}'\
| node -e "require('./json2xlsx.js').write('file.xlsx')"

# show additional info about writing (set env variable)
$ debug=1 echo'{ ...
````
###Read Excel
````
require('json2xlsx').read('file.xlsx');  // {...data...}

# or
node ./json2xlsx file.xlsx
````