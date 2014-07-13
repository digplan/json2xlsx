# write an Excel file
echo '{"work1": [["TRUE",2,3], [4,5,6]], "work2": [{"a": 1, "b":2},{"a":3, "b": 4}]}'\
| node -e "require('./json2xlsx.js').write('file.xlsx')"

# update existing
echo '{"work3": [["TRUE",2,3], [4,5,6]], "work4": [{"a": 1, "b":2},{"a":3, "b": 4}]}'\
| node -e "require('./json2xlsx.js').write('file.xlsx')"

# Read an Excel file
node -e "require('./json2xlsx.js').read('file.xlsx')"

# use node
node ./json2xlsx.js file.xlsx

# global (read xlsx)
json2xlsx file.xlsx
