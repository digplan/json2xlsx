echo '{"work1": [["TRUE",2,3], [4,5,6]], "work2": [{"a": 1, "b":2},{"a":3, "b": 4}]}'\
| node -e "require('./json2xlsx.js')('file.xlsx')"

# update existing
echo '{"work3": [["TRUE",2,3], [4,5,6]], "work4": [{"a": 1, "b":2},{"a":3, "b": 4}]}'\
| node -e "require('./json2xlsx.js')('file.xlsx')"