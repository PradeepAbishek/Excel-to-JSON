var express = require('express')
var app = express()
var xlsxtojson = require("xlsx-to-json");
var xlstojson = require("xls-to-json");
const fs = require('fs');
app.use(function(req, res, next) { //allow cross origin requests
    res.setHeader("Access-Control-Allow-Origin", "*");
    res.header("Access-Control-Allow-Methods", "POST, PUT, OPTIONS, DELETE, GET");
    res.header("Access-Control-Max-Age", "3600");
    res.header("Access-Control-Allow-Headers", "Content-Type, Access-Control-Allow-Headers, Authorization, X-Requested-With");
    next();
});
// configuration
app.get('/', function (req, res) {
  res.send('Welcome to Excel to JSON conversion')
})
app.post('/api/xlstojson', function(req, res) {
  fs
	.readdirSync(__dirname)
	.filter(file => {
		return (file.slice(-5) === '.xlsx');
	})
	.forEach(file => {
		xlsxtojson({ //Please change xlsx to xls based on the file extension
      input: __dirname+"/"+file,  // input file 
      output: __dirname+"/"+file.substring(0,(file.length)-5)+".json", // output json
      lowerCaseHeaders:true
    }, function(err, result) {
      if(err) {
        res.json(err);
      } else {
        res.json(result);
      }
    });
  });
  
  // Checks for the column has multiple values
  fs
  .readdirSync(__dirname)
	.filter(file => {
		return (file.slice(-5) === '.json');
  })
  .forEach(file => {
    fs.readFile(__dirname+"/"+file, 'utf8', function (err,data) {
      if (err) {
        return console.log(err);
      }
      var s = JSON.parse(data);
      for(var i=0; i<s.length; i++) {
        for(n in s[i]) {
          if(s[i][n].length > 0 && s[i][n].includes(",") ) {
            var t = s[i][n].split(',');
            s[i][n] = t;
          }
        }
        s[i] = JSON.stringify(s[i]);
      }
      fs.writeFile(__dirname+"/"+file, s, 'utf8', function (err) {
        if (err) return console.log(err);
      });
    });
  })
});
app.listen(3000)
console.log("Application running on the PORT 3000");