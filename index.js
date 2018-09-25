var express = require("express");
var app = express();
var multer = require("multer");
var Excel = require("exceljs");
var cors = require("cors");

app.use(cors());

var server = app.listen(process.env.PORT || 8080, function () {
    var port = server.address().port;
    console.log("App now running on port", port);
});

const storage = multer.diskStorage({
    destination: './uploads',
    filename(req, file, cb) {
      cb(null, `${file.originalname}`);
    },
});

const upload = multer({ storage });

app.post("/upload", upload.single('file'), (req, res) => {

    var arrCol = []; 
    var arrRow = [];

    var workbook = new Excel.Workbook();
    workbook.xlsx.readFile(req.file.path)
        .then(function() {
            workbook.eachSheet(function(worksheet, sheetId) {

                var dobCol = worksheet.getColumn(1);
                dobCol.eachCell(function(cell, rowNumber) {
                    var row = worksheet.getRow(rowNumber);
                    row.eachCell({ includeEmpty: true }, function(cell, colNumber) {
                        arrRow.push(cell.value);
                    });
                    arrCol.push(arrRow);
                    arrRow = [];
               });

               res.send(arrCol);

            });
        });

});
