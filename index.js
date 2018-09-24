var express = require("express");
var app = express();
var multer = require("multer");

app.use(function (req, res, next) {
    res.header("Access-Control-Allow-Origin", "*");
    res.header("Access-Control-Allow-Methods", "GET,HEAD,OPTIONS,POST,PUT");
    res.header("Access-Control-Allow-Headers", "Origin, X-Requested-With, contentType,Content-Type, Accept, Authorization");
    next();
});

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
    console.log(req.file);
    res.send(req.file.originalname);
});
