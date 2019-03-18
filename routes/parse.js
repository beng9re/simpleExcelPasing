var express = require('express');
var multer =require('multer');
var excel = require('excel');
var xlsx= require('xlsx');
var upload = multer();

const router = express.Router();

/* GET home page. */
router.get('/', function(req, res, next) {
  res.render('parse', { title: 'Express' });
});





router.post('/excelParse', upload.single("uploadfile"), function(req, res, next) {
    let work = xlsx.read(req.file.buffer);
    let workseet=work.Sheets["Sheet1"];
    let ref = workseet["!ref"].replace(/A/gi,'').split(":");
    console.log(ref);
    let tempArray = new Array();
    
    for(let i = ref[0];i<ref[1];i++){
        tempArray.push("A"+i);
    }
    console.log(tempArray);
    
    tempArray.forEach(s => {
        console.log(workseet[s].w);
    });
  
    //tempArray.push
    console.log("sad"+workseet.A2);
    
    res.json(workseet);
    // console.log(workseet);  
});




module.exports = router;
