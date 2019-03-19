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


const arrayJsonMerge = (array) => { 
    let resultArray=new Array();
    array.forEach(element => {
        element.forEach(lowerElement=>{
            resultArray.push(lowerElement);
        });
    });
    console.dir(resultArray);

    return resultArray;
};

const createExcel = (array) => {
    let workseet = new xlsx;
    
    console.log(xlsx.createExcel(array));
    
    return xlsx.writeFile;  //여기부터 수정할것
    

};




router.post('/excelParse', upload.single("uploadfile"), function(req, res, next) {
    let work = xlsx.read(req.file.buffer);
    const workseet=work.Sheets["Sheet1"];
    
    
    let ref = workseet["!ref"].replace(/A/gi,'').split(":");
    console.log(ref);
    let tempArray = new Array();
    
    for(let i = ref[0];i<ref[1];i++){
        tempArray.push("A"+i);
    }
    console.log(tempArray);
    let map = tempArray.map(s=>JSON.parse(workseet[s].w).data);
    
    map=arrayJsonMerge(map);
    
    /*
    createExcel(arrayJsonMerge(map));
     try {
        res.json(arrayJsonMerge(map));    
    } catch (error) {
        res.json(error);
    }*/
    
    res.render('result',{data:map,title:'파싱된다'});

});




module.exports = router;
