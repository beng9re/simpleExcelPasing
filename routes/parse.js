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
    return resultArray;
};


const jsonToHeader = (json) => Object.keys(json[0]);




const createExcel = (json) => {
    let workseet = new xlsx;
    
    console.log(xlsx.createExcel(array));
    
    return xlsx.writeFile;  //여기부터 수정할것
    

};

const createCVS = (json) =>{
    
}




router.post('/excelParse', upload.single("uploadfile"), function(req, res, next) {
    let work;
    try { 
        //버퍼 값 
        work = xlsx.read(req.file.buffer);
    }
    catch(e){
        res.render('result',{check:false,msg:`재시도 요청 ${e}`,title:'파싱된다'});
    }

    const workseet=work.Sheets["Sheet1"];
    let ref = workseet["!ref"].replace(/[A-Za-z]*/gi,'').split(":");
    //console.log(ref);
    let tempArray = new Array();
    
    for(let i = ref[0];i<ref[1];i++){
        tempArray.push("A"+i);
    }
    //console.log(tempArray);
    let map = tempArray.map(s=>JSON.parse(workseet[s].w).data);
    
    
    map=arrayJsonMerge(map);
    const key= Object.keys(map[0]); //header 값
    const values = map;             //value 값

    console.debug(typeof(values));
    res.render('result',{check:true,key:key,list:values,title:'파싱된다'});

});




module.exports = router;