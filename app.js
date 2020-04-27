var express    = require('express');
var mongoose   = require('mongoose');
var bodyParser = require('body-parser');
var path       = require('path');
var XLSX       = require('xlsx');
var multer     = require('multer');


//multer
var storage = multer.diskStorage({
    destination: function (req, file, cb) {
      cb(null, './public/uploads')
    },
    filename: function (req, file, cb) {
      cb(null, file.originalname)
    }
  });
  
  var upload = multer({ storage: storage });

//connect to db
mongoose.connect('mongodb://localhost:27017/Demoexcel',{useNewUrlParser:true})
.then(()=>{console.log('connected to db')})
.catch((error)=>{console.log('error',error)});

//init app
var app = express();

//set the template engine
app.set('view engine','ejs');

//fetch data from the request
app.use(bodyParser.urlencoded({extended:false}));

//static folder path
app.use(express.static(path.resolve(__dirname,'public')));

//collection schema
var excelSchema = new mongoose.Schema({
    firstname:String,
    lastname:String,  
    country:String
});

var excelModel = mongoose.model('excelData',excelSchema);


app.get('/',(req,res)=>{
   excelModel.find((err,data)=>{
       if(err){
           console.log(err)
           //res.status(500).send("required missing");
       }else{
           if(data!=''){
               res.render('home',{result:data});
           }else{
               res.render('home',{result:{}});
           }
       }
   });
});
// process.on('uncaughtException', function(err) {
//     console.log('Caught exception: ' + err);
//   });
  
app.post('/',upload.single('excel'),(req,res)=>{
  var workbook =  XLSX.readFile(req.file.path);
 
  var sheet_namelist = workbook.SheetNames;
  var x=0;
  sheet_namelist.forEach(element => {
      var xlData = XLSX.utils.sheet_to_json(workbook.Sheets[sheet_namelist[x]]);

      console.log(xlData);
      
      excelModel.insertMany(xlData,(err,data)=>{
          if(err){
              console.log(err);
            
         
          }else{
              console.log(data);
              
          }
      })
      x++;
  });
  res.redirect('/');
  
});

// app.get('/download/:id',(req,res)=>{
//   excelModel.find({_id:req.params.id},(err,data)=>{
//       if(err){
//           console.log(err)
//       } 
//       else{
//         var down = __dirname+'/public/exportdata.xlsx'
//         // var path= __dirname+'/public/'+data[0].picspath;
//          res.download(down);
//       }
//   })
// })
app.get('/download',(req,res)=>{
  var wb = XLSX.utils.book_new(); 
  excelModel.find((err,data)=>{
      if(err){
          console.log(err)
      }else{
          var temp = JSON.stringify(data);
          temp = JSON.parse(temp);
          var ws = XLSX.utils.json_to_sheet(temp);
          var down = __dirname+'/public/Test.xlsx'
         XLSX.utils.book_append_sheet(wb,ws,"sheet1");
         XLSX.writeFile(wb,down);
         res.download(down);
      }
  });
});
var port = process.env.PORT || 5000;
app.listen(port,()=>console.log('server run at '+port));