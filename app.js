const express = require('express')
const multer = require('multer')

const fs = require('fs')
const path = require('path')

const app = express()
const PORT = process.env.PORT || 8000;
var toPdf = require("office-to-pdf")

var dir = 'public'
var subDir = 'public/uploads'

const bodyParser = require('body-parser');


if(!fs.existsSync(dir)){

	fs.mkdirSync(dir);
	fs.mkdirSync(subDir);
}

app.use(bodyParser.urlencoded({extended:false}))
app.use(bodyParser.json())
app.use(express.static("public"))


const imageSize = require('image-size')
const sharp = require('sharp')

var width;
var height;
var format;

var storage = multer.diskStorage({
  destination: function (req, file, cb) {
    cb(null, "public/uploads");
  },
  filename: function (req, file, cb) {
    cb(
      null,
      file.fieldname + "-" + Date.now() + path.extname(file.originalname)
    );
  },
});

const imageFilter = function (req, file, cb) {
  if (
    file.mimetype == "image/png" ||
    file.mimetype == "image/jpg" ||
    file.mimetype == "image/jpeg"
  ) {
    cb(null, true);
  } else {
    cb(null, false);
    return cb(new Error("Only .png, .jpg and .jpeg format allowed!"));
  }
};


const FileFilter = function (req, file, cb) {
  
    var ext = path.extname(file.originalname)
    if(ext !== ".docx" && ext !== ".doc" &&
      ext !== ".pptx" && ext !== ".ppt") {
      return cb("Please Select the .docx or .doc File..")
    }
    cb(null,true);

};


const pdfFilter = function(req,file,cb){
       var ext = path.extname(file.originalname);
       if(ext!=='.pdf'){
        return cb("Please Select the pdf File Only..")
      }
      cb(null,true)
}

var uploadImg = multer({ storage: storage, fileFilter: imageFilter });
var uploadFile = multer({ storage: storage, fileFilter: FileFilter });
var pdfUpload = multer({ storage: storage, fileFilter: pdfFilter })


app.get('/uploadImg',(req,res)=>{
    res.sendFile(__dirname+"/index.html")
})

app.get('/pdfUpload',(req,res)=>{
    res.sendFile(__dirname+"/pdfs_merge_tool.html")
})


app.get('/',function(req,res){
	res.sendFile(__dirname+'/main.html');
})

/// #### upload and process Image ####

/*app.post('/processimage',uploadImg.single('file'),function(req,res){
  
     format = req.body.format;
     width = parseInt(req.body.width);
     height = parseInt(req.body.height);
     

     if(req.file){
     	console.log(req.file.path);
      
     	if(isNaN(width) || isNaN(height)){
     		var dimensions = imageSize(req.file.path);
     		console.log(dimensions)
     		width = parseInt(dimensions.width);
     		height = parseInt(dimensions.height);
        processImage(width,height,req,res)
     	}
     	else{
                processImage(width,height,req,res)
     	}


     }


});*/

var outputFilePath;
var ratio;

app.post('/processimage',uploadImg.single('file'),function(req,res){
  
     format = req.body.format;
     width = parseInt(req.body.width);
     height = parseInt(req.body.height);
     ratio = req.body.ratio;
     console.log(ratio)

     if(ratio=='yes'){
          

          if(req.file){
            console.log(req.file.path);

            if(isNaN(height)){
              var dimensions = imageSize(req.file.path);
              console.log(dimensions)
              width = parseInt(req.body.width);
              var rt = (dimensions.width/dimensions.height);
              height = Math.ceil(width/rt);

              

              console.log(rt+' '+width+' '+height)
              processImage(width,height,req,res)
              

             }

   }
}
     else{


      if(req.file){
      console.log(req.file.path);
      
      if(isNaN(width) || isNaN(height)){
        var dimensions = imageSize(req.file.path);
        console.log(dimensions)
        width = parseInt(dimensions.width);
        height = parseInt(dimensions.height);
        processImage(width,height,req,res)
      }
      else{
                processImage(width,height,req,res)
      }


     }

     }
     


});





app.get('/uploadFile',(req,res)=>{
  res.sendFile(__dirname+"/index_for_file.html")
})


var outputFile;
app.post('/docxprocess',uploadFile.single('file'),function(req,res){

  if(req.file){
    console.log(req.file.path)
  }

  const selected_file = fs.readFileSync(req.file.path);
  //console.log(selected_file);
  outputFile=Math.ceil(Math.random()*9999+1)+"outputPDFFile.pdf";

  toPdf(selected_file).then((pdfBuffer) => {

              fs.writeFileSync(outputFile, pdfBuffer)
              res.download(outputFile,(err)=>{

          if(err){
           fs.unlinkSync(req.file.path)
           fs.unlinkSync(outputFile)
           res.send("Error in download process...")
          }
             fs.unlinkSync(req.file.path)
               fs.unlinkSync(outputFile)

            }) //download method end

         }, (err) => {
                 console.log(err)
             
     })

  


})





// ############ PDF PRocessing ##################

const pdf_merge = require('easy-pdf-merge')

var outputFinalFile;
app.post('/processPdf',pdfUpload.array('file',20),(req,res)=>{

  const files = []
  outputFinalFile = Math.ceil(Math.random()*9999+1)+'-output.pdf';


  if(req.files){

       req.files.forEach(file=>{
          console.log(file.path)
          files.push(file.path)
       })


  pdf_merge(files,outputFinalFile,(err)=>{
         if(err){
            res.send(" Please Select the File then Press Download..")
         }
         res.download(outputFinalFile,(err)=>{
            if(err){
              res.send("error in Mergeing File in download try again later...")
            }
            fs.unlinkSync(outputFinalFile)
            

            return Promise.all(files.map(file=>{
                 new Promise((res,rej)=>{

                   try{
                      fs.unlink(file,(err)=>{
                          if(err) throw err;
                          console.log("Successfully Deleted")
                      })
                   }
                   catch(err){

                   }
                 })
            })
            );



          ///#####################
         })
         
  })
      
    

  }


  
})








//  ##########################

app.listen(PORT,function(){
	console.log(`App is Start at ${PORT}`)
});

var outputFilePath;
function processImage(width,height,req,res){

   outputFilePath = Date.now() + "output." + format;
  if (req.file) {
    sharp(req.file.path)
      .resize(width, height)
      .toFile(outputFilePath, (err, info) => {
        if (err) throw err;
        res.download(outputFilePath, (err) => {
          if (err) throw err;
          fs.unlinkSync(req.file.path);
          fs.unlinkSync(outputFilePath);
        });
      });
  }
     
}