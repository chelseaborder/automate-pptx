const express = require("express");
const app     = express();
const path    = require("path");

app.use(express.static(path.join(__dirname, 'public')));
app.use(express.static(path.join(__dirname, "js")));

app.get('/',function(req,res){
  res.sendFile(
    path.join(__dirname+'/index.html'));
});

app.listen(3000);

console.log("Running at Port 3000");
