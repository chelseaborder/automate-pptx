const fs = require('fs')
const https = require("https");

const options = {
   agent: false,
   hostname: 'www.wrike.com',
   port: 443,
   method: 'GET',
   headers: {
    Authorization: '' // Retrieve your development token at wrike
  },
};

options.path = '/api/v3/folders/'

// Make the HTTPS request
const req = https.request(options, function(res) {

  let body = []

  res.on('data', function(data) {
      body.push(data);
  });

  res.on('end', function() {

      body = Buffer.concat(body);

      folderName = JSON.parse(body);

      let folderNameList = []

      for (var i=0; i<folderName.data.length; i++) {

          let folderNames = {
            name: folderName.data[i].title,
            id: folderName.data[i].id,
            childIds: folderName.data[i].childIds
          }

          folderNameList.push(folderNames)
      };

      const folderNamesAsString = JSON.stringify(folderNameList,null,2)

      const outputFileContent = 'var folder = ' + folderNamesAsString

      fs.writeFile('folderNameList.js',outputFileContent,'utf-8',()=>{
        console.log('Done!')
      });

    });

}).on("error", (err) => {
  console.log("Error: " + err.message);
});

req.end();
