const requirejs      = require('requirejs');
const express      = require('express'); // Not core - Only required for streaming
const app          = express(); // Not core - Only required for streaming
const fs           = require('fs')
const https        = require('https');
const PptxGenJS    = require('pptxgenjs');
const passFolderId = require('./passFolderId');

requirejs.config({
    //Pass the top-level main.js/index.js require
    //function to requirejs so that node modules
    //are loaded relative to the top-level JS file.
    nodeRequire: require
});

// Authorization Credentials
const options = {
   agent: false,
   hostname: 'www.wrike.com',
   port: 443,
   method: 'GET',
   headers: {
    Authorization: ''
  },
};

function createPowerpoint() {

  let folder = {
    id: selectedFolderValue
  }

  options.path = '/api/v3/folders/' + folder.id + '/tasks/'

  // Make the HTTPS request
  const req = https.request(options, function(res) {

    let body = [];

    res.on('data', function(data) {
        body.push(data);
      });

    res.on('end', function() {

        body = Buffer.concat(body);

        wrikeData = JSON.parse(body);

        const pptx = new PptxGenJS();
        const slide = pptx.addNewSlide();
        slide.back = 'ECE8F2';

        // Status
        slide.addShape(pptx.shapes.RECTANGLE, {x:0.4, y:0.75, w:4, h:1.35, fill:'1632C2'} );
        slide.addText('Status Report', { shape: pptx.shapes.RECTANGLE, align:'l', x:0.6, y:1, w:2.19, h:0.3, bold:true, fontSize:16, fontFace:'Segoe UI', color:'ffffff' } );
        slide.addText('Lorem ipsum dolor sit amet, consectetur adipiscing elit. Proin tincidunt orci sit amet turpis dignissim pretium. Suspendisse potenti.', { shape: pptx.shapes.RECTANGLE, align:'l', x:0.6, y:1.3, w:3, h:0.56, fontSize:9, fontFace:'Segoe UI', color:'ffffff' } );

        // 1st white square
        slide.addShape(pptx.shapes.RECTANGLE, {x:4.7, y:0.75, w:1.5, h:1.35, fill:'ffffff'});
        slide.addText('Scope', { shape: pptx.shapes.RECTANGLE, align: 'l', x:4.8, y:1.3, w:1.13, h:0.3,  bold:true, fontSize:12, fontFace:'Segoe UI', color:'1447D6'});
        slide.addText('Lorem ipsum dolor sit amet', { shape: pptx.shapes.RECTANGLE, align: 'l', x:4.8, y:1.6, w:1.13, h:0.3, fontSize:8, fontFace:'Segoe UI', color:'1169F8'});

        // 2nd white square
        slide.addShape(pptx.shapes.RECTANGLE, {x:6.4, y:0.75, w:1.5, h:1.35, fill:'ffffff'});
        slide.addText('Timeline', { shape: pptx.shapes.RECTANGLE, align: 'l', x:6.5, y:1.3, w:1.13, h:0.3,  bold:true, fontSize:12, fontFace:'Segoe UI', color:'1447D6'});
        slide.addText('Lorem ipsum dolor sit amet', { shape: pptx.shapes.RECTANGLE, align: 'l', x:6.5, y:1.6, w:1.13, h:0.3, fontSize:8, fontFace:'Segoe UI', color:'1169F8'});

        // 3rd white square
        slide.addShape(pptx.shapes.RECTANGLE, {x:8.1, y:0.75, w:1.5, h:1.35, fill:'ffffff'});
        slide.addText('Risk', { shape: pptx.shapes.RECTANGLE, align: 'l', x:8.2, y:1.3, w:1.13, h:0.3,  bold:true, fontSize:12, fontFace:'Segoe UI', color:'1447D6'});
        slide.addText('Lorem ipsum dolor sit amet', { shape: pptx.shapes.RECTANGLE, align: 'l', x:8.2, y:1.6, w:1.13, h:0.3, fontSize:8, fontFace:'Segoe UI', color:'1169F8'});

        //
        slide.addShape(pptx.shapes.OVAL, { x:5.8, y:0.9, w:.25, h:.25, line:'FF7575', lineSize:2, fill:'ffffff' });
        slide.addShape(pptx.shapes.OVAL, { x:7.5, y:0.9, w:.25, h:.25, line:'75FFA3', lineSize:2, fill:'ffffff' });
        slide.addShape(pptx.shapes.OVAL, { x:9.2, y:0.9, w:.25, h:.25, line:'FFE375', lineSize:2, fill:'ffffff' });

        var rows = [['Status', 'Name', 'Importance', 'Due']]

        for (var i=0; i<wrikeData.data.length; i++) {

          var row = [
            wrikeData.data[i].status,
            wrikeData.data[i].title,
            wrikeData.data[i].importance,
            wrikeData.data[i].dates.due
          ]

          rows.push(row)

        };

        slide.addTable( rows, { x:0.4, y:2.45, w:9.2, fill:'FFFFFF', fontSize:8, fontFace:'Segoe UI', color:'1447D6', border: 'none', margin:[5,5,,10]} );
        pptx.save('Node_Demo');
  });

  }).on("error", (err) => {
    console.log("Error: " + err.message);
  });

  req.end();

};
