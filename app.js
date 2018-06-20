const express = require('express');
const app = express();
const fs = require('fs')
const https = require("https");

var PptxGenJS = require("pptxgenjs");
var pptx = new PptxGenJS();

var gConsoleLog = true;

if (gConsoleLog) console.log(`
-------------
CREATING PPTX
-------------
`);

var PptxGenJS = require("pptxgenjs");
var pptx = new PptxGenJS();

// Generate dev token at https://developers.wrike.com/
var permanent_token = '';

// Parameters will be passed through web interface... eventuallly...

var options = {
   agent: false,
   hostname: 'www.wrike.com',
   port: 443,
   method: 'GET',
   headers: {
    Authorization: 'bearer ' + permanent_token
  },
  path: ''
};

var req = https.request(options, function(res) {

  var body = [];

  res.on('data', function(data) {
      body.push(data);
  })

  // After data response, create a new presentation
  res.on('end', function() {
      body = Buffer.concat(body);

      wrikeData = JSON.parse(body);

      var exportName = 'TaskActivity';

      var slide = pptx.addNewSlide();
      slide.back = 'ECE8F2';

      // Status description
      slide.addShape(pptx.shapes.ROUNDED_RECTANGLE, {x:0.4, y:0.75, w:4, h:1.35, fill:'1632C2', rectRadius:.1, shadow:{ angle:60, offset:2, blur:12, opacity:0.18 }} );
      slide.addText('Status Report', { shape: pptx.shapes.RECTANGLE, align:'l', x:0.6, y:1, w:2.19, h:0.3, bold:true, fontSize:16, fontFace:'Segoe UI', color:'ffffff' } );
      slide.addText('Lorem ipsum dolor sit amet, consectetur adipiscing elit. Proin tincidunt orci sit amet turpis dignissim pretium. Suspendisse potenti.', { shape: pptx.shapes.RECTANGLE, align:'l', x:0.6, y:1.3, w:3, h:0.56, fontSize:9, fontFace:'Segoe UI', color:'ffffff' } );

      // Scope
      slide.addShape(pptx.shapes.ROUNDED_RECTANGLE, {x:4.7, y:0.75, w:1.5, h:1.35, fill:'ffffff', rectRadius:.1, shadow:{ angle:60, offset:2, blur:12, opacity:0.18 }});
      slide.addText('Scope', { shape: pptx.shapes.RECTANGLE, align: 'l', x:4.8, y:1.3, w:1.13, h:0.3,  bold:true, fontSize:12, fontFace:'Segoe UI', color:'1447D6'});
      slide.addText('Lorem ipsum dolor sit amet', { shape: pptx.shapes.RECTANGLE, align: 'l', x:4.8, y:1.6, w:1.13, h:0.3, fontSize:8, fontFace:'Segoe UI', color:'1169F8'});

      // Timeline
      slide.addShape(pptx.shapes.ROUNDED_RECTANGLE, {x:6.4, y:0.75, w:1.5, h:1.35, fill:'ffffff', rectRadius:.1, shadow:{ angle:60, offset:2, blur:12, opacity:0.18 }});
      slide.addText('Timeline', { shape: pptx.shapes.RECTANGLE, align: 'l', x:6.5, y:1.3, w:1.13, h:0.3,  bold:true, fontSize:12, fontFace:'Segoe UI', color:'1447D6'});
      slide.addText('Lorem ipsum dolor sit amet', { shape: pptx.shapes.RECTANGLE, align: 'l', x:6.5, y:1.6, w:1.13, h:0.3, fontSize:8, fontFace:'Segoe UI', color:'1169F8'});

      // Risk
      slide.addShape(pptx.shapes.ROUNDED_RECTANGLE, {x:8.1, y:0.75, w:1.5, h:1.35, fill:'ffffff', rectRadius:.1, shadow:{ angle:60, offset:2, blur:12, opacity:0.18 }});
      slide.addText('Risk', { shape: pptx.shapes.RECTANGLE, align: 'l', x:8.2, y:1.3, w:1.13, h:0.3,  bold:true, fontSize:12, fontFace:'Segoe UI', color:'1447D6'});
      slide.addText('Lorem ipsum dolor sit amet', { shape: pptx.shapes.RECTANGLE, align: 'l', x:8.2, y:1.6, w:1.13, h:0.3, fontSize:8, fontFace:'Segoe UI', color:'1169F8'});

      // Status indicators, logic for color selection coming soon...
      slide.addShape(pptx.shapes.OVAL, { x:5.8, y:0.9, w:.25, h:.25, line:'FF7575', lineSize:2, fill:'ffffff' });
      slide.addShape(pptx.shapes.OVAL, { x:7.5, y:0.9, w:.25, h:.25, line:'75FFA3', lineSize:2, fill:'ffffff' });
      slide.addShape(pptx.shapes.OVAL, { x:9.2, y:0.9, w:.25, h:.25, line:'FFE375', lineSize:2, fill:'ffffff' });

      // Create header row
      var rows = [['Status', 'Name', 'Importance', 'Due']]

      // Populate table with project tasks
      for (var i=0; i<wrikeData.data.length; i++) {

          var row = [
            wrikeData.data[i].status,
            wrikeData.data[i].title,
            wrikeData.data[i].importance,
            wrikeData.data[i].dates.due
          ]

          rows.push(row)

      };

      // Make table look cool
      slide.addShape(pptx.shapes.ROUNDED_RECTANGLE, {x:0.4, y:2.3, w:9.2, h:2.2, fill:'ffffff', rectRadius:.1, shadow:{ angle:60, offset:2, blur:12, opacity:0.18 }});
      slide.addTable( rows, { x:0.5, y:2.45, w:9, fill:'FFFFFF', fontSize:8, fontFace:'Segoe UI', color:'1447D6', border: 'none', margin:[10,10,10,10]} );

      pptx.save( exportName ); if (gConsoleLog) console.log('\nFile created:\n'+' * '+exportName);
  })
})

req.end();

if (gConsoleLog) console.log(`
--------------
DONE!
--------------
`);
