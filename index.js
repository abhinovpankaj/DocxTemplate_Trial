"use strict";

const docxTemplate = require( 'docx-templates');
const fs = require( 'fs');
const path = require('path');
const DocxMerger = require('docx-merger');
//<b>This</b> is a Test <u>text in html</u> <font color="#e31c1c">format</font> with different <span>colors</span> and <i>combinations</i>.
const generate =async ()=>{

const template = fs.readFileSync('Deck2AllData.docx');

const buffer = await docxTemplate.createReport({
  template,
   data:{
  //   project:{
  //     reportType:'Visual',
  //     name:'Test Project',
  //     address:'California',
  //     description:'Just a sample project for testing',
  //     createdAt:'8th Oct2023',
  //     createdBy:'Abhinov',
  //     url:'https://deckinspectors.blob.core.windows.net/xnebbd/CAP4773236617995329626.jpg'
  //   },
  section:{
    isUnitUnavailable :'false',
    reportType : 'Visual',
    buildingName: 'Test',
      parentType: 'Building',
      parentName: 'Project 1',
      name: 'Test locations',
      exteriorelements: 'abc, gha',
      waterproofing:'xyz, gahd, adad',
      visualreview:'Bad',
      signsofleak : 'Yes',
      furtherinvasive:'Yes',
      conditionalassesment:'Fail',
      additionalconsiderations:'Test dasd asd',
      eee:'0-1 Years',
      lbc:'0-1 Years',
      awe:'0-1 Years',
      furtherInvasiveRequired:'true',
      invasiverepairsinspectedandcompleted:'true',
      propowneragreed:'false',
      images:
      [
        'https://deckinspectors.blob.core.windows.net/xnebbd/CAP8294094054021236826.jpg',
        'https://deckinspectors.blob.core.windows.net/xnebbd/CAP4773236617995329626.jpg',
        'https://deckinspectors.blob.core.windows.net/xnebbd/CAP8294094054021236826.jpg',
        'https://deckinspectors.blob.core.windows.net/xnebbd/CAP4773236617995329626.jpg'
      ],
      invasiveImages:
      [
        'https://deckinspectors.blob.core.windows.net/xnebbd/CAP8294094054021236826.jpg',
        'https://deckinspectors.blob.core.windows.net/xnebbd/CAP4773236617995329626.jpg'
      ],
      conclusiveImages:
      [
        'https://deckinspectors.blob.core.windows.net/xnebbd/CAP8294094054021236826.jpg',
        'https://deckinspectors.blob.core.windows.net/xnebbd/CAP4773236617995329626.jpg'
      ]
      
  }
},
  additionalJsContext: {
                
    getChunks : async (imageArray,chunk_size=4) => {
      var index = 0;
      var arrayLength = imageArray.length;
      var tempArray = [];
      var myChunk;
      for (index = 0; index < arrayLength; index += chunk_size) {
          myChunk = imageArray.slice(index, index+chunk_size);
          
          tempArray.push(myChunk);
        }
      
      return tempArray;
    },
    getadditionalconsiderations: ()=>{
      console.log('inside html fetch');
      //return "sfsdfdsfsfS";
       return  '<font face="Nunito"><b>This</b> is a Test <u>text in html</u> <font color="#e31c1c">format</font> with different <span>colors</span> and <i>combinations</i>.</font>';     
    },

    tile: async (imageurl) => {
      console.log('inside image fetch');
      if (imageurl===undefined) {
        return;
      }
      const resp = await fetch(
        imageurl
      );
      const buffer = resp.arrayBuffer
        ? await resp.arrayBuffer()
        : await resp.buffer();
      const extension  = path.extname(imageurl);
      return { height: 6.2,width: 4.85,  data: buffer, extension: extension };
    },
  }
});

fs.writeFileSync('report.docx', buffer);
var file1 = fs
    .readFileSync(path.resolve(__dirname, 'DeckProjectHeader.docx'), 'binary');

var file2 = fs
    .readFileSync(path.resolve(__dirname, 'report.docx'), 'binary');

var docx = new DocxMerger({},[file1,file2]);
docx.save('nodebuffer',function (data) {
  // fs.writeFile("output.zip", data, function(err){/*...*/});
  fs.writeFile("output.docx", data, function(err){/*...*/});

});
}

generate();



