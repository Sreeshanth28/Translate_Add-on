/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office, Word */
// const translate = require('google-translate-extended-api');

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    document.getElementById("sideload-msg").style.display = "none";
    // document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = run;
    document.getElementById("find").onclick = find;
    document.getElementById('english').onclick = english;
    document.getElementById('telugu').onclick = telugu;
    document.getElementById('hindi').onclick = hindi;
  }
});

export async function run() {
 
  return Word.run(async (context) => {
    // var shan = `<div>
    //   Hello
    // </div>}`
    const paragraph = context.document.body.insertText(
      '££Hello World££ \n$$Hello World$$ \n¥¥Hello World¥¥ \n================================ \n'
      , Word.InsertLocation.end);
      
      
    paragraph.font.color = "black";
    paragraph.font.size = 20;

    await context.sync();
  });
}

// export async function run() {
//   return Word.run(async (context) => {
//     let shan = ''
//     const options = {
//       method: 'POST',
//       headers: {
//         'content-type': 'application/json',
//         'X-RapidAPI-Host': 'deep-translate1.p.rapidapi.com',
//         'X-RapidAPI-Key': '1af02cfdddmsh0114bdde996de50p1b5784jsn5fb0c48b2a29'
//       },
//       body: '{"q":"Sreeshanth","source":"en","target":"te"}'
//     };
    
//     const s = await fetch('https://deep-translate1.p.rapidapi.com/language/translate/v2', options)

    
//     const response = await s.json()
//     shan = response.data.translations.translatedText
//     console.log(shan)
//       const paragraph = context.document.body.insertText(
//           // '<div></div>'
//       shan
//       , Word.InsertLocation.end);

//         paragraph.font.color = "black";
//         paragraph.font.size = 20;

//         await context.sync();
//       });
// }


async function find() {
  return Word.run(async (context) => {

    var results = context.document.body.search("World"); //Search for the text to replace

    context.load(results);
    
    return context
      .sync()
      .then(function () {
        for (var i = 0; i < results.items.length; i++) {
          results.items[i].insertHtml("Shan", "replace");
          results.items[i].font.color = "blue"; 
          results.items[i].font.size = 20; 
        }
      })
      .then(context.sync);
  })
  .catch(function(e){
    console.log(e.message);
    if (error instanceof OfficeExtension.Error) {
      console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
  });
}



async function english() {
  document.getElementsByClassName('loader')[0].style.display = "block";
  return Word.run(async (context) => {
    var searchResults = context.document.body.search('££*££', {matchWildcards: true});
    context.load(searchResults);

    return context
    .sync()
    .then (async function(){
      for (let i = 0; i < searchResults.items.length; i++) {
        let a = searchResults.items[i].text;
        a = a.split('££').join('');
        console.log(a);

        let shan = ''
        const options = {
          method: 'POST',
          headers: {
            'content-type': 'application/json',
            'X-RapidAPI-Host': 'deep-translate1.p.rapidapi.com',
            'X-RapidAPI-Key': '1af02cfdddmsh0114bdde996de50p1b5784jsn5fb0c48b2a29'
          },
          body: `{"q": "${a}","source":"en","target":"en"}`
        };
        
        const s = await fetch('https://deep-translate1.p.rapidapi.com/language/translate/v2', options);
        const response = await s.json();
        // console.log("Hello");
        // console.log(response);
        shan = response.data.translations.translatedText;
        // console.log(shan)
        searchResults.items[i].insertHtml(shan, "replace");
        searchResults.items[i].font.color = "purple"; 
        searchResults.items[i].font.size = 20;
      }
      document.getElementsByClassName('loader')[0].style.display = "none";
    }).then(context.sync);
  })
  .catch(function(e){
    console.log(e.message);
    if (error instanceof OfficeExtension.Error) {
      // console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
  });
}



async function telugu() {
  document.getElementsByClassName('loader')[0].style.display = "block";
  return Word.run(async (context) => {
    
    var searchResults = context.document.body.search('$$*$$', {matchWildcards: true});
    context.load(searchResults);

    return context
    .sync()
    .then (async function(){
      for (let i = 0; i < searchResults.items.length; i++) {
        let a = searchResults.items[i].text;
        a = a.split('$$').join('');
        // console.log(a);

        let shan = ''
        const options = {
          method: 'POST',
          headers: {
            'content-type': 'application/json',
            'X-RapidAPI-Host': 'deep-translate1.p.rapidapi.com',
            'X-RapidAPI-Key': '1af02cfdddmsh0114bdde996de50p1b5784jsn5fb0c48b2a29'
          },
          body: `{"q": "${a}","source":"en","target":"te"}`
        };
        
        const s = await fetch('https://deep-translate1.p.rapidapi.com/language/translate/v2', options);

        
        const response = await s.json();
        // console.log("Hello");
        // console.log(response);
        shan = response.data.translations.translatedText;
        // console.log(shan)
        searchResults.items[i].insertHtml(shan, "replace");
        searchResults.items[i].font.color = "green"; 
        searchResults.items[i].font.size = 20;
      }
      document.getElementsByClassName('loader')[0].style.display = "none";
    }).then(context.sync);
    
  })
  .catch(function(e){
    // console.log(e.message);
    if (error instanceof OfficeExtension.Error) {
      // console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
  });
  
}

async function hindi() {
  document.getElementsByClassName('loader')[0].style.display = "block";
  return Word.run(async (context) => {
    var searchResults = context.document.body.search('¥¥*¥¥', {matchWildcards: true});
    context.load(searchResults);

    

    return context
    .sync()
    .then (async function(){
      for (let i = 0; i < searchResults.items.length; i++) {
        let a = searchResults.items[i].text;
        a = a.split('¥¥').join('');
        // console.log(a);

        let shan = ''
        const options = {
          method: 'POST',
          headers: {
            'content-type': 'application/json',
            'X-RapidAPI-Host': 'deep-translate1.p.rapidapi.com',
            'X-RapidAPI-Key': '1af02cfdddmsh0114bdde996de50p1b5784jsn5fb0c48b2a29'
          },
          body: `{"q": "${a}","source":"en","target":"hi"}`
        };
        
        const s = await fetch('https://deep-translate1.p.rapidapi.com/language/translate/v2', options);

        
        const response = await s.json();
        // console.log("Hello");
        // console.log(response);
        shan = response.data.translations.translatedText;
        // console.log(shan)


        searchResults.items[i].insertHtml(shan, "replace");
        searchResults.items[i].font.color = "red"; 
        searchResults.items[i].font.size = 20;
      }
      document.getElementsByClassName('loader')[0].style.display = "none";
    }).then(context.sync);
  })
  .catch(function(e){
    // console.log(e.message);
    if (error instanceof OfficeExtension.Error) {
      // console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
  });
}




// /*
//  * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
//  * See LICENSE in the project root for license information.
//  */

// /* global document, Office, Word */

// Office.onReady((info) => {
//   if (info.host === Office.HostType.Word) {
//     document.getElementById("sideload-msg").style.display = "none";
//     document.getElementById("app-body").style.display = "flex";
//     document.getElementById("run").onclick = run;
//     document.getElementById("find").onclick = find;
//     // Determine if the user's version of Office supports all the Office.js APIs that are used in the tutorial.

//     // Assign event handlers and other initialization logic.
//     document.getElementById("insert-paragraph").onclick = insertParagraph;
//   }
// });

// async function run() {
//   return Word.run(async (context) => {
//     /*Insert your Word code here*/
//     // insert a paragraph at the end of the document.
//     const paragraph = context.document.body.insertParagraph("Hello World", Word.InsertLocation.end);
//     // change the paragraph color to blue.
//     paragraph.font.color = "blue";
//     await context.sync();
//   });
// }
// var data = {
//   date: "22 Aug 2016",
//   sender: "Someone really important",
//   company1: "The Boss | Company 1 | Somewhere | Here | There | 12345",
//   company2: "The Bigger Boss | Company 2 | Somewhere else | Near | Canada | 98765",
// };
// async function find() {
//   return Word.run(async (context) => {
//     var results = context.document.body.search("[Recipient Name]"); //Search for the text to replace
//     context.load(results);

//     return context
//       .sync()
//       .then(function () {
//         for (var i = 0; i < results.items.length; i++) {
//           results.items[i].insertHtml("Marky The Receiver", "replace"); //Replace the text HERE
//         }
//       })
//       .then(context.sync)
//       .then(function () {
//         var results = context.document.body.search("[Date]"); //Search for the text to replace
//         context.load(results);

//         return context
//           .sync()
//           .then(function () {
//             for (var i = 0; i < results.items.length; i++) {
//               results.items[i].insertHtml(data.date, "replace"); //Replace the text HERE
//             }
//           })
//           .then(context.sync())
//           .then(function () {
//             var results = context.document.body.search("[Title] | [Company] |[Address] | [City] | [State] | [Zip]"); //Search for the text to replace
//             context.load(results);

//             return context
//               .sync()
//               .then(function () {
//                 results.items[0].insertHtml(data.company1, "replace"); //Replace the text HERE
//                 results.items[1].insertHtml(data.company2, "replace"); //Replace the text HERE
//               })
//               .then(context.sync())
//               .then(function () {
//                 var results = context.document.body.search("[Sender Name]"); //Search for the text to replace
//                 context.load(results);

//                 return context
//                   .sync()
//                   .then(function () {
//                     for (var i = 0; i < results.items.length; i++) {
//                       results.items[i].insertHtml(data.sender, "replace"); //Replace the text HERE
//                     }
//                   })
//                   .then(context.sync());
//               });
//           });
//       });
//   });
// }

// // async function find() {
// //   return Word.run(async (context) => {
// //     const doc = context.document.body.search("Microsoft");
// //     context.load(doc);
// //     await context.sync();
// //     // if (window.console) console.log("Found count: " + doc.items.length);
// //     for (var i = 0; i < doc.items.length; i++) {
// //       doc.items[i].insertHtml("Marky The Receiver", "replace"); //Replace the text HERE
// //     }
// //   });
// // }

// async function insertParagraph() {
//   await Word.run(async (context) => {
//     // TODO1: Queue commands to insert a paragraph into the document.
//     const docBody = context.document.body.insertParagraph(
//       "Office has several versions, including Office 2016, Microsoft 365 subscription, and Office on the web.",
//       Word.InsertLocation.start
//     );
//     docBody.font.color = "blue";
//     await context.sync();
//   });
// }
