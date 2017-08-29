
var server = require('./server');
var router = require('./router');
var request = require("request");

var microsoftGraph = require("@microsoft/microsoft-graph-client");
var fs = require('fs');
var FileReader = require('filereader')
var url = require('url');
var auth = require('./auth');
var request = require('request');
var qs = require('querystring');
var csvWriter = require('csv-write-stream')

var handle = {};
handle['/mail'] = userEmail;
handle['/calendar'] = calendar;
handle['/contacts'] = contacts;
handle['/photo'] = updateProfilePicture;
handle['/users'] = users;
handle['/groups'] = groups;
handle['/file'] = shareFile;

server.start(router.route, handle);

var token = "";

function saveToken(tok) {
  token = tok;
}

//If expired, request new token in the methods!
auth.getAccessToken().then(function (token) {
  // console.log(token);
    saveToken(token)
    .then(function (tok) {
    }, function (error) {
      console.error('>>> Error getting users: ' + error);
    });
}, function (error) {
  console.error('>>> Error getting access token: ' + error);
});



function userEmail(userId) {
  userId = "e97f274a-2a86-4280-997d-8ee4d2c52078";
  var client = microsoftGraph.Client.init({
    authProvider: (done) => {
      done(null, token);
    }
  });
  client
    .api('/users/' + userId + '/mail')
    .get((err, res) => {
      if (err) {
        console.log(err);
      } else {

        console.log(res.value);
      }
    });
}


function groups(response, request) {
  var client = microsoftGraph.Client.init({
    authProvider: (done) => {
      done(null, token);
    }
  });
  client.api("/groups")
    .get((err, res) => {
      if (err) {
        console.log(err);
        response.write('<p>ERROR: ' + err + '</p>');
        response.end();
      } else {

        console.log(res.statusCode);
        console.log(res.value);
        response.write(JSON.stringify(res.value));
        response.end();
      }
    });

}


function users(response, request) {
  console.log(request.method);
  var userId = "e97f274a-2a86-4280-997d-8ee4d2c52078"; //"30be01d3-8214-4f2d-aea0-7028a19581fc" ;//"e97f274a-2a86-4280-997d-8ee4d2c52078";
  var client = microsoftGraph.Client.init({
    authProvider: (done) => {
      // Just return the token
      done(null, token);
    }
  });
  //tlf nr funker! /mobilePhone
  //If you want a different set of properties, you can request them using the $select query parameter. E.g https://graph.microsoft.com/v1.0/users/e97f274a-2a86-4280-997d-8ee4d2c52078?$select=aboutMe
  //Når AD brukes er det ikke mulig å gjøre endringer! Man kan kun gjøre GET requests. Ellers må man oppdatere direkte i AD.
  //Azure Ad Graph Api kan brukes for å gjøre endringer på brukere, grupper og kontakter i AD.
  if (request.method == "POST") {
    client.api("/users/" + userId + "/displayName")
      .patch(
      { "value": "Test" },
      (err, res) => {
        if (err)
          console.log(err);
        else
          console.log("Profile Updated");
      });
  } else {
    client
      .api('https://graph.microsoft.com/v1.0/users')
      .get((err, res) => {
        if (err) {
          console.log(err);
          response.write('<p>ERROR: ' + err + '</p>');
          response.end();
        } else {
          console.log(response.statusCode);
          console.log(res.value);
          response.write(JSON.stringify(res.value));
          response.end();
        }
      });
    //getCurrentUserSP();
    //sharedWithMe();
  }
}

// function getCurrentUserSP() {
//   var ur = 'https://bouvetasa.sharepoint.com/_api/web/currentuser';
//   var opt = {
//     url: ur,
//       method: "GET",
//     header: {
//       'User-Agent': 'Super Agent/0.0.1',
//       'Content-Type': 'application/x-www-form-urlencoded',

//     }
//   }
//   request(opt, function (error, response, body) {
//     if (!error && response.statusCode == 200) {
//       console.log(error);
//       return error;
//     } else {
//       //response.statusCode +s
//       console.log( " " + response.value + body);
//       return response;
//     }
//   });
// }

// function sharedWithMe() {
//   var ur = 'https://bouvetasa.sharepoint.com/_api/search/query?querytext=%27(SharedWithUsersOWSUSER:trond.tufte@bouvet.no)%27';
//   var opt = {
//     url: ur,
//   //    method: "GET",
//     header: {
//       'User-Agent': 'Super Agent/0.0.1',
//       'Content-Type': 'application/x-www-form-urlencoded',     
//     }
//   }
//   request(opt, function (error, response, body) {
//     if (!error && response.statusCode == 200) {
//       console.log(error);
//       return error;
//     } else {
//       //response.statusCode +s
//       console.log( " " + response.value + body);
//       return response;
//     }
//   });
// }


function calendar(userId) {
  userId = "e97f274a-2a86-4280-997d-8ee4d2c52078";
  var client = microsoftGraph.Client.init({
    authProvider: (done) => {
      done(null, token);
    }
  });
  client
    .api('/users/' + userId + '/events')
    .get((err, res) => {
      if (err) {

        console.log(err);
      } else {
        console.log(res.value);
      }

    });
}

function contacts(userId) {
  userId = "e97f274a-2a86-4280-997d-8ee4d2c52078";

  var client = microsoftGraph.Client.init({
    authProvider: (done) => {
      done(null, token);
    }
  });
  client
    .api('/users/' + userId + '/contacts')
    .get((err, res) => {
      if (err) {

        console.log(err);
      } else {
        console.log(res.value);
      }

    });
}


function updateProfilePicture() {
  var client = microsoftGraph.Client.init({
    authProvider: (done) => {
      // Just return the token
      done(null, token);
    }
  });

  var userId = "30be01d3-8214-4f2d-aea0-7028a19581fc";  //(britt)    "8fa20769-13f0-4b67-b777-c262b174d93e"; (Eirik?)  //e97f274a-2a86-4280-997d-8ee4d2c52078  (min)
  var file = fs.readFileSync('./logo.png'); // fs.openSync("logo.png","r"); //new File("logo.png");       
  //var reader = new FileReader();
  client
    .api("/users/" + userId + "/photo/$value")
    .put(file, (err, res) => {
      if (err) {
        console.log(err);
        return;
      }
      console.log("Image updated!");
    });

  // reader.addEventListener("load", function () {
  // 	client
  // 		.api('/users/e97f274a-2a86-4280-997d-8ee4d2c52078/photo/$value')
  // 		.put(file, (err, res) => {
  // 			if (err) {
  // 				console.log(err);
  // 				return;
  // 			}
  // 			console.log("We've updated your picture!");
  // 		});
  // }, false);
  // if (file) {
  // 	//reader.readAsDataURL(file);
  // }

}

// function photoDownload(response, request, userId) {

//    userId = "e97f274a-2a86-4280-997d-8ee4d2c52078";
//     var client = microsoftGraph.Client.init({
//       authProvider: (done) => {
//         done(null, token);
//       }
//     });

//    client
//       .api('users/'+userId+'/photo/$value')
//       .responseType('blob')
//       .getStream((err, downloadStream) => {
//         let writeStream = fs.createWriteStream('../myPhoto.jpg');
//         downloadStream.pipe(writeStream).on('error', console.log);

//         if (err) {
//           console.log('error: ' + err);
//           response.write('<p>ERROR: ' + err + '</p>');
//           response.end();
//         } else {   

//       // let profilePhotoReadStream = fs.createReadStream('me.jpg');
//         //  console.log(downloadStream);
//         console.log("Image downloaded!")
//           response.end();
//         }
//       });

// }



function shareFile(response, request) {
  if (request.method == "POST") {

    var data = "";
    var body = "";

    var client = microsoftGraph.Client.init({
      authProvider: (done) => {
        done(null, token);
      }
    });

    request.on('data', function (input) {
      body += input;
      if (body.length > 1e6) {
        request.connection.destroy();
      }

    });

    response.write("200");
    response.end();

    request.on('end', function () {
      data = body;
      var dataArray = JSON.parse(data);
      var writer = csvWriter({ headers: ["DepartmentId", "DepartmentName", "ParentDepartment", "Navn"] })
      writer.pipe(fs.createWriteStream('out.csv', {flags: 'a'}))
      dataArray.forEach(function (element) {

        var depName = element["DepartmentName"];

        var depId = "";
        if(element["DepartmentId"] != "_Scurrenttime-department:departmentref"){
          depId = element["DepartmentId"];
        }
      
        var name = "";
        if(element["DepartmentHead"] != null){
          name = element["DepartmentHead"]["Navn"];
        }else {
          name = "NA"
        }

        var parentName = "";
        if (element["ParentDepartment"] != 0){
          parentName = element["ParentDepartment"][0]["ParentName"][0];
        }
     
        writer.write([depId, depName, parentName, name])
      }, this);
      writer.end()
    });


    fs.readFile("./out.csv", "utf8", function (err, data) {
      data = "\ufeff" + data;
      if (err) {
        throw err;
      } else {
        client
          .api('users/e97f274a-2a86-4280-997d-8ee4d2c52078/drive/root/children/out.csv/content')
          //  .api('users/e97f274a-2a86-4280-997d-8ee4d2c52078/drive/root/children/orgMap.csv/content')      
          //  .api('users/e97f274a-2a86-4280-997d-8ee4d2c52078/drive/items/01DP2XB3GMZQKCKZ6GKRFL5ZE3BCTVJJ5S/out.csv/content') 
          .top(10)    
          .put(data, (err, res) => {
            if (err) {
              console.log(err);
            }else {
              console.log("File updated!");
            }
           
          });
      }

    });
  }
}





  // var array = fs.readFileSync('./out.csv').toString().split("\n");
  // for(i in array) {
  //     console.log(array[i]);
  // }

  // fs.readFile('./out.csv', function read(err, data) { 
  //   if (err) {
  //         throw err;
  //     }
  //     content = data;
  //     // console.log(content);
  //     processFile(content);         
  // });



// function processFile(content) {
//     console.log(content);
// }