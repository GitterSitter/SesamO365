var server = require('./server');
var router = require('./router');
var request = require("request");
var reqUrl = require('request').defaults({ encoding: null });

var microsoftGraph = require("@microsoft/microsoft-graph-client");
var fs = require('fs');
var FileReader = require('filereader')
var url = require('url');
var auth = require('./auth');
var request = require('request');
var qs = require('querystring');
var csvWriter = require('csv-write-stream')
var token = "";
var userStatusArray = [];
var orgDataArray = [];
var checked = false;

var handle = {};
handle['/photo'] = updateProfilePicture;
handle['/users'] = users;
handle['/groups'] = groups;
handle['/file'] = shareFile;
handle['/status'] = userStatus;

server.start(router.route, handle);

//Requesting a new token every second hour as the old one expires
function refreshToken() {
  getNewToken();
}
setInterval(refreshToken, 60 * 60 * 1000);

function saveToken(tok) {
  token = tok;
}

//If expired, request new token in the methods!
auth.getAccessToken().then(function (token) {
  saveToken(token)
    .then(function (tok) {
    }, function (error) {
      console.error('>>> Error getting users: ' + error);
    });
}, function (error) {
  console.error('>>> Error getting access token: ' + error);
});


function getNewToken() {
  auth.getAccessToken().then(function (token) {
    saveToken(token)
      .then(function (tok) {
      }, function (error) {
        console.error('>>> Error getting users: ' + error);
      });
  }, function (error) {
    console.error('>>> Error getting access token: ' + error);
  });
}


function userStatus(response, request) {

  var client = microsoftGraph.Client.init({
    authProvider: (done) => {
      done(null, token);
    }
  });

  if (request.method == "POST") {
    var body = [];
    request.on('data', function (input) {
      body += input;
      if (body.length > 1e6) {
        request.connection.destroy();
      }

      var userMail = [];
      var userArray = JSON.parse(body);
      var counter = 0;

      if (userArray.length === 0) {
        response.writeHead(200, { "Content-Type": "application/json" });
        response.end("No data");
        return;
      }

      console.log("request batch size: " + userArray.length);
      userArray.forEach(function (element) {
        var id = element["id"];
        var name = element["displayName"];

        client.api("https://graph.microsoft.com/beta/users/" + id + "/mailboxSettings/automaticRepliesSetting?pretty=1")
          .get((err, res) => {
            if (err) {
              console.log(name + " has got no mail account!");
              ++counter;
            } else {

              if (res["status"] != "disabled") {
                res.id = id;
                userMail.push(res);
                userStatusArray.push(res);
              }
              ++counter;
            }
            if (counter === userArray.length) {
              console.log("Instances: " + userMail.length);
              console.log("200 OK");
              response.writeHead(200, { "Content-Type": "application/json" });
              response.end("200");
            }

          });

      });
    });


  } else if (request.method == "GET") {
    console.log("Amount of users with status: " + userStatusArray.length);
    if (userStatusArray.length > 0) {
      var batchResponse = [];

      if (userStatusArray.length < 100) {
        console.log("Reached last elements:" + userStatusArray.length);
        response.writeHead(200, { "Content-Type": "application/json" });
        response.end(JSON.stringify(userStatusArray));
        return;

      } else {
        var counter = userStatusArray;
        for (let element of userStatusArray) {
          batchResponse.push(element);
          counter.splice(element, 1);

          if (batchResponse.length == 100) {
            console.log(200);
            response.writeHead(200, { "Content-Type": "application/json" });
            response.end(JSON.stringify(batchResponse));
            batchResponse = [];

          }

          if (counter.length < 100) {
            response.writeHead(200, { "Content-Type": "application/json" });
            response.end(JSON.stringify(counter));
            return;
          }
        }

      }

    } else {
      console.log("No data");
      response.writeHead(200, { "Content-Type": "application/json" });
      response.end("No data");
      return;
    }
  }
}



function groups(response, request) {

  var client = microsoftGraph.Client.init({
    authProvider: (done) => {
      done(null, token);
    }
  });

  if (request.method == "GET") {
    client
      .api("https://graph.microsoft.com/beta/groups")
      .top(999)
      .get((err, res) => {
        if (err) {
          console.log(err);
          response.writeHead(500, { "Content-Type": "application/json" });
          response.end(res.statusCode + " - " + err);

        } else if ('@odata.nextLink' in res) {
          var data = [];
          getNextPage(res, response, client, data);

        } else {
          console.log("200 OK");
          response.writeHead(200, { "Content-Type": "application/json" });
          response.end(JSON.stringify(res.value));
        }
      });
  }
}


function users(response, request) {

  var client = microsoftGraph.Client.init({
    authProvider: (done) => {
      done(null, token);
    }
  });

  if (request.method == "POST") {
    //tlf nr funker! /mobilePhone
    //If you want a different set of properties, you can request them using the $select query parameter. E.g https://graph.microsoft.com/v1.0/users/e97f274a-2a86-4280-997d-8ee4d2c52078?$select=aboutMe
    //Når AD brukes er det ikke mulig å gjøre endringer! Man kan kun gjøre GET requests. Ellers må man oppdatere direkte i AD.
    //Azure Ad Graph Api kan brukes for å gjøre endringer på brukere, grupper og kontakter i AD.
    // On premise AD hindrer updates

    var userId = request.data;
    client.api("/users/" + userId + "/displayName")
      .patch(
      { "value": "Test" },
      (err, res) => {
        if (err)
          console.log(err);
        else
          console.log("Profile Updated");
      });
  } else if (request.method == "GET") {
    client
      .api('https://graph.microsoft.com/beta/users?$filter=accountEnabled eq true')
      .top(999)
      .get((err, res) => {
        if (err) {
          console.log(err);
          response.writeHead(500, { "Content-Type": "application/json" });
          response.end();
        } else if ('@odata.nextLink' in res) {
          var data = [];
          getNextPage(res, response, client, data);
        } else {
          console.log("200 OK");
          response.writeHead(200, { "Content-Type": "application/json" });
          response.end(JSON.stringify(res.value));
        }
      });
  }
}


function getNextPage(result, response, client, data) {
  var completeResult = data;
  completeResult = data.concat(result.value);

  if (result['@odata.nextLink']) {
    client.api(result['@odata.nextLink'])
      .get((err, res) => {
        if (err) {
          console.log(err);
          response.writeHead(500, { "Content-Type": "application/json" });
          response.end();
          return;
        } else {

          completeResult.concat(res.value);
          getNextPage(res, response, client, completeResult)
        }
      });

  } else {
    console.log("200 OK");
    response.writeHead(200, { "Content-Type": "application/json" });
    response.end(JSON.stringify(completeResult));
    return;
  }

}



function updateProfilePicture(response, request) {
  if (request.method == "POST") {
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

      var data = JSON.parse(body);

      if (data.length === 0) {
        response.end("no data");
        return;
      }

      data.forEach(function (element) {
        var image = "";
        var test = "";
        var userId = element["id"];
        var userName = element["name"];

        if (element["image"] != null) {
          test = element["image"];
          if (test["fit_thumb"]["url"] != null) {
            image = test["fit_thumb"]["url"];
          }
        }

        if (image === null || image === "") {

          response.end("is skipped because of no picture");
          console.log(userName + " is skipped because of no picture");
        } else {
          download(image, userId + '.png', function () {
            var img = fs.readFile(userId + '.png', function (err, data) {
              if (err) {
                console.log(+"Error downloading file: " + err);
                return;

              } else {
                console.log("Image downloaded!");
                client.api("/users/" + userId + "/photo/$value")
                  .put(data, (err, res) => {
                    if (err) {
                      console.log(err);
                      console.log("Error setting downloaded profile image for user " + userName);
                      response.end("Error setting downloaded profile image for user " + userName);
                      return;
                    } else {
                      response.end("image updated!");
                      console.log(userName + "s image updated!");

                      fs.unlink("./" + userId + '.png', function (err) {
                        if (err) {
                          console.log("Cant remove file!");
                        } else {
                          console.log(userId + '.png' + " deleted");
                        }

                      });
                    }
                  });

              }

            });
          });

        }

      });

    });

  }

}


var download = function (uri, filename, callback) {
  request.head(uri, function (err, res, body) {
    request(uri).pipe(fs.createWriteStream(filename)).on('close', callback);
  });
};


function shareFile(response, request) {
  if (request.method == "POST") {
    var body = "";
    var is_last = false;

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

      is_last = request.url.includes("is_last=true");
      var dataArray = JSON.parse(body);
      if (dataArray.length != 0) {
        orgDataArray = orgDataArray.concat(dataArray);

      }

    
      if (is_last) {

        var writer;
        if(!checked){
           writer = csvWriter({headers: ["DepartmentId", "DepartmentName", "ParentDepartment", "Navn"] });
        }else {
           writer = csvWriter({headers: [" ", " ", " ", " "] });
        }
        writer.pipe(fs.createWriteStream('orgMap.csv', { flags: 'a' }));
        orgDataArray = orgDataArray.filter(function (item, index, inputArray) {
          return inputArray.indexOf(item) == index;
        });
  
        orgDataArray.forEach(function (element) {
          var parentName = "No Department Parent";
          var depId = "No Department Id";
          var nameDepartmentHead = "No Department Head";
          var depName = "No Department Name";

          if (element["DepartmentName"] != null) {
            depName = element["DepartmentName"];
          }

          if (element["DepartmentId"] != "_Scurrenttime-department:departmentref" && element["DepartmentId"] != null) {
            depId = element["DepartmentId"];
          }

          if (element["DepartmentHead"] != null) {
            nameDepartmentHead = element["DepartmentHead"]["Navn"];
          }

          if (typeof element["ParentDepartment"][0] != 'undefined' && element["ParentDepartment"][0]["ParentName"][0] != null) {
            parentName = element["ParentDepartment"][0]["ParentName"][0];
          }
          writer.write([depId, depName, parentName, nameDepartmentHead]);

        }, this);
        checked = true;
        writer.end();
        console.log("Ended Writing..");
      }

    });

    response.write("200");
    response.end();
    request.on('end', function () {
      if (is_last) {
        console.log("Started reading..");
        readOrgFile(client);
      }
    });
  }
}

function readOrgFile(client) {
  fs.readFile("./orgMap.csv", "utf8", function (err, data) {
    data = "\ufeff" + data;
    if (err) {
      console.log(err);
    } else {
      client
        .api('groups/2fe68adf-397c-4c85-90bb-4fd64544680d/drive/root/children/orgMap.csv/content')
        .put(data, (err, res) => {
          if (err) {
            console.log(err);
          } else {
            orgDataArray = [];
            console.log("File updated!");
          }
        });
    }
  });
}