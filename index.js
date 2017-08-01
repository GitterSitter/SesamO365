// {
//   "_id": "salesforce",
//   "type": "system:microservice",
//   "name": "Salesforce",
//   "authentication": "basic",
//   "connect_timeout": 60,
//   "docker": {
//     "environment": {
//       "VERSION": "1"
//     },
//     "image": "sesambuild/salesforce:latest",
//     "port": 5000
//   },
//   "password": "$SECRET(salesforce-password)",
//   "read_timeout": 7200,
//   "use_https": false,
//   "username": "$SECRET(salesforce-secret)\\$SECRET(salesforce-user)",
//   "verify_ssl": false
// }


var server = require('./server');
var router = require('./router');
var request = require("request");
var microsoftGraph = require("@microsoft/microsoft-graph-client");
var fs = require('fs');
var FileReader = require('filereader')
var url = require('url');
var auth = require('./auth');
var handle = {};
handle['/mail'] = userEmail;
handle['/calendar'] = calendar;
handle['/contacts'] = contacts;
handle['/photo'] =  updateProfilePicture; //photoDownload;
handle['/users'] = users;


server.start(router.route, handle);

var token = "";
function saveToken(tok){
token = tok;
}

auth.getAccessToken().then(function (token) {
  // Get all of the users in the tenant.
   // console.log(token);
  saveToken(token)
    .then(function (tok) {    
      // Create an event on each user's calendar.
     // graph.createEvent(token, users);
    }, function (error) {
  
      console.error('>>> Error getting users: ' + error);
    });
}, function (error) {
  console.error('>>> Error getting access token: ' + error);
});


function userEmail(userId){
userId = "e97f274a-2a86-4280-997d-8ee4d2c52078";
  // Create a Graph client
  var client = microsoftGraph.Client.init({
    authProvider: (done) => {
      // Just return the token
      done(null, token);
    }
  });
  // Get the Graph /Me endpoint to get user email address
  client
    .api('/users/'+ userId+ '/mail')
    .get((err, res) => {
      if (err) {
      
       console.log(err);
      } else {
      
       console.log(res.value);
      
      }
     
    });

}


function users(response, request) {
  console.log(request.method);

  var userId =  "e97f274a-2a86-4280-997d-8ee4d2c52078"; //"30be01d3-8214-4f2d-aea0-7028a19581fc" ;//"e97f274a-2a86-4280-997d-8ee4d2c52078";
  var client = microsoftGraph.Client.init({
        authProvider: (done) => {
          // Just return the token
          done(null, token);
        }
      });
 //users/30be01d3-8214-4f2d-aea0-7028a19581fc/mobilePhone               
  //request.method

  //tlf nr funker! /mobilePhone
//If you want a different set of properties, you can request them using the $select query parameter. E.g https://graph.microsoft.com/v1.0/users/e97f274a-2a86-4280-997d-8ee4d2c52078?$select=aboutMe
  //Når AD brukes er det ikke mulig å gjøre endringer! Man kan kun gjøre GET requests. Ellers må man oppdatere direkte i AD.
  //Azure Ad Graph Api kan brukes for å gjøre endringer på brukere, grupper og kontakter i AD.
  if("POST" == "POST"){
    client.api("/users/"+userId+"/displayName")  //Skaflestad britt.skaflestad@bouvet.no
       .patch(
        {"value": "Test"},
        (err, res) => {
            if (err)
                console.log(err);
            else
                console.log("Profile Updated");
        });
  }else{
      client
        .api('/users')
       // .header('X-AnchorMailbox', email)
        // .top(20)
        // .select('subject,from,receivedDateTime,isRead')
        // .orderby('receivedDateTime DESC')
        .get((err, res) => {
          if (err) {
            console.log('getUsers returned an error: ' + err);
            response.write('<p>ERROR: ' + err + '</p>');
            response.end();
          } else {
          
            response.write('');
            // res.value.forEach(function(message) {
            //   console.log('User: ' + message);
              
            //   response.write(message);
            // });
           // console.log(res.value);
              
            response.end();
          }
        });

        }
}


function calendar(userId) {
  userId = "e97f274a-2a86-4280-997d-8ee4d2c52078";

   var client = microsoftGraph.Client.init({
        authProvider: (done) => {
          // Just return the token
          done(null, token);
        }
      });
      client
        .api('/users/'+userId+'/events')
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
          // Just return the token
          done(null, token);
        }
      });
      client
        .api('/users/'+userId+'/contacts')
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

     var userId =  "30be01d3-8214-4f2d-aea0-7028a19581fc";  //(britt)    "8fa20769-13f0-4b67-b777-c262b174d93e"; (Eirik?)  //e97f274a-2a86-4280-997d-8ee4d2c52078  (min)
     var file = fs.readFileSync('./logo.png'); // fs.openSync("logo.png","r"); //new File("logo.png");       
     //var reader = new FileReader();
    	client
					.api("/users/"+userId+"/photo/$value")
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
    
function photoDownload(response, request, userId) {

  // Get the profile photo of the current user (from the user's mailbox on Exchange Online).
  // This operation in version 1.0 supports only work or school mailboxes, not personal mailboxes.
  
   userId = "e97f274a-2a86-4280-997d-8ee4d2c52078";
    var client = microsoftGraph.Client.init({
      authProvider: (done) => {
        // Just return the token
        done(null, token);
      }
    });

   client
      .api('users/'+userId+'/photo/$value')
      .responseType('blob')
      //.get((err, res,rawResponse) => {
      .getStream((err, downloadStream) => {
        let writeStream = fs.createWriteStream('../myPhoto.jpg');
        downloadStream.pipe(writeStream).on('error', console.log);
    
        if (err) {
          console.log('error: ' + err);
          response.write('<p>ERROR: ' + err + '</p>');
          response.end();
        } else {   

      // let profilePhotoReadStream = fs.createReadStream('me.jpg');
        //  console.log(downloadStream);
        console.log("Image downloaded!")
          response.end();
        }
      });
}


