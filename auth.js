
var request = require('request');
var Q = require('q');

var tokenEndpoint = 'https://login.windows.net/c317fa72-b393-44ea-a87c-ea272e8d963d/oauth2/token';
var clientId = 'b2e9e676-4110-4340-ae4c-21742e848f3d';
var clientSecret = process.env.Token_Node_Office;

// The auth module object.
var auth = {};

// @name getAccessToken
// @desc Makes a request for a token using client credentials.
auth.getAccessToken = function () {
  var deferred = Q.defer();

var result = "";
  // These are the parameters necessary for the OAuth 2.0 Client Credentials Grant Flow.
  // For more information, see Service to Service Calls Using Client Credentials (https://msdn.microsoft.com/library/azure/dn645543.aspx).
  var requestParams = {
    grant_type: 'client_credentials',
    client_id: clientId,
    client_secret: clientSecret,
    resource: 'https://graph.microsoft.com'
  };

  // Make a request to the token issuing endpoint.
  request.post({ url: tokenEndpoint, form: requestParams }, function (err, response, body) {
    var parsedBody = JSON.parse(body);
    if (err) {
     deferred.reject(err);
      result =  err; // deferred.reject(err);
    } else if (parsedBody.error) {
     deferred.reject(parsedBody.error_description);
      result = parsedBody.error_description; //deferred.reject(parsedBody.error_description);
    } else {
      // If successful, return the access token.
      deferred.resolve(parsedBody.access_token);

      result = parsedBody.access_token; //parsedBody.access_token;
      
    }

  });

  return deferred.promise;
};

module.exports = auth;
