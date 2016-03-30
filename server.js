/*
 * Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

var path = require('path');
var fs = require('fs');
var https = require('https');
var express = require('express');
var app = express();

// Set the address and the certificate.
var options = {
    hostname: 'localhost',
    key: fs.readFileSync('server.key'),
    cert: fs.readFileSync('server.crt'),
    ca: fs.readFileSync('ca.crt')
};

// Define the port. The service uses 'localhost' as the host address.
// Set the host member in the options object to set a custom host domain name or IP address.
var port = 8080;

// Set the front-end folder to serve public assets.
app.use(express.static(__dirname + '/scripts'));

// Set the front-end folder to serve public assets.
app.use(express.static(__dirname + '/media'));

// Set the front-end folder to serve public assets.
app.use(express.static(__dirname + '/css'));

// Set the front-end folder to serve public assets.
app.use(express.static(__dirname));

// Set up route to get the spec template.
app.get('/gettemplate', function(req, res) {

    // Create path to get the file. 'specs' is the directory where the spec templates are stored.
    // MiniSpecTemplate.docx is the name of the spec template document.
    var pathToFile = path.join(__dirname, 'spec', 'MiniSpecTemplate.docx');

    // Read the file, convert it to base64, and return the base64 file in the response to the add-in.
    fs.readFile(pathToFile, function(err, data) {
        var fileData = new Buffer(data).toString('base64');
        res.send(fileData);
    });
});

// Set up route to get the blacklist.
app.get('/blacklist', function(req, res) {

    // Create path to get the file. 'data' is the directory where the blacklist is stored.
    var pathToFile = path.join(__dirname, 'data', 'blacklist.json');

    // Read the file, and return the JSON in the response to the add-in.
    fs.readFile(pathToFile, function(err, data) {
        var fileData = new Buffer(data).toString();
        res.send(fileData);
    });
});

// Set up GET route to get the boilerplate json.
app.get('/boilerplate', function(req, res) {

    // Create path to get the file. 'data' is the directory where the boilerplate is stored.
    var pathToFile = path.join(__dirname, 'data', 'boilerplate.json');

    // Read the file, and return the JSON in the response to the add-in.
    fs.readFile(pathToFile, function(err, data) {
        var fileData = new Buffer(data).toString();
        res.send(fileData);
    });
});

// Set up POST route to save the boilerplate json.
app.post('/boilerplate', function(req, res) {

    // Create path to get the file. 'data' is the directory where the boilerplate is stored.
    var pathToFile = path.join(__dirname, 'data', 'boilerplate.json');

    req.on('data', function(data) {

        fs.writeFile(pathToFile, data.toString(), function(err) {
            if (err) {
                console.log(err);
                res.sendStatus(500);
            } else {
                console.log('Boilerplate has been saved.');
                res.sendStatus(200);
            }
        });
    });
});

// Set the route to the index.html file.
app.get('/', function(req, res) {
    var homepage = path.join(__dirname, 'index.html');
    res.sendFile(homepage);
});

// Set the route for the HTML served to the dialog API call.
app.get('/dialog', function(req, res) {
    var homepage = path.join(__dirname, 'dialog.html');
    res.sendFile(homepage);
});

// Start the server.
https.createServer(options, app).listen(port, function() {
    console.log('Listening on https://localhost:' + port + '...');
});
