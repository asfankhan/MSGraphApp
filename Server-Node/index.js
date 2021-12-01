var express = require('express');
const cors = require('cors');
const path = require('path');
const bodyParser = require('body-parser')
var fs = require('fs');
var http = require('http');
var https = require('https');

var jsonParser = bodyParser.json()

const HTTP_SERVER_PORT = process.env.PORT || 80 ;
const HTTPS_SERVER_PORT = process.env.PORT || 443;

const SERVER_PORT = process.env.PORT || 8000;
const REDIRECT_URI = "http://localhost:8000/redirect";

// Create Self Sign keys
//  openssl req -x509 -nodes -days 365 -newkey rsa:2048 -keyout ./selfsigned.key -out selfsigned.crt

var privateKey = fs.readFileSync('selfsigned.key');
var certificate = fs.readFileSync('selfsigned.crt');
var credentials = {key: privateKey, cert: certificate};

// Create Express App and Routes
var app = express();
var httpServer = http.createServer(app);
var httpsServer = https.createServer(credentials, app);
app.use(cors());
app.options('*', cors());


// app.use(express.staticProvider(__dirname + '/build'));
app.get('/', (req, res) => {
    // res.sendFile(path.join(__dirname , '../Frontend-React/build/index.html'));
    res.sendStatus(200);
});

app.get('/redirect', (req, res) => {
    const tokenRequest = {
        code: req.query.code,
        scopes: ["user.read"],
        redirectUri: REDIRECT_URI,
    };

    pca.acquireTokenByCode(tokenRequest).then((response) => {
        console.log("\nResponse: \n:", response);
        res.sendStatus(200);
        // callMSGraph("https://graph.microsoft.com", "v1.0", "me", response.accessToken)
    }).catch((error) => {
        console.log(error);
        res.status(500).send(error);
    });
});


app.post('/GraphCall',jsonParser, (req, res) => {
    console.log('Got body:', req.body);
    res.sendStatus(200);
});

httpServer.listen(HTTP_SERVER_PORT, () => console.log(`Http Server listening on ${HTTP_SERVER_PORT}!`))
httpsServer.listen(HTTPS_SERVER_PORT, () => console.log(`Https Server listening on ${HTTPS_SERVER_PORT}!`))