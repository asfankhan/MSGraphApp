const http = require('http');
var express = require('express');

var app = express();
const port = process.env.PORT || 1337;

var httpServer = http.createServer(app);

app.get('/', (request, response) => {
    response.writeHead(200, {"Content-Type": "text/plain"});
    response.end("Hello version 3!");
});


app.post('/GraphCall', async (req, res) => {

    if (req.query && req.query.validationToken) {
        res.set('Content-Type', 'text/plain');
        res.send(req.query.validationToken);
        return;
    }
    if (areTokensValid) {
        for (let i = 0; i < req.body.value.length; i++) {
          const notification = req.body.value[i];
    
          // Verify the client state matches the expected value
          if (notification.clientState == process.env.SUBSCRIPTION_CLIENT_STATE) {
            // Verify we have a matching subscription record in the database
            const subscription = await dbHelper.getSubscription(
              notification.subscriptionId
            );
            if (subscription) {
              // If notification has encrypted content, process that
              if (notification.encryptedContent) {
                processEncryptedNotification(notification);
              } else {
                await processNotification(
                  notification,
                  req.app.locals.msalClient,
                  subscription.userAccountId
                );
              }
            }
          }
        }
    }
    
    res.status(202).end();

});

httpServer.listen(port, () => console.log(`Http Server listening on ${port}!`))

console.log("Server running at http://localhost:%d", port);
