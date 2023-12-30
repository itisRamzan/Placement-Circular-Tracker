const express = require("express");
const app = express();
const circularTracker = require("./script");
const port = process.env.PORT || 5000;

circularTracker();

app.listen(port, () => console.log(`Server up and running on port ${port}!`));