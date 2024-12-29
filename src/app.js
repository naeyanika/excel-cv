const express = require('express');
const convertRouter = require('./routes/convertRouter');

const app = express();

app.use(express.static('public'));
app.use(convertRouter);

module.exports = app;