const express = require("express");
let { PythonShell } = require("python-shell");
const schedule = require("node-schedule");

const app = express();
const port = 3000;

//API
app.get("/procesar", (req, res) => {
  PythonShell.run("main.py", null, function (err,results) {
    if (err) throw err;
    res.send(`finish scheduler${new Date()}`);
    console.log('results: %j', results);

  });
});

app.get("/", (req, res) => {
  res.send(`RPA está corriendo`);
});

app.listen(port, () => {
  console.log(`listening at http://localhost:${port}`);
});

//SCHEDULER
const job = schedule.scheduleJob("0 3,7,11,15,19 * * *", function (fireDate) {
  console.log(`running scheduler${new Date()}`);
  PythonShell.run("main.py", null, function (err) {
    if (err) throw err;
    console.log(`finish scheduler${new Date()}`);
  });
});
