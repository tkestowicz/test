var express = require('express')
var app = express();
var parser = require('body-parser');
var DocxGen = require('./generator');
var path = require('path');
var fs = require('fs');
var cors = require('express-cors');

app.use(cors({
    allowedOrigins: [
        'http://localhost:4200',
        'http://api.trenerkakobiet.pl',
    ]
}))

app.use(parser.json());

var monthNames = ["Styczeń", "Luty", "Marzec", "Kwiecień", "Maj", "Czerwiec", "Lipiec", "Sierpień", "Wrzesień", "Październik", "Listopad", "Grudzień"];

app.get('/schedule/download/:uid', function (req, res) {

    var uid = req.params.uid;

    var month = req.query.month;

    res.set({
        "Content-Disposition": `'attachment; filename="grafik-${monthNames[month].toLowerCase()}.docx"`,
    });

    var filename = path.resolve(__dirname, 'schedules', uid+'.docx');

    res.sendFile(filename);

    //var file = fs.createReadStream(filename);
    //file.on('end', function () {
    //    fs.unlink(filename, function () {
    //        
    //    });
    //});

    //file.pipe(res);
});

function weekCount(year, month_number) {

    // month_number is in the range 1..12

    var firstOfMonth = new Date(year, month_number - 1, 1);
    var lastOfMonth = new Date(year, month_number, 0);

    var used = firstOfMonth.getDay() + 6 + lastOfMonth.getDate();

    return Math.ceil(used / 7);
}

function daysCount(year, month_number) {
    return new Date(year, month_number, 0).getDate();
}

function getWeekNo(date) {

    var day = date.getDate()

    //get weekend date
    day += (date.getDay() == 0 ? 0 : 7 - date.getDay());

    return Math.ceil(parseFloat(day) / 7);
}

app.post('/schedule/generate', function (req, res) {

    var input = req.body;

    var byMonth = [];

    input.map(x => {

        x.day = new Date(x.day);

        var month = x.day.getMonth(),
            year = x.day.getFullYear();

        var found = byMonth.find(m => m.month === x.day.getMonth());

        if (found) {

            found.items.push(x);

        } else {

            byMonth.push({
                month: month,
                weeks: weekCount(year, month),
                days: daysCount(year, month + 1),
                year: year,
                items: [x]
            });

        }
    });

    var dayKeys = ["sun", "mon", "tue", "wed", "thu", "fri", "sat"];

    var output = [];

    const monday = 1;

    byMonth.map(function (month) {

        var days = 1,
            monthIndex = byMonth.indexOf(month);

        output[monthIndex] = {
            month: month.month,
            weeks: []
        };

        for (var day = 1, currentWeek = -1; day <= month.days; day++) {
            var item = month.items.find(x => x.day.getDate() == day);

            var dayOfMonth = new Date(month.year, month.month, day);

            var formattedDay = dayOfMonth.toLocaleDateString('pl-PL', {
                day: '2-digit',
                month: '2-digit'
            }).split(' ').join('.');

            var meeting = {};

            if (item) {

                var time = (item.time)? ` - ${item.time}` : '';

                meeting = {
                    title: (item.title || '').toUpperCase(),
                    day: `${formattedDay}${time}`,
                    desc: item.desc || ''
                };
            } else {

                meeting = {
                    title: '',
                    day: `${formattedDay}`,
                    desc: ''
                };
            }

            if (currentWeek < getWeekNo(dayOfMonth) - 1) {
                currentWeek = getWeekNo(dayOfMonth) - 1;

                output[monthIndex].weeks[currentWeek] = {};
            }

            var week = output[monthIndex].weeks[currentWeek];

            if (!week) {
                week = {};
            }

            week[dayKeys[dayOfMonth.getDay()]] = meeting;

            output[monthIndex].weeks[currentWeek] = week;
        }
    });

    output.map(month => month.weeks.map(week => {
        dayKeys.map(day => {
            if (Object.keys(week).indexOf(day) < 0) {
                return week[day] = {
                    title: '',
                    day: '',
                    desc: ''
                };
            }
        })
    }));

    try {

        var reports = [];

        output.map(month => {
            var uid = Date.now();
            var outputFile = uid + '.docx';
            var outputDir = 'schedules';

            var gen = new DocxGen();

            gen.generate('template.docx', {
                month: monthNames[month.month].toUpperCase(),
                weeks: month.weeks
            });

            if(!fs.existsSync(path.resolve(__dirname, outputDir))) {
                fs.mkdirSync(path.resolve(__dirname, outputDir));
            }

            gen.saveAs(path.resolve(__dirname, outputDir, outputFile));

            reports.push({
                uid: uid,
                month: month.month
            });
        });

        res.json(reports);
    }
    catch (error) {
        var e = {
            message: error.message,
            name: error.name,
            stack: error.stack,
            properties: error.properties,
        }
        console.log(JSON.stringify({ error: e }));
        // The error thrown here contains additional information when logged with JSON.stringify (it contains a property object).
        res.sendStatus(500);
    }
})

var debug = require("debug")("express:server");
var http = require("http");

//get port from environment and store in Express.
var port = normalizePort(process.env.PORT || 8080);
app.set("port", port);

//create http server
var server = http.createServer(app);

//listen on provided ports
server.listen(port);

//add error handler
server.on("error", onError);

//start listening on port
server.on("listening", onListening);

/**
 * Normalize a port into a number, string, or false.
 */
function normalizePort(val) {
  var port = parseInt(val, 10);

  if (isNaN(port)) {
    // named pipe
    return val;
  }

  if (port >= 0) {
    // port number
    return port;
  }

  return false;
}

/**
 * Event listener for HTTP server "error" event.
 */
function onError(error) {
  if (error.syscall !== "listen") {
    throw error;
  }

  var bind = typeof port === "string"
    ? "Pipe " + port
    : "Port " + port;

  // handle specific listen errors with friendly messages
  switch (error.code) {
    case "EACCES":
      console.error(bind + " requires elevated privileges");
      process.exit(1);
      break;
    case "EADDRINUSE":
      console.error(bind + " is already in use");
      process.exit(1);
      break;
    default:
      throw error;
  }
}

/**
 * Event listener for HTTP server "listening" event.
 */
function onListening() {
  var addr = server.address();
  var bind = typeof addr === "string"
    ? "pipe " + addr
    : "port " + addr.port;
  debug("Listening on " + bind);
}

module.exports = app;