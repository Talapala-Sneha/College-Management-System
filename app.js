const Nexmo = require('nexmo');
const express = require('express');
var XLSX = require('xlsx');
var workbook = XLSX.readFile('C:/Users/jayvi/msgProject/test.xlsx');
var nodemailer = require('nodemailer');
const cors = require('cors')
const BodyParser = require("body-parser");

const app = express();
app.use(cors());
app.use(BodyParser.json());
app.use(BodyParser.urlencoded({ extended: true }));
const nexmo = new Nexmo({
    apiKey: '054070e8',
    apiSecret: 'RNXjWHtR0eh3FGbK',
});
app.use(express.json());
//PORT
const port = process.env.PORT || 5000;

var sheet_name_list = workbook.SheetNames;
var data = [];
var result = [];

sendMessage = (number,res) => {
    const from = 'SSNCOE';
    const to = number;
    const text = 'Hello from SSN COLLEGE OF ENGINEERING \n';
    nexmo.message.sendSms(from, to, text);
    console.log("message Sent called and executed" + number);
    res.send("SUCCESS");
}

var transporter = nodemailer.createTransport({
    service: 'gmail',
    auth: {
        user: 'jayvishaalj.01@gmail.com',
        pass: 'jayvishaal144'
    }
});

var mailOptions = {}
sendMail = (sendTo,res) => {
        mailOptions = {
            from: 'jayvishaalj.01@gmail.com',
            to: sendTo,
            subject: 'Hi Test',
            text: 'Testing Mail',

        };
        console.log("Mail Part " + sendTo);
        transporter.sendMail(mailOptions, function (error, info) {
            if (error) {
                console.log(error);
            }
            else {
                console.log('Email Sent' + info.response);
                res.send("Sucess");
                
            }
        });

}

getData = () => {
    sheet_name_list.forEach(function (y) {
        var worksheet = workbook.Sheets[y];
        var headers = {};
        for (z in worksheet) {
            if (z[0] === '!') continue;
            //parse out the column, row, and value
            var tt = 0;
            for (var i = 0; i < z.length; i++) {
                if (!isNaN(z[i])) {
                    tt = i;
                    break;
                }
            };
            var col = z.substring(0, tt);
            var row = parseInt(z.substring(tt));
            var value = worksheet[z].v;

            //store header names
            if (row == 1 && value) {
                headers[col] = value;
                continue;
            }

            if (!data[row]) data[row] = {};
            data[row][headers[col]] = value;
        }
        //drop those first two rows which are empty
        data.shift();
        data.shift();
    });
}


sendMailFromExcel = (sendTo) => {
    sheet_name_list.forEach(function (y) {
        var worksheet = workbook.Sheets[y];
        var headers = {};
        for (z in worksheet) {
            if (z[0] === '!') continue;
            //parse out the column, row, and value
            var tt = 0;
            for (var i = 0; i < z.length; i++) {
                if (!isNaN(z[i])) {
                    tt = i;
                    break;
                }
            };
            var col = z.substring(0, tt);
            var row = parseInt(z.substring(tt));
            var value = worksheet[z].v;

            //store header names
            if (row == 1 && value) {
                headers[col] = value;
                continue;
            }

            if (!data[row]) data[row] = {};
            data[row][headers[col]] = value;
        }
        //drop those first two rows which are empty
        data.shift();
        data.shift();
        console.log(data);
        console.log(data.length);
        for (i = 0; i < data.length; i++) {
            console.log((data[i]['Phno']));
            console.log((data[i]['Email']));
            sendMessage(data[i]['Phno']);
            sendMail(data[i]['Email']);
        }
    });
}

app.get('/api/Class/:Id', (req, res) => {
    getData();
    flag = 0;
    result = [];
    for (i = 0; i < data.length; i++) {
        if (data[i]['ClassId'] == req.params.Id) {
            JsonData = { "Name": data[i]['Name'], "Id": data[i]['Id'], "Phno": data[i]['Phno'], "Email": data[i]['Email'] }
            result.push(JsonData)
            flag = 1;
        }
    }
    if (flag == 1) {
        res.send(result);
    }
    else {
        res.send("Not Found");
    }
    console.log(data);


});

app.get('/api/Student/:Id', (req, res) => {
    getData();
    for (i = 0; i < data.length; i++) {
        if (data[i]['Id'] == req.params.Id) {
            res.send({ "Name": data[i]['Name'], "Phno": data[i]['Phno'], "Email": data[i]['Email'] });
        }
    }
    console.log(data);
    // response = { 'Key': 'Value' };
    res.send("Not Found");

});

app.get('/api/SendMail/:Id', (req, res) => {
    getData();
    for (i = 0; i < data.length; i++) {
        if (data[i]['Id'] == req.params.Id) {
            resp = sendMail(data[i]['Email'],res);

        }
    }
});

app.get('/api/SendMsg/:Id', (req, res) => {
    getData();
    for (i = 0; i < data.length; i++) {
        if (data[i]['Id'] == req.params.Id) {
            resp = sendMessage(data[i]['Phno'],res);

        }
    }
});

app.listen(port, () => {
    console.log(`Listening on Port ${port}..`);
});