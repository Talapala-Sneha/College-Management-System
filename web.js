const Nexmo = require('nexmo');
const express = require('express');
var XLSX = require('xlsx');
var workbook = XLSX.readFile('C:/Users/jayvi/msgProject/test.xlsx');
var nodemailer = require('nodemailer');

const app = express();
const nexmo = new Nexmo({
  apiKey: '054070e8',
  apiSecret: 'RNXjWHtR0eh3FGbK',
});

var sheet_name_list = workbook.SheetNames;

sendMessage = (number) =>{
  const from = 'SSNCOE';
  const to = number;
  const text = 'Hello from SSN COLLEGE OF ENGINEERING \n';
  nexmo.message.sendSms(from, to, text);
  console.log("message Sent called and executed"+number);
}

//set this link
// https://myaccount.google.com/lesssecureapps
var transporter = nodemailer.createTransport({
  service: 'gmail',
  auth: {
    user: 'jayvishaalj.01@gmail.com',
    pass: 'jayvishaal144'
  }
});

var mailOptions ={}

sendMailFromExcel = (sendTo) =>{
  mailOptions = {
    from: 'jayvishaalj.01@gmail.com',
    to: sendTo,
    subject: 'Hi Test',
    text: 'Testing Mail',
  
  };
  console.log("Mail PArt "+sendTo);
  transporter.sendMail(mailOptions,function(error,info){
    if(error){
      console.log.log(error);
    }
    else{
      console.log('Email Sent'+info.response);
    }
  });
}


sheet_name_list.forEach(function(y) {
    var worksheet = workbook.Sheets[y];
    var headers = {};
    var data = [];
    for(z in worksheet) {
        if(z[0] === '!') continue;
        //parse out the column, row, and value
        var tt = 0;
        for (var i = 0; i < z.length; i++) {
            if (!isNaN(z[i])) {
                tt = i;
                break;
            }
        };
        var col = z.substring(0,tt);
        var row = parseInt(z.substring(tt));
        var value = worksheet[z].v;

        //store header names
        if(row == 1 && value) {
            headers[col] = value;
            continue;
        }

        if(!data[row]) data[row]={};
        data[row][headers[col]] = value;
    }
    //drop those first two rows which are empty
    data.shift();
    data.shift();
    console.log(data);
    console.log(data.length);
    for(i=0;i<data.length;i++)
    {
      console.log((data[i]['Phno']));
      console.log((data[i]['Email']));
      sendMessage(data[i]['Phno']);
      sendMailFromExcel(data[i]['Email']);
    }  
  });




