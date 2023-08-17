
const fs = require('fs');
const express = require('express');
const mongoose = require('mongoose');
const multer = require('multer');
const path = require('path');
const userModel = require('./models/userModel');
const excelToJson = require('convert-excel-to-json');
const bodyParser = require('body-parser');
const async = require('async');
const { log } = require('console');

const storage = multer.diskStorage({
  destination: (req, file, cb) => {
    cb(null, 'public/uploads/');
  },
  filename: (req, file, cb) => {
    cb(null, file.originalname);
  }
});

const uploads = multer({ storage: storage });

const uri = "mongodb://127.0.0.1:27017/tommaacgae";

// Connect to MongoDB
mongoose.connect(uri, { useNewUrlParser: true }).then(() => console.log('Connected to MongoDB')).catch((err) => console.log(err))


const app = express();


app.use(bodyParser.urlencoded({ extended: true }));
app.use(express.static(path.join(__dirname, 'public')));


app.get('/', (req, res) => {
  res.sendFile(__dirname + '/index.html');
});


app.post('/uploadfile', uploads.single("uploadfile"), (req, res) => {
  const filePath = __dirname + '/public/uploads/' + req.file.filename;

  // -> Read Excel File to Json Data
  const excelData = excelToJson({
    sourceFile: filePath,
    sheets: [{
      // Excel Sheet Name
      name: 'Sheet1',

      header: {
        rows: 1
      },
      // Mapping columns to keys
      columnToKey: {
        A: 'name',
        B: 'Email',
        C: 'MobileNo',
        D: 'DateOfBirth',
        E: 'WorkExperience',
        F: 'ResumeTitle',
        G: 'CurrentLocation',
        H: 'PostalAddress',
        I: 'CurrentEmployer',
        J: 'CurrentDesignation'
      }
    }]
  });

  // print excel data on console
  console.log(excelData);

  const dataArray = excelData.Sheet1;
 

  let uniqueEmails = new Set();

  const callback = (err) => {
    if (err) {
      console.log("Error occurred while saving data to database:", err);
    } else {
      console.log("All records saved successfully!");
      res.redirect('/');
    }
    fs.unlinkSync(filePath);
  }


  async.eachSeries(dataArray, async (data,callback) => {

    const email = data.Email;

    if (uniqueEmails.has(email)) {
      console.log(`Duplicate email found: ${email}`);
      return;
    }

    uniqueEmails.add(email);

    try {

      const existingUser = await userModel.findOne({ Email: email });
      if (existingUser) {
        console.log(`${email} is already inserted into database`);
        return;
      }

      const user = new userModel(data);
      await user.save();
      console.log(`Successfully inserted record for ${email}`);
      console.log(user+" "+" user");
   //   return;

    }
    catch (err) {
      console.log(err);
      callback(err);
    }
  },callback);
});


// Assign port
const port = process.env.PORT || 3000;

app.listen(port, () => console.log(`server running on ${port}`));


