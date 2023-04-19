const express = require('express');
const multer = require('multer');
const xlsx = require('xlsx');
const nodemailer = require('nodemailer');
const validator = require('email-validator');

const app = express();
const upload = multer({ dest: 'uploads/' });

// Serve the HTML form for uploading a file
app.get('/', (req, res) => {
  res.send(`
    <form method="POST" enctype="multipart/form-data">
      <input type="file" name="file" />
      <button>Upload</button>
    </form>
  `);
});

// Handle the file upload and email verification
app.post('/', upload.single('file'), async (req, res) => {
  const { file } = req;

  const transporter = nodemailer.createTransport({
    host: 'smtp.gmail.com',
    port: 587,
    secure: false,
    auth: {
      user: 'minimane138@gmail.com',
      pass: 'wavdjfffjhpjnczm'
    }
  });

  try {
    const workbook = xlsx.readFile(file.path);
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const emailAddresses = xlsx.utils.sheet_to_json(sheet, { header: ['email'] });

    const results = [];

    for (const { email } of emailAddresses) {
      // Check if email is valid
      if (!validator.validate(email)) {
        console.log(`Email address ${email} is invalid`);
        results.push({ email, message: 'Invalid email address' });
        continue;
      }

      // Step 1: Initiate SMTP session
      const info = await transporter.verify();

      // Step 2-8: Send SMTP commands and receive responses
      const response = await transporter.sendMail({
        from: 'minimane138@gmail.com',
        to: email
      });

      // Step 9: Close SMTP session
      transporter.close();

      // Check the response from the email server to see if the email address exists
      if (response.accepted.includes(email)) {
        console.log(`Email address ${email} exists`);
        results.push({ email, message: 'Email address exists' });
      } else {
        console.log(`Email address ${email} does not exist`);
        results.push({ email, message: 'Email address does not exist' });
      }
    }

    res.send(`
      <h1>Email verification complete</h1>
      <ul>
        ${results.map(({ email, message }) => `<li>${email}: ${message}</li>`).join('')}
      </ul>
    `);
  } catch (error) {
    console.log(`Error: ${error}`);
    res.status(500).send('Internal server error');
  }
});

app.listen(3000, () => {
  console.log('Server running on port 3000');
});
