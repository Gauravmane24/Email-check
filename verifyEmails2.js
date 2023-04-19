const { SMTPConnection } = require('smtp-connection');

app.post('/', upload.single('file'), async (req, res) => {
  const { file } = req;

  const connection = new SMTPConnection({
    host: 'smtp.gmail.com', // Replace with the recipient's mail server
    port: 25
  });

  try {
    // Step 1: Connect to the SMTP server
    await connection.connect();

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

      // Parse the email address to get the mailbox and domain parts
      const [mailbox, domain] = email.split('@');

      try {
        // Step 2: Send EHLO command
        const ehloResponse = await connection.sendCommand(`EHLO ${domain}`);

        // Step 3: Send MAIL FROM command
        const mailFromResponse = await connection.sendCommand(`MAIL FROM:<${mailbox}@${domain}>`);

        // Step 4: Send RCPT TO command
        const rcptToResponse = await connection.sendCommand(`RCPT TO:<${email}>`);

        // Check the response from the email server to see if the email address exists
        if (rcptToResponse.status === 250) {
          console.log(`Email address ${email} exists`);
          results.push({ email, message: 'Email address exists' });
        } else {
          console.log(`Email address ${email} does not exist`);
          results.push({ email, message: 'Email address does not exist' });
        }
      } catch (error) {
        console.log(`Error verifying email address ${email}: ${error.message}`);
        results.push({ email, message: 'Error verifying email address' });
      }
    }

    // Step 5: Close the SMTP connection
    await connection.quit();

    // Generate HTML content based on the results
    const html = `
      <h1>Email verification complete</h1>
      <ul>
        ${results.map(({ email, message }) => `<li>${email}: ${message}</li>`).join('')}
      </ul>
    `;
    res.send(html);
  } catch (error) {
    console.log(`Error: ${error}`);
    res.status(500).send('Internal server error');
  }
});
