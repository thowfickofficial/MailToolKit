const nodemailer = require('nodemailer');

const transporter = nodemailer.createTransport({
  host: 'smtp.gmail.com',
  port: 587,
  secure: true,
  auth: {
    user: 'example@gmail.com',
    pass: ' xpwl nnvy nnvy nxyd'
  }
});

const sendRegistrationEmail = (userEmail) => {
  const mailOptions = {
    from: 'YOUR-EMAIL-ID',
    to: userEmail,
    subject: 'Registration Successful',
    html: `
      <html>
      <head>
        <title>Email Template</title>
        <style>
          body {
            font-family: Arial, sans-serif;
            background-color: #f2f2f2;
            margin: 0;
            padding: 0;
          }
      
          .container {
            max-width: 600px;
            margin: 0 auto;
            padding: 20px;
            background-color: #ffffff;
            box-shadow: 0 0 10px rgba(0, 0, 0, 0.1);
          }
      
          h1 {
            color: #333333;
            font-size: 24px;
            margin-bottom: 20px;
          }
      
          p {
            color: #666666;
            font-size: 16px;
            line-height: 1.5;
          }
        </style>
      </head>
      <body>
        <div class="container">
          <h1>Dear User,</h1>
          <p>Thank you for registering an account. We appreciate your support!</p>
        </div>
      </body>
      </html>
    `
  };

  transporter.sendMail(mailOptions, (error, info) => {
    if (error) {
      console.log('Error sending email: ', error);
    } else {
      console.log('Email sent: ', info.response);
    }
  });
};

const userEmail = 'example@example.com';
sendRegistrationEmail(userEmail);
