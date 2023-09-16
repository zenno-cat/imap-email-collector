const { ImapFlow } = require("imapflow");
require("dotenv").config();

const client = new ImapFlow({
  host: process.env.HOST,
  port: process.env.PORT,
  secure: true,
  auth: {
    user: process.env.USER,
    pass: process.env.PASS,
  },
  tls: {
    rejectUnauthorized: false,
  },
  logger: false,
});

module.exports = client;
