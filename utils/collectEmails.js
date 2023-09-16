const validator = require("validator");

/**
 * @param {Object [AsyncGenerator]} messages - Async generator with messages
 */

module.exports = async (messages) => {
  try {
    const emails = [];

    for await (let msg of messages) {
      const body = msg.source.toString("utf-8");

      const emailsRaw = body.match(
        /([A-Za-z0-9._-]+@[A-Za-z0-9._-]+\.[A-Za-z0-9._-]+)/gim
      );

      if (emailsRaw && emailsRaw.length) {
        emailsRaw.forEach((email) => {
          if (validator.isEmail(email) && !emails.includes(email)) {
            emails.push(email);
          }
        });
      }
    }

    return emails;
  } catch (err) {
    console.log("💥ERROR:", err);
  }
};
