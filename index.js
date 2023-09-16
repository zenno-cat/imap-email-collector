const client = require("./utils/imapClient");
const collectEmails = require("./utils/collectEmails");
const writeToExcel = require("./utils/writeToExcel");

const main = async () => {
  // Wait until client connects and authorizes
  await client
    .connect()
    .then(() =>
      console.log(`${process.env.USER} Успешное подключение по IMAP!`)
    );

  // Select and lock a mailbox. Throws if mailbox does not exist

  let tree = await client.listTree();

  if (!tree && !tree.length)
    return console.log("В почтовом ящике нету сообщений");

  tree = tree.folders.map((mailbox) => mailbox.path);

  for (let box of tree) {
    console.log(`Поиск эмейлов в ящике: ${box}`);
    let lock = await client.getMailboxLock(box).catch((err) => {
      console.log(`Ошибка: ${err.message}, ${err.responseText}`);
      console.log(`Не удалось открыть ящик: ${box}`);
    });

    if (!lock) continue;

    try {
      const messages = client.fetch("1:*", {
        source: true,
      });

      const emails = await collectEmails(messages);

      if (emails && emails.length) {
        await writeToExcel(emails, box, process.env.USER);
      }
    } finally {
      // Make sure lock is released, otherwise next `getMailboxLock()` never returns

      lock.release();
    }
  }

  console.log("Результаты сохранены в папку с программой");
  // log out and close connection
  await client.logout();
};

main().catch((err) => console.error("💥ERROR:", err));
