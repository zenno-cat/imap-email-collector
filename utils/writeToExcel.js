const ExcelJS = require("exceljs");
const moment = require("moment");
const fs = require("fs");

/**
 *
 * @param {Array} data - Array of object with emails
 * @param {string} box - Mailbox name
 * @param {string} user - Mailbox address
 */

module.exports = async (data, box, user) => {
  const resultPath = __dirname + "/../result";
  try {
    if (!fs.existsSync(resultPath)) {
      fs.mkdirSync(resultPath);
    }
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet("Сообщения");
    const timestamp = moment().format("YYYY-MM-DD_HH-mm-ss");
    const boxName = box.replaceAll(
      /[-’\/`~!#*$@_%+=.,'^&(){}[\]|;:”<>?\\]/gm,
      "_"
    );
    const filename = `${user}_${boxName}_${timestamp}.xlsx`;

    worksheet.columns = [
      { header: "Отправитель", key: "sender", width: 20 },
      { header: "Название директории", key: "directory", width: 30 },
    ];

    data.forEach((item) => {
      worksheet.addRow({
        sender: item,
        directory: box,
      });
    });

    await workbook.xlsx.writeFile(`${resultPath}/${filename}`);
  } catch (err) {
    console.log("💥ERROR:", err);
  }
};
