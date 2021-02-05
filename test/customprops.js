const main = require("../index");
const path = require("path");
const root = path.resolve("./test/files/customprops.pptx");
function init() {
  main.fromFilePath(root, (err, data) => {
    if (err) console.error(err);
    console.log("RESULT:\r\n");
    console.log(data);
  });
}

init();
