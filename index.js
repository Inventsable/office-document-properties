//
// office-document-properties
// https://github.com/swestmoreland/office-document-properties
//
"use strict";

var _ = require("lodash"),
  parse = require("xml2js").parseString,
  yauzl = require("yauzl");

var APP_XML = require("./app.xml.json"),
  CORE_XML = require("./core.xml.json");

function fromBuffer(buffer, cb) {
  if (buffer && buffer instanceof Buffer && typeof cb === "function") {
    yauzl.fromBuffer(buffer, { lazyEntries: true }, function (err, zipfile) {
      if (err) return cb(err, null);

      readEntries(zipfile, cb);
    });
  } else {
    if (typeof cb === "function") {
      cb(new Error("Incorrect parameters."), null);
    } else {
      console.error("Incorrect parameters.");
    }
  }
}

function fromFilePath(filePath, cb) {
  if (typeof filePath === "string" && typeof cb === "function") {
    yauzl.open(filePath, { lazyEntries: true }, function (err, zipfile) {
      if (err) return cb(err, null);

      readEntries(zipfile, cb);
    });
  } else {
    if (typeof cb === "function") {
      cb(new Error("Incorrect parameters."), null);
    } else {
      console.error("Incorrect parameters.");
    }
  }
}

function readEntries(zipfile, cb) {
  var data = {};

  zipfile.readEntry();

  zipfile.on("end", function () {
    cb(null, sortByKeys(data));
  });

  zipfile.on("entry", function (entry) {
    switch (entry.fileName) {
      case "docProps/app.xml":
        readEntryStreamXML(zipfile, entry, function (err, result) {
          _.assign(data, getDocumentProperties(result, APP_XML));
          zipfile.readEntry();
        });
        break;
      case "docProps/core.xml":
        readEntryStreamXML(zipfile, entry, function (err, result) {
          _.assign(data, getDocumentProperties(result, CORE_XML));
          zipfile.readEntry();
        });
        break;
      case "docProps/custom.xml":
        readEntryStreamXML(zipfile, entry, function (err, result) {
          _.assign(data, getCustomProps(result));
          zipfile.readEntry();
        });
        break;
      default:
        zipfile.readEntry();
    }
  });
}

// Adding function to Object.assign custom properties from file to original package
function getCustomProps(result) {
  let temp = {};
  if (
    !result.Properties ||
    !result.Properties.property ||
    !result.Properties.property.length
  ) {
    return temp;
  }
  result.Properties.property.forEach((prop) => {
    let key = prop["$"].name;
    let val =
      prop[
        Object.keys(prop).filter((key) => {
          return !/\$/.test(key);
        })[0]
      ][0];
    temp[
      `${key[0].toLowerCase() + key.substring(1)}`
        .replace(/(?:^\w|[A-Z]|\b\w)/g, function (word, index) {
          return index === 0 ? word.toLowerCase() : word.toUpperCase();
        })
        .replace(/\s+/g, "")
    ] = val;
  });
  return temp;
}

function readEntryStreamXML(zipfile, entry, cb) {
  zipfile.openReadStream(entry, function (err, readStream) {
    var data = "";

    if (err) return cb(err, null);

    readStream.on("data", function (chunk) {
      data += chunk;
    });

    readStream.on("end", function () {
      parse(data, function (err, result) {
        if (err) return cb(err, null);

        cb(null, result);
      });
    });
  });
}

function getDocumentProperties(obj, props) {
  var data = {};

  _.forEach(props, function (prop) {
    var val;

    if (_.has(obj, prop.path)) {
      switch (prop.type) {
        case "number":
          val = _.toNumber(_.get(obj, prop.path));
          break;
        case "string":
        default:
          val = _.toString(_.get(obj, prop.path));
      }

      if (prop.type == "string" && _.isEmpty(val)) return;

      _.set(data, prop.name, val);
    }
  });

  return data;
}

var sortByKeys = (object) => {
  const keys = Object.keys(object);
  const sortedKeys = _.sortBy(keys);

  return _.fromPairs(_.map(sortedKeys, (key) => [key, object[key]]));
};

module.exports = {
  fromBuffer: fromBuffer,
  fromFilePath: fromFilePath,
};
