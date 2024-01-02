const ImageModule = require('docxtemplater-image-module-free');
const PizZip = require("pizzip");
const Docxtemplater = require("docxtemplater");
const fs = require("fs");
const path = require("path");
const { promisify } = require("util");

const convertTemplate = async ({ templateDocx, parameters, outFile }) => {
  const content = await promisify(fs.readFile)(
    path.resolve(__dirname, templateDocx),
    "binary"
  );
  var opts = {}
  opts.centered = false; //Set to true to always center images
  opts.fileType = "docx"; //Or pptx

  //Pass your image loader
  opts.getImage = function(tagValue, tagName) {
    //tagValue is 'examples/image.png'
    //tagName is 'image'
    return fs.readFileSync(tagValue);
  }

  opts.getSize = function(img, tagValue, tagName) {
    //img is the image returned by opts.getImage()
    //tagValue is 'examples/image.png'
    //tagName is 'image'
    //tip: you can use node module 'image-size' here
    return [50, 50];
  }
  const imageModule = new ImageModule(opts);
  const zip = new PizZip(content);
  const doc = new Docxtemplater(zip, {
    paragraphLoop: true,
    linebreaks: true,
    modules: [imageModule],
  });
  doc.render(parameters);
  const buf = doc.getZip().generate({
    type: "nodebuffer",
    compression: "DEFLATE",
  });
  if (!outFile) return buf;
  await promisify(fs.writeFile)(path.resolve(__dirname, outFile), buf);
  return outFile;
};

module.exports = function (RED) {
  function docxtemplater(config) {
    RED.nodes.createNode(this, config);
    const node = this;

    node.on("input", async function (msg) {
      const templateDocx = config.templateDocx || msg.templateDocx;
      const outFile = config.outFile || msg.outFile;
      const parameters = msg.payload || {};
      try {
        const convertedTemplate = await convertTemplate({
          templateDocx,
          parameters,
          outFile,
        });
        msg.payload = convertedTemplate;
        node.send(msg);
      } catch (error) {
        node.error(error);
      }
    });
  }
  RED.nodes.registerType("docxtemplater", docxtemplater);
};
