exports.read = read;
exports._findPartPaths = findPartPaths;

var promises = require("../promises");
var documents = require("../documents");
var Result = require("../results").Result;
var zipfile = require("../zipfile");

var readXmlFromZipFile = require("./office-xml-reader").readXmlFromZipFile;
var createBodyReader = require("./body-reader").createBodyReader;
var DocumentXmlReader = require("./document-xml-reader").DocumentXmlReader;
var relationshipsReader = require("./relationships-reader");
var contentTypesReader = require("./content-types-reader");
var numberingXml = require("./numbering-xml");
var stylesReader = require("./styles-reader");
var notesReader = require("./notes-reader");
var commentsReader = require("./comments-reader");
var Files = require("./files").Files;


function read(docxFile, input = {}) {
  return promises.props({
      contentTypes: readContentTypesFromZipFile(docxFile),
      partPaths: findPartPaths(docxFile),
      docxFile: docxFile,
      files: input.path ? Files.relativeToFile(input.path) : new Files(null)
  }).then(result => {
      return readStylesFromZipFile(docxFile, result.partPaths.styles).then(styles => {
          result.styles = styles; // Attach styles to the result object
          return result;
      });
  }).then(result => {
      return readNumberingFromZipFile(docxFile, result.partPaths.numbering, result.styles).then(numbering => {
          result.numbering = numbering; // Attach numbering to the result object
          return result;
      });
  }).then(result => {
      return Promise.all([
          readXmlFileWithBody(result.partPaths.footnotes, result, notesReader.createFootnotesReader),
          readXmlFileWithBody(result.partPaths.endnotes, result, notesReader.createEndnotesReader),
          readXmlFileWithBody(result.partPaths.comments, result, commentsReader.createCommentsReader)
      ]).then(([footnotes, endnotes, comments]) => {
          result.footnotes = footnotes;
          result.endnotes = endnotes;
          result.comments = comments;
          result.notes = new documents.Notes([...footnotes, ...endnotes]);
          return result;
      });
  }).then(result => {
      return readXmlFileWithBody(result.partPaths.mainDocument, result, function(bodyReader, xml) {
          const reader = new DocumentXmlReader({
              bodyReader: bodyReader,
              notes: result.notes,
              comments: result.comments
          });
          return reader.convertXmlToDocument(xml);
      }).then(document => {
          return Object.assign(document, { styles: result.styles });
      });
  });
}

function findPartPaths(docxFile) {
  return readPackageRelationships(docxFile).then(packageRelationships => {
      var mainDocumentPath = findPartPath({
          docxFile: docxFile,
          relationships: packageRelationships,
          relationshipType: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument",
          basePath: "",
          fallbackPath: "word/document.xml"
      });

      if (!docxFile.exists(mainDocumentPath)) {
          throw new Error("Could not find main document part. Are you sure this is a valid .docx file?");
      }

      return xmlFileReader({
          filename: relationshipsFilename(mainDocumentPath),
          readElement: relationshipsReader.readRelationships,
          defaultValue: relationshipsReader.defaultValue
      })(docxFile).then(documentRelationships => {
          function findPartRelatedToMainDocument(name) {
              return findPartPath({
                  docxFile: docxFile,
                  relationships: documentRelationships,
                  relationshipType: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/" + name,
                  basePath: zipfile.splitPath(mainDocumentPath).dirname,
                  fallbackPath: "word/" + name + ".xml"
              });
          }

          return {
              mainDocument: mainDocumentPath,
              comments: findPartRelatedToMainDocument("comments"),
              endnotes: findPartRelatedToMainDocument("endnotes"),
              footnotes: findPartRelatedToMainDocument("footnotes"),
              numbering: findPartRelatedToMainDocument("numbering"),
              styles: findPartRelatedToMainDocument("styles")
          };
      });
  });
}

function findPartPath(options) {
    var docxFile = options.docxFile;
    var relationships = options.relationships;
    var relationshipType = options.relationshipType;
    var basePath = options.basePath;
    var fallbackPath = options.fallbackPath;

    var targets = relationships.findTargetsByType(relationshipType);
    var normalisedTargets = targets.map(function(target) {
        return stripPrefix(zipfile.joinPath(basePath, target), "/");
    });
    var validTargets = normalisedTargets.filter(function(target) {
        return docxFile.exists(target);
    });
    if (validTargets.length === 0) {
        return fallbackPath;
    } else {
        return validTargets[0];
    }
}

function stripPrefix(value, prefix) {
    if (value.substring(0, prefix.length) === prefix) {
        return value.substring(prefix.length);
    } else {
        return value;
    }
}

function xmlFileReader(options) {
    return function(zipFile) {
        return readXmlFromZipFile(zipFile, options.filename)
            .then(function(element) {
                return element ? options.readElement(element) : options.defaultValue;
            });
    };
}

function readXmlFileWithBody(filename, options, func) {
    var readRelationshipsFromZipFile = xmlFileReader({
        filename: relationshipsFilename(filename),
        readElement: relationshipsReader.readRelationships,
        defaultValue: relationshipsReader.defaultValue
    });

    return readRelationshipsFromZipFile(options.docxFile).then(function(relationships) {
        var bodyReader = new createBodyReader({
            relationships: relationships,
            contentTypes: options.contentTypes,
            docxFile: options.docxFile,
            numbering: options.numbering,
            styles: options.styles,
            files: options.files
        });
        return readXmlFromZipFile(options.docxFile, filename)
            .then(function(xml) {
                return func(bodyReader, xml);
            });
    });
}

function relationshipsFilename(filename) {
    var split = zipfile.splitPath(filename);
    return zipfile.joinPath(split.dirname, "_rels", split.basename + ".rels");
}

var readContentTypesFromZipFile = xmlFileReader({
    filename: "[Content_Types].xml",
    readElement: contentTypesReader.readContentTypesFromXml,
    defaultValue: contentTypesReader.defaultContentTypes
});

function readNumberingFromZipFile(zipFile, path, styles) {
    return xmlFileReader({
        filename: path,
        readElement: function(element) {
            return numberingXml.readNumberingXml(element, {styles: styles});
        },
        defaultValue: numberingXml.defaultNumbering
    })(zipFile);
}

function readStylesFromZipFile(zipFile, path) {
    return xmlFileReader({
        filename: path,
        readElement: stylesReader.readStylesXml,
        defaultValue: stylesReader.defaultStyles
    })(zipFile);
}

var readPackageRelationships = xmlFileReader({
    filename: "_rels/.rels",
    readElement: relationshipsReader.readRelationships,
    defaultValue: relationshipsReader.defaultValue
});
