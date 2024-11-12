exports.readStylesXml = readStylesXml;
exports.Styles = Styles;
exports.defaultStyles = new Styles({}, {});

function Styles(paragraphStyles, characterStyles, tableStyles, numberingStyles, customStyles) {
    return {
        findParagraphStyleById: function(styleId) {
            return paragraphStyles[styleId];
        },
        findCharacterStyleById: function(styleId) {
            return characterStyles[styleId];
        },
        findTableStyleById: function(styleId) {
            return tableStyles[styleId];
        },
        findNumberingStyleById: function(styleId) {
            return numberingStyles[styleId];
        },
        getCustomStyles: function() {
            return customStyles;
        }
    };
}

Styles.EMPTY = new Styles({}, {}, {}, {}, {});

function readStylesXml(root) {
    var paragraphStyles = {};
    var characterStyles = {};
    var tableStyles = {};
    var numberingStyles = {};
    var customStyles = {};

    var styles = {
        "paragraph": paragraphStyles,
        "character": characterStyles,
        "table": tableStyles,
    };

    root.getElementsByTagName("w:style").forEach(function(styleElement) {
        var style = readStyleElement(styleElement);
        style.element = styleElement; // Store the raw XML element for inheritance purposes

        if (style.type === "numbering") {
            numberingStyles[style.styleId] = readNumberingStyleElement(styleElement);
        } else {
            var styleSet = styles[style.type];
            if (styleSet) {
                styleSet[style.styleId] = style;
            }
        }

        if (isCustomStyle(styleElement)) {
            customStyles[style.styleId] = {
                name: style.name,
                type: style.type,
                properties: extractStyleProperties(styleElement, {
                    paragraph: paragraphStyles,
                    character: characterStyles,
                    table: tableStyles,
                    numbering: numberingStyles,
                })
            };
        }
    });

    var returnObject = new Styles(paragraphStyles, characterStyles, tableStyles, numberingStyles, customStyles);

    return returnObject;
}

function readStyleElement(styleElement) {
    var type = styleElement.attributes["w:type"];
    var styleId = styleElement.attributes["w:styleId"];
    var name = styleName(styleElement);
    return { type: type, styleId: styleId, name: name };
}

function styleName(styleElement) {
    var nameElement = styleElement.first("w:name");
    return nameElement ? nameElement.attributes["w:val"] : null;
}

function readNumberingStyleElement(styleElement) {
  var numberingElement = styleElement.firstOrEmpty("w:pPr").firstOrEmpty("w:numPr");
    var numId = numberingElement
        .firstOrEmpty("w:numId")
        .attributes["w:val"];
        var level = numberingElement
        .firstOrEmpty("w:ilvl")
        .attributes["w:val"];

    return { 
      numId: numId,
      level: level,
    };
}

function isCustomStyle(styleElement) {
  return !!styleElement.attributes["w:customStyle"];
}

// Helper to find style by styleId in the various style sets
function findStyleById(styleId, styles) {
  return (
      styles.paragraph[styleId] ||
      styles.character[styleId] ||
      styles.table[styleId] ||
      null
  );
}

// Helper function to extract properties of a style (e.g., font size, color, etc.)
function extractStyleProperties(styleElement, styles) {
  // Start with an empty properties object for each call
  var properties = {};

  // Clear properties that should not be inherited unless explicitly set
  properties.fontSize = undefined;
  properties.numbering = undefined;

  // Recursively fetch and merge base style properties if this style is based on another style
  var basedOnElement = styleElement.first("w:basedOn");
  if (basedOnElement) {
    var baseStyleId = basedOnElement.attributes["w:val"];
    var baseStyle = findStyleById(baseStyleId, styles);
    if (baseStyle && baseStyle.element) {
      console.log('baseStyle', baseStyleId, 'original', styleElement.attributes["w:styleId"]);
      var baseProperties = extractStyleProperties(baseStyle.element, styles);
      console.log('recursiveRes', baseProperties);
      for (var key in baseProperties) {
        if (baseProperties.hasOwnProperty(key)) {
          properties[key] = baseProperties[key];
        }
      }
    }
  }

  // Read numbering style if present in the current style
  var numberingProperties = readNumberingStyleElement(styleElement);
  if (numberingProperties.numId || numberingProperties.level) {
    properties.numbering = numberingProperties;
  }

  // Extract the current style's properties
  var rPr = styleElement.first("w:rPr");
  var pPr = styleElement.first("w:pPr");

  console.log('rPr', rPr, styleElement.attributes["w:styleId"]);

  if (rPr) {
    var fontSizeString = rPr.firstOrEmpty("w:sz").attributes["w:val"];
    console.log('fontSizeString', fontSizeString, styleElement.attributes["w:styleId"]);
    properties.fontSize = /^[0-9]+$/.test(fontSizeString) ? parseInt(fontSizeString, 10) / 2 : undefined;

    var colorElement = rPr.first("w:color");
    properties.color = colorElement ? colorElement.attributes["w:val"] : undefined;

    // Underline: only true if w:u exists and w:val is not "none"
    var underlineElement = rPr.first("w:u");
    properties.underline = underlineElement && underlineElement.attributes["w:val"] !== "none" ? true : undefined;
    
    // Bold: only true if w:b exists and w:val is not "false"
    var boldElement = rPr.first("w:b");
    properties.bold = boldElement && boldElement.attributes["w:val"] !== "false" ? true : undefined;

    // Italic: only true if w:i exists and w:val is not "false"
    var italicElement = rPr.first("w:i");
    properties.italic = italicElement && italicElement.attributes["w:val"] !== "false" ? true : undefined;

    // Strike-through: only true if w:strike exists and w:val is not "false"
    var strikeElement = rPr.first("w:strike");
    properties.strike = strikeElement && strikeElement.attributes["w:val"] !== "false" ? true : undefined;

    // All Caps: only true if w:caps exists and w:val is not "false"
    var capsElement = rPr.first("w:caps");
    properties.isAllCaps = capsElement && capsElement.attributes["w:val"] !== "false" ? true : undefined;

    // Small Caps: only true if w:smallCaps exists and w:val is not "false"
    var smallCapsElement = rPr.first("w:smallCaps");
    properties.isSmallCaps = smallCapsElement && smallCapsElement.attributes["w:val"] !== "false" ? true : undefined;

    var verticalAlignmentElement = rPr.first("w:vertAlign");
    properties.verticalAlignment = verticalAlignmentElement ? verticalAlignmentElement.attributes["w:val"] : undefined;

    var fontElement = rPr.first("w:rFonts");
    properties.font = fontElement ? fontElement.attributes["w:ascii"] : undefined;

    var highlightElement = rPr.first("w:highlight");
    properties.highlight = highlightElement ? highlightElement.attributes["w:val"] : undefined;

    var shadingElement = rPr.first("w:shd");
    properties.shading = shadingElement ? shadingElement.attributes["w:fill"] : undefined;
  }

  if (pPr) {
    var spacingElement = pPr.first("w:spacing");
    properties.spacing = spacingElement
      ? {
          before: spacingElement.attributes["w:before"],
          after: spacingElement.attributes["w:after"],
          line: spacingElement.attributes["w:line"],
          lineRule: spacingElement.attributes["w:lineRule"]
        }
      : undefined;

    var alignmentElement = pPr.first("w:jc");
    properties.alignment = alignmentElement ? alignmentElement.attributes["w:val"] : undefined;

    var indElement = pPr.first("w:ind");
    properties.indent = indElement
      ? {
          left: indElement.attributes["w:left"],
          right: indElement.attributes["w:right"],
          start: indElement.attributes["w:start"],
          end: indElement.attributes["w:end"],
          hanging: indElement.attributes["w:hanging"],
          firstLine: indElement.attributes["w:firstLine"]
        }
      : undefined;
  }

  var cleanedProperties = {};
  for (var prop in properties) {
    if (properties[prop] !== undefined) {
      cleanedProperties[prop] = properties[prop];
    }
  }

  return cleanedProperties;
}
