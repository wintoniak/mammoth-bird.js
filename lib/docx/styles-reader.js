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

  // Recursively fetch and merge base style properties if this style is based on another style
  var basedOnElement = styleElement.first("w:basedOn");
  if (basedOnElement) {
    var baseStyleId = basedOnElement.attributes["w:val"];
    var baseStyle = findStyleById(baseStyleId, styles);
    if (baseStyle && baseStyle.element) {
      // Recursively get the base style properties using its stored element
      var baseProperties = extractStyleProperties(baseStyle.element, styles);
      properties = { ...baseProperties, ...properties }; // Merge base properties first
    }
  }

  // Read numbering style if present in the current style
  var numberingProperties = readNumberingStyleElement(styleElement);
  properties.numbering = (numberingProperties.numId || numberingProperties.level) ? numberingProperties : undefined;

  // Extract the current style's properties
  var rPr = styleElement.first("w:rPr");
  var pPr = styleElement.first("w:pPr");

  // Run properties
  if (rPr) {
    // Font Size
    var fontSizeString = rPr.firstOrEmpty("w:sz").attributes["w:val"];
    properties.fontSize = /^[0-9]+$/.test(fontSizeString) ? parseInt(fontSizeString, 10) / 2 : undefined;

    // Color
    var colorElement = rPr.first("w:color");
    properties.color = colorElement ? colorElement.attributes["w:val"] : undefined;

    // Other Text Properties
    properties.bold = rPr.first("w:b") ? true : undefined;
    properties.italic = rPr.first("w:i") ? true : undefined;
    properties.underline = rPr.first("w:u") ? true : undefined;
    properties.strike = rPr.first("w:strike") ? true : undefined;
    properties.isAllCaps = rPr.first("w:caps") ? true : undefined;
    properties.isSmallCaps = rPr.first("w:smallCaps") ? true : undefined;

    // Vertical Alignment
    var verticalAlignmentElement = rPr.first("w:vertAlign");
    properties.verticalAlignment = verticalAlignmentElement ? verticalAlignmentElement.attributes["w:val"] : undefined;

    // Font Family
    var fontElement = rPr.first("w:rFonts");
    properties.font = fontElement ? fontElement.attributes["w:ascii"] : undefined;

    // Highlight
    var highlightElement = rPr.first("w:highlight");
    properties.highlight = highlightElement ? highlightElement.attributes["w:val"] : undefined;

    // Shading
    var shadingElement = rPr.first("w:shd");
    properties.shading = shadingElement ? shadingElement.attributes["w:fill"] : undefined;
  }

  // Paragraph properties
  if (pPr) {
    // Spacing
    var spacingElement = pPr.first("w:spacing");
    properties.spacing = spacingElement
      ? {
          before: spacingElement.attributes["w:before"],
          after: spacingElement.attributes["w:after"],
          line: spacingElement.attributes["w:line"],
          lineRule: spacingElement.attributes["w:lineRule"]
        }
      : undefined;

    // Alignment
    var alignmentElement = pPr.first("w:jc");
    properties.alignment = alignmentElement ? alignmentElement.attributes["w:val"] : undefined;

    // Indentation
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

  // Return the final properties object with only defined properties
  return Object.fromEntries(Object.entries(properties).filter(([_, v]) => v !== undefined));
}

