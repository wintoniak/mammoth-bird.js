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
              properties: extractStyleProperties(styleElement)
            };
        }
    });

    const returnObject = new Styles(paragraphStyles, characterStyles, tableStyles, numberingStyles, customStyles);

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
    var numId = styleElement
        .firstOrEmpty("w:pPr")
        .firstOrEmpty("w:numPr")
        .firstOrEmpty("w:numId")
        .attributes["w:val"];
    return { numId: numId };
}

function isCustomStyle(styleElement) {
  return !!styleElement.attributes["w:customStyle"];
}

// Helper function to extract properties of a style (e.g., font size, color, etc.)
function extractStyleProperties(styleElement) {
    var properties = {};
    var rPr = styleElement.first("w:rPr");

    if (rPr) {
        // Extract font size
        var fontSizeElement = rPr.first("w:sz");
        if (fontSizeElement) {
            properties.fontSize = fontSizeElement.attributes["w:val"];
        }

        // Extract color
        var colorElement = rPr.first("w:color");
        if (colorElement) {
            properties.color = colorElement.attributes["w:val"];
        }

        // Extract more properties as needed (e.g., bold, italic, underline, etc.)
        var boldElement = rPr.first("w:b");
        if (boldElement) {
            properties.bold = true;
        }

        var italicElement = rPr.first("w:i");
        if (italicElement) {
            properties.italic = true;
        }

        var underlineElement = rPr.first("w:u");
        if (underlineElement) {
            properties.underline = true;
        }

        var strikeElement = rPr.first("w:strike");
        if (strikeElement) {
            properties.strike = true;
        }

        var isAllCapsElement = rPr.first("w:caps");
        if (isAllCapsElement) {
            properties.isAllCaps = true;
        }

        var isSmallCapsElement = rPr.first("w:smallCaps");
        if (isSmallCapsElement) {
            properties.isSmallCaps = true;
        }

        var verticalAlignmentElement = rPr.first("w:vertAlign");
        if (verticalAlignmentElement) {
            properties.verticalAlignment = verticalAlignmentElement.attributes["w:val"];
        }

        var fontElement = rPr.first("w:rFonts");
        if (fontElement) {
            properties.font = fontElement.attributes["w:ascii"];
        }

        var highlightElement = rPr.first("w:highlight");
        if (highlightElement) {
            properties.highlight = highlightElement.attributes["w:val"];
        }

        var shadingElement = rPr.first("w:shd");
        if (shadingElement) {
            properties.shading = shadingElement.attributes["w:fill"];
        }

        var alignmentElement = rPr.first("w:jc");
        if (alignmentElement) {
            properties.alignment = alignmentElement.attributes["w:val"];
        }

        // Extract numbering properties
        var numberingProperties = {};
        var numPrElement = styleElement.firstOrEmpty("w:numPr");
        var numIdElement = numPrElement.firstOrEmpty("w:numId");
        var ilvlElement = numPrElement.firstOrEmpty("w:ilvl");
        if (numIdElement) {
            numberingProperties.numId = numIdElement.attributes["w:val"];
        }
        if (ilvlElement) {
            numberingProperties.ilvl = ilvlElement.attributes["w:val"];
        }
        properties.numbering = numberingProperties;

        // Extract indent properties
        var indentProperties = {};
        var indElement = styleElement.firstOrEmpty("w:ind");
        if (indElement) {
            indentProperties.left = indElement.attributes["w:left"];
            indentProperties.right = indElement.attributes["w:right"];
            indentProperties.hanging = indElement.attributes["w:hanging"];
            indentProperties.firstLine = indElement.attributes["w:firstLine"];
        }
        properties.indent = indentProperties;
    }

    return properties;
}
