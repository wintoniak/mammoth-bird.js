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

Styles.EMPTY = new Styles({}, {}, {}, {}, []);

function readStylesXml(root) {
    var paragraphStyles = {};
    var characterStyles = {};
    var tableStyles = {};
    var numberingStyles = {};
    var customStyles = [];

    var styles = {
        "paragraph": paragraphStyles,
        "character": characterStyles,
        "table": tableStyles
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
            customStyles.push({
                styleId: style.styleId,
                type: style.type,
                name: style.name,
                properties: extractStyleProperties(styleElement)
            });
        }
    });

    return new Styles(paragraphStyles, characterStyles, tableStyles, numberingStyles, customStyles);
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
  return styleElement.attributes["w:customStyle"] === 1;
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
    }

    return properties;
}
