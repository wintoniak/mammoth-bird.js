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
                    table: tableStyles
                })
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
    let properties = {};

    // Recursively fetch and merge base style properties if this style is based on another style
    const basedOnElement = styleElement.first("w:basedOn");
    if (basedOnElement) {
        const baseStyleId = basedOnElement.attributes["w:val"];
        const baseStyle = findStyleById(baseStyleId, styles);
        if (baseStyle && baseStyle.element) {
            // Recursively get the base style properties using its stored element
            const baseProperties = extractStyleProperties(baseStyle.element, styles);
            properties = { ...baseProperties, ...properties }; // Merge base properties first
        }
    }

    // Extract the current style's properties
    const rPr = styleElement.first("w:rPr");
    const pPr = styleElement.first("w:pPr");

    // Run properties
    if (rPr) {
        const fontSizeString = rPr.firstOrEmpty("w:sz").attributes["w:val"];
        const fontSize = /^[0-9]+$/.test(fontSizeString) ? parseInt(fontSizeString, 10) / 2 : null;
        if (fontSize) {
            properties.fontSize = fontSize;
        }

        const colorElement = rPr.first("w:color");
        if (colorElement) {
            properties.color = colorElement.attributes["w:val"];
        }

        if (rPr.first("w:b")) properties.bold = true;
        if (rPr.first("w:i")) properties.italic = true;
        if (rPr.first("w:u")) properties.underline = true;
        if (rPr.first("w:strike")) properties.strike = true;
        if (rPr.first("w:caps")) properties.isAllCaps = true;
        if (rPr.first("w:smallCaps")) properties.isSmallCaps = true;

        const verticalAlignmentElement = rPr.first("w:vertAlign");
        if (verticalAlignmentElement) {
            properties.verticalAlignment = verticalAlignmentElement.attributes["w:val"];
        }

        const fontElement = rPr.first("w:rFonts");
        if (fontElement) {
            properties.font = fontElement.attributes["w:ascii"];
        }

        const highlightElement = rPr.first("w:highlight");
        if (highlightElement) {
            properties.highlight = highlightElement.attributes["w:val"];
        }

        const shadingElement = rPr.first("w:shd");
        if (shadingElement) {
            properties.shading = shadingElement.attributes["w:fill"];
        }
    }

    // Paragraph properties
    if (pPr) {
        const spacingElement = pPr.first("w:spacing");
        if (spacingElement) {
            properties.spacing = {
                before: spacingElement.attributes["w:before"],
                after: spacingElement.attributes["w:after"],
            };
        }

        const alignmentElement = pPr.first("w:jc");
        if (alignmentElement) {
            properties.alignment = alignmentElement.attributes["w:val"];
        }

        const indElement = pPr.first("w:ind");
        if (indElement) {
            properties.indent = {
                left: indElement.attributes["w:left"],
                right: indElement.attributes["w:right"],
                hanging: indElement.attributes["w:hanging"],
                firstLine: indElement.attributes["w:firstLine"],
            };
        }
    }

    return properties;
}
