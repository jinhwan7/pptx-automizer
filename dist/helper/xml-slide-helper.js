"use strict";
var __awaiter = (this && this.__awaiter) || function (thisArg, _arguments, P, generator) {
    function adopt(value) { return value instanceof P ? value : new P(function (resolve) { resolve(value); }); }
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : adopt(result.value).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
Object.defineProperty(exports, "__esModule", { value: true });
exports.XmlSlideHelper = exports.mapUriType = exports.nsMain = void 0;
const xml_helper_1 = require("./xml-helper");
exports.nsMain = 'http://schemas.openxmlformats.org/presentationml/2006/main';
exports.mapUriType = {
    'http://schemas.openxmlformats.org/drawingml/2006/table': 'table',
    'http://schemas.openxmlformats.org/drawingml/2006/chart': 'chart',
    'http://schemas.microsoft.com/office/drawing/2014/chartex': 'chartEx',
    'http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject': 'oleObject',
    'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink': 'hyperlink',
};
/**
 * Class that represents an XML slide helper
 */
class XmlSlideHelper {
    /**
     * Constructor for the XmlSlideHelper class.
     * @param {XmlDocument} slideXml - The slide XML document to be used by the helper.
     * @param hasShapes
     */
    constructor(slideXml, hasShapes) {
        /**
         * Fetches an XML file from the given path and extracts the dimensions.
         *
         * @param {string} path - The path of the XML file.
         * @returns {Promise<{ width: number; height: number } | null>} - A promise that resolves with an object containing the width and height, or `null` if there was an error.
         */
        this.getAndExtractDimensions = (path) => __awaiter(this, void 0, void 0, function* () {
            try {
                const xml = yield xml_helper_1.XmlHelper.getXmlFromArchive(this.hasShapes.sourceTemplate.archive, path);
                if (!xml)
                    return null;
                const sldSz = xml.getElementsByTagName('p:sldSz')[0];
                if (sldSz) {
                    const width = XmlSlideHelper.parseCoordinate(sldSz, 'cx');
                    const height = XmlSlideHelper.parseCoordinate(sldSz, 'cy');
                    return { width, height };
                }
                return null;
            }
            catch (error) {
                console.warn(`Error while fetching XML from path ${path}: ${error}`);
                return null;
            }
        });
        if (!slideXml) {
            throw Error('Slide XML is not defined');
        }
        this.slideXml = slideXml;
        this.hasShapes = hasShapes;
    }
    getSlideCreationId() {
        const creationIdItem = this.slideXml
            .getElementsByTagName('p14:creationId')
            .item(0);
        if (!creationIdItem) {
            return;
        }
        const creationIdSlide = creationIdItem.getAttribute('val');
        if (!creationIdSlide) {
            return;
        }
        return Number(creationIdSlide);
    }
    /**
     * Get an array of ElementInfo objects for all named elements on a slide.
     * @param selector
     */
    getElement(selector) {
        return __awaiter(this, void 0, void 0, function* () {
            const shapeNode = xml_helper_1.XmlHelper.isElementCreationId(selector)
                ? xml_helper_1.XmlHelper.findByCreationId(this.slideXml, selector)
                : xml_helper_1.XmlHelper.findByName(this.slideXml, selector);
            return XmlSlideHelper.getElementInfo(shapeNode);
        });
    }
    getElementByUniqueSelector(selector) {
        return __awaiter(this, void 0, void 0, function* () {
            const shapeNode = xml_helper_1.XmlHelper.findByNameAndId(this.slideXml, selector.name, selector.id);
            return XmlSlideHelper.getElementInfo(shapeNode);
        });
    }
    /**
     * Get an array of ElementInfo objects for all named elements on a slide.
     * @param filterTags Use an array of strings to filter the output array
     */
    getAllElements(filterTags) {
        const elementInfo = [];
        try {
            const shapeNodes = this.getNamedElements(filterTags);
            shapeNodes.forEach((shapeNode) => {
                elementInfo.push(XmlSlideHelper.getElementInfo(shapeNode));
            });
        }
        catch (error) {
            console.error(error);
            throw new Error(`Failed to retrieve elements: ${error.message}`);
        }
        return elementInfo;
    }
    /**
     * Get all text element IDs from the slide.
     * @return {string[]} An array of text element IDs.
     */
    getAllTextElementIds(useCreationIds) {
        const elementIds = [];
        try {
            elementIds.push(...this.getAllElements(['sp'])
                .filter((element) => element.hasTextBody)
                .map((element) => (useCreationIds ? element.id : element.name)));
        }
        catch (error) {
            console.error(error);
            throw new Error(`Failed to retrieve text element IDs: ${error.message}`);
        }
        return elementIds;
    }
    getAllTextElementUniqueSelectors() {
        const elementSelectors = [];
        try {
            elementSelectors.push(...this.getAllElements(['sp'])
                .filter((element) => element.hasTextBody)
                .map((element) => element.uniqueSelector));
        }
        catch (error) {
            console.error(error);
            throw new Error(`Failed to retrieve text element IDs: ${error.message}`);
        }
        return elementSelectors;
    }
    static getElementInfo(slideElement) {
        return {
            name: XmlSlideHelper.getElementName(slideElement),
            id: XmlSlideHelper.getElementCreationId(slideElement),
            uniqueSelector: XmlSlideHelper.getElementUniqueSelector(slideElement),
            type: XmlSlideHelper.getElementType(slideElement),
            position: XmlSlideHelper.parseShapeCoordinates(slideElement),
            hasTextBody: !!XmlSlideHelper.getTextBody(slideElement),
            getXmlElement: () => slideElement,
            getText: () => XmlSlideHelper.parseTextFragments(slideElement),
            getTableInfo: () => XmlSlideHelper.readTableInfo(slideElement),
            getAltText: () => XmlSlideHelper.getImageAltText(slideElement),
            getTextByLine: () => XmlSlideHelper.getTextByLine(slideElement),
        };
    }
    /**
     * Retrieves a list of all named elements on a slide. Automation requires at least a name.
     * @param filterTags Use an array of strings to filter the output array
     */
    getNamedElements(filterTags) {
        const skipTags = ['spTree'];
        const nvPrs = this.slideXml.getElementsByTagNameNS(exports.nsMain, 'cNvPr');
        const namedElements = [];
        xml_helper_1.XmlHelper.modifyCollection(nvPrs, (nvPr) => {
            const parentNode = nvPr.parentNode.parentNode;
            const parentTag = parentNode.localName;
            if (!skipTags.includes(parentTag) &&
                (!(filterTags === null || filterTags === void 0 ? void 0 : filterTags.length) || filterTags.includes(parentTag))) {
                namedElements.push(parentNode);
            }
        });
        return namedElements;
    }
    static getTextBody(shapeNode) {
        return shapeNode.getElementsByTagNameNS(exports.nsMain, 'txBody').item(0);
    }
    static parseTextFragments(shapeNode) {
        const txBody = XmlSlideHelper.getTextBody(shapeNode);
        const textFragments = [];
        const texts = txBody.getElementsByTagName('a:t');
        for (let t = 0; t < texts.length; t++) {
            textFragments.push(texts.item(t).textContent);
        }
        return textFragments;
    }
    static getTextByLine(shapeNode) {
        const txBody = XmlSlideHelper.getTextBody(shapeNode);
        if (!txBody) {
            return [];
        }
        const allLines = [];
        const paragraphs = txBody.getElementsByTagName('a:p');
        for (let p = 0; p < paragraphs.length; p++) {
            const paragraph = paragraphs.item(p);
            let currentLine = '';
            for (const childNode of Array.from(paragraph.childNodes)) {
                const element = childNode;
                if (element.nodeName === 'a:br') {
                    allLines.push(currentLine);
                    currentLine = '';
                    continue;
                }
                //childeNode가 a:r이라면 또 그 안에서 a:t를 찾아서 텍스트를 추가한다.
                if (element.nodeName === 'a:r') {
                    const textNodes = element.getElementsByTagName('a:t');
                    for (let t = 0; t < textNodes.length; t++) {
                        currentLine += textNodes.item(t).textContent;
                    }
                }
            }
            allLines.push(currentLine);
        }
        return allLines;
    }
    static getNonVisibleProperties(shapeNode) {
        return shapeNode.getElementsByTagNameNS(exports.nsMain, 'cNvPr').item(0);
    }
    static getImageAltText(slideElement) {
        const cNvPr = XmlSlideHelper.getNonVisibleProperties(slideElement);
        if (cNvPr) {
            return cNvPr.getAttribute('descr');
        }
    }
    static getElementName(slideElement) {
        const cNvPr = XmlSlideHelper.getNonVisibleProperties(slideElement);
        if (cNvPr) {
            return cNvPr.getAttribute('name');
        }
    }
    static getElementUniqueSelector(slideElement) {
        const cNvPr = XmlSlideHelper.getNonVisibleProperties(slideElement);
        if (cNvPr) {
            return {
                name: cNvPr.getAttribute('name'),
                id: cNvPr.getAttribute('id'),
            };
        }
    }
    static getElementCreationId(slideElement) {
        const cNvPr = XmlSlideHelper.getNonVisibleProperties(slideElement);
        if (cNvPr) {
            const creationIdElement = cNvPr
                .getElementsByTagName('a16:creationId')
                .item(0);
            if (creationIdElement) {
                return creationIdElement.getAttribute('id');
            }
        }
    }
    /**
     * Parses local tag name to specify element type in case it is a 'graphicFrame'.
     * @param slideElementParent
     */
    static getElementType(slideElementParent) {
        let type = slideElementParent.localName;
        switch (type) {
            case 'graphicFrame':
                const graphicData = slideElementParent.getElementsByTagName('a:graphicData')[0];
                const uri = graphicData.getAttribute('uri');
                type = exports.mapUriType[uri] ? exports.mapUriType[uri] : type;
                break;
            case 'oleObj':
                type = 'OLEObject';
                break;
        }
        // Check for hyperlinks
        const hasHyperlink = slideElementParent.getElementsByTagName('a:hlinkClick');
        if (hasHyperlink.length > 0) {
            type = 'Hyperlink';
        }
        return type;
    }
    static parseShapeCoordinates(slideElementParent) {
        const xFrmsA = slideElementParent.getElementsByTagName('a:xfrm');
        const xFrmsP = slideElementParent.getElementsByTagName('p:xfrm');
        const xFrms = xFrmsP.item(0) ? xFrmsP : xFrmsA;
        const position = {
            x: 0,
            y: 0,
            cx: 0,
            cy: 0,
        };
        if (!xFrms.item(0)) {
            return position;
        }
        const xFrm = xFrms.item(0);
        const Off = xFrm.getElementsByTagName('a:off').item(0);
        const Ext = xFrm.getElementsByTagName('a:ext').item(0);
        position.x = XmlSlideHelper.parseCoordinate(Off, 'x');
        position.y = XmlSlideHelper.parseCoordinate(Off, 'y');
        position.cx = XmlSlideHelper.parseCoordinate(Ext, 'cx');
        position.cy = XmlSlideHelper.parseCoordinate(Ext, 'cy');
        return position;
    }
    /**
     * Asynchronously retrieves the dimensions of a slide.
     * Tries to find the dimensions from the slide XML, then from the layout, master, and presentation XMLs in order.
     *
     * @returns {Promise<{ width: number, height: number }>} The dimensions of the slide.
     * @throws Error if unable to determine dimensions.
     */
    getDimensions() {
        return __awaiter(this, void 0, void 0, function* () {
            try {
                const dimensions = yield this.getAndExtractDimensions('ppt/presentation.xml');
                if (dimensions)
                    return dimensions;
            }
            catch (error) {
                console.error(`Error while fetching slide dimensions: ${error}`);
                throw error;
            }
        });
    }
}
exports.XmlSlideHelper = XmlSlideHelper;
XmlSlideHelper.parseCoordinate = (element, attributeName) => {
    return parseInt(element.getAttribute(attributeName), 10);
};
XmlSlideHelper.readTableInfo = (element) => {
    const info = [];
    const rows = element.getElementsByTagName('a:tr');
    if (!rows) {
        console.error("Can't find a table row.");
        return info;
    }
    for (let r = 0; r < rows.length; r++) {
        const row = rows.item(r);
        const columns = row.getElementsByTagName('a:tc');
        for (let c = 0; c < columns.length; c++) {
            const cell = columns.item(c);
            const gridSpan = cell.getAttribute('gridSpan');
            const hMerge = cell.getAttribute('hMerge');
            const texts = cell.getElementsByTagName('a:t');
            const text = [];
            for (let t = 0; t < texts.length; t++) {
                text.push(texts.item(t).textContent);
            }
            info.push({
                row: r,
                column: c,
                rowXml: row,
                columnXml: cell,
                text: text,
                textContent: text.join(''),
                gridSpan: Number(gridSpan),
                hMerge: Number(hMerge),
            });
        }
    }
    return info;
};
//# sourceMappingURL=xml-slide-helper.js.map