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
exports.Shape = void 0;
const xml_helper_1 = require("../helper/xml-helper");
const general_helper_1 = require("../helper/general-helper");
const content_type_map_1 = require("../enums/content-type-map");
class Shape {
    constructor(shape, targetType) {
        this.shape = shape;
        this.mode = shape.mode;
        this.name = shape.name;
        this.targetType = targetType;
        this.sourceArchive = shape.sourceArchive;
        this.sourceSlideNumber = shape.sourceSlideNumber;
        this.sourceSlideFile = `ppt/slides/slide${this.sourceSlideNumber}.xml`;
        this.sourceElement = shape.sourceElement;
        this.hasCreationId = shape.hasCreationId;
        this.callbacks = general_helper_1.GeneralHelper.arrayify(shape.callback);
        this.contentTypeMap = content_type_map_1.ContentTypeMap;
        if (shape.target) {
            this.sourceNumber = shape.target.number;
            this.sourceRid = shape.target.rId;
            this.subtype = shape.target.subtype;
            this.target = shape.target;
        }
    }
    setTarget(targetTemplate, targetSlideNumber) {
        return __awaiter(this, void 0, void 0, function* () {
            const targetType = this.targetType;
            this.targetTemplate = targetTemplate;
            this.targetArchive = yield this.targetTemplate.archive;
            this.targetSlideNumber = targetSlideNumber;
            this.targetSlideFile = `ppt/${targetType}s/${targetType}${this.targetSlideNumber}.xml`;
            this.targetSlideRelFile = `ppt/${targetType}s/_rels/${targetType}${this.targetSlideNumber}.xml.rels`;
        });
    }
    setTargetElement() {
        return __awaiter(this, void 0, void 0, function* () {
            if (!this.sourceElement) {
                // If we don't have a source element, we might be trying to remove a hyperlink
                console.log(`Warning: No source element for shape ${this.name}. Creating empty element for operations.`);
                if (this.shape && this.shape.mode === 'remove' && this.targetArchive) {
                    // For remove operations, we don't need a source element
                    // Just continue without setting targetElement
                    return;
                }
                // For non-remove operations or if other conditions aren't met, throw the error
                console.log(this.shape);
                throw `No source element for shape ${this.name}`;
            }
            this.targetElement = this.sourceElement.cloneNode(true);
        });
    }
    appendToSlideTree() {
        return __awaiter(this, void 0, void 0, function* () {
            const targetSlideXml = yield xml_helper_1.XmlHelper.getXmlFromArchive(this.targetArchive, this.targetSlideFile);
            targetSlideXml
                .getElementsByTagName('p:spTree')[0]
                .appendChild(this.targetElement);
            // Process hyperlinks in the element if this is a hyperlink element
            if (this.relRootTag === 'a:hlinkClick') {
                yield this.processHyperlinks(targetSlideXml);
            }
            xml_helper_1.XmlHelper.writeXmlToArchive(this.targetArchive, this.targetSlideFile, targetSlideXml);
        });
    }
    // Process hyperlinks in the element
    processHyperlinks(targetSlideXml) {
        return __awaiter(this, void 0, void 0, function* () {
            if (!this.targetElement)
                return;
            // Scenario 1: Update r:id in <p:nvSpPr><p:cNvPr><a:hlinkClick r:id="..." />
            // This is for hyperlinks applied to the shape itself.
            const nvSpPr = this.targetElement.getElementsByTagName('p:nvSpPr')[0];
            if (nvSpPr) {
                const cNvPr = nvSpPr.getElementsByTagName('p:cNvPr')[0];
                if (cNvPr) {
                    const shapeHlinks = cNvPr.getElementsByTagName('a:hlinkClick');
                    for (let k = 0; k < shapeHlinks.length; k++) {
                        const shapeHlink = shapeHlinks[k];
                        const currentRid = shapeHlink.getAttribute('r:id');
                        if (this.createdRid && currentRid) {
                            shapeHlink.setAttribute('r:id', this.createdRid);
                            shapeHlink.setAttribute('xmlns:r', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships');
                        }
                    }
                }
            }
            // Scenario 2: Update r:id in <p:txBody>...<a:rPr><a:hlinkClick r:id="..." />
            // This is for hyperlinks applied to text runs within the shape.
            const runs = this.targetElement.getElementsByTagName('a:r');
            for (let i = 0; i < runs.length; i++) {
                const run = runs[i];
                const rPr = run.getElementsByTagName('a:rPr')[0];
                if (rPr) {
                    const hlinkClicks = rPr.getElementsByTagName('a:hlinkClick');
                    for (let j = 0; j < hlinkClicks.length; j++) {
                        const hlinkClick = hlinkClicks[j];
                        const currentRid = hlinkClick.getAttribute('r:id');
                        if (this.createdRid && currentRid) {
                            hlinkClick.setAttribute('r:id', this.createdRid);
                            hlinkClick.setAttribute('xmlns:r', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships');
                        }
                    }
                }
            }
        });
    }
    replaceIntoSlideTree() {
        return __awaiter(this, void 0, void 0, function* () {
            yield this.modifySlideTree(true);
        });
    }
    removeFromSlideTree() {
        return __awaiter(this, void 0, void 0, function* () {
            yield this.modifySlideTree(false);
        });
    }
    modifySlideTree(insertBefore) {
        return __awaiter(this, void 0, void 0, function* () {
            const archive = this.targetArchive;
            const slideFile = this.targetSlideFile;
            const targetSlideXml = yield xml_helper_1.XmlHelper.getXmlFromArchive(archive, slideFile);
            const findMethod = this.hasCreationId ? 'findByCreationId' : 'findByName';
            const sourceElementOnTargetSlide = yield xml_helper_1.XmlHelper[findMethod](targetSlideXml, this.name);
            if (!(sourceElementOnTargetSlide === null || sourceElementOnTargetSlide === void 0 ? void 0 : sourceElementOnTargetSlide.parentNode)) {
                console.error(`Can't modify slide tree for ${this.name}`);
                return;
            }
            if (insertBefore === true && this.targetElement) {
                sourceElementOnTargetSlide.parentNode.insertBefore(this.targetElement, sourceElementOnTargetSlide);
            }
            sourceElementOnTargetSlide.parentNode.removeChild(sourceElementOnTargetSlide);
            // Process hyperlinks in the element if this is a hyperlink element
            if (this.relRootTag === 'a:hlinkClick') {
                yield this.processHyperlinks(targetSlideXml);
            }
            xml_helper_1.XmlHelper.writeXmlToArchive(archive, slideFile, targetSlideXml);
        });
    }
    updateElementsRelId() {
        return __awaiter(this, void 0, void 0, function* () {
            const targetSlideXml = yield xml_helper_1.XmlHelper.getXmlFromArchive(this.targetArchive, this.targetSlideFile);
            const targetElements = yield this.getElementsByRid(targetSlideXml, this.sourceRid);
            targetElements.forEach((targetElement) => {
                this.relParent(targetElement)
                    .getElementsByTagName(this.relRootTag)[0]
                    .setAttribute(this.relAttribute, this.createdRid);
            });
            xml_helper_1.XmlHelper.writeXmlToArchive(this.targetArchive, this.targetSlideFile, targetSlideXml);
        });
    }
    /*
     * This will find all elements with a matching rId on a
     * <p:cSld>, including related images at <p:bg> and <p:spTree>.
     */
    getElementsByRid(slideXml, rId) {
        return __awaiter(this, void 0, void 0, function* () {
            const sourceList = slideXml
                .getElementsByTagName('p:cSld')[0]
                .getElementsByTagName(this.relRootTag);
            return xml_helper_1.XmlHelper.findByAttributeValue(sourceList, this.relAttribute, rId);
        });
    }
    updateTargetElementRelId() {
        return __awaiter(this, void 0, void 0, function* () {
            this.targetElement
                .getElementsByTagName(this.relRootTag)
                .item(0)
                .setAttribute(this.relAttribute, this.createdRid);
        });
    }
    applyCallbacks(callbacks, element, relation) {
        callbacks.forEach((callback) => {
            if (typeof callback === 'function') {
                try {
                    callback(element, relation);
                }
                catch (e) {
                    console.warn(e);
                }
            }
        });
    }
    applyChartCallbacks(callbacks, element, chart, workbook) {
        callbacks.forEach((callback) => {
            if (typeof callback === 'function') {
                try {
                    callback(element, chart, workbook);
                }
                catch (e) {
                    console.warn(e);
                }
            }
        });
    }
    appendImageExtensionToContentType(extension) {
        return xml_helper_1.XmlHelper.appendImageExtensionToContentType(this.targetArchive, extension);
    }
}
exports.Shape = Shape;
//# sourceMappingURL=shape.js.map