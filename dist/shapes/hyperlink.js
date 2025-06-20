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
exports.Hyperlink = void 0;
const xml_helper_1 = require("../helper/xml-helper");
const shape_1 = require("../classes/shape");
const content_tracker_1 = require("../helper/content-tracker");
class Hyperlink extends shape_1.Shape {
    constructor(shape, targetType, sourceArchive, hyperlinkType = 'external', hyperlinkTarget) {
        super(shape, targetType);
        this.sourceArchive = sourceArchive;
        this.hyperlinkType = hyperlinkType;
        this.hyperlinkTarget = hyperlinkTarget || '';
        this.relRootTag = 'a:hlinkClick';
        this.relAttribute = 'r:id';
    }
    modify(targetTemplate, targetSlideNumber) {
        return __awaiter(this, void 0, void 0, function* () {
            yield this.prepare(targetTemplate, targetSlideNumber);
            yield this.editTargetHyperlinkRel();
            yield this.replaceIntoSlideTree();
            return this;
        });
    }
    append(targetTemplate, targetSlideNumber) {
        return __awaiter(this, void 0, void 0, function* () {
            yield this.prepare(targetTemplate, targetSlideNumber);
            yield this.setTargetElement();
            yield this.appendToSlideTree();
            yield this.editTargetHyperlinkRel();
            yield this.updateHyperlinkInSlide();
            return this;
        });
    }
    remove(targetTemplate, targetSlideNumber) {
        return __awaiter(this, void 0, void 0, function* () {
            yield this.prepare(targetTemplate, targetSlideNumber);
            if (this.target && this.target.rId) {
                this.sourceRid = this.target.rId;
            }
            yield this.removeFromSlideTree();
            yield this.removeFromSlideRels();
            return this;
        });
    }
    prepare(targetTemplate, targetSlideNumber) {
        return __awaiter(this, void 0, void 0, function* () {
            yield this.setTarget(targetTemplate, targetSlideNumber);
            if (!this.createdRid) {
                const baseId = yield xml_helper_1.XmlHelper.getNextRelId(this.targetArchive, this.targetSlideRelFile);
                // Strip '-created' suffix more robustly
                this.createdRid = baseId.endsWith('-created')
                    ? baseId.slice(0, -8)
                    : baseId;
            }
            if (this.shape && this.shape.target && this.shape.target.rId) {
                this.sourceRid = this.shape.target.rId;
            }
            // If hyperlinkTarget is not set, try to get it from the original rel target
            if (!this.hyperlinkTarget && this.shape && this.shape.target && this.shape.target.file) {
                this.hyperlinkTarget = this.shape.target.file;
                this.hyperlinkType = (this.shape.target.isExternal || this.shape.target.type === 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink') ? 'external' : 'internal';
            }
        });
    }
    editTargetHyperlinkRel() {
        return __awaiter(this, void 0, void 0, function* () {
            const targetRelFile = this.targetSlideRelFile;
            const relXml = yield xml_helper_1.XmlHelper.getXmlFromArchive(this.targetArchive, targetRelFile);
            const relationships = relXml.getElementsByTagName('Relationship');
            // Check if the relationship already exists
            let relationshipExists = false;
            for (let i = 0; i < relationships.length; i++) {
                if (relationships[i].getAttribute('Id') === this.createdRid) {
                    this.updateHyperlinkRelation(relationships[i]);
                    relationshipExists = true;
                    break;
                }
            }
            // If the relationship doesn't exist, create it
            if (!relationshipExists) {
                const newRel = relXml.createElement('Relationship');
                newRel.setAttribute('Id', this.createdRid);
                newRel.setAttribute('Type', this.getRelationshipType());
                newRel.setAttribute('Target', this.getRelationshipTarget());
                if (this.hyperlinkType === 'external') {
                    newRel.setAttribute('TargetMode', 'External');
                }
                relXml.documentElement.appendChild(newRel);
                // Track the relationship for content integrity
                content_tracker_1.contentTracker.trackRelation(targetRelFile, {
                    Id: this.createdRid,
                    Target: this.getRelationshipTarget(),
                    Type: this.getRelationshipType(),
                });
            }
            xml_helper_1.XmlHelper.writeXmlToArchive(this.targetArchive, targetRelFile, relXml);
        });
    }
    // Add a method to update the hyperlink in the slide XML
    updateHyperlinkInSlide() {
        return __awaiter(this, void 0, void 0, function* () {
            if (!this.targetElement && this.sourceRid && this.createdRid) {
                const slideXml = yield xml_helper_1.XmlHelper.getXmlFromArchive(this.targetArchive, this.targetSlideFile);
                const allHyperlinkElements = slideXml.getElementsByTagName('a:hlinkClick');
                let foundAndUpdatedInSlide = false;
                for (let i = 0; i < allHyperlinkElements.length; i++) {
                    const hlinkClick = allHyperlinkElements[i];
                    if (hlinkClick.getAttribute('r:id') === this.sourceRid) {
                        hlinkClick.setAttribute('r:id', this.createdRid);
                        hlinkClick.setAttribute('xmlns:r', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships');
                        foundAndUpdatedInSlide = true;
                        break;
                    }
                }
                if (foundAndUpdatedInSlide) {
                    xml_helper_1.XmlHelper.writeXmlToArchive(this.targetArchive, this.targetSlideFile, slideXml);
                }
            }
        });
    }
    updateHyperlinkRelation(element) {
        element.setAttribute('Type', this.getRelationshipType());
        element.setAttribute('Target', this.getRelationshipTarget());
        if (this.hyperlinkType === 'external') {
            element.setAttribute('TargetMode', 'External');
        }
        else if (element.hasAttribute('TargetMode')) {
            element.removeAttribute('TargetMode');
        }
        content_tracker_1.contentTracker.trackRelation(this.targetSlideRelFile, {
            Id: element.getAttribute('Id') || '',
            Target: this.getRelationshipTarget(),
            Type: this.getRelationshipType(),
        });
    }
    getRelationshipType() {
        if (this.hyperlinkType === 'internal') {
            return 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide';
        }
        return 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink';
    }
    getRelationshipTarget() {
        var _a, _b;
        if (this.hyperlinkType === 'internal') {
            // Enhanced internal slide link handling
            const slideNumber = ((_b = (_a = this.hyperlinkTarget) === null || _a === void 0 ? void 0 : _a.match(/\d+/)) === null || _b === void 0 ? void 0 : _b[0]) || this.targetSlideNumber.toString();
            // Ensure proper slide path format
            return `../slides/slide${slideNumber}.xml`;
        }
        return this.hyperlinkTarget || 'https://example.com';
    }
    removeFromSlideRels() {
        return __awaiter(this, void 0, void 0, function* () {
            const targetRelFile = this.targetSlideRelFile;
            const relXml = yield xml_helper_1.XmlHelper.getXmlFromArchive(this.targetArchive, targetRelFile);
            const relationships = relXml.getElementsByTagName('Relationship');
            const ridToRemove = this.sourceRid || this.createdRid;
            if (ridToRemove) {
                for (let i = relationships.length - 1; i >= 0; i--) {
                    if (relationships[i].getAttribute('Id') === ridToRemove) {
                        relationships[i].parentNode.removeChild(relationships[i]);
                        break;
                    }
                }
                xml_helper_1.XmlHelper.writeXmlToArchive(this.targetArchive, targetRelFile, relXml);
            }
        });
    }
    static getAllOnSlide(archive, relsPath) {
        return __awaiter(this, void 0, void 0, function* () {
            const hyperlinkRelType = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink';
            const slideRelType = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide';
            return xml_helper_1.XmlHelper.getRelationshipItems(archive, relsPath, (element, rels) => {
                const type = element.getAttribute('Type');
                if (type === hyperlinkRelType || type === slideRelType) {
                    rels.push({
                        rId: element.getAttribute('Id'),
                        type: element.getAttribute('Type'),
                        file: element.getAttribute('Target'),
                        filename: element.getAttribute('Target'),
                        element: element,
                        isExternal: element.getAttribute('TargetMode') === 'External' || type === hyperlinkRelType,
                    });
                }
            });
        });
    }
    modifyOnAddedSlide(targetTemplate, targetSlideNumber) {
        return __awaiter(this, void 0, void 0, function* () {
            if (!this.target || !this.target.rId) {
                console.warn('modifyOnAddedSlide called on Hyperlink without a valid source target/rId.');
                return;
            }
            this.sourceRid = this.target.rId;
            this.hyperlinkTarget = this.target.file;
            this.hyperlinkType = (this.target.isExternal || this.target.type === 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink') ? 'external' : 'internal';
            yield this.prepare(targetTemplate, targetSlideNumber);
            // 1. Modify the copied _rels file:
            const targetRelFile = this.targetSlideRelFile;
            const relXml = yield xml_helper_1.XmlHelper.getXmlFromArchive(this.targetArchive, targetRelFile);
            const relationships = relXml.getElementsByTagName('Relationship');
            let relFoundAndUpdated = false;
            for (let i = 0; i < relationships.length; i++) {
                const relElement = relationships[i];
                if (relElement.getAttribute('Id') === this.sourceRid) { // Find by original rId
                    relElement.setAttribute('Id', this.createdRid); // Update Id to new unique rId
                    relElement.setAttribute('Target', this.getRelationshipTarget());
                    if (this.hyperlinkType === 'external') {
                        relElement.setAttribute('TargetMode', 'External');
                    }
                    else {
                        if (relElement.hasAttribute('TargetMode'))
                            relElement.removeAttribute('TargetMode');
                    }
                    relFoundAndUpdated = true;
                    content_tracker_1.contentTracker.trackRelation(targetRelFile, {
                        Id: this.createdRid,
                        Target: relElement.getAttribute('Target') || '',
                        Type: relElement.getAttribute('Type') || '',
                    });
                    break;
                }
            }
            if (!relFoundAndUpdated) {
                console.warn(`modifyOnAddedSlide: Relationship with sourceRId ${this.sourceRid} not found in target rels ${targetRelFile}. It might have been processed by another instance or was missing in the copied rels.`);
                const newRel = relXml.createElement('Relationship');
                newRel.setAttribute('Id', this.createdRid);
                newRel.setAttribute('Type', this.getRelationshipType());
                newRel.setAttribute('Target', this.getRelationshipTarget());
                if (this.hyperlinkType === 'external') {
                    newRel.setAttribute('TargetMode', 'External');
                }
                relXml.documentElement.appendChild(newRel);
                content_tracker_1.contentTracker.trackRelation(targetRelFile, {
                    Id: this.createdRid, Target: this.getRelationshipTarget(), Type: this.getRelationshipType()
                });
            }
            yield xml_helper_1.XmlHelper.writeXmlToArchive(this.targetArchive, targetRelFile, relXml);
            // 2. Modify the copied slide content XML
            yield this.updateHyperlinkInSlide();
        });
    }
    // Helper method to create a hyperlink in a shape
    static addHyperlinkToShape(archive, slidePath, slideRelsPath, shapeId, hyperlinkTarget) {
        return __awaiter(this, void 0, void 0, function* () {
            const slideXml = yield xml_helper_1.XmlHelper.getXmlFromArchive(archive, slidePath);
            // Find the shape by ID or name
            const shape = xml_helper_1.XmlHelper.isElementCreationId(shapeId)
                ? xml_helper_1.XmlHelper.findByCreationId(slideXml, shapeId)
                : xml_helper_1.XmlHelper.findByName(slideXml, shapeId);
            if (!shape) {
                throw new Error(`Shape with ID/name "${shapeId}" not found`);
            }
            // Create a new relationship ID
            const relId = yield xml_helper_1.XmlHelper.getNextRelId(archive, slideRelsPath);
            // Add the hyperlink relationship to the slide relationships
            const relXml = yield xml_helper_1.XmlHelper.getXmlFromArchive(archive, slideRelsPath);
            const newRel = relXml.createElement('Relationship');
            newRel.setAttribute('Id', relId);
            // Improved internal link detection
            const isInternalLink = typeof hyperlinkTarget === 'number' ||
                (typeof hyperlinkTarget === 'string' && /^\d+$/.test(hyperlinkTarget));
            if (isInternalLink) {
                // Enhanced internal slide link handling
                const slideNumber = typeof hyperlinkTarget === 'number' ?
                    hyperlinkTarget :
                    parseInt(hyperlinkTarget, 10);
                newRel.setAttribute('Type', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide');
                newRel.setAttribute('Target', `../slides/slide${slideNumber}.xml`);
            }
            else {
                newRel.setAttribute('Type', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink');
                newRel.setAttribute('Target', hyperlinkTarget.toString());
                newRel.setAttribute('TargetMode', 'External');
            }
            relXml.documentElement.appendChild(newRel);
            yield xml_helper_1.XmlHelper.writeXmlToArchive(archive, slideRelsPath, relXml);
            // Add the hyperlink to the shape
            const txBody = shape.getElementsByTagName('p:txBody')[0];
            if (txBody) {
                const paragraphs = txBody.getElementsByTagName('a:p');
                if (paragraphs.length > 0) {
                    const paragraph = paragraphs[0];
                    const runs = paragraph.getElementsByTagName('a:r');
                    if (runs.length > 0) {
                        const run = runs[0];
                        const rPr = run.getElementsByTagName('a:rPr')[0];
                        if (rPr) {
                            const hlinkClick = slideXml.createElement('a:hlinkClick');
                            hlinkClick.setAttribute('r:id', relId);
                            if (isInternalLink) {
                                hlinkClick.setAttribute('action', 'ppaction://hlinksldjump');
                            }
                            hlinkClick.setAttribute('xmlns:r', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships');
                            rPr.appendChild(hlinkClick);
                        }
                    }
                }
            }
            yield xml_helper_1.XmlHelper.writeXmlToArchive(archive, slidePath, slideXml);
            return relId;
        });
    }
}
exports.Hyperlink = Hyperlink;
//# sourceMappingURL=hyperlink.js.map