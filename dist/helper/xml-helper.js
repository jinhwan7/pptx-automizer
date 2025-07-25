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
exports.XmlHelper = void 0;
const xmldom_1 = require("@xmldom/xmldom");
const constants_1 = require("../constants/constants");
const xml_pretty_print_1 = require("./xml-pretty-print");
const general_helper_1 = require("./general-helper");
const content_tracker_1 = require("./content-tracker");
const content_type_map_1 = require("../enums/content-type-map");
class XmlHelper {
    static modifyXmlInArchive(archive, file, callbacks) {
        return __awaiter(this, void 0, void 0, function* () {
            const fileProxy = yield archive;
            const xml = yield XmlHelper.getXmlFromArchive(fileProxy, file);
            let i = 0;
            for (const callback of callbacks) {
                yield callback(xml, i++, fileProxy);
            }
            XmlHelper.writeXmlToArchive(yield archive, file, xml);
        });
    }
    static getXmlFromArchive(archive, file) {
        return __awaiter(this, void 0, void 0, function* () {
            return archive.readXml(file);
        });
    }
    static writeXmlToArchive(archive, file, xml) {
        archive.writeXml(file, xml);
    }
    static appendIf(element) {
        return __awaiter(this, void 0, void 0, function* () {
            const xml = yield XmlHelper.getXmlFromArchive(element.archive, element.file);
            return element.clause !== undefined && !element.clause(xml)
                ? false
                : XmlHelper.append(element);
        });
    }
    static append(element) {
        return __awaiter(this, void 0, void 0, function* () {
            const xml = yield XmlHelper.getXmlFromArchive(element.archive, element.file);
            const newElement = xml.createElement(element.tag);
            for (const attribute in element.attributes) {
                const value = element.attributes[attribute];
                const setValue = typeof value === 'function' ? value(xml) : value;
                newElement.setAttribute(attribute, setValue);
            }
            content_tracker_1.contentTracker.trackRelation(element.file, element.attributes);
            if (element.assert) {
                element.assert(xml);
            }
            const parent = element.parent(xml);
            parent.appendChild(newElement);
            XmlHelper.writeXmlToArchive(element.archive, element.file, xml);
            return newElement;
        });
    }
    static removeIf(element) {
        return __awaiter(this, void 0, void 0, function* () {
            const xml = yield XmlHelper.getXmlFromArchive(element.archive, element.file);
            const collection = xml.getElementsByTagName(element.tag);
            const toRemove = [];
            XmlHelper.modifyCollection(collection, (item, index) => {
                if (element.clause(xml, item)) {
                    toRemove.push(item);
                }
            });
            toRemove.forEach((item) => {
                XmlHelper.remove(item);
            });
            XmlHelper.writeXmlToArchive(element.archive, element.file, xml);
            return toRemove;
        });
    }
    static getNextRelId(rootArchive, file) {
        return __awaiter(this, void 0, void 0, function* () {
            const presentationRelsXml = yield XmlHelper.getXmlFromArchive(rootArchive, file);
            const increment = (max) => 'rId' + max;
            const relationNodes = presentationRelsXml.documentElement.childNodes;
            const rid = XmlHelper.getMaxId(relationNodes, 'Id', true);
            return increment(rid) + '-created';
        });
    }
    static getMaxId(rels, attribute, increment, minId) {
        let max = 0;
        for (const i in rels) {
            const rel = rels[i];
            if (rel.getAttribute !== undefined) {
                const id = Number(rel
                    .getAttribute(attribute)
                    .replace('rId', '')
                    .replace('-created', ''));
                max = id > max ? id : max;
            }
        }
        switch (typeof increment) {
            case 'boolean':
                ++max;
                break;
        }
        if (max < minId) {
            return minId;
        }
        return max;
    }
    static getRelationshipTargetsByPrefix(archive, path, prefix) {
        return __awaiter(this, void 0, void 0, function* () {
            const prefixes = typeof prefix === 'string' ? [prefix] : prefix;
            return XmlHelper.getRelationshipItems(archive, path, (element, targets) => {
                prefixes.forEach((prefix) => {
                    const target = XmlHelper.parseRelationTarget(element, prefix);
                    if (target.prefix) {
                        targets.push(target);
                    }
                });
            });
        });
    }
    static parseRelationTarget(element, prefix) {
        const type = element.getAttribute('Type');
        const file = element.getAttribute('Target');
        const last = (arr) => arr[arr.length - 1];
        const filename = last(file.split('/'));
        const subtype = last(prefix.split('/'));
        const relType = last(type.split('/'));
        const rId = element.getAttribute('Id');
        const filenameExt = last(filename.split('.'));
        const filenameMatch = filename
            .replace('.' + filenameExt, '')
            .match(/^(.+?)(\d+)*$/);
        const filenameBase = filenameMatch && filenameMatch[1] ? filenameMatch[1] : filename;
        const number = filenameMatch && filenameMatch[2] ? Number(filenameMatch[2]) : 0;
        const target = {
            rId,
            type,
            file,
            filename,
            filenameBase,
            number,
            subtype,
            relType,
            element,
        };
        if (prefix &&
            XmlHelper.targetMatchesRelationship(relType, subtype, file, prefix)) {
            return Object.assign(Object.assign({}, target), { prefix });
        }
        if (prefix && prefix.indexOf('../') === 0) {
            // Try again with absolute path instead of relative
            return XmlHelper.parseRelationTarget(element, prefix.replace('../', '/ppt/'));
        }
        return target;
    }
    static targetMatchesRelationship(relType, subtype, file, prefix) {
        if (relType === 'package')
            return true;
        // pptgenjs uses absolute paths in "Target" attributes
        if (file.indexOf('/ppt/') === 0) {
            file = file.replace('/ppt/', '../');
        }
        return relType === subtype && file.indexOf(prefix) === 0;
    }
    static getTargetsByRelationshipType(archive, path, type) {
        return __awaiter(this, void 0, void 0, function* () {
            return yield XmlHelper.getRelationshipItems(archive, path, (element, rels) => {
                const target = element.getAttribute('Type');
                if (target === type) {
                    rels.push({
                        file: element.getAttribute('Target'),
                        rId: element.getAttribute('Id'),
                        element: element,
                    });
                }
            });
        });
    }
    static getRelationshipItems(archive, path, cb, tag) {
        return __awaiter(this, void 0, void 0, function* () {
            tag = tag || 'Relationship';
            const xml = yield XmlHelper.getXmlFromArchive(archive, path);
            const relationshipItems = xml.getElementsByTagName(tag);
            const rels = [];
            for (const i in relationshipItems) {
                if (relationshipItems[i].getAttribute) {
                    cb(relationshipItems[i], rels);
                }
            }
            return rels;
        });
    }
    static findByAttribute(xml, tagName, attributeName, attributeValue) {
        const elements = xml.getElementsByTagName(tagName);
        for (const i in elements) {
            const element = elements[i];
            if (element.getAttribute !== undefined) {
                if (element.getAttribute(attributeName) === attributeValue) {
                    return true;
                }
            }
        }
        return false;
    }
    static replaceAttribute(archive, path, tagName, attributeName, attributeValue, replaceValue, replaceAttributeName) {
        return __awaiter(this, void 0, void 0, function* () {
            const xml = yield XmlHelper.getXmlFromArchive(archive, path);
            const elements = xml.getElementsByTagName(tagName);
            for (const i in elements) {
                const element = elements[i];
                if (element.getAttribute !== undefined &&
                    element.getAttribute(attributeName) === attributeValue) {
                    element.setAttribute(replaceAttributeName || attributeName, replaceValue);
                }
                if (element.getAttribute !== undefined) {
                    content_tracker_1.contentTracker.trackRelation(path, {
                        Id: element.getAttribute('Id'),
                        Target: element.getAttribute('Target'),
                        Type: element.getAttribute('Type'),
                    });
                }
            }
            XmlHelper.writeXmlToArchive(archive, path, xml);
        });
    }
    static getTargetByRelId(archive, relsPath, element, type) {
        var _a;
        return __awaiter(this, void 0, void 0, function* () {
            const params = constants_1.TargetByRelIdMap[type];
            // For elements that need to search all instances (like hyperlinks)
            if (params.findAll) {
                // Find all hyperlink elements
                const hyperlinks = element.getElementsByTagName(params.relRootTag);
                if (hyperlinks.length > 0) {
                    // Use the first hyperlink found
                    const sourceRid = hyperlinks[0].getAttribute(params.relAttribute);
                    // Get all relationships
                    const allRels = yield XmlHelper.getRelationshipItems(archive, relsPath, (element, rels) => {
                        rels.push({
                            rId: element.getAttribute('Id'),
                            type: element.getAttribute('Type'),
                            file: element.getAttribute('Target'),
                            filename: element.getAttribute('Target'),
                            element: element,
                            isExternal: element.getAttribute('TargetMode') === 'External',
                        });
                    });
                    // Find the matching relationship
                    const target = allRels.find((rel) => rel.rId === sourceRid);
                    return target;
                }
            }
            else {
                // Standard behavior for other element types
                const sourceRid = (_a = element
                    .getElementsByTagName(params.relRootTag)
                    .item(0)) === null || _a === void 0 ? void 0 : _a.getAttribute(params.relAttribute);
                if (!sourceRid) {
                    throw 'No sourceRid for ' + params.relRootTag;
                }
                const shapeRels = yield XmlHelper.getRelationshipTargetsByPrefix(archive, relsPath, params.prefix);
                const target = shapeRels.find((rel) => {
                    return rel.rId === sourceRid;
                });
                return target;
            }
        });
    }
    // Determine whether a given string is a creationId or a shape name
    // Example creationId: '{EFC74B4C-D832-409B-9CF4-73C1EFF132D8}'
    static isElementCreationId(selector) {
        return selector.indexOf('{') === 0 && selector.split('-').length === 5;
    }
    static findByElementCreationId(archive, path, creationId) {
        return __awaiter(this, void 0, void 0, function* () {
            const slideXml = yield XmlHelper.getXmlFromArchive(archive, path);
            return XmlHelper.findByCreationId(slideXml, creationId);
        });
    }
    static findByElementNameAndId(archive, path, name, id) {
        return __awaiter(this, void 0, void 0, function* () {
            const slideXml = yield XmlHelper.getXmlFromArchive(archive, path);
            return XmlHelper.findByNameAndId(slideXml, name, id);
        });
    }
    static findByElementName(archive, path, name) {
        return __awaiter(this, void 0, void 0, function* () {
            const slideXml = yield XmlHelper.getXmlFromArchive(archive, path);
            return XmlHelper.findByName(slideXml, name);
        });
    }
    static findByName(doc, name) {
        const names = doc.getElementsByTagName('p:cNvPr');
        for (const i in names) {
            if (names[i].getAttribute && names[i].getAttribute('name') === name) {
                return names[i].parentNode.parentNode;
            }
        }
        return null;
    }
    static findByNameAndId(doc, name, id) {
        const names = doc.getElementsByTagName('p:cNvPr');
        for (const i in names) {
            if (names[i].getAttribute &&
                names[i].getAttribute('name') === name &&
                names[i].getAttribute('id') === id) {
                return names[i].parentNode.parentNode;
            }
        }
        return null;
    }
    static findByCreationId(doc, creationId) {
        const creationIds = doc.getElementsByTagName('a16:creationId');
        for (const i in creationIds) {
            if (creationIds[i].getAttribute &&
                creationIds[i].getAttribute('id') === creationId) {
                return creationIds[i].parentNode.parentNode.parentNode.parentNode
                    .parentNode;
            }
        }
        return null;
    }
    static findFirstByAttributeValue(nodes, attributeName, attributeValue) {
        for (const i in nodes) {
            const node = nodes[i];
            if (node.getAttribute &&
                node.getAttribute(attributeName) === attributeValue) {
                return node;
            }
        }
        return null;
    }
    static findByAttributeValue(nodes, attributeName, attributeValue) {
        const matchingNodes = [];
        for (const i in nodes) {
            const node = nodes[i];
            if (node.getAttribute &&
                node.getAttribute(attributeName) === attributeValue) {
                matchingNodes.push(node);
            }
        }
        return matchingNodes;
    }
    static createContentTypeChild(archive, attributes) {
        return {
            archive,
            file: `[Content_Types].xml`,
            parent: (xml) => xml.getElementsByTagName('Types')[0],
            tag: 'Override',
            attributes,
        };
    }
    static createRelationshipChild(archive, targetRelFile, attributes) {
        content_tracker_1.contentTracker.trackRelation(targetRelFile, attributes);
        return {
            archive,
            file: targetRelFile,
            parent: (xml) => xml.getElementsByTagName('Relationships')[0],
            tag: 'Relationship',
            attributes,
        };
    }
    static appendImageExtensionToContentType(targetArchive, extension) {
        const contentType = content_type_map_1.ContentTypeMap[extension]
            ? content_type_map_1.ContentTypeMap[extension]
            : 'image/' + extension;
        return XmlHelper.appendIf(Object.assign(Object.assign({}, XmlHelper.createContentTypeChild(targetArchive, {
            Extension: extension,
            ContentType: contentType,
        })), { tag: 'Default', clause: (xml) => !XmlHelper.findByAttribute(xml, 'Default', 'Extension', extension) }));
    }
    static appendSharedString(sharedStrings, stringValue) {
        const strings = sharedStrings.getElementsByTagName('sst')[0];
        const newLabel = sharedStrings.createTextNode(stringValue);
        const newText = sharedStrings.createElement('t');
        newText.appendChild(newLabel);
        const newString = sharedStrings.createElement('si');
        newString.appendChild(newText);
        strings.appendChild(newString);
        return strings.getElementsByTagName('si').length - 1;
    }
    static insertAfter(newNode, referenceNode) {
        return referenceNode.parentNode.insertBefore(newNode, referenceNode.nextSibling);
    }
    static sliceCollection(collection, length, from) {
        if (from !== undefined) {
            for (let i = from; i < length; i++) {
                XmlHelper.remove(collection[i]);
            }
        }
        else {
            for (let i = collection.length; i > length; i--) {
                XmlHelper.remove(collection[i - 1]);
            }
        }
    }
    static getClosestParent(tag, element) {
        if (element.parentNode) {
            if (element.parentNode.nodeName === tag) {
                return element.parentNode;
            }
            return XmlHelper.getClosestParent(tag, element.parentNode);
        }
    }
    static remove(toRemove) {
        if (toRemove === null || toRemove === void 0 ? void 0 : toRemove.parentNode) {
            toRemove.parentNode.removeChild(toRemove);
        }
    }
    static moveChild(childToMove, insertBefore) {
        const parent = childToMove.parentNode;
        parent.insertBefore(childToMove, insertBefore);
    }
    static appendClone(childToClone, parent) {
        const clone = childToClone.cloneNode(true);
        parent.appendChild(clone);
        return clone;
    }
    static sortCollection(collection, order, callback) {
        if (collection.length === 0) {
            return;
        }
        const parent = collection[0].parentNode;
        order.forEach((index, i) => {
            if (!collection[index]) {
                (0, general_helper_1.log)('sortCollection index not found' + index, 1);
                return;
            }
            const item = collection[index];
            if (callback) {
                callback(item, i);
            }
            parent.appendChild(item);
        });
    }
    static modifyCollection(collection, callback) {
        for (let i = 0; i < collection.length; i++) {
            const item = collection[i];
            callback(item, i);
        }
    }
    static modifyCollectionAsync(collection, callback) {
        return __awaiter(this, void 0, void 0, function* () {
            for (let i = 0; i < collection.length; i++) {
                const item = collection[i];
                yield callback(item, i);
            }
        });
    }
    static dump(element) {
        const s = new xmldom_1.XMLSerializer();
        const xmlBuffer = s.serializeToString(element);
        const p = new xml_pretty_print_1.XmlPrettyPrint(xmlBuffer);
        p.dump();
    }
}
exports.XmlHelper = XmlHelper;
//# sourceMappingURL=xml-helper.js.map