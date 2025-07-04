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
var __importDefault = (this && this.__importDefault) || function (mod) {
    return (mod && mod.__esModule) ? mod : { "default": mod };
};
Object.defineProperty(exports, "__esModule", { value: true });
exports.Chart = void 0;
const file_helper_1 = require("../helper/file-helper");
const xml_helper_1 = require("../helper/xml-helper");
const shape_1 = require("../classes/shape");
const path_1 = __importDefault(require("path"));
const content_tracker_1 = require("../helper/content-tracker");
const general_helper_1 = require("../helper/general-helper");
class Chart extends shape_1.Shape {
    constructor(shape, targetType) {
        super(shape, targetType);
        this.relRootTag = this.subtype === 'chart' ? 'c:chart' : 'cx:chart';
        this.relAttribute = 'r:id';
        this.relParent =
            this.subtype === 'chart'
                ? (element) => element.parentNode.parentNode.parentNode
                : (element) => element.parentNode.parentNode.parentNode.parentNode
                    .parentNode;
        this.wbEmbeddingsPath = `../embeddings/`;
        this.wbExtension = '.xlsx';
        this.relTypeChartColorStyle =
            'http://schemas.microsoft.com/office/2011/relationships/chartColorStyle';
        this.relTypeChartStyle =
            'http://schemas.microsoft.com/office/2011/relationships/chartStyle';
        this.relTypeChartImage =
            'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image';
        this.relTypeChartThemeOverride =
            'http://schemas.openxmlformats.org/officeDocument/2006/relationships/themeOverride';
        this.styleRelationFiles = {};
        this.hasWorkbook = true;
    }
    modify(targetTemplate, targetSlideNumber) {
        return __awaiter(this, void 0, void 0, function* () {
            yield this.prepare(targetTemplate, targetSlideNumber);
            yield this.clone();
            yield this.replaceIntoSlideTree();
            return this;
        });
    }
    append(targetTemplate, targetSlideNumber) {
        return __awaiter(this, void 0, void 0, function* () {
            yield this.prepare(targetTemplate, targetSlideNumber);
            yield this.clone();
            yield this.appendToSlideTree();
            return this;
        });
    }
    remove(targetTemplate, targetSlideNumber) {
        return __awaiter(this, void 0, void 0, function* () {
            yield this.prepare(targetTemplate, targetSlideNumber);
            yield this.removeFromSlideTree();
            return this;
        });
    }
    modifyOnAddedSlide(targetTemplate, targetSlideNumber) {
        return __awaiter(this, void 0, void 0, function* () {
            yield this.prepare(targetTemplate, targetSlideNumber);
            yield this.updateElementsRelId();
            return this;
        });
    }
    prepare(targetTemplate, targetSlideNumber) {
        return __awaiter(this, void 0, void 0, function* () {
            yield this.setTarget(targetTemplate, targetSlideNumber);
            this.targetNumber = this.targetTemplate.incrementCounter('charts');
            this.wbRelsPath = `ppt/charts/_rels/${this.subtype}${this.sourceNumber}.xml.rels`;
            yield this.copyFiles();
            yield this.copyChartStyleFiles();
            yield this.appendTypes();
            yield this.appendToSlideRels();
        });
    }
    clone() {
        return __awaiter(this, void 0, void 0, function* () {
            yield this.setTargetElement();
            yield this.modifyChartData();
            yield this.updateTargetElementRelId();
        });
    }
    modifyChartData() {
        return __awaiter(this, void 0, void 0, function* () {
            if (!this.hasWorkbook) {
                return;
            }
            const chartXml = yield xml_helper_1.XmlHelper.getXmlFromArchive(this.targetArchive, `ppt/charts/${this.subtype}${this.targetNumber}.xml`);
            const workbook = yield this.readWorkbook();
            this.applyChartCallbacks(this.callbacks, this.targetElement, chartXml, workbook);
            xml_helper_1.XmlHelper.writeXmlToArchive(this.targetArchive, `ppt/charts/${this.subtype}${this.targetNumber}.xml`, chartXml);
            yield this.writeWorkbook(workbook);
        });
    }
    readWorkbook() {
        return __awaiter(this, void 0, void 0, function* () {
            const workbookFilename = `ppt/embeddings/${this.worksheetFilePrefix}${this.targetWorksheet}${this.wbExtension}`;
            const archive = yield this.targetArchive.extract(workbookFilename);
            const sheet = yield xml_helper_1.XmlHelper.getXmlFromArchive(archive, 'xl/worksheets/sheet1.xml');
            const table = file_helper_1.FileHelper.fileExistsInArchive(archive, 'xl/tables/table1.xml')
                ? yield xml_helper_1.XmlHelper.getXmlFromArchive(archive, 'xl/tables/table1.xml')
                : undefined;
            const sharedStrings = yield xml_helper_1.XmlHelper.getXmlFromArchive(archive, 'xl/sharedStrings.xml');
            return {
                archive,
                sheet,
                sharedStrings,
                table,
            };
        });
    }
    writeWorkbook(workbook) {
        return __awaiter(this, void 0, void 0, function* () {
            xml_helper_1.XmlHelper.writeXmlToArchive(workbook.archive, 'xl/worksheets/sheet1.xml', workbook.sheet);
            if (workbook.table) {
                xml_helper_1.XmlHelper.writeXmlToArchive(workbook.archive, 'xl/tables/table1.xml', workbook.table);
            }
            xml_helper_1.XmlHelper.writeXmlToArchive(workbook.archive, 'xl/sharedStrings.xml', workbook.sharedStrings);
            const worksheet = yield workbook.archive.getContent({});
            yield this.targetArchive.write(`ppt/embeddings/${this.worksheetFilePrefix}${this.targetWorksheet}${this.wbExtension}`, worksheet);
        });
    }
    copyFiles() {
        return __awaiter(this, void 0, void 0, function* () {
            yield this.copyChartFiles();
            this.worksheetFilePrefix = yield this.getWorksheetFilePrefix(this.wbRelsPath);
            if (this.hasWorkbook) {
                const worksheets = yield xml_helper_1.XmlHelper.getRelationshipTargetsByPrefix(this.sourceArchive, this.wbRelsPath, `${this.wbEmbeddingsPath}${this.worksheetFilePrefix}`);
                const worksheet = worksheets[0];
                this.sourceWorksheet = worksheet.number === 0 ? '' : worksheet.number;
                this.targetWorksheet = '-created-' + this.targetNumber;
                yield this.copyWorksheetFile();
            }
            else {
                (0, general_helper_1.log)('Chart has no worksheet: ' + this.wbRelsPath, 2);
            }
            yield this.editTargetWorksheetRel();
        });
    }
    getWorksheetFilePrefix(targetRelFile) {
        return __awaiter(this, void 0, void 0, function* () {
            const relationTargets = yield xml_helper_1.XmlHelper.getRelationshipTargetsByPrefix(this.sourceArchive, targetRelFile, this.wbEmbeddingsPath);
            if (!relationTargets[0]) {
                this.hasWorkbook = false;
                return '';
            }
            return relationTargets[0].filenameBase;
        });
    }
    appendTypes() {
        return __awaiter(this, void 0, void 0, function* () {
            yield this.appendChartExtensionToContentType();
            yield this.appendChartToContentType();
            yield this.appendColorToContentType();
            yield this.appendStyleToContentType();
            yield this.appendThemeOverrideToContentType();
        });
    }
    copyChartFiles() {
        return __awaiter(this, void 0, void 0, function* () {
            yield file_helper_1.FileHelper.zipCopy(this.sourceArchive, `ppt/charts/${this.subtype}${this.sourceNumber}.xml`, this.targetArchive, `ppt/charts/${this.subtype}${this.targetNumber}.xml`);
            yield file_helper_1.FileHelper.zipCopy(this.sourceArchive, `ppt/charts/_rels/${this.subtype}${this.sourceNumber}.xml.rels`, this.targetArchive, `ppt/charts/_rels/${this.subtype}${this.targetNumber}.xml.rels`);
        });
    }
    copyChartStyleFiles() {
        var _a, _b, _c;
        return __awaiter(this, void 0, void 0, function* () {
            yield this.getChartStyles();
            if ((_a = this.styleRelationFiles.relTypeChartStyle) === null || _a === void 0 ? void 0 : _a.length) {
                yield file_helper_1.FileHelper.zipCopy(this.sourceArchive, `ppt/charts/${this.styleRelationFiles.relTypeChartStyle[0]}`, this.targetArchive, `ppt/charts/style${this.targetNumber}.xml`);
            }
            if ((_b = this.styleRelationFiles.relTypeChartColorStyle) === null || _b === void 0 ? void 0 : _b.length) {
                yield file_helper_1.FileHelper.zipCopy(this.sourceArchive, `ppt/charts/${this.styleRelationFiles.relTypeChartColorStyle[0]}`, this.targetArchive, `ppt/charts/colors${this.targetNumber}.xml`);
            }
            if (this.styleRelationFiles.relTypeChartImage) {
                for (const relTypeChartImage of this.styleRelationFiles
                    .relTypeChartImage) {
                    const imageInfo = this.getTargetChartImageUri(relTypeChartImage);
                    yield this.appendImageExtensionToContentType(imageInfo.extension);
                    yield file_helper_1.FileHelper.zipCopy(this.sourceArchive, imageInfo.source, this.targetArchive, imageInfo.target);
                }
            }
            if ((_c = this.styleRelationFiles.relTypeChartThemeOverride) === null || _c === void 0 ? void 0 : _c.length) {
                const sourceFile = this.styleRelationFiles.relTypeChartThemeOverride[0].replace('../theme/', '');
                yield file_helper_1.FileHelper.zipCopy(this.sourceArchive, `ppt/theme/${sourceFile}`, this.targetArchive, `ppt/theme/themeOverride${this.targetNumber}.xml`);
            }
        });
    }
    getChartStyles() {
        return __awaiter(this, void 0, void 0, function* () {
            const styleTypes = [
                'relTypeChartStyle',
                'relTypeChartColorStyle',
                'relTypeChartImage',
                'relTypeChartThemeOverride',
            ];
            for (const i in styleTypes) {
                const styleType = styleTypes[i];
                const styleRelation = yield xml_helper_1.XmlHelper.getTargetsByRelationshipType(this.sourceArchive, this.wbRelsPath, this[styleType]);
                this.styleRelationFiles[styleType] =
                    this.styleRelationFiles[styleType] || [];
                if (styleRelation.length) {
                    styleRelation.forEach((styleRelation) => {
                        this.styleRelationFiles[styleType].push(styleRelation.file);
                    });
                }
            }
        });
    }
    appendToSlideRels() {
        return __awaiter(this, void 0, void 0, function* () {
            this.createdRid = yield xml_helper_1.XmlHelper.getNextRelId(this.targetArchive, this.targetSlideRelFile);
            const type = this.subtype === 'chart'
                ? 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart'
                : 'http://schemas.microsoft.com/office/2014/relationships/chartEx';
            const attributes = {
                Id: this.createdRid,
                Type: type,
                Target: `../charts/${this.subtype}${this.targetNumber}.xml`,
            };
            return xml_helper_1.XmlHelper.append(xml_helper_1.XmlHelper.createRelationshipChild(this.targetArchive, this.targetSlideRelFile, attributes));
        });
    }
    editTargetWorksheetRel() {
        return __awaiter(this, void 0, void 0, function* () {
            const targetRelFile = `ppt/charts/_rels/${this.subtype}${this.targetNumber}.xml.rels`;
            const relXml = yield xml_helper_1.XmlHelper.getXmlFromArchive(this.targetArchive, targetRelFile);
            const relations = relXml.getElementsByTagName('Relationship');
            Object.keys(relations)
                .map((key) => relations[key])
                .filter((element) => element.getAttribute)
                .forEach((element) => {
                const type = element.getAttribute('Type');
                switch (type) {
                    case 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/package':
                        this.updateTargetWorksheetRelation(targetRelFile, element, 'Target', `${this.wbEmbeddingsPath}${this.worksheetFilePrefix}${this.targetWorksheet}${this.wbExtension}`);
                        break;
                    case this.relTypeChartColorStyle:
                        this.updateTargetWorksheetRelation(targetRelFile, element, 'Target', `colors${this.targetNumber}.xml`);
                        break;
                    case this.relTypeChartStyle:
                        this.updateTargetWorksheetRelation(targetRelFile, element, 'Target', `style${this.targetNumber}.xml`);
                        break;
                    case this.relTypeChartImage:
                        this.updateTargetWorksheetRelation(targetRelFile, element, 'Target', this.getTargetChartImageUri(element.getAttribute('Target')).rel);
                        break;
                    case this.relTypeChartThemeOverride:
                        this.updateTargetWorksheetRelation(targetRelFile, element, 'Target', `../theme/themeOverride${this.targetNumber}.xml`);
                        break;
                }
                content_tracker_1.contentTracker.trackRelation(targetRelFile, {
                    Id: element.getAttribute('Id'),
                    Target: element.getAttribute('Target'),
                    Type: element.getAttribute('Type'),
                });
            });
            xml_helper_1.XmlHelper.writeXmlToArchive(this.targetArchive, targetRelFile, relXml);
        });
    }
    updateTargetWorksheetRelation(targetRelFile, element, attribute, value) {
        element.setAttribute(attribute, value);
    }
    getTargetChartImageUri(origin) {
        const file = origin.replace('../media/', '');
        const extension = path_1.default
            .extname(file)
            .replace('.', '');
        return {
            source: `ppt/media/${file}`,
            target: `ppt/media/${file}-chart-${this.targetNumber}.${extension}`,
            rel: `../media/${file}-chart-${this.targetNumber}.${extension}`,
            extension: extension,
        };
    }
    copyWorksheetFile() {
        return __awaiter(this, void 0, void 0, function* () {
            const targetFile = `ppt/embeddings/${this.worksheetFilePrefix}${this.targetWorksheet}${this.wbExtension}`;
            yield file_helper_1.FileHelper.zipCopy(this.sourceArchive, `ppt/embeddings/${this.worksheetFilePrefix}${this.sourceWorksheet}${this.wbExtension}`, this.targetArchive, targetFile);
        });
    }
    appendChartExtensionToContentType() {
        return xml_helper_1.XmlHelper.appendIf(Object.assign(Object.assign({}, xml_helper_1.XmlHelper.createContentTypeChild(this.targetArchive, {
            Extension: `xlsx`,
            ContentType: `application/vnd.openxmlformats-officedocument.spreadsheetml.sheet`,
        })), { tag: 'Default', clause: (xml) => !xml_helper_1.XmlHelper.findByAttribute(xml, 'Default', 'Extension', 'xlsx') }));
    }
    appendChartToContentType() {
        const contentType = this.subtype === 'chart'
            ? 'application/vnd.openxmlformats-officedocument.drawingml.chart+xml'
            : 'application/vnd.ms-office.chartex+xml';
        return xml_helper_1.XmlHelper.append(xml_helper_1.XmlHelper.createContentTypeChild(this.targetArchive, {
            PartName: `/ppt/charts/${this.subtype}${this.targetNumber}.xml`,
            ContentType: contentType,
        }));
    }
    appendColorToContentType() {
        return xml_helper_1.XmlHelper.append(xml_helper_1.XmlHelper.createContentTypeChild(this.targetArchive, {
            PartName: `/ppt/charts/colors${this.targetNumber}.xml`,
            ContentType: `application/vnd.ms-office.chartcolorstyle+xml`,
        }));
    }
    appendStyleToContentType() {
        return xml_helper_1.XmlHelper.append(xml_helper_1.XmlHelper.createContentTypeChild(this.targetArchive, {
            PartName: `/ppt/charts/style${this.targetNumber}.xml`,
            ContentType: `application/vnd.ms-office.chartstyle+xml`,
        }));
    }
    appendThemeOverrideToContentType() {
        return xml_helper_1.XmlHelper.append(xml_helper_1.XmlHelper.createContentTypeChild(this.targetArchive, {
            PartName: `/ppt/theme/themeOverride${this.targetNumber}.xml`,
            ContentType: `application/vnd.openxmlformats-officedocument.themeOverride+xml`,
        }));
    }
    static getAllOnSlide(archive, relsPath) {
        return __awaiter(this, void 0, void 0, function* () {
            return yield xml_helper_1.XmlHelper.getRelationshipTargetsByPrefix(archive, relsPath, [
                '../charts/chart',
                '../charts/chartEx',
            ]);
        });
    }
}
exports.Chart = Chart;
//# sourceMappingURL=chart.js.map