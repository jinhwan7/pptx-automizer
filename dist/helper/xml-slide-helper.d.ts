import { ElementInfo, ElementType, XmlDocument, XmlElement } from '../types/xml-types';
import HasShapes from '../classes/has-shapes';
import { TableInfo } from '../types/table-types';
export declare const nsMain = "http://schemas.openxmlformats.org/presentationml/2006/main";
export declare const mapUriType: {
    'http://schemas.openxmlformats.org/drawingml/2006/table': string;
    'http://schemas.openxmlformats.org/drawingml/2006/chart': string;
    'http://schemas.microsoft.com/office/drawing/2014/chartex': string;
    'http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject': string;
    'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink': string;
};
/**
 * Class that represents an XML slide helper
 */
export declare class XmlSlideHelper {
    private slideXml;
    protected hasShapes: HasShapes;
    /**
     * Constructor for the XmlSlideHelper class.
     * @param {XmlDocument} slideXml - The slide XML document to be used by the helper.
     * @param hasShapes
     */
    constructor(slideXml: XmlDocument, hasShapes?: HasShapes);
    getSlideCreationId(): number | undefined;
    /**
     * Get an array of ElementInfo objects for all named elements on a slide.
     * @param selector
     */
    getElement(selector: string): Promise<ElementInfo>;
    getElementByUniqueSelector(selector: {
        name: string;
        id: string;
    }): Promise<ElementInfo>;
    /**
     * Get an array of ElementInfo objects for all named elements on a slide.
     * @param filterTags Use an array of strings to filter the output array
     */
    getAllElements(filterTags?: string[]): ElementInfo[];
    /**
     * Get all text element IDs from the slide.
     * @return {string[]} An array of text element IDs.
     */
    getAllTextElementIds(useCreationIds?: boolean): string[];
    getAllTextElementUniqueSelectors(): {
        name: string;
        id: string;
    }[];
    static getElementInfo(slideElement: XmlElement): ElementInfo;
    /**
     * Retrieves a list of all named elements on a slide. Automation requires at least a name.
     * @param filterTags Use an array of strings to filter the output array
     */
    getNamedElements(filterTags?: string[]): XmlElement[];
    static getTextBody(shapeNode: XmlElement): XmlElement;
    static parseTextFragments(shapeNode: XmlElement): string[];
    static getTextByLine(shapeNode: XmlElement): string[];
    static getNonVisibleProperties(shapeNode: XmlElement): XmlElement;
    static getImageAltText(slideElement: XmlElement): string;
    static getElementName(slideElement: XmlElement): string;
    static getElementUniqueSelector(slideElement: XmlElement): {
        name: string;
        id: string;
    };
    static getElementCreationId(slideElement: XmlElement): string | undefined;
    /**
     * Parses local tag name to specify element type in case it is a 'graphicFrame'.
     * @param slideElementParent
     */
    static getElementType(slideElementParent: XmlElement): ElementType;
    static parseShapeCoordinates(slideElementParent: XmlElement): {
        x: number;
        y: number;
        cx: number;
        cy: number;
    };
    static parseCoordinate: (element: XmlElement, attributeName: string) => number;
    /**
     * Asynchronously retrieves the dimensions of a slide.
     * Tries to find the dimensions from the slide XML, then from the layout, master, and presentation XMLs in order.
     *
     * @returns {Promise<{ width: number, height: number }>} The dimensions of the slide.
     * @throws Error if unable to determine dimensions.
     */
    getDimensions(): Promise<{
        width: number;
        height: number;
    }>;
    /**
     * Fetches an XML file from the given path and extracts the dimensions.
     *
     * @param {string} path - The path of the XML file.
     * @returns {Promise<{ width: number; height: number } | null>} - A promise that resolves with an object containing the width and height, or `null` if there was an error.
     */
    getAndExtractDimensions: (path: string) => Promise<{
        width: number;
        height: number;
    } | null>;
    static readTableInfo: (element: XmlElement) => TableInfo[];
}
