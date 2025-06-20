import { ShapeModificationCallback } from '../types/types';
/**
 * Helper class for modifying hyperlinks in PowerPoint elements
 */
export default class ModifyHyperlinkHelper {
    /**
     * Set the target URL of a hyperlink (WIP)
     *
     * @param target The new target URL for the hyperlink
     * @param isExternal Whether the hyperlink is external (true) or internal (false)
     * @returns A callback function that modifies the hyperlink
     */
    static setHyperlinkTarget: (target: string, isExternal?: boolean) => ShapeModificationCallback;
    /**
     * Remove hyperlinks from an element (WIP)
     *
     * @returns A callback function that removes hyperlinks
     */
    static removeHyperlink: () => ShapeModificationCallback;
    /**
     * Add a hyperlink to an element
     *
     * @param target The target URL for external links, or slide number for internal links
     * @returns A callback function that adds a hyperlink
     */
    static addHyperlink: (target: string | number) => ShapeModificationCallback;
}
