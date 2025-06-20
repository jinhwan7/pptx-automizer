import { ShapeTargetType, SourceIdentifier } from '../types/types';
import { ISlide } from '../interfaces/islide';
import { IPresentationProps } from '../interfaces/ipresentation-props';
import { PresTemplate } from '../interfaces/pres-template';
import { RootPresTemplate } from '../interfaces/root-pres-template';
import HasShapes from './has-shapes';
export declare class Slide extends HasShapes implements ISlide {
    targetType: ShapeTargetType;
    constructor(params: {
        presentation: IPresentationProps;
        template: PresTemplate;
        slideIdentifier: SourceIdentifier;
    });
    /**
     * Appends slide
     * @internal
     * @param targetTemplate
     * @returns append
     */
    append(targetTemplate: RootPresTemplate): Promise<void>;
    /**
     * Use another slide layout.
     * @param targetLayoutId
     */
    useSlideLayout(layoutId?: number | string): this;
    /**
     * The current slide will be fully calculated, but removed from slide
     * sortation.
     */
    drop(): void;
    /**
     * Find another slide layout by name.
     * @param targetLayoutName
     */
    useNamedSlideLayout(targetLayoutName: string): Promise<number>;
    /**
     * Use another slide layout by index or detect original index.
     * @param targetLayoutIndex
     */
    useIndexedSlideLayout(targetLayoutIndex?: number): Promise<number>;
    autoImportSourceSlideMaster(templateName: string, sourceLayoutId: number): Promise<any>;
    /**
     * Copys slide files
     * @internal
     */
    copySlideFiles(): Promise<void>;
    /**
     * Remove a slide from presentation's slide list.
     * ToDo: Find the current count for this slide;
     * ToDo: this.targetNumber is undefined at this point.
     */
    remove(slide: number): void;
}
