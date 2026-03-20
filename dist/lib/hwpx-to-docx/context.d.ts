import { CharProperty, ParaProperty, BorderFill, ImageInfo, ImageRel } from './types';

export declare class ConversionContext {
    ns: {
        hp: string;
        hh: string;
        hc: string;
        hs: string;
        ha: string;
        opf: string;
    };
    charProperties: Map<string, CharProperty>;
    paraProperties: Map<string, ParaProperty>;
    borderFills: Map<string, BorderFill>;
    fontFaces: {
        [lang: string]: {
            id: string;
            face: string;
        }[];
    };
    images: Map<string, ImageInfo>;
    imageRels: ImageRel[];
    relIdCounter: number;
    pageWidth: number;
    pageHeight: number;
    marginTop: number;
    marginBottom: number;
    marginLeft: number;
    marginRight: number;
    marginHeader: number;
    marginFooter: number;
}
