import { CharProperty, ParaProperty, StyleEntry, ImageInfo, BorderFill } from './types';

export declare class MdConversionContext {
    ns: {
        hp: string;
        hh: string;
        hc: string;
    };
    charProperties: Map<string, CharProperty>;
    paraProperties: Map<string, ParaProperty>;
    styles: Map<string, StyleEntry>;
    images: Map<string, ImageInfo>;
    borderFills: Map<string, BorderFill>;
}
