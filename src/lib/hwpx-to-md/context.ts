import { CharProperty, ParaProperty, StyleEntry, ImageInfo, BorderFill } from './types';

export class MdConversionContext {
    public ns = {
        hp: 'http://www.hancom.co.kr/hwpml/2011/paragraph',
        hh: 'http://www.hancom.co.kr/hwpml/2011/head',
        hc: 'http://www.hancom.co.kr/hwpml/2011/core',
    };

    public charProperties = new Map<string, CharProperty>();
    public paraProperties = new Map<string, ParaProperty>();
    public styles = new Map<string, StyleEntry>();
    public images = new Map<string, ImageInfo>();
    public borderFills = new Map<string, BorderFill>();
}
