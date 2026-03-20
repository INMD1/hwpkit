import { CharProperty, ParaProperty, BorderFill, ImageInfo, ImageRel } from './types';

export class ConversionContext {
    public ns = {
        hp: 'http://www.hancom.co.kr/hwpml/2011/paragraph',
        hh: 'http://www.hancom.co.kr/hwpml/2011/head',
        hc: 'http://www.hancom.co.kr/hwpml/2011/core',
        hs: 'http://www.hancom.co.kr/hwpml/2011/section',
        ha: 'http://www.hancom.co.kr/hwpml/2011/app',
        opf: 'http://www.idpf.org/2007/opf/'
    };

    public charProperties = new Map<string, CharProperty>();
    public paraProperties = new Map<string, ParaProperty>();
    public borderFills = new Map<string, BorderFill>();
    public fontFaces: { [lang: string]: { id: string, face: string }[] } = {};
    public images = new Map<string, ImageInfo>();
    public imageRels: ImageRel[] = [];
    public relIdCounter = 1;

    // 페이지 설정
    public pageWidth = 11906;
    public pageHeight = 16838;
    public marginTop = 1440;
    public marginBottom = 1440;
    public marginLeft = 1800;
    public marginRight = 1800;
    public marginHeader = 851;
    public marginFooter = 851;
}
