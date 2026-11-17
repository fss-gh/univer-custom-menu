import { getXmlAttibute } from "../common/method";
import { FrozenData } from "./types";
import { Element } from "./ReadXml";

export class LuckyFreezen implements FrozenData {
    horizen: number;
    vertical: number;
    constructor(pane: Element) {
        const xSplit = getXmlAttibute(pane.attributeList, "xSplit", "0");
        const ySplit = getXmlAttibute(pane.attributeList, "ySplit", "0");
        this.horizen = Number(ySplit);
        this.vertical = Number(xSplit);
    }
}