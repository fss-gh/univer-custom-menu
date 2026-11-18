import JSZip from "@progress/jszip-esm";

/**
 * 对 ExcelJS 生成的 xlsx buffer 做一次后处理，避免 Excel 打开时提示“文件已损坏/修复”：
 *
 * 1. ExcelJS（@zwight/exceljs）在导出时总会向 xl/_rels/workbook.xml.rels 写入
 *    一个指向 xl/cellimages.xml 的 WPS cellImage 关系，但只有真正使用 addCellImage
 *    时才会生成 xl/cellimages.xml。没有 cell image 的工作簿会因此带上一个“悬空关系”，
 *    Excel 打开时会提示文件损坏，甚至修复后丢失内容。
 *
 * 2. x14 扩展数据条（x14:dataBar）中，Excel 保存的是 autoMin/autoMax，
 *    而 ExcelJS 写的是 min/max。这里统一转成 Excel 原生写法。
 *
 * 3. ExcelJS 写 x14:dataBar 颜色子节点的顺序是 borderColor、negativeBorderColor、
 *    negativeFillColor、axisColor，但 OOXML/Excel 实际期望的顺序是 borderColor、
 *    negativeFillColor、negativeBorderColor、axisColor。顺序不对也会触发 Excel 修复提示，
 *    这里统一重排为 Excel 原生顺序。
 *
 * 4. ExcelJS 把 x14:conditionalFormattings 的 extLst 放在 <legacyDrawing/> 之前，
 *    但 CT_Worksheet 要求 extLst 必须是最后一个子元素；顺序错误同样会让 Excel
 *    认为文件损坏。这里把该 extLst 移动到 </worksheet> 之前。
 */
function moveX14ConditionalFormattingsToEnd(xml: string): string {
    const extLstRe = /<extLst>([\s\S]*?)<\/extLst>/g;
    let moved = "";
    let result = xml.replace(extLstRe, (whole, inner) => {
        if (inner.includes("78C0D931-6437-407d-A8EE-F0AAD7539E65")) {
            moved = whole;
            return "";
        }
        return whole;
    });

    if (moved) {
        result = result.replace(/<\/worksheet>/, moved + "</worksheet>");
    }

    return result;
}
function normalizeX14DataBar(xml: string): string {
    let result = xml
        .replace(/(<x14:cfvo type=")min(")/g, "$1autoMin$2")
        .replace(/(<x14:cfvo type=")max(")/g, "$1autoMax$2");

    result = result.replace(
        /(<x14:dataBar\b[^>]*>)([\s\S]*?)(<\/x14:dataBar>)/g,
        (whole, open: string, inner: string, close: string) => {
            const cfvos = (
                inner.match(/<x14:cfvo\b[\s\S]*?<\/x14:cfvo>|<x14:cfvo\b[^>]*\/>/g) || []
            ).join("");

            const pickColor = (tag: string) => {
                const reg = new RegExp(`<${tag}\\b[^>]*\\/>`, "g");
                return (inner.match(reg) || []).join("");
            };

            const borderColor = pickColor("x14:borderColor");
            const negativeFillColor = pickColor("x14:negativeFillColor");
            const negativeBorderColor = pickColor("x14:negativeBorderColor");
            const axisColor = pickColor("x14:axisColor");

            const newInner = [
                cfvos,
                borderColor,
                negativeFillColor,
                negativeBorderColor,
                axisColor,
            ]
                .filter(Boolean)
                .join("");

            return open + newInner + close;
        }
    );

    return result;
}

export async function sanitizeExportedXml(buffer: ArrayBuffer): Promise<ArrayBuffer> {
    const zip = new JSZip();
    await zip.loadAsync(buffer);

    let changed = false;

    // 1. 清理悬空的 WPS cellImage 关系
    if (!zip.file("xl/cellimages.xml")) {
        const relPath = "xl/_rels/workbook.xml.rels";
        const relFile = zip.file(relPath);
        if (relFile) {
            const relXml = await relFile.async("string");
            const cleaned = relXml.replace(
                /<Relationship\s+[^>]*Type="http:\/\/www\.wps\.cn\/officeDocument\/2020\/cellImage"[^>]*\/?>/g,
                ""
            );
            if (cleaned !== relXml) {
                zip.file(relPath, cleaned);
                changed = true;
            }
        }
    }

    // 2. x14 数据条规范化：autoMin/autoMax + 颜色子节点顺序
    const sheetPaths = Object.keys(zip.files).filter((path) =>
        /^xl\/worksheets\/sheet\d+\.xml$/.test(path) && !zip.files[path].dir
    );
    for (const path of sheetPaths) {
        const sheetFile = zip.files[path];
        const xml = await sheetFile.async("string");
        const normalized = moveX14ConditionalFormattingsToEnd(normalizeX14DataBar(xml));
        if (normalized !== xml) {
            zip.file(path, normalized);
            changed = true;
        }
    }

    if (!changed) {
        return buffer;
    }

    return zip.generateAsync({
        type: "arraybuffer",
        compression: "DEFLATE",
        compressionOptions: { level: 6 },
    });
}
