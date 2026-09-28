const JSZip = require('jszip');

const CONTENT_TYPES_PATH = '[Content_Types].xml';
const CONTENT_TYPES_NS = 'http://schemas.openxmlformats.org/package/2006/content-types';
const PRESENTATION_NS = 'http://schemas.openxmlformats.org/presentationml/2006/main';
const DRAWING_NS = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const PPTX_MIME = 'application/vnd.openxmlformats-officedocument.presentationml.presentation';

function getXmlApi() {
    if (typeof window !== 'undefined' && window.DOMParser && window.XMLSerializer) {
        return { DOMParser: window.DOMParser, XMLSerializer: window.XMLSerializer };
    }
    return require('@xmldom/xmldom');
}

function parseXml(xml, path) {
    const { DOMParser } = getXmlApi();
    const document = new DOMParser().parseFromString(xml, 'application/xml');
    if (!document || document.getElementsByTagName('parsererror').length) {
        throw new Error(`PPTX 中的 XML 无法解析：${path}`);
    }
    return document;
}

function serializeXml(document) {
    const { XMLSerializer } = getXmlApi();
    return new XMLSerializer().serializeToString(document);
}

function elementChildren(node) {
    return Array.prototype.filter.call(node.childNodes || [], child => child.nodeType === 1);
}

function elementsByLocalName(node, localName) {
    return Array.prototype.filter.call(node.getElementsByTagName('*'), child => child.localName === localName);
}

function hasZipPart(zip, path) {
    return Boolean(zip.files[path] && !zip.files[path].dir);
}

function normalizePackagePath(path) {
    const normalized = [];
    String(path || '').split('/').forEach((part) => {
        if (!part || part === '.') return;
        if (part === '..') {
            if (!normalized.length) throw new Error(`PPTX Relationship 路径越界：${path}`);
            normalized.pop();
            return;
        }
        normalized.push(part);
    });
    return normalized.join('/');
}

function relationshipBaseDirectory(relsPath) {
    if (relsPath.indexOf('_rels/') === 0) return '';
    const marker = '/_rels/';
    const markerIndex = relsPath.indexOf(marker);
    return markerIndex === -1 ? '' : relsPath.slice(0, markerIndex);
}

function resolveRelationshipTarget(relsPath, target) {
    if (String(target).charAt(0) === '/') return normalizePackagePath(String(target).slice(1));
    const base = relationshipBaseDirectory(relsPath);
    return normalizePackagePath(base ? `${base}/${target}` : target);
}

async function readZipText(zip, path) {
    const file = zip.file(path);
    if (!file) throw new Error(`PPTX 缺少必要文件：${path}`);
    return file.async('string');
}

async function writeXml(zip, path, transform) {
    const document = parseXml(await readZipText(zip, path), path);
    transform(document);
    zip.file(path, serializeXml(document));
}

async function removeDanglingContentTypeOverrides(zip) {
    await writeXml(zip, CONTENT_TYPES_PATH, (document) => {
        elementsByLocalName(document, 'Override').forEach((override) => {
            const partName = override.getAttribute('PartName');
            const packagePath = String(partName || '').replace(/^\//, '');
            if (packagePath && !hasZipPart(zip, packagePath)) {
                override.parentNode.removeChild(override);
            }
        });
    });
}

async function ensureNotesMasterTheme(zip) {
    const relPath = 'ppt/notesMasters/_rels/notesMaster1.xml.rels';
    const theme1Path = 'ppt/theme/theme1.xml';
    const theme2Path = 'ppt/theme/theme2.xml';

    if (!hasZipPart(zip, relPath)) return;
    if (!hasZipPart(zip, theme1Path)) {
        throw new Error('PPTX notesMaster 存在，但缺少 theme1.xml。');
    }
    if (!hasZipPart(zip, theme2Path)) {
        zip.file(theme2Path, await readZipText(zip, theme1Path));
    }

    await writeXml(zip, relPath, (document) => {
        const themeRelationship = elementsByLocalName(document, 'Relationship').find((relationship) => (
            /\/theme$/.test(relationship.getAttribute('Type') || '')
        ));
        if (!themeRelationship) {
            throw new Error('PPTX notesMaster 缺少 theme Relationship。');
        }
        themeRelationship.setAttribute('Target', '../theme/theme2.xml');
    });

    await writeXml(zip, CONTENT_TYPES_PATH, (document) => {
        const hasTheme2Override = elementsByLocalName(document, 'Override').some((override) => (
            override.getAttribute('PartName') === '/ppt/theme/theme2.xml'
        ));
        if (hasTheme2Override) return;

        const override = document.createElementNS(CONTENT_TYPES_NS, 'Override');
        override.setAttribute('PartName', '/ppt/theme/theme2.xml');
        override.setAttribute(
            'ContentType',
            'application/vnd.openxmlformats-officedocument.theme+xml',
        );
        document.documentElement.appendChild(override);
    });
}

function createEmptyTextBody(document) {
    const textBody = document.createElementNS(PRESENTATION_NS, 'p:txBody');
    const bodyProperties = document.createElementNS(DRAWING_NS, 'a:bodyPr');
    const listStyle = document.createElementNS(DRAWING_NS, 'a:lstStyle');
    const paragraph = document.createElementNS(DRAWING_NS, 'a:p');
    const endParagraphProperties = document.createElementNS(DRAWING_NS, 'a:endParaRPr');

    endParagraphProperties.setAttribute('lang', 'zh-CN');
    paragraph.appendChild(endParagraphProperties);
    textBody.appendChild(bodyProperties);
    textBody.appendChild(listStyle);
    textBody.appendChild(paragraph);
    return textBody;
}

async function addMissingShapeTextBodies(zip) {
    const slidePaths = Object.keys(zip.files).filter(path => /^ppt\/slides\/slide\d+\.xml$/.test(path));
    await Promise.all(slidePaths.map(async (slidePath) => {
        await writeXml(zip, slidePath, (document) => {
            elementsByLocalName(document, 'sp').forEach((shape) => {
                const children = elementChildren(shape);
                if (children.some(child => child.localName === 'txBody')) return;

                const extensionList = children.find(child => child.localName === 'extLst');
                shape.insertBefore(createEmptyTextBody(document), extensionList || null);
            });
        });
    }));
}

async function validatePptxPackage(zip) {
    const contentTypes = parseXml(await readZipText(zip, CONTENT_TYPES_PATH), CONTENT_TYPES_PATH);
    elementsByLocalName(contentTypes, 'Override').forEach((override) => {
        const packagePath = String(override.getAttribute('PartName') || '').replace(/^\//, '');
        if (packagePath && !hasZipPart(zip, packagePath)) {
            throw new Error(`PPTX Content Type 指向不存在的文件：${packagePath}`);
        }
    });

    const xmlPaths = Object.keys(zip.files).filter(path => /\.(xml|rels)$/.test(path) && !zip.files[path].dir);
    await Promise.all(xmlPaths.map(async (path) => {
        const xml = await readZipText(zip, path);
        const document = parseXml(xml, path);
        if (!/\.rels$/.test(path)) return;

        elementsByLocalName(document, 'Relationship').forEach((relationship) => {
            if ((relationship.getAttribute('TargetMode') || '').toLowerCase() === 'external') return;
            const target = relationship.getAttribute('Target');
            const packagePath = resolveRelationshipTarget(path, target);
            if (!hasZipPart(zip, packagePath)) {
                throw new Error(`PPTX Relationship 指向不存在的文件：${path} -> ${packagePath}`);
            }
        });
    }));
}

async function repairPptxPackage(input, options = {}) {
    const zip = await JSZip.loadAsync(input);

    await removeDanglingContentTypeOverrides(zip);
    await ensureNotesMasterTheme(zip);
    await addMissingShapeTextBodies(zip);
    await validatePptxPackage(zip);

    return zip.generateAsync({
        type: options.outputType || 'blob',
        mimeType: PPTX_MIME,
        compression: 'DEFLATE',
    });
}

function downloadPptxBlob(blob, fileName) {
    const requestedFileName = String(fileName || 'Presentation.pptx');
    const normalizedFileName = requestedFileName.toLowerCase().endsWith('.pptx')
        ? requestedFileName
        : `${requestedFileName}.pptx`;
    const url = window.URL.createObjectURL(blob);
    const anchor = document.createElement('a');

    anchor.href = url;
    anchor.download = normalizedFileName;
    document.body.appendChild(anchor);
    anchor.click();
    anchor.remove();
    window.setTimeout(() => window.URL.revokeObjectURL(url), 0);
}

module.exports = {
    downloadPptxBlob,
    repairPptxPackage,
};
