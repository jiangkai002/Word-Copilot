/** Word.js insertOoxml 使用的最小 Flat OPC 文档包。 */
const PKG_NS = "http://schemas.microsoft.com/office/2006/xmlPackage";
const REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships";
const OFFICE_DOC_REL =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";
const MAIN_DOCUMENT_CONTENT_TYPE =
  "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml";
const W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const M_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math";

/** 把正文级 OOXML 包装为 Range/Body.insertOoxml 可稳定导入的完整文档包。 */
export function bodyContentToFlatOpc(bodyContent: string): string {
  return (
    `<pkg:package xmlns:pkg="${PKG_NS}">` +
    `<pkg:part pkg:name="/_rels/.rels"` +
    ` pkg:contentType="application/vnd.openxmlformats-package.relationships+xml"` +
    ` pkg:padding="512"><pkg:xmlData>` +
    `<Relationships xmlns="${REL_NS}">` +
    `<Relationship Id="rId1" Type="${OFFICE_DOC_REL}" Target="word/document.xml"/>` +
    `</Relationships></pkg:xmlData></pkg:part>` +
    `<pkg:part pkg:name="/word/document.xml" pkg:contentType="${MAIN_DOCUMENT_CONTENT_TYPE}">` +
    `<pkg:xmlData><w:document xmlns:w="${W_NS}" xmlns:m="${M_NS}">` +
    `<w:body>${bodyContent}</w:body></w:document></pkg:xmlData></pkg:part>` +
    `</pkg:package>`
  );
}
