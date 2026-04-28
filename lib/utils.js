const { XMLSerializer } = require('xmldom')

module.exports.contentIsXML = (content) => {
  if (!Buffer.isBuffer(content) && typeof content !== 'string') {
    return false
  }

  const str = content.toString()

  return str.startsWith('<?xml') || (/^\s*<[\s\S]*>/).test(str)
}

module.exports.nodeListToArray = (nodes) => {
  const arr = []
  for (let i = 0; i < nodes.length; i++) {
    arr.push(nodes[i])
  }
  return arr
}

module.exports.pxToEMU = (val) => {
  return Math.round(val * 914400 / 96)
}

module.exports.cmToEMU = (val) => {
  // cm to dxa
  const dxa = val * 567.058823529411765
  // dxa to EMU
  return Math.round(dxa * 914400 / 72 / 20)
}

module.exports.serializeXml = (doc) => new XMLSerializer().serializeToString(doc).replace(/xmlns(:[a-z0-9]+)?="" ?/g, '')

/**
 * Returns a fresh rId number that does not collide with any existing
 * `Id="rIdN"` in the supplied list of <Relationship> elements.
 *
 * The previous logic in postprocess/image.js used `relsCount + 1`, which
 * collided whenever the rels file already contained non-sequential rIds. For
 * example a template carrying 21 relationships whose ids reach up to rId23
 * would receive a new rId22 that conflicts with the existing rId22. Microsoft
 * Word tolerates duplicate `Id` attributes (it picks the relationship whose
 * Type matches the consumer context), but the OOXML spec requires Ids to be
 * unique and stricter consumers do not. In particular LibreOffice — which is
 * frequently used to convert DOCX to PDF (e.g. via Gotenberg) — resolves
 * `<a:blip r:embed="rId22">` to the *first* declaration of rId22 in the rels
 * file. When that first declaration is not an image (often it is a font
 * table, theme, or header part), LibreOffice silently drops the image during
 * PDF rendering while still emitting the rest of the document, leaving users
 * with mysteriously image-less PDFs that look fine in Word.
 *
 * The allocator below preserves the historical starting point (`relsCount + 1`)
 * to keep media filenames stable for templates that never collided, then
 * advances past any rId already in use.
 *
 * @param {Array} relsElements - existing <Relationship> elements (any order)
 * @returns {number} a numeric rId suffix that is not yet used in the document
 */
module.exports.allocateFreshRid = (relsElements) => {
  const used = new Set()
  for (let i = 0; i < relsElements.length; i++) {
    const id = relsElements[i].getAttribute('Id')
    if (id) {
      used.add(id)
    }
  }
  let n = relsElements.length + 1
  while (used.has(`rId${n}`)) {
    n += 1
  }
  return n
}
