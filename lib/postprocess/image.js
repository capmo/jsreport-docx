const sizeOf = require('image-size')
const axios = require('axios')
const pLimit = require('p-limit')
// const path = require('path')
// const fs = require('fs')

const stringReplaceAsync = require('string-replace-async')
const { nodeListToArray, pxToEMU, cmToEMU } = require('../utils')
const { DOMParser, XMLSerializer } = require('xmldom')

const getDimension = value => {
  const regexp = /^(\d+(.\d+)?)(cm|px)$/
  const match = regexp.exec(value)

  if (match) {
    return {
      value: parseFloat(match[1]),
      unit: match[3]
    }
  }

  return null
}

/**
 * Use module wide request limitation
 */
const limit = pLimit(64)

module.exports = async files => {
  const contentTypesFile = files.find(f => f.path === '[Content_Types].xml')
  const types = contentTypesFile.doc.getElementsByTagName('Types')[0]

  let pngDefault = nodeListToArray(types.getElementsByTagName('Default')).find(
    d => d.getAttribute('Extension') === 'png'
  )

  if (!pngDefault) {
    const defaultPng = contentTypesFile.doc.createElement('Default')
    defaultPng.setAttribute('Extension', 'png')
    defaultPng.setAttribute('ContentType', 'image/png')
    types.appendChild(defaultPng)
  }

  const relsDoc = files.find(f => f.path === 'word/_rels/document.xml.rels')
    .doc

  const documentFile = files.find(f => f.path === 'word/document.xml')

  documentFile.data = await stringReplaceAsync(
    documentFile.data.toString(),
    /<w:drawing>[^]*?(?=<\/w:drawing>)<\/w:drawing>/g,
    async val => {
      const elDrawing = new DOMParser().parseFromString(val)
      const elLinkClicks = elDrawing.getElementsByTagName('a:hlinkClick')
      const elLinkClick = elLinkClicks[0]

      if (!elLinkClick) {
        return val
      }

      const tooltip = elLinkClick.getAttribute('tooltip')

      if (tooltip == null || !tooltip.includes('$docxImage')) {
        return
      }

      const match = tooltip.match(/\$docxImage([^$]*)\$/)
      elLinkClick.setAttribute('tooltip', tooltip.replace(match[0], ''))
      const imageConfig = JSON.parse(match[1])

      // somehow there are duplicated hlinkclick els produced by word, we need to clean them up
      for (let i = 1; i < elLinkClicks.length; i++) {
        const elLinkClick = elLinkClicks[i]
        const match = tooltip.match(/\$docxImage([^$]*)\$/)
        elLinkClick.setAttribute('tooltip', tooltip.replace(match[0], ''))
      }

      let imageBuffer
      let imageExtensions
      if (imageConfig.src && imageConfig.src.startsWith('data:')) {
        const imageSrc = imageConfig.src
        imageExtensions = imageSrc.split(';')[0].split('/')[1]
        imageBuffer = Buffer.from(
          imageSrc.split(';')[1].substring('base64,'.length),
          'base64'
        )
      } else {
        const response = await limit(() => axios({
          url: imageConfig.src,
          responseType: 'arraybuffer',
          method: 'GET'
        }))
        const contentType =
          response.headers['content-type'] || response.headers['Content-Type']
        imageExtensions = contentType.split('/')[1]
        imageBuffer = Buffer.from(response.data)

        // check if the file is valid
        try {
          sizeOf(imageBuffer)
        } catch (e) {
          const imageSrc = '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'
          imageExtensions = 'png'
          imageBuffer = Buffer.from(imageSrc, 'base64')
          console.warn('Image from ' + imageConfig.src + ' throws the following error: ' + e)
        }
      }

      const relsElements = nodeListToArray(
        relsDoc.getElementsByTagName('Relationship')
      )
      const relsCount = relsElements.length
      const id = relsCount + 1
      const relEl = relsDoc.createElement('Relationship')
      relEl.setAttribute('Id', `rId${id}`)
      relEl.setAttribute(
        'Type',
        'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image'
      )
      relEl.setAttribute('Target', `media/imageDocx${id}.${imageExtensions}`)

      files.push({
        path: `word/media/imageDocx${id}.${imageExtensions}`,
        data: imageBuffer
      })

      relsDoc.getElementsByTagName('Relationships')[0].appendChild(relEl)

      const relPlaceholder = elDrawing.getElementsByTagName('a:blip')[0]
      const wpExtendEl = elDrawing.getElementsByTagName('wp:extent')[0]

      let imageWidthEMU
      let imageHeightEMU

      if (imageConfig.width != null || imageConfig.height != null) {
        const imageDimension = sizeOf(imageBuffer)
        const targetWidth = getDimension(imageConfig.width)
        const targetHeight = getDimension(imageConfig.height)

        if (targetWidth) {
          imageWidthEMU =
            targetWidth.unit === 'cm'
              ? cmToEMU(targetWidth.value)
              : pxToEMU(targetWidth.value)
        }

        if (targetHeight) {
          imageHeightEMU =
            targetHeight.unit === 'cm'
              ? cmToEMU(targetHeight.value)
              : pxToEMU(targetHeight.value)
        }

        if (imageWidthEMU != null && imageHeightEMU == null) {
          // adjust height based on aspect ratio of image
          imageHeightEMU = Math.round(
            imageWidthEMU *
              (pxToEMU(imageDimension.height) / pxToEMU(imageDimension.width))
          )
        } else if (imageHeightEMU != null && imageWidthEMU == null) {
          // adjust width based on aspect ratio of image
          imageWidthEMU = Math.round(
            imageHeightEMU *
              (pxToEMU(imageDimension.width) / pxToEMU(imageDimension.height))
          )
        }
      } else if (imageConfig.usePlaceholderSize) {
        // taking existing size defined in word
        imageWidthEMU = parseFloat(wpExtendEl.getAttribute('cx'))
        imageHeightEMU = parseFloat(wpExtendEl.getAttribute('cy'))
      } else {
        const imageDimension = sizeOf(imageBuffer)
        imageWidthEMU = pxToEMU(imageDimension.width)
        imageHeightEMU = pxToEMU(imageDimension.height)
      }

      relPlaceholder.setAttribute('r:embed', `rId${id}`)

      wpExtendEl.setAttribute('cx', imageWidthEMU)
      wpExtendEl.setAttribute('cy', imageHeightEMU)
      const aExtEl = elDrawing
        .getElementsByTagName('a:xfrm')[0]
        .getElementsByTagName('a:ext')[0]
      aExtEl.setAttribute('cx', imageWidthEMU)
      aExtEl.setAttribute('cy', imageHeightEMU)

      return new XMLSerializer()
        .serializeToString(elDrawing)
        .replace(/xmlns(:[a-z0-9]+)?="" ?/g, '')
    }
  )
}
