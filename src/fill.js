import tinycolor from 'tinycolor2'
import { getSchemeColorFromTheme } from './schemeColor'
import {
  applyShade,
  applyTint,
  applyLumOff,
  applyLumMod,
  applyHueMod,
  applySatMod,
  hslToRgb,
  getColorName2Hex,
} from './color'

import {
  base64ArrayBuffer,
  getTextByPathList,
  angleToDegrees,
  escapeHtml,
  getMimeType,
  toHex,
} from './utils'

export function getFillType(node) {
  let fillType = ''
  if (node['a:noFill']) fillType = 'NO_FILL'
  if (node['a:solidFill']) fillType = 'SOLID_FILL'
  if (node['a:gradFill']) fillType = 'GRADIENT_FILL'
  if (node['a:pattFill']) fillType = 'PATTERN_FILL'
  if (node['a:blipFill']) fillType = 'PIC_FILL'
  if (node['a:grpFill']) fillType = 'GROUP_FILL'

  return fillType
}

export async function getPicFill(type, node, warpObj, uploadFun) {
  if (!node) return ''

  let img
  const rId = getTextByPathList(node, ['a:blip', 'attrs', 'r:embed'])
  let imgPath
  if (type === 'slideBg' || type === 'slide') {
    imgPath = getTextByPathList(warpObj, ['slideResObj', rId, 'target'])
  }
  else if (type === 'slideLayoutBg') {
    imgPath = getTextByPathList(warpObj, ['layoutResObj', rId, 'target'])
  }
  else if (type === 'slideMasterBg') {
    imgPath = getTextByPathList(warpObj, ['masterResObj', rId, 'target'])
  }
  else if (type === 'themeBg') {
    imgPath = getTextByPathList(warpObj, ['themeResObj', rId, 'target'])
  }
  else if (type === 'diagramBg') {
    imgPath = getTextByPathList(warpObj, ['diagramResObj', rId, 'target'])
  }
  if (!imgPath) return imgPath

  img = getTextByPathList(warpObj, ['loaded-images', imgPath])
  if (!img) {
    imgPath = escapeHtml(imgPath)

    const imgExt = imgPath.split('.').pop()
    if (imgExt === 'xml') return ''

    const imgArrayBuffer = await warpObj['zip'].file(imgPath).async('arraybuffer')
    const imgMimeType = getMimeType(imgExt)
    if (uploadFun) {
      const uploadResp = await uploadFun(new Blob([imgArrayBuffer], { type: imgMimeType }), imgExt)
      img = uploadResp.url
    }
    else {
      img = `data:${imgMimeType};base64,${base64ArrayBuffer(imgArrayBuffer)}`
    }
    const loadedImages = warpObj['loaded-images'] || {}
    loadedImages[imgPath] = img
    warpObj['loaded-images'] = loadedImages
  }
  return img
}

export function getPicFillOpacity(node) {
  const aBlipNode = node['a:blip']

  const aphaModFixNode = getTextByPathList(aBlipNode, ['a:alphaModFix', 'attrs'])
  let opacity = 1
  if (aphaModFixNode && aphaModFixNode['amt'] && aphaModFixNode['amt'] !== '') {
    opacity = parseInt(aphaModFixNode['amt']) / 100000
  }

  return opacity
}

export function getPicFilters(node) {
  if (!node) return null

  const aBlipNode = node['a:blip']
  if (!aBlipNode) return null

  const filters = {}

  // 从a:extLst中获取滤镜效果（Microsoft Office 2010+扩展）
  const extLstNode = aBlipNode['a:extLst']
  if (extLstNode && extLstNode['a:ext']) {
    const extNodes = Array.isArray(extLstNode['a:ext']) ? extLstNode['a:ext'] : [extLstNode['a:ext']]

    for (const extNode of extNodes) {
      if (!extNode['a14:imgProps'] || !extNode['a14:imgProps']['a14:imgLayer']) continue

      const imgLayerNode = extNode['a14:imgProps']['a14:imgLayer']
      const imgEffects = imgLayerNode['a14:imgEffect']

      if (!imgEffects) continue

      const effectArray = Array.isArray(imgEffects) ? imgEffects : [imgEffects]

      for (const effect of effectArray) {
        // 饱和度
        if (effect['a14:saturation']) {
          const satAttr = getTextByPathList(effect, ['a14:saturation', 'attrs', 'sat'])
          if (satAttr) {
            filters.saturation = parseInt(satAttr) / 100000
          }
        }

        // 亮度、对比度
        if (effect['a14:brightnessContrast']) {
          const brightAttr = getTextByPathList(effect, ['a14:brightnessContrast', 'attrs', 'bright'])
          const contrastAttr = getTextByPathList(effect, ['a14:brightnessContrast', 'attrs', 'contrast'])

          if (brightAttr) {
            filters.brightness = parseInt(brightAttr) / 100000
          }
          if (contrastAttr) {
            filters.contrast = parseInt(contrastAttr) / 100000
          }
        }

        // 锐化/柔化
        if (effect['a14:sharpenSoften']) {
          const amountAttr = getTextByPathList(effect, ['a14:sharpenSoften', 'attrs', 'amount'])
          if (amountAttr) {
            const amount = parseInt(amountAttr) / 100000
            if (amount > 0) {
              filters.sharpen = amount
            }
            else {
              filters.soften = Math.abs(amount)
            }
          }
        }

        // 色温
        if (effect['a14:colorTemperature']) {
          const tempAttr = getTextByPathList(effect, ['a14:colorTemperature', 'attrs', 'colorTemp'])
          if (tempAttr) {
            filters.colorTemperature = parseInt(tempAttr)
          }
        }
      }
    }
  }

  return Object.keys(filters).length > 0 ? filters : null
}

export async function getBgPicFill(bgPr, sorce, warpObj, uploadFun) {
  const picBase64 = await getPicFill(sorce, bgPr['a:blipFill'], warpObj, uploadFun)
  const aBlipNode = bgPr['a:blipFill']['a:blip']

  const aphaModFixNode = getTextByPathList(aBlipNode, ['a:alphaModFix', 'attrs'])
  let opacity = 1
  if (aphaModFixNode && aphaModFixNode['amt'] && aphaModFixNode['amt'] !== '') {
    opacity = parseInt(aphaModFixNode['amt']) / 100000
  }

  return {
    picBase64,
    opacity,
  }
}

export function getGradientFill(node, warpObj) {
  const gsLst = node['a:gsLst']['a:gs']
  const colors = []
  for (let i = 0; i < gsLst.length; i++) {
    const lo_color = getSolidFill(gsLst[i], undefined, undefined, warpObj)
    const pos = getTextByPathList(gsLst[i], ['attrs', 'pos'])
    
    colors[i] = {
      pos: pos ? (pos / 1000 + '%') : '',
      color: lo_color,
    }
  }
  const lin = node['a:lin']
  let rot = 0
  let pathType = 'line'
  if (lin) rot = angleToDegrees(lin['attrs']['ang'])
  else {
    const path = node['a:path']
    if (path && path['attrs'] && path['attrs']['path']) pathType = path['attrs']['path'] 
  }
  return {
    rot,
    path: pathType,
    colors: colors.sort((a, b) => parseInt(a.pos) - parseInt(b.pos)),
  }
}

export function getPatternFill(node, warpObj) {
  if (!node) return null

  const pattFill = node['a:pattFill']
  if (!pattFill) return null

  const type = getTextByPathList(pattFill, ['attrs', 'prst'])

  const fgColorNode = pattFill['a:fgClr']
  const bgColorNode = pattFill['a:bgClr']

  let foregroundColor = '#000000'
  let backgroundColor = '#FFFFFF'

  if (fgColorNode) {
    foregroundColor = getSolidFill(fgColorNode, undefined, undefined, warpObj)
  }

  if (bgColorNode) {
    backgroundColor = getSolidFill(bgColorNode, undefined, undefined, warpObj)
  }

  return {
    type,
    foregroundColor,
    backgroundColor,
  }
}

export function getBgGradientFill(bgPr, phClr, slideMasterContent, warpObj) {
  if (bgPr) {
    const grdFill = bgPr['a:gradFill']
    const gsLst = grdFill['a:gsLst']['a:gs']
    const colors = []
    
    for (let i = 0; i < gsLst.length; i++) {
      const lo_color = getSolidFill(gsLst[i], slideMasterContent['p:sldMaster']['p:clrMap']['attrs'], phClr, warpObj)
      const pos = getTextByPathList(gsLst[i], ['attrs', 'pos'])

      colors[i] = {
        pos: pos ? (pos / 1000 + '%') : '',
        color: lo_color,
      }
    }
    const lin = grdFill['a:lin']
    let rot = 0
    let pathType = 'line'
    if (lin) rot = angleToDegrees(lin['attrs']['ang']) + 0
    else {
      const path = grdFill['a:path']
      if (path && path['attrs'] && path['attrs']['path']) pathType = path['attrs']['path'] 
    }
    return {
      rot,
      path: pathType,
      colors: colors.sort((a, b) => parseInt(a.pos) - parseInt(b.pos)),
    }
  }
  else if (phClr) {
    return phClr.indexOf('#') === -1 ? `#${phClr}` : phClr
  }
  return null
}

export async function getSlideBackgroundFill(warpObj, uploadFun) {
  const slideContent = warpObj['slideContent']
  const slideLayoutContent = warpObj['slideLayoutContent']
  const slideMasterContent = warpObj['slideMasterContent']
  
  let bgPr = getTextByPathList(slideContent, ['p:sld', 'p:cSld', 'p:bg', 'p:bgPr'])
  let bgRef = getTextByPathList(slideContent, ['p:sld', 'p:cSld', 'p:bg', 'p:bgRef'])

  let background = '#fff'
  let backgroundType = 'color'

  if (bgPr) {
    const bgFillTyp = getFillType(bgPr)

    if (bgFillTyp === 'SOLID_FILL') {
      const sldFill = bgPr['a:solidFill']
      let clrMapOvr
      const sldClrMapOvr = getTextByPathList(slideContent, ['p:sld', 'p:clrMapOvr', 'a:overrideClrMapping', 'attrs'])
      if (sldClrMapOvr) clrMapOvr = sldClrMapOvr
      else {
        const sldClrMapOvr = getTextByPathList(slideLayoutContent, ['p:sldLayout', 'p:clrMapOvr', 'a:overrideClrMapping', 'attrs'])
        if (sldClrMapOvr) clrMapOvr = sldClrMapOvr
        else clrMapOvr = getTextByPathList(slideMasterContent, ['p:sldMaster', 'p:clrMap', 'attrs'])
      }
      const sldBgClr = getSolidFill(sldFill, clrMapOvr, undefined, warpObj)
      background = sldBgClr
    }
    else if (bgFillTyp === 'GRADIENT_FILL') {
      const gradientFill = getBgGradientFill(bgPr, undefined, slideMasterContent, warpObj)
      if (typeof gradientFill === 'string') {
        background = gradientFill
      }
      else if (gradientFill) {
        background = gradientFill
        backgroundType = 'gradient'
      }
    }
    else if (bgFillTyp === 'PIC_FILL') {
      background = await getBgPicFill(bgPr, 'slideBg', warpObj, uploadFun)
      backgroundType = 'image'
    }else if (bgFillTyp === 'PATTERN_FILL') {
      const patternFill = getPatternFill(bgPr, warpObj)
      if (patternFill) {
        background = patternFill
        backgroundType = 'pattern'
      }
    }
  }
  else if (bgRef) {
    let clrMapOvr
    const sldClrMapOvr = getTextByPathList(slideContent, ['p:sld', 'p:clrMapOvr', 'a:overrideClrMapping', 'attrs'])
    if (sldClrMapOvr) clrMapOvr = sldClrMapOvr
    else {
      const sldClrMapOvr = getTextByPathList(slideLayoutContent, ['p:sldLayout', 'p:clrMapOvr', 'a:overrideClrMapping', 'attrs'])
      if (sldClrMapOvr) clrMapOvr = sldClrMapOvr
      else clrMapOvr = getTextByPathList(slideMasterContent, ['p:sldMaster', 'p:clrMap', 'attrs'])
    }
    const phClr = getSolidFill(bgRef, clrMapOvr, undefined, warpObj)
    const idx = Number(bgRef['attrs']['idx'])

    if (idx > 1000) {
      const trueIdx = idx - 1000
      const bgFillLst = warpObj['themeContent']['a:theme']['a:themeElements']['a:fmtScheme']['a:bgFillStyleLst']
      const sortblAry = []
      Object.keys(bgFillLst).forEach(key => {
        const bgFillLstTyp = bgFillLst[key]
        if (key !== 'attrs') {
          if (bgFillLstTyp.constructor === Array) {
            for (let i = 0; i < bgFillLstTyp.length; i++) {
              const obj = {}
              obj[key] = bgFillLstTyp[i]
              if (bgFillLstTyp[i]['attrs']) {
                obj['idex'] = bgFillLstTyp[i]['attrs']['order']
                obj['attrs'] = {
                  'order': bgFillLstTyp[i]['attrs']['order']
                }
              }
              sortblAry.push(obj)
            }
          } 
          else {
            const obj = {}
            obj[key] = bgFillLstTyp
            if (bgFillLstTyp['attrs']) {
              obj['idex'] = bgFillLstTyp['attrs']['order']
              obj['attrs'] = {
                'order': bgFillLstTyp['attrs']['order']
              }
            }
            sortblAry.push(obj)
          }
        }
      })
      const sortByOrder = sortblAry.slice(0)
      sortByOrder.sort((a, b) => a.idex - b.idex)
      const bgFillLstIdx = sortByOrder[trueIdx - 1]
      const bgFillTyp = getFillType(bgFillLstIdx)
      if (bgFillTyp === 'SOLID_FILL') {
        const sldFill = bgFillLstIdx['a:solidFill']
        const sldBgClr = getSolidFill(sldFill, clrMapOvr, undefined, warpObj)
        background = sldBgClr
      } 
      else if (bgFillTyp === 'GRADIENT_FILL') {
        const gradientFill = getBgGradientFill(bgFillLstIdx, phClr, slideMasterContent, warpObj)
        if (typeof gradientFill === 'string') {
          background = gradientFill
        }
        else if (gradientFill) {
          background = gradientFill
          backgroundType = 'gradient'
        }
      }
    }
  }
  else {
    bgPr = getTextByPathList(slideLayoutContent, ['p:sldLayout', 'p:cSld', 'p:bg', 'p:bgPr'])
    bgRef = getTextByPathList(slideLayoutContent, ['p:sldLayout', 'p:cSld', 'p:bg', 'p:bgRef'])

    let clrMapOvr
    const sldClrMapOvr = getTextByPathList(slideLayoutContent, ['p:sldLayout', 'p:clrMapOvr', 'a:overrideClrMapping', 'attrs'])
    if (sldClrMapOvr) clrMapOvr = sldClrMapOvr
    else clrMapOvr = getTextByPathList(slideMasterContent, ['p:sldMaster', 'p:clrMap', 'attrs'])

    if (bgPr) {
      const bgFillTyp = getFillType(bgPr)
      if (bgFillTyp === 'SOLID_FILL') {
        const sldFill = bgPr['a:solidFill']
        const sldBgClr = getSolidFill(sldFill, clrMapOvr, undefined, warpObj)
        background = sldBgClr
      }
      else if (bgFillTyp === 'GRADIENT_FILL') {
        const gradientFill = getBgGradientFill(bgPr, undefined, slideMasterContent, warpObj)
        if (typeof gradientFill === 'string') {
          background = gradientFill
        }
        else if (gradientFill) {
          background = gradientFill
          backgroundType = 'gradient'
        }
      }
      else if (bgFillTyp === 'PIC_FILL') {
        background = await getBgPicFill(bgPr, 'slideLayoutBg', warpObj)
        backgroundType = 'image'
      }else if (bgFillTyp === 'PATTERN_FILL') {
        const patternFill = getPatternFill(bgPr, warpObj)
        if (patternFill) {
          background = patternFill
          backgroundType = 'pattern'
        }
      }
    }
    else if (bgRef) {
      const phClr = getSolidFill(bgRef, clrMapOvr, undefined, warpObj)
      const idx = Number(bgRef['attrs']['idx'])
  
      if (idx > 1000) {
        const trueIdx = idx - 1000
        const bgFillLst = warpObj['themeContent']['a:theme']['a:themeElements']['a:fmtScheme']['a:bgFillStyleLst']
        const sortblAry = []
        Object.keys(bgFillLst).forEach(key => {
          const bgFillLstTyp = bgFillLst[key]
          if (key !== 'attrs') {
            if (bgFillLstTyp.constructor === Array) {
              for (let i = 0; i < bgFillLstTyp.length; i++) {
                const obj = {}
                obj[key] = bgFillLstTyp[i]
                if (bgFillLstTyp[i]['attrs']) {
                  obj['idex'] = bgFillLstTyp[i]['attrs']['order']
                  obj['attrs'] = {
                    'order': bgFillLstTyp[i]['attrs']['order']
                  }
                }
                sortblAry.push(obj)
              }
            } 
            else {
              const obj = {}
              obj[key] = bgFillLstTyp
              if (bgFillLstTyp['attrs']) {
                obj['idex'] = bgFillLstTyp['attrs']['order']
                obj['attrs'] = {
                  'order': bgFillLstTyp['attrs']['order']
                }
              }
              sortblAry.push(obj)
            }
          }
        })
        const sortByOrder = sortblAry.slice(0)
        sortByOrder.sort((a, b) => a.idex - b.idex)
        const bgFillLstIdx = sortByOrder[trueIdx - 1]
        const bgFillTyp = getFillType(bgFillLstIdx)
        if (bgFillTyp === 'SOLID_FILL') {
          const sldFill = bgFillLstIdx['a:solidFill']
          const sldBgClr = getSolidFill(sldFill, clrMapOvr, undefined, warpObj)
          background = sldBgClr
        } 
        else if (bgFillTyp === 'GRADIENT_FILL') {
          const gradientFill = getBgGradientFill(bgFillLstIdx, phClr, slideMasterContent, warpObj)
          if (typeof gradientFill === 'string') {
            background = gradientFill
          }
          else if (gradientFill) {
            background = gradientFill
            backgroundType = 'gradient'
          }
        }
        else if (bgFillTyp === 'PIC_FILL') {
          background = await getBgPicFill(bgFillLstIdx, 'themeBg', warpObj)
          backgroundType = 'image'
        }else if (bgFillTyp === 'PATTERN_FILL') {
        const patternFill = getPatternFill(bgPr, warpObj)
        if (patternFill) {
          background = patternFill
          backgroundType = 'pattern'
        }
      }
      }
    }
    else {
      bgPr = getTextByPathList(slideMasterContent, ['p:sldMaster', 'p:cSld', 'p:bg', 'p:bgPr'])
      bgRef = getTextByPathList(slideMasterContent, ['p:sldMaster', 'p:cSld', 'p:bg', 'p:bgRef'])

      const clrMap = getTextByPathList(slideMasterContent, ['p:sldMaster', 'p:clrMap', 'attrs'])
      if (bgPr) {
        const bgFillTyp = getFillType(bgPr)
        if (bgFillTyp === 'SOLID_FILL') {
          const sldFill = bgPr['a:solidFill']
          const sldBgClr = getSolidFill(sldFill, clrMap, undefined, warpObj)
          background = sldBgClr
        }
        else if (bgFillTyp === 'GRADIENT_FILL') {
          const gradientFill = getBgGradientFill(bgPr, undefined, slideMasterContent, warpObj)
          if (typeof gradientFill === 'string') {
            background = gradientFill
          }
          else if (gradientFill) {
            background = gradientFill
            backgroundType = 'gradient'
          }
        }
        else if (bgFillTyp === 'PIC_FILL') {
          background = await getBgPicFill(bgPr, 'slideMasterBg', warpObj)
          backgroundType = 'image'
        }else if (bgFillTyp === 'PATTERN_FILL') {
        const patternFill = getPatternFill(bgPr, warpObj)
        if (patternFill) {
          background = patternFill
          backgroundType = 'pattern'
        }
      }
      }
      else if (bgRef) {
        const phClr = getSolidFill(bgRef, clrMap, undefined, warpObj)
        const idx = Number(bgRef['attrs']['idx'])
    
        if (idx > 1000) {
          const trueIdx = idx - 1000
          const bgFillLst = warpObj['themeContent']['a:theme']['a:themeElements']['a:fmtScheme']['a:bgFillStyleLst']
          const sortblAry = []
          Object.keys(bgFillLst).forEach(key => {
            const bgFillLstTyp = bgFillLst[key]
            if (key !== 'attrs') {
              if (bgFillLstTyp.constructor === Array) {
                for (let i = 0; i < bgFillLstTyp.length; i++) {
                  const obj = {}
                  obj[key] = bgFillLstTyp[i]
                  if (bgFillLstTyp[i]['attrs']) {
                    obj['idex'] = bgFillLstTyp[i]['attrs']['order']
                    obj['attrs'] = {
                      'order': bgFillLstTyp[i]['attrs']['order']
                    }
                  }
                  sortblAry.push(obj)
                }
              } 
              else {
                const obj = {}
                obj[key] = bgFillLstTyp
                if (bgFillLstTyp['attrs']) {
                  obj['idex'] = bgFillLstTyp['attrs']['order']
                  obj['attrs'] = {
                    'order': bgFillLstTyp['attrs']['order']
                  }
                }
                sortblAry.push(obj)
              }
            }
          })
          const sortByOrder = sortblAry.slice(0)
          sortByOrder.sort((a, b) => a.idex - b.idex)
          const bgFillLstIdx = sortByOrder[trueIdx - 1]
          const bgFillTyp = getFillType(bgFillLstIdx)
          if (bgFillTyp === 'SOLID_FILL') {
            const sldFill = bgFillLstIdx['a:solidFill']
            const sldBgClr = getSolidFill(sldFill, clrMapOvr, undefined, warpObj)
            background = sldBgClr
          } 
          else if (bgFillTyp === 'GRADIENT_FILL') {
            const gradientFill = getBgGradientFill(bgFillLstIdx, phClr, slideMasterContent, warpObj)
            if (typeof gradientFill === 'string') {
              background = gradientFill
            }
            else if (gradientFill) {
              background = gradientFill
              backgroundType = 'gradient'
            }
          }
          else if (bgFillTyp === 'PIC_FILL') {
            background = await getBgPicFill(bgFillLstIdx, 'themeBg', warpObj)
            backgroundType = 'image'
          }else if (bgFillTyp === 'PATTERN_FILL') {
        const patternFill = getPatternFill(bgPr, warpObj)
        if (patternFill) {
          background = patternFill
          backgroundType = 'pattern'
        }
      }
        }
      }
    }
  }
  return {
    type: backgroundType,
    value: background,
  }
}

/**
 * 获取形状的填充信息
 * @param {Object} node - 形状节点对象
 * @param {Object} pNode - 父节点对象
 * @param {boolean} isSvgMode - 是否为SVG模式
 * @param {Object} warpObj - 包装对象，包含各种资源和配置信息
 * @param {string} source - 资源来源类型
 * @param {Array} groupHierarchy - 组层次结构数组，默认为空数组
 * @returns {Object|string} 返回填充信息对象或字符串
 */
export async function getShapeFill(node, pNode, isSvgMode, warpObj, source, groupHierarchy = []) {
  const fillType = getFillType(getTextByPathList(node, ['p:spPr']))
  let type = 'color'
  let fillValue = ''
  // 根据不同的填充类型处理填充效果
  if (fillType === 'NO_FILL') {
    return isSvgMode ? 'none' : ''
  } 
  else if (fillType === 'SOLID_FILL') {
    const shpFill = node['p:spPr']['a:solidFill']
    fillValue = getSolidFill(shpFill, undefined, undefined, warpObj)
    type = 'color'
  }
  else if (fillType === 'GRADIENT_FILL') {
    const shpFill = node['p:spPr']['a:gradFill']
    fillValue = getGradientFill(shpFill, warpObj)
    type = 'gradient'
  }
  else if (fillType === 'PIC_FILL') {
    const shpFill = node['p:spPr']['a:blipFill']
    const picBase64 = await getPicFill(source, shpFill, warpObj)
    const opacity = getPicFillOpacity(shpFill)
    fillValue = {
      picBase64,
      opacity,
    }
    type = 'image'
  }else if (fillType === 'PATTERN_FILL') {
    const shpFill = node['p:spPr']['a:pattFill']
    fillValue = getPatternFill({ 'a:pattFill': shpFill }, warpObj)
    type = 'pattern'
  }
  else if (fillType === 'GROUP_FILL') {
    return findFillInGroupHierarchy(groupHierarchy, warpObj, source)
  }


  if (!fillType || fillType === 'NO_FILL') {
    const txBoxVal = getTextByPathList(node, ['p:nvSpPr', 'p:cNvSpPr', 'attrs', 'txBox'])
    if (txBoxVal === '1') {
      // 对于 文本框（txBox="1"），默认行为是 “无背景填充”（即透明）
      return ''
    }
  }
  // 当没有获取到填充值时，尝试从样式中获取填充引用
  if (!fillValue) {
    const clrName = getTextByPathList(node, ['p:style', 'a:fillRef'])
    fillValue = getSolidFill(clrName, undefined, undefined, warpObj)
    if (fillValue) {
      type = 'color'
    }
  }

  // 当仍未获取到填充值且父节点存在且填充类型为无填充时，返回相应值
  if (!fillValue && pNode && fillType === 'NO_FILL') {
    return isSvgMode ? 'none' : ''
  }

  return {
    type,
    value: fillValue,
  } 
}

async function findFillInGroupHierarchy(groupHierarchy, warpObj, source) {
  for (const groupNode of groupHierarchy) {
    if (!groupNode || !groupNode['p:grpSpPr']) continue

    const grpSpPr = groupNode['p:grpSpPr']
    const fillType = getFillType(grpSpPr)

    if (fillType === 'SOLID_FILL') {
      const shpFill = grpSpPr['a:solidFill']
      const fillValue = getSolidFill(shpFill, undefined, undefined, warpObj)
      if (fillValue) {
        return {
          type: 'color',
          value: fillValue,
        }
      }
    }
    else if (fillType === 'GRADIENT_FILL') {
      const shpFill = grpSpPr['a:gradFill']
      const fillValue = getGradientFill(shpFill, warpObj)
      if (fillValue) {
        return {
          type: 'gradient',
          value: fillValue,
        }
      }
    }
    else if (fillType === 'PIC_FILL') {
      const shpFill = grpSpPr['a:blipFill']
      const picBase64 = await getPicFill(source, shpFill, warpObj)
      const opacity = getPicFillOpacity(shpFill)
      if (picBase64) {
        return {
          type: 'image',
          value: {
            picBase64,
            opacity,
          },
        }
      }
    }
    else if (fillType === 'PATTERN_FILL') {
      const shpFill = grpSpPr['a:pattFill']
      const fillValue = getPatternFill({ 'a:pattFill': shpFill }, warpObj)
      if (fillValue) {
        return {
          type: 'pattern',
          value: fillValue,
        }
      }
    }
  }

  return null
}

/**
 * 获取纯色填充的颜色值
 * @param {Object} solidFill - 包含颜色定义的对象
 * @param {Object} clrMap - 颜色映射对象
 * @param {string} phClr - 占位符颜色
 * @param {Object} warpObj - 包装对象，包含主题和其他相关信息
 * @returns {string} 返回十六进制颜色值或RGBA颜色值
 */
export function getSolidFill(solidFill, clrMap, phClr, warpObj) {
  if (!solidFill) return ''

  let color = ''
  let clrNode

  // 处理RGB颜色
  if (solidFill['a:srgbClr']) {
    clrNode = solidFill['a:srgbClr']
    color = getTextByPathList(clrNode, ['attrs', 'val'])
  } 
  // 处理主题颜色方案
  else if (solidFill['a:schemeClr']) {
    clrNode = solidFill['a:schemeClr']
    const schemeClr = 'a:' + getTextByPathList(clrNode, ['attrs', 'val'])
    color = getSchemeColorFromTheme(schemeClr, warpObj, clrMap, phClr) || ''
  }
  // 处理ScRGB颜色
  else if (solidFill['a:scrgbClr']) {
    clrNode = solidFill['a:scrgbClr']
    const defBultColorVals = clrNode['attrs']
    const red = (defBultColorVals['r'].indexOf('%') !== -1) ? defBultColorVals['r'].split('%').shift() : defBultColorVals['r']
    const green = (defBultColorVals['g'].indexOf('%') !== -1) ? defBultColorVals['g'].split('%').shift() : defBultColorVals['g']
    const blue = (defBultColorVals['b'].indexOf('%') !== -1) ? defBultColorVals['b'].split('%').shift() : defBultColorVals['b']
    color = toHex(255 * (Number(red) / 100)) + toHex(255 * (Number(green) / 100)) + toHex(255 * (Number(blue) / 100))    
  } 
  // 处理预设颜色
  else if (solidFill['a:prstClr']) {
    clrNode = solidFill['a:prstClr']
    const prstClr = getTextByPathList(clrNode, ['attrs', 'val'])
    color = getColorName2Hex(prstClr)
  } 
  // 处理HSL颜色
  else if (solidFill['a:hslClr']) {
    clrNode = solidFill['a:hslClr']
    const defBultColorVals = clrNode['attrs']
    const hue = Number(defBultColorVals['hue']) / 100000
    const sat = Number((defBultColorVals['sat'].indexOf('%') !== -1) ? defBultColorVals['sat'].split('%').shift() : defBultColorVals['sat']) / 100
    const lum = Number((defBultColorVals['lum'].indexOf('%') !== -1) ? defBultColorVals['lum'].split('%').shift() : defBultColorVals['lum']) / 100
    const hsl2rgb = hslToRgb(hue, sat, lum)
    color = toHex(hsl2rgb.r) + toHex(hsl2rgb.g) + toHex(hsl2rgb.b)
  } 
  // 处理系统颜色
  else if (solidFill['a:sysClr']) {
    clrNode = solidFill['a:sysClr']
    const sysClr = getTextByPathList(clrNode, ['attrs', 'lastClr'])
    if (sysClr) color = sysClr
  }

  // 应用透明度效果
  let isAlpha = false
  const alpha = parseInt(getTextByPathList(clrNode, ['a:alpha', 'attrs', 'val'])) / 100000
  if (!isNaN(alpha)) {
    const al_color = tinycolor(color)
    al_color.setAlpha(alpha)
    color = al_color.toHex8()
    isAlpha = true
  }

  // 应用色调调整
  const hueMod = parseInt(getTextByPathList(clrNode, ['a:hueMod', 'attrs', 'val'])) / 100000
  if (!isNaN(hueMod)) {
    color = applyHueMod(color, hueMod, isAlpha)
  }
  
  // 应用亮度调整
  const lumMod = parseInt(getTextByPathList(clrNode, ['a:lumMod', 'attrs', 'val'])) / 100000
  if (!isNaN(lumMod)) {
    color = applyLumMod(color, lumMod, isAlpha)
  }
  
  // 应用亮度偏移
  const lumOff = parseInt(getTextByPathList(clrNode, ['a:lumOff', 'attrs', 'val'])) / 100000
  if (!isNaN(lumOff)) {
    color = applyLumOff(color, lumOff, isAlpha)
  }
  
  // 应用饱和度调整
  const satMod = parseInt(getTextByPathList(clrNode, ['a:satMod', 'attrs', 'val'])) / 100000
  if (!isNaN(satMod)) {
    color = applySatMod(color, satMod, isAlpha)
  }
  
  // 应用阴影效果
  const shade = parseInt(getTextByPathList(clrNode, ['a:shade', 'attrs', 'val'])) / 100000
  if (!isNaN(shade)) {
    color = applyShade(color, shade, isAlpha)
  }
  
  // 应用高亮效果
  const tint = parseInt(getTextByPathList(clrNode, ['a:tint', 'attrs', 'val'])) / 100000
  if (!isNaN(tint)) {
    color = applyTint(color, tint, isAlpha)
  }

  // 确保颜色值以#开头
  if (color && color.indexOf('#') === -1) color = '#' + color

  return color
}

export function createGradientText(colors, rot = 90) {
  const gradientStops = colors
    .map((stop) => {
      return `${stop.color} ${stop.pos}`
    })
    .join(', ')

  const gradientStyle = `linear-gradient(${rot}deg, ${gradientStops})`

  return `background: ${gradientStyle};
    -webkit-background-clip: text;
    -webkit-text-fill-color: transparent;
    background-clip: text;
    color: transparent;`
}
export function getGradFill(solidFill, clrMap, phClr, warpObj) {
  const grdFill = solidFill['a:gradFill']
  const gsLst = grdFill['a:gsLst']['a:gs']
  const color_ary = []

  for (let i = 0; i < gsLst.length; i++) {
    const lo_color = getSolidFill(gsLst[i], clrMap, phClr, warpObj)
    const pos = getTextByPathList(gsLst[i], ['attrs', 'pos'])

    color_ary[i] = {
      pos: pos ? pos / 1000 + '%' : '',
      color: lo_color,
    }
  }
  const lin = grdFill['a:lin']
  let rot = 90
  if (lin) {
    rot = angleToDegrees(lin['attrs']['ang'])
    rot = rot + 90
  }
  const colors = color_ary.sort((a, b) => parseInt(a.pos) - parseInt(b.pos))
  return createGradientText(colors, rot)
}