import { getTextByPathList } from './utils'
import { getShadow } from './shadow'
import { getFillType, getSolidFill } from './fill'
/**
 * 获取节点的字体类型
 * 
 * 该函数首先尝试从节点本身获取字体类型，如果未找到，
 * 则根据节点类型从主题方案中获取相应的字体：
 * - 标题类型（title/subTitle/ctrTitle）使用主字体（majorFont）
 * - 正文类型（body）使用次字体（minorFont）
 * - 其他类型默认使用次字体（minorFont）
 * 
 * @param {Object} node - 文本节点对象
 * @param {string} type - 节点类型，如 'title', 'subTitle', 'ctrTitle', 'body' 等
 * @param {Object} warpObj - 包含主题内容的包装对象
 * @returns {string} 字体类型名称，如果未找到则返回空字符串
 */
export function getFontType(node, type, warpObj) {
  let typeface = getTextByPathList(node, ['a:rPr', 'a:latin', 'attrs', 'typeface'])

  // 如果节点中没有直接定义字体，则从主题方案中获取
  if (!typeface) {
    const fontSchemeNode = getTextByPathList(warpObj['themeContent'], ['a:theme', 'a:themeElements', 'a:fontScheme'])

    if (type === 'title' || type === 'subTitle' || type === 'ctrTitle') {
      // 标题类使用主字体
      typeface = getTextByPathList(fontSchemeNode, ['a:majorFont', 'a:latin', 'attrs', 'typeface'])
    } 
    else if (type === 'body') {
      // 正文使用次字体
      typeface = getTextByPathList(fontSchemeNode, ['a:minorFont', 'a:latin', 'attrs', 'typeface'])
    } 
    else {
      // 其他类型默认使用次字体
      typeface = getTextByPathList(fontSchemeNode, ['a:minorFont', 'a:latin', 'attrs', 'typeface'])
    }
  }

  return typeface || ''
}

/**
 * 获取文本的字体颜色
 * 
 * 按照以下优先级顺序查找字体颜色：
 * 1. 直接在文本节点中定义的颜色
 * 2. 列表样式中对应级别的默认颜色
 * 3. 父节点的字体引用颜色或传入的默认字体样式颜色
 * 
 * @param {Object} node - 当前文本节点
 * @param {Object} pNode - 父节点（段落节点）
 * @param {Object} lstStyle - 列表样式对象
 * @param {Object} pFontStyle - 默认字体样式
 * @param {number} lvl - 列表级别
 * @param {Object} warpObj - 包装对象，用于传递上下文数据
 * @returns {string} 字体颜色值，如果未找到则返回空字符串
 */
export function getFontColor(node, pNode, lstStyle, pFontStyle, lvl, warpObj) {
  const rPrNode = getTextByPathList(node, ['a:rPr'])
  let filTyp, color
  // 尝试从当前文本节点获取颜色
  if (rPrNode) {
    filTyp = getFillType(rPrNode)
    if (filTyp === 'SOLID_FILL') {
      const solidFillNode = rPrNode['a:solidFill']
      color = getSolidFill(solidFillNode, undefined, undefined, warpObj)
    }
  }
  // 如果当前节点没有颜色，则尝试从列表样式中获取
  if (!color && getTextByPathList(lstStyle, ['a:lvl' + lvl + 'pPr', 'a:defRPr'])) {
    const lstStyledefRPr = getTextByPathList(lstStyle, ['a:lvl' + lvl + 'pPr', 'a:defRPr'])
    filTyp = getFillType(lstStyledefRPr)
    if (filTyp === 'SOLID_FILL') {
      const solidFillNode = lstStyledefRPr['a:solidFill']
      color = getSolidFill(solidFillNode, undefined, undefined, warpObj)
    }
  }
  // 如果仍未获取到颜色，则尝试从父节点样式或默认字体样式中获取
  if (!color) {
    const sPstyle = getTextByPathList(pNode, ['p:style', 'a:fontRef'])
    if (sPstyle) color = getSolidFill(sPstyle, undefined, undefined, warpObj)
    if (!color && pFontStyle) color = getSolidFill(pFontStyle, undefined, undefined, warpObj)
  }
  return color || ''
}
/**
 * 获取字体大小
 * 
 * 按照 PowerPoint 的层级结构逐层查找字体大小设置：
 * 1. 首先检查直接应用于文本的字体大小
 * 2. 检查段落结束标记中的字体大小
 * 3. 检查文本框列表样式中的字体大小
 * 4. 检查幻灯片版式中的默认字体大小
 * 5. 检查幻灯片母版中的字体大小
 * 6. 根据元素类型应用默认字体大小
 * 
 * @param {Object} node - 当前文本运行节点（包含具体文本内容的节点）
 * @param {Object} slideLayoutSpNode - 幻灯片版式中的形状节点
 * @param {string} type - 文本框类型（如 'title', 'subTitle', 'body', 'dt', 'sldNum' 等）
 * @param {Object} slideMasterTextStyles - 幻灯片母版文本样式
 * @param {Object} textBodyNode - 文本主体节点
 * @param {Object} pNode - 段落节点
 * @returns {string} 字体大小，格式为 'Xpt'
 */
export function getFontSize(node, slideLayoutSpNode, type, slideMasterTextStyles, textBodyNode, pNode) {
  let fontSize
  // 字体数字应该除以100
  if (getTextByPathList(node, ['a:rPr', 'attrs', 'sz'])) fontSize = getTextByPathList(node, ['a:rPr', 'attrs', 'sz']) / 100

  // 如果当前节点没有设置字体大小，尝试从段落的结束标记中获取
  if ((isNaN(fontSize) || !fontSize) && pNode) {
    if (getTextByPathList(pNode, ['a:endParaRPr', 'attrs', 'sz'])) {
      fontSize = getTextByPathList(pNode, ['a:endParaRPr', 'attrs', 'sz']) / 100
    }
  }

  // 如果仍未找到字体大小，尝试从文本主体的列表样式中获取
  if ((isNaN(fontSize) || !fontSize) && textBodyNode) {
    const lstStyle = getTextByPathList(textBodyNode, ['a:lstStyle'])
    if (lstStyle) {
      let lvl = 1
      if (pNode) {
        const lvlNode = getTextByPathList(pNode, ['a:pPr', 'attrs', 'lvl'])
        if (lvlNode !== undefined) lvl = parseInt(lvlNode) + 1
      }

      const sz = getTextByPathList(lstStyle, [`a:lvl${lvl}pPr`, 'a:defRPr', 'attrs', 'sz'])
      if (sz) fontSize = parseInt(sz) / 100
    }
  }

  // 如果仍未找到字体大小，使用幻灯片版式中第一级别的默认字体大小
  if ((isNaN(fontSize) || !fontSize)) {
    const sz = getTextByPathList(slideLayoutSpNode, ['p:txBody', 'a:lstStyle', 'a:lvl1pPr', 'a:defRPr', 'attrs', 'sz'])
    if (sz) fontSize = parseInt(sz) / 100
  }

  // 如果仍未找到字体大小，根据段落级别从幻灯片版式中获取
  if ((isNaN(fontSize) || !fontSize) && slideLayoutSpNode) {
    let lvl = 1
    if (pNode) {
      const lvlNode = getTextByPathList(pNode, ['a:pPr', 'attrs', 'lvl'])
      if (lvlNode !== undefined) lvl = parseInt(lvlNode) + 1
    }
    const layoutSz = getTextByPathList(slideLayoutSpNode, ['p:txBody', 'a:lstStyle', `a:lvl${lvl}pPr`, 'a:defRPr', 'attrs', 'sz'])
    if (layoutSz) fontSize = parseInt(layoutSz) / 100
  }

  // 如果仍未找到字体大小，尝试从段落属性中获取
  if ((isNaN(fontSize) || !fontSize) && pNode) {
    const paraSz = getTextByPathList(pNode, ['a:pPr', 'a:defRPr', 'attrs', 'sz'])
    if (paraSz) fontSize = parseInt(paraSz) / 100
  }

  // 如果仍未找到字体大小，根据元素类型从幻灯片母版中获取默认值
  if (isNaN(fontSize) || !fontSize) {
    let sz
    if (type === 'title' || type === 'subTitle' || type === 'ctrTitle') {
      sz = getTextByPathList(slideMasterTextStyles, ['p:titleStyle', 'a:lvl1pPr', 'a:defRPr', 'attrs', 'sz'])
    } 
    else if (type === 'body') {
      sz = getTextByPathList(slideMasterTextStyles, ['p:bodyStyle', 'a:lvl1pPr', 'a:defRPr', 'attrs', 'sz'])
    } 
    else if (type === 'dt' || type === 'sldNum') {
      sz = '1200'
    } 
    else if (!type) {
      sz = getTextByPathList(slideMasterTextStyles, ['p:otherStyle', 'a:lvl1pPr', 'a:defRPr', 'attrs', 'sz'])
    }
    if (sz) fontSize = parseInt(sz) / 100
  }

  // 处理基线调整对字体大小的影响
  const baseline = getTextByPathList(node, ['a:rPr', 'attrs', 'baseline'])
  if (baseline && !isNaN(fontSize)) fontSize -= 10

  // 设置默认字体大小为18pt
  fontSize = (isNaN(fontSize) || !fontSize) ? 18 : fontSize

  return fontSize + 'pt'
}

export function getFontBold(node) {
  return getTextByPathList(node, ['a:rPr', 'attrs', 'b']) === '1' ? 'bold' : ''
}

export function getFontItalic(node) {
  return getTextByPathList(node, ['a:rPr', 'attrs', 'i']) === '1' ? 'italic' : ''
}

export function getFontDecoration(node) {
  return getTextByPathList(node, ['a:rPr', 'attrs', 'u']) === 'sng' ? 'underline' : ''
}

export function getFontDecorationLine(node) {
  return getTextByPathList(node, ['a:rPr', 'attrs', 'strike']) === 'sngStrike' ? 'line-through' : ''
}

export function getFontSpace(node) {
  const spc = getTextByPathList(node, ['a:rPr', 'attrs', 'spc'])
  return spc ? (parseInt(spc) / 100 + 'pt') : ''
}

export function getFontSubscript(node) {
  const baseline = getTextByPathList(node, ['a:rPr', 'attrs', 'baseline'])
  if (!baseline) return ''
  return parseInt(baseline) > 0 ? 'super' : 'sub'
}

export function getFontShadow(node, warpObj) {
  const txtShadow = getTextByPathList(node, ['a:rPr', 'a:effectLst', 'a:outerShdw'])
  if (txtShadow) {
    const shadow = getShadow(txtShadow, warpObj)
    if (shadow) {
      const { h, v, blur, color } = shadow
      if (!isNaN(v) && !isNaN(h)) {
        return h + 'pt ' + v + 'pt ' + (blur ? blur + 'pt' : '') + ' ' + color
      }
    }
  }
  return ''
}