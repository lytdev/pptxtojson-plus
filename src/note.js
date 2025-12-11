import {getFontColor, getFontBold, getFontDecoration, getFontSubscript, getFontItalic} from './fontStyle'
import { getTextByPathList } from './utils'
/**
 * 从备注内容中提取富文本
 * @param {Object} noteContent - 包含笔记内容的XML对象
 * @returns {string} 提取的文本内容
 */
export function getNote(noteContent) {
  // 初始化空字符串以存储提取的文本
  let text = ''
  // 获取所有形状节点
  let spNodes = getTextByPathList(noteContent, ['p:notes', 'p:cSld', 'p:spTree', 'p:sp'])
  // 如果没有找到形状节点，则返回空字符串
  if (!spNodes) return ''
  // 确保spNodes是数组格式
  if (spNodes.constructor !== Array) spNodes = [spNodes]
  // 遍历所有形状节点
  for (const spNode of spNodes) {
    // 获取段落中的文本运行节点
    let apNodes = getTextByPathList(spNode, ['p:txBody', 'a:p'])
    // 如果没有找到文本运行节点，则跳过当前形状节点
    if (!apNodes) continue

    // 确保rNodes是数组格式
    if (apNodes.constructor !== Array) apNodes = [apNodes]
    // 遍历所有文本运行节点
    for (const apNode of apNodes) {
      text += '<p>'
      
      // 提取实际文本内容
      let arNodes = getTextByPathList(apNode, ['a:r'])
      if (arNodes.constructor !== Array) arNodes = [arNodes]
      for (const arNode of arNodes) {
        let style = ''
        let htmlTag = '' 
        // 加粗
        const fontColor = getFontColor(arNode)
        if (fontColor) {
          style += 'color: ' + fontColor + ';'
        }
        // 加粗
        const fontBold = getFontBold(arNode)
        if (fontBold) {
          style += 'font-weight: ' + fontBold + ';'
        }
        // 斜体
        const fontItalic = getFontItalic(arNode)
        if (fontItalic) {
          style += 'font-style: ' + fontItalic + ';'
        }
        // 下划线
        const fontDecoration = getFontDecoration(arNode)
        if (fontDecoration) {
          style += 'text-decoration: ' + fontDecoration + ';'
        }
        // 上标或者下标
        const fontSubscript = getFontSubscript(arNode)
        if (fontSubscript) {
          htmlTag = fontSubscript
        }
        const t = getTextByPathList(arNode, ['a:t'])
        // 如果文本存在且为字符串类型，则将其添加到结果中
        if (t && typeof t === 'string') {
          if (style) {
            if (!htmlTag) {
              htmlTag = 'span'
            }
            const html = `<${htmlTag} style="${style}">${t}</${htmlTag}>`
            text += html
          }
          else if (htmlTag) {
            const html = `<${htmlTag}>${t}</${htmlTag}>`
            text += html
          }
          else {
            text += t
          }
        
        }
      }
      text += '</p>'
     
    }
  }
  // 返回提取的完整文本
  return text
}