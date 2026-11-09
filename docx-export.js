/* Browser-side DOCX export from the existing rendered preview. No API calls. */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory;
  else root.DocxExport = { create: (html) => factory(root.docx, root.DOMParser).create(html) };
})(typeof window !== 'undefined' ? window : globalThis, function (d, Parser) {
  if (!d || !Parser) throw Error('Word 导出依赖未加载');
  const warnings = [];
  const text = (node) => node.textContent || '';
  const originalFormula = (node) => node.querySelector('annotation[encoding="application/x-tex"]')?.textContent || text(node);
  const mathChildren = (node) => Array.from(node.children).filter((child) => child.localName !== 'annotation').flatMap(mathNode);
  function mathNode(node) {
    switch (node.localName) {
      case 'math': case 'semantics': case 'mrow': case 'mstyle': return mathChildren(node);
      case 'mi': case 'mn': case 'mo': case 'mtext': case 'ms': return [new d.MathRun(text(node))];
      case 'mfrac': return [new d.MathFraction({ numerator: mathNode(node.children[0]), denominator: mathNode(node.children[1]) })];
      case 'msqrt': return [new d.MathRadical({ children: mathChildren(node) })];
      case 'mroot': return [new d.MathRadical({ children: mathNode(node.children[0]), degree: mathNode(node.children[1]) })];
      case 'msup': return [new d.MathSuperScript({ children: mathNode(node.children[0]), superScript: mathNode(node.children[1]) })];
      case 'msub': return [new d.MathSubScript({ children: mathNode(node.children[0]), subScript: mathNode(node.children[1]) })];
      case 'msubsup': return [new d.MathSubSuperScript({ children: mathNode(node.children[0]), subScript: mathNode(node.children[1]), superScript: mathNode(node.children[2]) })];
      case 'munder': return [new d.MathSubScript({ children: mathNode(node.children[0]), subScript: mathNode(node.children[1]) })];
      case 'mover': return [new d.MathSuperScript({ children: mathNode(node.children[0]), superScript: mathNode(node.children[1]) })];
      case 'munderover': return [new d.MathSubSuperScript({ children: mathNode(node.children[0]), subScript: mathNode(node.children[1]), superScript: mathNode(node.children[2]) })];
      default: throw Error('不支持的公式结构：' + node.localName);
    }
  }
  function formula(node) {
    try {
      const children = mathNode(node);
      if (!children.length) throw Error('空公式');
      return new d.Math({ children });
    } catch (e) {
      warnings.push('部分公式不支持原生转换，已保留 LaTeX 原文');
      return new d.TextRun({ text: originalFormula(node), color: 'A22A22' });
    }
  }
  function runs(node, style = {}) {
    if (node.nodeType === 3) return text(node) ? [new d.TextRun({ text: text(node), ...style })] : [];
    if (node.nodeType !== 1) return [];
    const tag = node.localName;
    if (tag === 'math') return [formula(node)];
    if (tag === 'br') return [new d.TextRun({ break: 1 })];
    if (tag === 'img') {
      warnings.push('远程图片未嵌入，已保留图片说明及地址');
      return [new d.TextRun({ text: '[图片：' + (node.getAttribute('alt') || '无说明') + '] ' + (node.getAttribute('src') || ''), ...style })];
    }
    if (tag === 'a') {
      const url = node.getAttribute('href') || '';
      if (/^https?:\/\//i.test(url)) return [new d.ExternalHyperlink({ link: url, children: [new d.TextRun({ text: text(node), style: 'Hyperlink' })] })];
    }
    if (tag === 'span' && node.classList.contains('math-raw')) warnings.push('公式渲染器不可用，公式以原文导出');
    const next = { ...style };
    if (tag === 'strong' || tag === 'b') next.bold = true;
    if (tag === 'em' || tag === 'i') next.italics = true;
    if (tag === 'del' || tag === 's') next.strike = true;
    if (tag === 'code') { next.font = 'Consolas'; next.color = '34495E'; }
    return Array.from(node.childNodes).flatMap((child) => runs(child, next));
  }
  function createParagraph(node, options = {}) {
    return new d.Paragraph({ children: runs(node), spacing: { after: 160, line: 360 }, ...options });
  }
  async function create(html) {
    warnings.length = 0;
    const document = new Parser().parseFromString('<!doctype html><html><body>' + html + '</body></html>', 'text/html');
    const content = [], numbering = [];
    const headings = { h1: d.HeadingLevel.HEADING_1, h2: d.HeadingLevel.HEADING_2, h3: d.HeadingLevel.HEADING_3, h4: d.HeadingLevel.HEADING_4, h5: d.HeadingLevel.HEADING_5, h6: d.HeadingLevel.HEADING_6 };
    for (const node of document.body.children) {
      const tag = node.localName;
      if (headings[tag]) { content.push(createParagraph(node, { heading: headings[tag] })); continue; }
      if (tag === 'ul' || tag === 'ol') {
        const reference = 'list-' + numbering.length;
        numbering.push({ reference, levels: [{ level: 0, format: tag === 'ol' ? d.LevelFormat.DECIMAL : d.LevelFormat.BULLET, text: tag === 'ol' ? '%1.' : '•', start: tag === 'ol' ? Math.max(1, Number(node.getAttribute('start')) || 1) : 1, alignment: d.AlignmentType.START, style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] });
        for (const item of node.children) if (item.localName === 'li') content.push(createParagraph(item, { numbering: { reference, level: 0 } }));
        continue;
      }
      if (tag === 'table') {
        const rows = Array.from(node.querySelectorAll('tr'));
        const count = Math.max(1, ...rows.map((r) => r.children.length));
        const width = Math.floor(9000 / count);
        content.push(new d.Table({ columnWidths: Array(count).fill(width), rows: rows.map((r, rowIndex) => new d.TableRow({ children: Array.from({ length: count }, (_, i) => new d.TableCell({ width: { size: width, type: d.WidthType.DXA }, shading: rowIndex === 0 ? { fill: 'EDF2F6', type: d.ShadingType.CLEAR } : undefined, children: [createParagraph(r.children[i] || r, { spacing: { after: 0 } })] })) })) }));
        continue;
      }
      if (tag === 'pre') { content.push(new d.Paragraph({ children: text(node).split('\n').flatMap((line, i) => [new d.TextRun({ text: line, font: 'Consolas', size: 20, ...(i ? { break: 1 } : {}) })]), spacing: { before: 120, after: 160 } })); continue; }
      if (tag === 'blockquote') { content.push(createParagraph(node, { indent: { left: 480 }, border: { left: { color: '8194A7', size: 12, space: 8 } } })); continue; }
      if (tag === 'div' && node.classList.contains('math-display')) { content.push(createParagraph(node, { alignment: d.AlignmentType.CENTER })); continue; }
      content.push(createParagraph(node));
    }
    const doc = new d.Document({ numbering: { config: numbering }, sections: [{ properties: {}, children: content.length ? content : [new d.Paragraph('')] }] });
    return { blob: await d.Packer.toBlob(doc), warnings: [...new Set(warnings)] };
  }
  return { create };
});
