/* Transcript merge engine.
 * Turns a line-per-utterance transcript (DOC_01 style):
 *   01:00:08:00    Transcribed    Contributor 2 - Right, are we rolling?
 * into speaker-block layout (DOC_02 style):
 *   [bold]Contributor 2 [/bold][grey][01:00:08] [/grey]Right, are we rolling?
 * Every source line stays its own paragraph, word for word (default).
 * Timecode is shown as [HH:MM:SS], or full HH:MM:SS:FF with {frames:true}.
 * {merge:true} joins consecutive same-speaker lines, as the reference DOC_02 does.
 * Works on the raw WordprocessingML so every other part of the .docx
 * (styles, page setup, fonts) passes through untouched.
 */
(function (root) {
  // groups: 1 full timecode, 2 HH:MM:SS, 3 speaker, 4 text
  var LINE_RX = /^((\d{1,2}:\d{2}:\d{2})[:;.]\d{2,3})\s+\S+\s+(.+?)\s+-\s?(.*)$/;
  var PARA_RX = /<w:p(?:\s[^>]*)?\/>|<w:p(?:\s[^>]*)?>[\s\S]*?<\/w:p>/g;
  var TOKEN_RX = /<w:t(?:\s[^>]*)?>([^<]*)<\/w:t>|<w:tab\/>|<w:br\/>|<w:t(?:\s[^>]*)?\/>/g;
  var DEFAULT_FONTS = '<w:rFonts w:ascii="Helvetica" w:cs="Helvetica" w:eastAsia="Helvetica" w:hAnsi="Helvetica"/>';

  function decode(s) {
    return s.replace(/&(amp|lt|gt|quot|apos|#\d+|#x[0-9a-f]+);/gi, function (m, e) {
      var k = e.toLowerCase();
      if (k === 'amp') return '&'; if (k === 'lt') return '<'; if (k === 'gt') return '>';
      if (k === 'quot') return '"'; if (k === 'apos') return "'";
      return String.fromCodePoint(k[1] === 'x' ? parseInt(k.slice(2), 16) : parseInt(k.slice(1), 10));
    });
  }
  function encode(s) {
    return s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;').replace(/'/g, '&apos;');
  }
  function paraText(xml) {
    var out = '', m;
    TOKEN_RX.lastIndex = 0;
    while ((m = TOKEN_RX.exec(xml))) {
      if (m[0] === '<w:tab/>') out += '\t';
      else if (m[0] === '<w:br/>') out += '\n';
      else if (m[1] !== undefined) out += decode(m[1]);
    }
    return out;
  }
  function run(fonts, extra, text, br) {
    return '<w:r><w:rPr>' + extra + fonts + '</w:rPr>' + (br ? '<w:br/>' : '') +
      '<w:t xml:space="preserve">' + encode(text) + '</w:t></w:r>';
  }
  // Joins a speaker's lines. Lines without closing punctuation get a full stop,
  // which is what DOC_02 does (it even turns "there," into "there,.").
  function joinLines(lines, tidy) {
    return lines.map(function (s, i) {
      s = s.trim();
      if (i < lines.length - 1 && !/[.?!]$/.test(s)) {
        if (tidy && /[,;:—-]$/.test(s)) s = s.replace(/[,;:—-]+$/, '');
        s += '.';
      }
      return s;
    }).join(' ');
  }

  /**
   * @param {string} xml  word/document.xml
   * @param {{merge?: boolean, frames?: boolean, tidyJoins?: boolean}} opts
   * @returns {{xml: string, stats: {lines:number, blocks:number, speakers:string[], passthrough:number}, preview: Array}}
   */
  function convertDocumentXml(xml, opts) {
    opts = opts || {};
    var bodyStart = xml.indexOf('<w:body>');
    var sect = xml.lastIndexOf('<w:sectPr');
    var bodyEnd = xml.lastIndexOf('</w:body>');
    if (bodyStart < 0 || bodyEnd < 0) throw new Error('This file has no document body. Is it a Word .docx?');
    bodyStart += '<w:body>'.length;
    var contentEnd = sect > bodyStart ? sect : bodyEnd;
    var body = xml.slice(bodyStart, contentEnd);

    // Tokenise body into paragraphs and anything in between (kept verbatim).
    var items = [], last = 0, m;
    PARA_RX.lastIndex = 0;
    while ((m = PARA_RX.exec(body))) {
      if (m.index > last) items.push({ raw: body.slice(last, m.index), para: false });
      var t = paraText(m[0]);
      items.push({ raw: m[0], para: true, text: t, line: LINE_RX.exec(t.trim()) });
      last = m.index + m[0].length;
    }
    if (last < body.length) items.push({ raw: body.slice(last), para: false });

    var first = -1;
    for (var i = 0; i < items.length; i++) if (items[i].para && items[i].line) { first = i; break; }
    if (first < 0) throw new Error('No transcript lines found. Expected lines like "01:00:07:05    Transcribed    Contributor 1 - text".');

    // Header: everything before the first transcript line, minus trailing blank paragraphs.
    var head = items.slice(0, first);
    while (head.length && head[head.length - 1].para && !head[head.length - 1].text.trim()) head.pop();

    var fontsMatch = /<w:rFonts[^>]*\/>/.exec(items[first].raw);
    var fonts = fontsMatch ? fontsMatch[0] : DEFAULT_FONTS;

    var out = head.map(function (h) { return h.raw; });
    var blocks = [], cur = null, lines = 0, passthrough = 0, speakers = {};
    function flush() {
      if (!cur) return;
      var text = joinLines(cur.lines, opts.tidyJoins);
      out.push('<w:p>' +
        run(fonts, '<w:b/><w:bCs/>', cur.speaker + ' ', true) +
        run(fonts, '<w:color w:val="777777"/>', '[' + cur.tc + '] ', false) +
        run(fonts, '', text, false) + '</w:p>');
      blocks.push({ speaker: cur.speaker, tc: cur.tc, text: text, count: cur.lines.length });
      cur = null;
    }
    for (var j = first; j < items.length; j++) {
      var it = items[j];
      if (!it.para) { flush(); out.push(it.raw); continue; }
      if (it.line) {
        lines++;
        var tc = opts.frames ? it.line[1] : it.line[2];
        if (/^\d:/.test(tc)) tc = '0' + tc;
        var spk = it.line[3].trim();
        speakers[spk] = (speakers[spk] || 0) + 1;
        if (opts.merge && cur && cur.speaker === spk) cur.lines.push(it.line[4]);
        else { flush(); cur = { speaker: spk, tc: tc, lines: [it.line[4]] }; }
      } else if (it.text.trim()) {
        flush(); out.push(it.raw); passthrough++;
      }
      // blank paragraphs inside the transcript are dropped, as in DOC_02
    }
    flush();

    return {
      xml: xml.slice(0, bodyStart) + out.join('') + xml.slice(contentEnd),
      stats: { lines: lines, blocks: blocks.length, speakers: Object.keys(speakers).sort(), passthrough: passthrough },
      blocks: blocks,
      header: head.filter(function (h) { return h.para; }).map(function (h) { return h.text; })
    };
  }

  /** Converts a .docx (ArrayBuffer/Blob) with JSZip. Returns {blob, stats, blocks, header}. */
  function convertDocx(JSZip, data, opts) {
    return JSZip.loadAsync(data).then(function (zip) {
      var f = zip.file('word/document.xml');
      if (!f) throw new Error('This file is not a Word .docx (no word/document.xml inside).');
      return f.async('string').then(function (xml) {
        var r = convertDocumentXml(xml, opts);
        zip.file('word/document.xml', r.xml);
        var type = (typeof Blob !== 'undefined' && !(opts && opts.nodebuffer)) ? 'blob' : 'nodebuffer';
        return zip.generateAsync({
          type: type, compression: 'DEFLATE',
          mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
        }).then(function (out) { r.file = out; delete r.xml; return r; });
      });
    });
  }

  function outputName(name) {
    var base = name.replace(/\.docx$/i, '');
    return base + '_FORMATTED.docx';
  }

  var api = { convertDocumentXml: convertDocumentXml, convertDocx: convertDocx, outputName: outputName, joinLines: joinLines };
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  else root.TranscriptMerge = api;
})(this);
