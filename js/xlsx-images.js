/*
 * Payslip Generator - read pictures from an .xlsx file and embed pictures into one.
 * SheetJS (community) ignores images, so this works directly on the zip package with JSZip.
 * Pure functions: usable from the browser and from Node tests.
 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) {
    module.exports = factory();
  } else {
    root.XlsxImages = factory();
  }
})(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  var NS_REL_IMAGE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image';
  var NS_REL_DRAWING = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing';
  var NS_REL_WORKSHEET = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet';

  var MIME = { png: 'image/png', jpg: 'image/jpeg', jpeg: 'image/jpeg', gif: 'image/gif', bmp: 'image/bmp', emf: 'image/x-emf', wmf: 'image/x-wmf', tif: 'image/tiff', tiff: 'image/tiff', svg: 'image/svg+xml' };

  function decodeXmlEntities(s) {
    return String(s).replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&amp;/g, '&');
  }

  function escapeXml(s) {
    return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
  }

  // Tiny attribute reader: returns map of attributes for every <tag ...> occurrence.
  function findTags(xml, tagName) {
    var re = new RegExp('<(?:[\\w.-]+:)?' + tagName + '\\b([^>]*?)/?>', 'g');
    var out = [];
    var m;
    while ((m = re.exec(xml))) {
      var attrs = {};
      var ar = /([\w:.-]+)\s*=\s*"([^"]*)"/g;
      var a;
      while ((a = ar.exec(m[1]))) attrs[a[1]] = decodeXmlEntities(a[2]);
      out.push({ attrs: attrs, index: m.index });
    }
    return out;
  }

  function dirname(p) {
    var i = p.lastIndexOf('/');
    return i < 0 ? '' : p.slice(0, i);
  }

  function resolvePath(baseDir, target) {
    if (target.charAt(0) === '/') return target.slice(1);
    var parts = (baseDir ? baseDir.split('/') : []);
    target.split('/').forEach(function (seg) {
      if (seg === '..') parts.pop();
      else if (seg !== '.' && seg !== '') parts.push(seg);
    });
    return parts.join('/');
  }

  function relsPathFor(partPath) {
    return dirname(partPath) + (dirname(partPath) ? '/' : '') + '_rels/' + partPath.slice(partPath.lastIndexOf('/') + 1) + '.rels';
  }

  function readText(zip, path) {
    var f = zip.file(path);
    return f ? f.async('string') : Promise.resolve(null);
  }

  function readRels(zip, partPath) {
    return readText(zip, relsPathFor(partPath)).then(function (xml) {
      var map = {};
      if (!xml) return map;
      findTags(xml, 'Relationship').forEach(function (t) {
        map[t.attrs.Id] = { type: t.attrs.Type, target: t.attrs.Target, mode: t.attrs.TargetMode };
      });
      return map;
    });
  }

  function imageDimensions(bytes, mime) {
    try {
      if (mime === 'image/png' && bytes.length > 24) {
        var dv = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
        return { width: dv.getUint32(16), height: dv.getUint32(20) };
      }
      if (mime === 'image/jpeg') {
        var i = 2;
        while (i < bytes.length) {
          if (bytes[i] !== 0xff) { i++; continue; }
          var marker = bytes[i + 1];
          if (marker >= 0xc0 && marker <= 0xcf && marker !== 0xc4 && marker !== 0xc8 && marker !== 0xcc) {
            return { height: (bytes[i + 5] << 8) | bytes[i + 6], width: (bytes[i + 7] << 8) | bytes[i + 8] };
          }
          var len = (bytes[i + 2] << 8) | bytes[i + 3];
          i += 2 + len;
        }
      }
      if (mime === 'image/gif' && bytes.length > 10) {
        return { width: bytes[6] | (bytes[7] << 8), height: bytes[8] | (bytes[9] << 8) };
      }
    } catch (e) { /* fall through */ }
    return { width: 0, height: 0 };
  }

  // Lists worksheets as [{ name, path }] in workbook order.
  function listSheets(zip) {
    return Promise.all([readText(zip, 'xl/workbook.xml'), readRels(zip, 'xl/workbook.xml')]).then(function (r) {
      var wbXml = r[0], rels = r[1];
      if (!wbXml) return [];
      return findTags(wbXml, 'sheet').map(function (t) {
        var rid = t.attrs['r:id'] || t.attrs.id;
        var rel = rels[rid];
        var target = rel ? rel.target : null;
        return { name: t.attrs.name, path: target ? resolvePath('xl', target) : null };
      }).filter(function (s) { return s.path; });
    });
  }

  /**
   * Extract pictures from an .xlsx ArrayBuffer/Uint8Array.
   * Resolves to [{ sheet, order, row, col, path, mime, bytes(Uint8Array), width, height }]
   * sorted by sheet order then by anchor position (top-left first).
   */
  function extractImages(JSZip, data) {
    var zip;
    return JSZip.loadAsync(data).then(function (z) {
      zip = z;
      return listSheets(zip);
    }).then(function (sheets) {
      var tasks = sheets.map(function (sheet, sheetIndex) {
        return readRels(zip, sheet.path).then(function (rels) {
          var drawings = Object.keys(rels).filter(function (k) { return rels[k].type === NS_REL_DRAWING; }).map(function (k) {
            return resolvePath(dirname(sheet.path), rels[k].target);
          });
          return Promise.all(drawings.map(function (drawingPath) {
            return Promise.all([readText(zip, drawingPath), readRels(zip, drawingPath)]).then(function (r) {
              var xml = r[0], drels = r[1];
              if (!xml) return [];
              // every anchor: locate its blip and its "from" position
              var anchorRe = /<xdr:(oneCellAnchor|twoCellAnchor|absoluteAnchor)\b[\s\S]*?<\/xdr:\1>/g;
              var found = [];
              var m;
              while ((m = anchorRe.exec(xml))) {
                var block = m[0];
                var blip = /<a:blip\b[^>]*\br:(?:embed|link)="([^"]+)"/.exec(block);
                if (!blip) continue;
                var rel = drels[blip[1]];
                if (!rel || rel.mode === 'External') continue;
                var rowM = /<xdr:from>[\s\S]*?<xdr:row>(\d+)<\/xdr:row>/.exec(block);
                var colM = /<xdr:from>[\s\S]*?<xdr:col>(\d+)<\/xdr:col>/.exec(block);
                var nameM = /<xdr:cNvPr\b[^>]*\bname="([^"]*)"/.exec(block);
                found.push({
                  sheet: sheet.name,
                  sheetIndex: sheetIndex,
                  row: rowM ? +rowM[1] : 0,
                  col: colM ? +colM[1] : 0,
                  name: nameM ? decodeXmlEntities(nameM[1]) : '',
                  path: resolvePath(dirname(drawingPath), rel.target)
                });
              }
              return found;
            });
          }));
        });
      });
      return Promise.all(tasks);
    }).then(function (perSheet) {
      var flat = [];
      perSheet.forEach(function (byDrawing) {
        byDrawing.forEach(function (list) { flat = flat.concat(list); });
      });
      flat.sort(function (a, b) { return a.sheetIndex - b.sheetIndex || a.row - b.row || a.col - b.col; });
      return Promise.all(flat.map(function (img, i) {
        var f = zip.file(img.path);
        if (!f) return null;
        return f.async('uint8array').then(function (bytes) {
          var ext = img.path.slice(img.path.lastIndexOf('.') + 1).toLowerCase();
          var mime = MIME[ext] || 'application/octet-stream';
          var dim = imageDimensions(bytes, mime);
          return { sheet: img.sheet, order: i, row: img.row, col: img.col, name: img.name, path: img.path, mime: mime, bytes: bytes, width: dim.width, height: dim.height };
        });
      }));
    }).then(function (list) {
      return list.filter(Boolean);
    });
  }

  function bytesToBase64(bytes) {
    if (typeof Buffer !== 'undefined' && Buffer.from) return Buffer.from(bytes).toString('base64');
    var s = '';
    for (var i = 0; i < bytes.length; i += 0x8000) s += String.fromCharCode.apply(null, bytes.subarray(i, i + 0x8000));
    return btoa(s);
  }

  function dataUrlToBytes(dataUrl) {
    var m = /^data:([^;,]+)?(;base64)?,(.*)$/.exec(dataUrl);
    if (!m) throw new Error('Not a data URL');
    var mime = m[1] || 'application/octet-stream';
    var b64 = m[3];
    var bin;
    if (typeof Buffer !== 'undefined' && Buffer.from) bin = new Uint8Array(Buffer.from(b64, 'base64'));
    else {
      var raw = atob(b64);
      bin = new Uint8Array(raw.length);
      for (var i = 0; i < raw.length; i++) bin[i] = raw.charCodeAt(i);
    }
    return { mime: mime, bytes: bin };
  }

  /**
   * Embed pictures in a sheet of an existing .xlsx (ArrayBuffer/Uint8Array produced by SheetJS).
   * images: [{ dataUrl | bytes+mime, width, height, name, row, col, maxPx }]
   * Resolves to a Uint8Array of the new file.
   */
  function embedImages(JSZip, data, sheetName, images) {
    if (!images || !images.length) return Promise.resolve(data instanceof Uint8Array ? data : new Uint8Array(data));
    var zip;
    var sheetPath;
    return JSZip.loadAsync(data).then(function (z) {
      zip = z;
      return listSheets(zip);
    }).then(function (sheets) {
      var sheet = sheets.filter(function (s) { return s.name === sheetName; })[0] || sheets[0];
      if (!sheet) throw new Error('Workbook has no sheets');
      sheetPath = sheet.path;
      return Promise.all([readText(zip, sheetPath), readText(zip, relsPathFor(sheetPath)), readText(zip, '[Content_Types].xml')]);
    }).then(function (r) {
      var sheetXml = r[0], sheetRels = r[1], contentTypes = r[2];

      // Find a free drawing number
      var n = 1;
      while (zip.file('xl/drawings/drawing' + n + '.xml')) n++;
      var drawingPath = 'xl/drawings/drawing' + n + '.xml';

      // Media files + drawing rels
      var drawingRels = '';
      var anchors = '';
      var mediaIdx = 1;
      while (zip.file('xl/media/image' + mediaIdx + '.png') || zip.file('xl/media/image' + mediaIdx + '.jpeg')) mediaIdx++;
      var usedExt = {};
      images.forEach(function (img, i) {
        var bytes, mime;
        if (img.dataUrl) {
          var d = dataUrlToBytes(img.dataUrl);
          bytes = d.bytes;
          mime = d.mime;
        } else {
          bytes = img.bytes;
          mime = img.mime;
        }
        var ext = mime === 'image/jpeg' ? 'jpeg' : 'png';
        usedExt[ext] = true;
        var mediaName = 'image' + (mediaIdx + i) + '.' + ext;
        zip.file('xl/media/' + mediaName, bytes);
        var rId = 'rId' + (i + 1);
        drawingRels += '<Relationship Id="' + rId + '" Type="' + NS_REL_IMAGE + '" Target="../media/' + mediaName + '"/>';
        var w = img.width || 300, h = img.height || 100;
        var maxPx = img.maxPx || 220;
        var scale = Math.min(1, maxPx / Math.max(w, h));
        var cx = Math.round(w * scale * 9525), cy = Math.round(h * scale * 9525);
        anchors += '<xdr:oneCellAnchor>' +
          '<xdr:from><xdr:col>' + (img.col || 0) + '</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>' + (img.row || 0) + '</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>' +
          '<xdr:ext cx="' + cx + '" cy="' + cy + '"/>' +
          '<xdr:pic><xdr:nvPicPr><xdr:cNvPr id="' + (i + 2) + '" name="' + escapeXml(img.name || ('Picture ' + (i + 1))) + '"/><xdr:cNvPicPr><a:picLocks noChangeAspect="1"/></xdr:cNvPicPr></xdr:nvPicPr>' +
          '<xdr:blipFill><a:blip r:embed="' + rId + '"/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill>' +
          '<xdr:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="' + cx + '" cy="' + cy + '"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr></xdr:pic>' +
          '<xdr:clientData/></xdr:oneCellAnchor>';
      });

      zip.file(drawingPath, '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
        '<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
        anchors + '</xdr:wsDr>');
      zip.file(relsPathFor(drawingPath), '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' + drawingRels + '</Relationships>');

      // Sheet rels: add drawing relationship
      var relId = 'rId1';
      if (sheetRels) {
        var ids = findTags(sheetRels, 'Relationship').map(function (t) { return parseInt(String(t.attrs.Id).replace(/\D/g, ''), 10) || 0; });
        relId = 'rId' + (Math.max.apply(null, ids.concat([0])) + 1);
        sheetRels = sheetRels.replace('</Relationships>', '<Relationship Id="' + relId + '" Type="' + NS_REL_DRAWING + '" Target="../drawings/drawing' + n + '.xml"/></Relationships>');
      } else {
        sheetRels = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
          '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
          '<Relationship Id="' + relId + '" Type="' + NS_REL_DRAWING + '" Target="../drawings/drawing' + n + '.xml"/></Relationships>';
      }
      zip.file(relsPathFor(sheetPath), sheetRels);

      // Sheet xml: add <drawing r:id=.../> at the right position
      if (sheetXml.indexOf('xmlns:r=') < 0) {
        sheetXml = sheetXml.replace('<worksheet ', '<worksheet xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ');
      }
      var drawingTag = '<drawing r:id="' + relId + '"/>';
      var inserted = false;
      ['<legacyDrawing', '<legacyDrawingHF', '<picture', '<oleObjects', '<controls', '<webPublishItems', '<tableParts', '<extLst'].forEach(function (tag) {
        if (!inserted && sheetXml.indexOf(tag) >= 0) {
          sheetXml = sheetXml.replace(tag, drawingTag + tag);
          inserted = true;
        }
      });
      if (!inserted) sheetXml = sheetXml.replace('</worksheet>', drawingTag + '</worksheet>');
      zip.file(sheetPath, sheetXml);

      // Content types
      var ct = contentTypes;
      Object.keys(usedExt).forEach(function (ext) {
        if (ct.indexOf('Extension="' + ext + '"') < 0) {
          ct = ct.replace('</Types>', '<Default Extension="' + ext + '" ContentType="' + (ext === 'jpeg' ? 'image/jpeg' : 'image/png') + '"/></Types>');
        }
      });
      ct = ct.replace('</Types>', '<Override PartName="/' + drawingPath + '" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/></Types>');
      zip.file('[Content_Types].xml', ct);

      return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
    });
  }

  function toDataUrl(img) {
    return 'data:' + img.mime + ';base64,' + bytesToBase64(img.bytes);
  }

  return {
    extractImages: extractImages,
    embedImages: embedImages,
    toDataUrl: toDataUrl,
    dataUrlToBytes: dataUrlToBytes,
    imageDimensions: imageDimensions
  };
});
