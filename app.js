/* Universal file converter — 100% client-side.
   Supports: xlsx/xls/csv/pdf/png/jpg/webp/gif/bmp/txt
   Outputs:  pdf/png/jpg/xlsx/csv/txt (depending on input)
*/
(function () {
  "use strict";
  var $ = function (id) { return document.getElementById(id); };
  var files = []; // {file, kind, ext}
  var lastWorkbooks = {}; // fileName -> workbook (cached for sheet list)

  function extOf(name) {
    var p = (name || "").split(".");
    return p.length > 1 ? p.pop().toLowerCase() : "";
  }
  function detectKind(file) {
    var e = extOf(file.name);
    var t = (file.type || "").toLowerCase();
    if (e === "xlsx" || e === "xls") return "spreadsheet";
    if (e === "csv") return "csv";
    if (e === "pdf") return "pdf";
    if (e === "txt" || e === "tsv") return "text";
    if (["png", "jpg", "jpeg", "webp", "gif", "bmp"].indexOf(e) >= 0) return "image";
    if (t.indexOf("image/") === 0) return "image";
    if (t === "application/pdf") return "pdf";
    return "unknown";
  }
  function kindLabel(k) {
    return { spreadsheet: "xlsx", csv: "csv", pdf: "pdf", image: "image", text: "text", unknown: "file" }[k] || k;
  }
  function optionsForKind(k) {
    if (k === "spreadsheet" || k === "csv") return ["pdf", "png", "jpg", "xlsx", "csv", "txt"];
    if (k === "pdf") return ["png", "jpg", "xlsx", "csv", "txt"];
    if (k === "image") return ["pdf", "png", "jpg", "xlsx", "csv", "txt"];
    if (k === "text") return ["pdf", "png", "jpg", "xlsx", "csv", "txt"];
    return ["pdf", "png", "jpg", "txt"];
  }
  function hintFor(out) {
    var m = {
      pdf: "Best for sharing and printing. Tables keep their layout.",
      png: "High-quality image. Multi-page PDFs become a zip of images.",
      jpg: "Smaller image file. Multi-page PDFs become a zip of images.",
      xlsx: "Spreadsheet workbook. PDFs and images use extracted text (enable OCR for photos).",
      csv: "Plain comma-separated values, opens in Excel.",
      txt: "Plain text, one row per line."
    };
    return m[out] || "";
  }

  function refreshUI() {
    var list = $("fileList");
    list.innerHTML = "";
    files.forEach(function (f, i) {
      var row = document.createElement("div");
      row.className = "fileRow";
      var left = document.createElement("span");
      left.textContent = f.file.name + " (" + kindLabel(f.kind) + ", " + fmtBytes(f.file.size) + ")";
      var rm = document.createElement("button");
      rm.textContent = "remove";
      rm.onclick = function () { files.splice(i, 1); refreshUI(); };
      row.appendChild(left);
      row.appendChild(rm);
      list.appendChild(row);
    });
    var sel = $("outputSelect");
    var prev = sel.value;
    sel.innerHTML = "";
    var opts;
    if (!files.length) {
      opts = ["pdf", "png", "jpg", "xlsx", "csv", "txt"];
    } else {
      opts = optionsForKind(files[0].kind);
      for (var i = 1; i < files.length; i++) {
        var o2 = optionsForKind(files[i].kind);
        opts = opts.filter(function (o) { return o2.indexOf(o) >= 0; });
      }
      if (!opts.length) opts = ["pdf", "png", "jpg", "txt"];
    }
    opts.forEach(function (o) {
      var el = document.createElement("option");
      el.value = o;
      el.textContent = o.toUpperCase();
      sel.appendChild(el);
    });
    if (opts.indexOf(prev) >= 0) sel.value = prev;
    $("outputHint").textContent = hintFor(sel.value);
    $("convertBtn").disabled = !files.length;
    refreshSheetSelect();
  }

  function fmtBytes(n) {
    if (n < 1024) return n + " B";
    if (n < 1048576) return (n / 1024).toFixed(1) + " KB";
    return (n / 1048576).toFixed(1) + " MB";
  }

  function refreshSheetSelect() {
    var s = $("sheetSelect");
    var cur = s.value;
    s.innerHTML = "";
    var o = document.createElement("option");
    o.value = "";
    o.textContent = "auto — first sheet";
    s.appendChild(o);
    Object.keys(lastWorkbooks).forEach(function (key) {
      try {
        lastWorkbooks[key].SheetNames.forEach(function (n) {
          var el = document.createElement("option");
          el.value = n;
          el.textContent = key + " : " + n;
          s.appendChild(el);
        });
      } catch (e) {}
    });
    s.value = cur || "";
  }

  async function cacheWorkbook(f) {
    if (f.kind !== "spreadsheet" && f.kind !== "csv") return;
    if (lastWorkbooks[f.file.name]) return;
    try {
      var ab = await f.file.arrayBuffer();
      var wb = XLSX.read(ab, { type: "array" });
      lastWorkbooks[f.file.name] = wb;
    } catch (e) {}
  }

  function setStatus(html, spinning) {
    $("statusEl").innerHTML = (spinning ? '<span class="spinner"></span>' : "") + (html || "");
  }
  function clearResults() {
    $("resultsEl").innerHTML = "";
    $("previewEl").innerHTML = "";
  }
  function addResult(blob, filename) {
    var url = URL.createObjectURL(blob);
    var row = document.createElement("div");
    row.className = "resultRow";
    var name = document.createElement("span");
    name.textContent = filename + " (" + fmtBytes(blob.size) + ")";
    var a = document.createElement("a");
    a.href = url;
    a.download = filename;
    a.textContent = "Download";
    row.appendChild(name);
    row.appendChild(a);
    $("resultsEl").appendChild(row);
    var ext = extOf(filename);
    if (["png", "jpg", "jpeg", "webp", "gif", "bmp"].indexOf(ext) >= 0) {
      var img = document.createElement("img");
      img.src = url;
      img.alt = filename;
      $("previewEl").appendChild(img);
    }
    return url;
  }

  function canvasToBlob(canvas, type, q) {
    return new Promise(function (res, rej) {
      canvas.toBlob(function (b) { b ? res(b) : rej(new Error("canvas export failed")); }, type, q);
    });
  }
  function loadImageFromFile(file) {
    return new Promise(function (res, rej) {
      var url = URL.createObjectURL(file);
      var img = new Image();
      img.onload = function () { res({ img: img, url: url }); };
      img.onerror = function () { URL.revokeObjectURL(url); rej(new Error("cannot read image " + file.name)); };
      img.src = url;
    });
  }
  function baseNameOf(name) {
    var i = name.lastIndexOf(".");
    return i > 0 ? name.slice(0, i) : name;
  }
  function ensureScript(src, globalName) {
    return new Promise(function (res, rej) {
      if (globalName && window[globalName]) return res();
      var s = document.createElement("script");
      s.src = src;
      s.onload = function () { res(); };
      s.onerror = function () { rej(new Error("failed to load " + src)); };
      document.head.appendChild(s);
    });
  }

  function initPdfJs() {
    try {
      if (window.pdfjsLib) {
        window.pdfjsLib.GlobalWorkerOptions.workerSrc =
          "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js";
      }
    } catch (e) {}
  }

  // ---------- spreadsheet helpers ----------
  function aoaOfSheet(ws) {
    return XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: "" });
  }
  function spreadsheetToPdfBlob(wb, onlySheet) {
    var ns = window.jspdf && window.jspdf.jsPDF;
    if (!ns) throw new Error("PDF library not loaded (network needed once for CDN).");
    var doc = new ns({ unit: "pt", compress: true });
    var names = onlySheet && wb.SheetNames.indexOf(onlySheet) >= 0 ? [onlySheet] : wb.SheetNames;
    var first = true;
    names.forEach(function (sn) {
      var aoa = aoaOfSheet(wb.Sheets[sn]);
      if (!aoa.length) return;
      if (!first) doc.addPage();
      first = false;
      doc.setFontSize(12);
      doc.text(String(sn).slice(0, 80), 40, 40);
      var head = [aoa[0].map(function (c) { return String(c); })];
      var body = aoa.slice(1, 1001).map(function (r) { return r.map(function (c) { return String(c); }); });
      doc.autoTable({ startY: 55, head: head, body: body, styles: { fontSize: 8 }, headStyles: { fillColor: [40, 90, 160] } });
    });
    return doc.output("blob");
  }
  function sheetToCanvas(aoa, scale) {
    scale = scale || 2;
    var maxRows = Math.min(aoa.length, 200);
    var maxCols = 0;
    for (var r = 0; r < maxRows; r++) maxCols = Math.max(maxCols, aoa[r].length);
    maxCols = Math.min(Math.max(maxCols, 1), 20);
    var widths = [];
    for (var c = 0; c < maxCols; c++) {
      var m = 6;
      for (var r2 = 0; r2 < maxRows; r2++) m = Math.max(m, String(aoa[r2][c] == null ? "" : aoa[r2][c]).length);
      widths[c] = Math.min(46, Math.max(10, m)) * 7.2;
    }
    var rh = 24, padH = 44;
    var W = Math.ceil(widths.reduce(function (a, b) { return a + b; }, 0) + padH);
    var H = Math.ceil(maxRows * rh + padH);
    var cv = document.createElement("canvas");
    cv.width = Math.min(4000, Math.ceil(W * scale / 2));
    cv.height = Math.min(4000, Math.ceil(H * scale / 2));
    var sx = cv.width / W, sy = cv.height / H;
    var ctx = cv.getContext("2d");
    var dark = window.matchMedia && window.matchMedia("(prefers-color-scheme: dark)").matches;
    ctx.fillStyle = dark ? "#1b1e24" : "#ffffff";
    ctx.fillRect(0, 0, cv.width, cv.height);
    ctx.scale(sx, sy);
    ctx.font = "12px system-ui, sans-serif";
    var y = padH / 2;
    for (var r3 = 0; r3 < maxRows; r3++) {
      var x = padH / 2, hrow = r3 === 0;
      if (hrow) { ctx.fillStyle = dark ? "#2c3e57" : "#e8eef7"; ctx.fillRect(x - 4, y - 16, W - padH + 8, rh); }
      for (var c3 = 0; c3 < maxCols; c3++) {
        ctx.fillStyle = hrow ? (dark ? "#fff" : "#123") : (dark ? "#e8e8e8" : "#222");
        if (hrow) ctx.font = "bold 12px system-ui, sans-serif"; else ctx.font = "12px system-ui, sans-serif";
        var txt = String(aoa[r3][c3] == null ? "" : aoa[r3][c3]);
        if (txt.length > 46) txt = txt.slice(0, 45) + "…";
        ctx.fillText(txt, x, y);
        x += widths[c3];
      }
      ctx.strokeStyle = dark ? "#444" : "#ddd";
      ctx.beginPath(); ctx.moveTo(padH / 2 - 4, y + 8); ctx.lineTo(W - padH / 2 + 4, y + 8); ctx.stroke();
      y += rh;
    }
    ctx.strokeStyle = dark ? "#666" : "#bbb";
    ctx.strokeRect(padH / 2 - 4, padH / 2 - 20, W - padH + 8, maxRows * rh + 8);
    return cv;
  }

  // ---------- pdf helpers ----------
  async function pdfFileToImages(file, imgFmt, scale) {
    if (!window.pdfjsLib) throw new Error("PDF engine not loaded.");
    initPdfJs();
    var ab = await file.arrayBuffer();
    var pdf = await window.pdfjsLib.getDocument({ data: ab }).promise;
    var outs = [];
    var mime = imgFmt === "jpg" ? "image/jpeg" : "image/png";
    for (var p = 1; p <= Math.min(pdf.numPages, 50); p++) {
      var page = await pdf.getPage(p);
      var vp = page.getViewport({ scale: scale || 2 });
      var cv = document.createElement("canvas");
      cv.width = Math.ceil(vp.width);
      cv.height = Math.ceil(vp.height);
      var ctx = cv.getContext("2d");
      if (imgFmt === "jpg") { ctx.fillStyle = "#fff"; ctx.fillRect(0, 0, cv.width, cv.height); }
      await page.render({ canvasContext: ctx, viewport: vp }).promise;
      var blob = await canvasToBlob(cv, mime, 0.92);
      outs.push({ blob: blob, filename: baseNameOf(file.name) + "-p" + p + "." + imgFmt });
    }
    return outs;
  }
  async function pdfFileToRows(file) {
    if (!window.pdfjsLib) throw new Error("PDF engine not loaded.");
    initPdfJs();
    var ab = await file.arrayBuffer();
    var pdf = await window.pdfjsLib.getDocument({ data: ab }).promise;
    var all = []; // {page, text}
    for (var p = 1; p <= Math.min(pdf.numPages, 50); p++) {
      var page = await pdf.getPage(p);
      var tc = await page.getTextContent();
      var buckets = {};
      tc.items.forEach(function (it) {
        var y = Math.round(it.transform[5] / 3) * 3;
        (buckets[y] = buckets[y] || []).push(it);
      });
      var ys = Object.keys(buckets).map(Number).sort(function (a, b) { return b - a; });
      ys.forEach(function (y) {
        var line = buckets[y].sort(function (a, b) { return a.transform[4] - b.transform[4]; })
          .map(function (x) { return x.str; }).join(" ").replace(/\s+/g, " ").trim();
        if (line) all.push({ page: p, line: line });
      });
      if (!ys.length) all.push({ page: p, line: "" });
    }
    return all;
  }
  function rowsToWorkbook(rows) {
    var aoa = [["page", "text"]];
    rows.forEach(function (r) {
      var parts = String(r.line).split(/\s{2,}\|\s{2,}|\s{2,}/).map(function (s) { return s.trim(); }).filter(Boolean);
      if (parts.length > 1) aoa.push([r.page].concat(parts));
      else aoa.push([r.page, r.line]);
    });
    var wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(aoa), "extracted");
    return wb;
  }

  // ---------- image helpers ----------
  async function imagesToPdfBlob(imgFiles) {
    var ns = window.jspdf && window.jspdf.jsPDF;
    if (!ns) throw new Error("PDF library not loaded.");
    var doc = null;
    for (var i = 0; i < imgFiles.length; i++) {
      var info = await loadImageFromFile(imgFiles[i]);
      try {
        var w = info.img.naturalWidth, h = info.img.naturalHeight;
        var orient = w > h ? "l" : "p";
        if (!doc) doc = new ns({ unit: "pt", format: "a4", orientation: orient });
        else doc.addPage("a4", orient);
        var pw = doc.internal.pageSize.getWidth(), ph = doc.internal.pageSize.getHeight();
        var s = Math.min((pw - 40) / w, (ph - 40) / h);
        var dw = w * s, dh = h * s;
        var fmt = /\.png$/i.test(imgFiles[i].name) ? "PNG" : "JPEG";
        doc.addImage(info.img, fmt, (pw - dw) / 2, (ph - dh) / 2, dw, dh);
      } finally { URL.revokeObjectURL(info.url); }
    }
    return doc.output("blob");
  }
  async function convertSingleImage(file, outFmt) {
    var info = await loadImageFromFile(file);
    try {
      var cv = document.createElement("canvas");
      cv.width = info.img.naturalWidth;
      cv.height = info.img.naturalHeight;
      var ctx = cv.getContext("2d");
      var mime = outFmt === "jpg" ? "image/jpeg" : "image/png";
      if (outFmt === "jpg") { ctx.fillStyle = "#fff"; ctx.fillRect(0, 0, cv.width, cv.height); }
      ctx.drawImage(info.img, 0, 0);
      return await canvasToBlob(cv, mime, 0.92);
    } finally { URL.revokeObjectURL(info.url); }
  }
  async function imagesToWorkbook(imgFiles, useOcr, onTick) {
    var ocrTexts = {};
    if (useOcr) {
      await ensureScript("https://cdn.jsdelivr.net/npm/tesseract.js@5.1.1/dist/tesseract.min.js", "Tesseract");
      for (var i = 0; i < imgFiles.length; i++) {
        try {
          if (onTick) onTick("OCR " + (i + 1) + "/" + imgFiles.length + " …");
          var r = await window.Tesseract.recognize(imgFiles[i], "eng");
          ocrTexts[imgFiles[i].name] = (r && r.data && r.data.text ? r.data.text : "").trim();
        } catch (e) { ocrTexts[imgFiles[i].name] = ""; }
      }
    }
    try {
      await ensureScript("https://cdn.jsdelivr.net/npm/exceljs@4.4.0/dist/exceljs.min.js", "ExcelJS");
      var wb = new window.ExcelJS.Workbook();
      var ws = wb.addWorksheet("images");
      ws.columns = [{ header: "file", width: 30 }, { header: "width", width: 10 }, { header: "height", width: 10 }, { header: "ocr_text", width: 80 }];
      var rowY = 1;
      for (var k = 0; k < imgFiles.length; k++) {
        var f = imgFiles[k];
        var info = await loadImageFromFile(f);
        try {
          var buf = await f.arrayBuffer();
          var b64 = arrayBufferToBase64(buf);
          var ext = extOf(f.name) === "png" ? "png" : "jpeg";
          var imgId = wb.addImage({ base64: "data:image/" + ext + ";base64," + b64, extension: ext });
          var maxW = 520, sh = Math.min(1, maxW / info.img.naturalWidth);
          ws.addRow([f.name, info.img.naturalWidth, info.img.naturalHeight, ocrTexts[f.name] || ""]);
          rowY++;
          var hPx = Math.min(380, info.img.naturalHeight * sh);
          ws.addImage(imgId, { tl: { col: 4, row: rowY - 1 }, ext: { width: Math.round(info.img.naturalWidth * sh), height: Math.round(hPx) } });
          ws.getRow(rowY).height = Math.max(18, hPx * 0.75);
        } finally { URL.revokeObjectURL(info.url); }
      }
      var out = await wb.xlsx.writeBuffer();
      return new Blob([out], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
    } catch (e) {
      var aoa = [["file", "width", "height", "ocr_text"]];
      for (var m = 0; m < imgFiles.length; m++) {
        var fi = imgFiles[m];
        try {
          var inf = await loadImageFromFile(fi);
          aoa.push([fi.name, inf.inf || inf.img.naturalWidth, inf.img.naturalHeight, ocrTexts[fi.name] || ""]);
          URL.revokeObjectURL(inf.url);
        } catch (err) { aoa.push([fi.name, "", "", ocrTexts[fi.name] || ""]); }
      }
      var wb2 = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb2, XLSX.utils.aoa_to_sheet(aoa), "images");
      var arr = XLSX.write(wb2, { bookType: "xlsx", type: "array" });
      return new Blob([arr], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
    }
  }
  function arrayBufferToBase64(buf) {
    var b = new Uint8Array(buf), s = "";
    for (var i = 0; i < b.length; i += 8192) s += String.fromCharCode.apply(null, b.subarray(i, i + 8192));
    return btoa(s);
  }

  async function textFileToWorkbook(file) {
    var t = await file.text();
    var aoa = [["line"]];
    t.split(/\r?\n/).slice(0, 5000).forEach(function (ln) { aoa.push([ln]); });
    var wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(aoa), "text");
    return wb;
  }

  function workbookToBlob(wb, fmt) {
    if (fmt === "xlsx") {
      var arr = XLSX.write(wb, { bookType: "xlsx", type: "array" });
      return new Blob([arr], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
    }
    var sn = wb.SheetNames[0];
    if (fmt === "csv") return new Blob([XLSX.utils.sheet_to_csv(wb.Sheets[sn])], { type: "text/csv" });
    var aoa = aoaOfSheet(wb.Sheets[sn]);
    var txt = aoa.map(function (r) { return r.join("\t"); }).join("\n");
    return new Blob([txt], { type: "text/plain" });
  }

  async function convertOne(f, outFmt, opts) {
    var base = baseNameOf(f.file.name);
    if ((f.kind === "spreadsheet" || f.kind === "csv") && (outFmt === "pdf")) {
      var ab = await f.file.arrayBuffer();
      var wb = XLSX.read(ab, { type: "array" });
      var only = opts.sheet && wb.SheetNames.indexOf(opts.sheet) >= 0 ? opts.sheet : null;
      return [{ blob: spreadsheetToPdfBlob(wb, only), filename: base + ".pdf" }];
    }
    if ((f.kind === "spreadsheet" || f.kind === "csv") && (outFmt === "png" || outFmt === "jpg")) {
      var ab2 = await f.file.arrayBuffer();
      var wb2 = XLSX.read(ab2, { type: "array" });
      var sn2 = opts.sheet && wb2.SheetNames.indexOf(opts.sheet) >= 0 ? opts.sheet : wb2.SheetNames[0];
      var cv = sheetToCanvas(aoaOfSheet(wb2.Sheets[sn2]), opts.scale);
      var mime = outFmt === "jpg" ? "image/jpeg" : "image/png";
      return [{ blob: await canvasToBlob(cv, mime, 0.92), filename: base + "." + outFmt }];
    }
    if ((f.kind === "spreadsheet" || f.kind === "csv") && ["xlsx", "csv", "txt"].indexOf(outFmt) >= 0) {
      var ab3 = await f.file.arrayBuffer();
      var wb3 = XLSX.read(ab3, { type: "array" });
      return [{ blob: workbookToBlob(wb3, outFmt), filename: base + "." + outFmt }];
    }
    if (f.kind === "pdf" && (outFmt === "png" || outFmt === "jpg")) {
      return await pdfFileToImages(f.file, outFmt, opts.scale);
    }
    if (f.kind === "pdf" && ["xlsx", "csv", "txt"].indexOf(outFmt) >= 0) {
      var rows = await pdfFileToRows(f.file);
      var wbP = rowsToWorkbook(rows);
      return [{ blob: workbookToBlob(wbP, outFmt), filename: base + "." + outFmt }];
    }
    if (f.kind === "text" && ["xlsx", "csv", "txt"].indexOf(outFmt) >= 0) {
      var wbT = await textFileToWorkbook(f.file);
      return [{ blob: workbookToBlob(wbT, outFmt), filename: base + "." + outFmt }];
    }
    if (f.kind === "text" && outFmt === "pdf") {
      var ns = window.jspdf && window.jspdf.jsPDF;
      if (!ns) throw new Error("PDF library not loaded.");
      var doc = new ns({ unit: "pt" });
      var t = await f.file.text();
      var lines = doc.splitTextToSize(t.slice(0, 20000), 515);
      doc.text(lines, 40, 50);
      return [{ blob: doc.output("blob"), filename: base + ".pdf" }];
    }
    if (f.kind === "text" && (outFmt === "png" || outFmt === "jpg")) {
      var aoaT = [["line"]].concat((await f.file.text()).split(/\r?\n/).slice(0, 200).map(function (l) { return [l]; }));
      var cvT = sheetToCanvas(aoaT, opts.scale);
      return [{ blob: await canvasToBlob(cvT, outFmt === "jpg" ? "image/jpeg" : "image/png", 0.92), filename: base + "." + outFmt }];
    }
    throw new Error("unsupported conversion for " + f.file.name);
  }

  async function convertAll() {
    var outFmt = $("outputSelect").value;
    var scale = parseFloat($("imgScaleInput").value || "2");
    var sheet = $("sheetSelect").value || null;
    var useOcr = $("ocrCheck").checked;
    clearResults();

    if (!window.XLSX) throw new Error("Spreadsheet library failed to load. Check connection and reload.");
    if (files.some(function (f) { return f.kind === "unknown"; })) throw new Error("One file type is not supported.");

    // multi-image fast paths
    var imgFiles = files.filter(function (f) { return f.kind === "image"; }).map(function (f) { return f.file; });
    var nonImg = files.filter(function (f) { return f.kind !== "image"; });
    var results = [];

    if (imgFiles.length && nonImg.length === 0) {
      if (outFmt === "pdf") {
        setStatus("Building PDF …", true);
        results.push({ blob: await imagesToPdfBlob(imgFiles), filename: "images.pdf" });
      } else if (outFmt === "png" || outFmt === "jpg") {
        for (var i = 0; i < files.length; i++) {
          setStatus("Converting image " + (i + 1) + "/" + files.length + " …", true);
          results.push({ blob: await convertSingleImage(files[i].file, outFmt), filename: baseNameOf(files[i].file.name) + "." + outFmt });
        }
      } else if (outFmt === "xlsx" || outFmt === "csv" || outFmt === "txt") {
        setStatus("Building spreadsheet" + (useOcr ? " with OCR …" : " …"), true);
        var wbBlob = await imagesToWorkbook(imgFiles, useOcr, setStatus);
        if (outFmt === "xlsx") results.push({ blob: wbBlob, filename: "images.xlsx" });
        else {
          var abI = await wbBlob.arrayBuffer();
          var wbI = XLSX.read(abI, { type: "array" });
          results.push({ blob: workbookToBlob(wbI, outFmt), filename: "images." + outFmt });
        }
      }
    } else {
      for (var k = 0; k < files.length; k++) {
        setStatus("Converting " + (k + 1) + "/" + files.length + " …", true);
        var outs = await convertOne(files[k], outFmt, { scale: scale, sheet: sheet });
        outs.forEach(function (o) { results.push(o); });
      }
    }

    setStatus("");
    if (!results.length) throw new Error("Nothing produced.");
    if (results.length === 1) {
      addResult(results[0].blob, results[0].filename);
      setStatus("Done — 1 file ready.");
    } else {
      if (!window.JSZip) throw new Error("Zip library not loaded.");
      setStatus("Zipping " + results.length + " files …", true);
      var zip = new window.JSZip();
      results.forEach(function (r) { zip.file(r.filename, r.blob); });
      var zb = await zip.generateAsync({ type: "blob" });
      addResult(zb, "converted-" + outFmt + ".zip");
      results.forEach(function (r) { addResult(r.blob, r.filename); });
      setStatus("Done — " + results.length + " files ready (plus zip).");
    }
  }

  function addFiles(fileList) {
    Array.prototype.forEach.call(fileList, function (file) {
      if (file.size > 80 * 1048576) return;
      var kind = detectKind(file);
      if (kind === "unknown") {
        setStatus("Skipped unsupported file: " + file.name);
        return;
      }
      files.push({ file: file, kind: kind, ext: extOf(file.name) });
      cacheWorkbook(files[files.length - 1]).then(refreshSheetSelect);
    });
    refreshUI();
  }

  function init() {
    initPdfJs();
    var dz = $("dropZone"), fi = $("fileInput");
    dz.addEventListener("click", function () { fi.click(); });
    fi.addEventListener("change", function () { addFiles(fi.files); fi.value = ""; });
    ["dragover", "dragenter"].forEach(function (ev) {
      dz.addEventListener(ev, function (e) { e.preventDefault(); dz.classList.add("over"); });
    });
    ["dragleave", "drop"].forEach(function (ev) {
      dz.addEventListener(ev, function (e) { e.preventDefault(); dz.classList.remove("over"); });
    });
    dz.addEventListener("drop", function (e) {
      if (e.dataTransfer && e.dataTransfer.files) addFiles(e.dataTransfer.files);
    });
    $("outputSelect").addEventListener("change", function () {
      $("outputHint").textContent = hintFor($("outputSelect").value);
    });
    $("convertBtn").addEventListener("click", function () {
      $("convertBtn").disabled = true;
      setStatus("Starting …", true);
      convertAll().catch(function (err) {
        setStatus("Error: " + (err && err.message ? err.message : err));
      }).then(function () { $("convertBtn").disabled = !files.length; });
    });
    refreshUI();
  }
  if (document.readyState !== "loading") init();
  else document.addEventListener("DOMContentLoaded", init);
})();
