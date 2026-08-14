(function () {
  const $ = (id) => document.getElementById(id);
  const dash = "\u2014";
  const idPattern = /^\s*(B0[A-Z0-9]{8,})\s*,\s*([^,\n]+?-[^,\n]+?)\s*$/i;
  let selected = null;
  let rows = [];

  const text = (v) => v == null ? "" : String(v).trim();
  const parseNumber = (v) => {
    if (typeof v === "number" && Number.isFinite(v)) return v;
    const s = text(v).replace(/,/g, "");
    if (!s || s === dash || s === "-") return null;
    const m = s.match(/-?\d+(?:\.\d+)?/);
    return m ? Number(m[0]) : null;
  };
  const titleCandidate = (s) => s &&
    !/^(?:亚马逊|更改\s*SKU|有货|推荐\s*ASIN)$/i.test(s) &&
    !/^(?:无库存|非在售|您的|参考价|最小值|最大值)/.test(s) &&
    !/^-?\d+(?:[.,]\d+)?$/.test(s) && s !== dash && s !== "-" &&
    !/[:：]\s*€?\s*-?\d/.test(s) && /[A-Za-zÀ-ÿ\u4e00-\u9fff]/.test(s);

  function clean(values) {
    const result = [];
    let current = null;
    let pendingTitle = "";
    let waitingActivity = false;
    let waitingMax = false;
    let referenceState = "none";
    const finish = () => {
      if (current) result.push([current.title, current.asin, current.sku, current.original, current.max]);
      current = null; waitingActivity = false; waitingMax = false; referenceState = "none";
    };
    for (const raw of values) {
      const s = text(raw);
      const id = s.match(idPattern);
      if (id) {
        const title = pendingTitle;
        finish();
        current = { title, asin: id[1], sku: id[2], original: null, max: null };
        pendingTitle = "";
        continue;
      }
      if (!current) { if (s) pendingTitle = s; continue; }
      // A closed item may leave the value row truly blank instead of using
      // an em dash after the 参考价 label. Treat that blank as a missing
      // reference price so the next numeric row can become 活动价最大值.
      if (!s) {
        if (referenceState === "await") referenceState = "missing";
        continue;
      }
      if (/商品价格\s*[:：]?/.test(s) && !/最低|最高/.test(s)) { waitingActivity = true; continue; }
      const inlineMax = s.match(/最大值\s*[:：]?\s*[^\d-]*(-?\d+(?:[.,]\d+)?)/);
      if (inlineMax) { current.max = parseNumber(inlineMax[1].replace(",", ".")); waitingMax = false; referenceState = "done"; continue; }
      if (/最大值\s*[:：]?/.test(s)) { waitingMax = true; continue; }
      if (/参考价\s*[:：]?/.test(s)) { referenceState = "await"; continue; }
      if (referenceState === "await") {
        const n = parseNumber(raw);
        if (n !== null) referenceState = "seen";
        else if (s === dash || s === "-") referenceState = "missing";
        continue;
      }
      // Some browser copies omit the separate “最大值” label but keep the
      // next price row. In the source layout that first number after a
      // non-empty reference price is the activity-price maximum.
      if (referenceState === "seen" && current.max === null) {
        const n = parseNumber(raw);
        if (n !== null) { current.max = n; referenceState = "done"; continue; }
      }
      // Closed ASINs may omit the 最大值 label. In that case the first
      // number after a blank reference price is the discounted-price maximum.
      if (referenceState === "missing" && current.max === null) {
        const n = parseNumber(raw);
        if (n !== null) { current.max = n; referenceState = "done"; continue; }
      }
      if (waitingActivity) {
        const n = parseNumber(raw);
        if (n !== null) { current.original = n; waitingActivity = false; }
        else if (s === dash || s === "-") waitingActivity = false;
        continue;
      }
      if (waitingMax) {
        const n = parseNumber(raw);
        if (n !== null) { current.max = n; waitingMax = false; }
        continue;
      }
      if (titleCandidate(s)) pendingTitle = s;
    }
    finish();
    return result;
  }

  // A few WPS exports are valid ZIP files but contain central-directory
  // records that make JSZip reject the whole archive. Read only the entries
  // needed by XLSX with the browser's native deflate decoder instead.
  async function readZipEntriesNative(buffer, wanted) {
    const bytes = new Uint8Array(buffer);
    const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
    const u16 = (offset) => view.getUint16(offset, true);
    const u32 = (offset) => view.getUint32(offset, true);
    let eocd = -1;
    for (let i = bytes.length - 22; i >= Math.max(0, bytes.length - 65558); i--) {
      if (u32(i) === 0x06054b50) { eocd = i; break; }
    }
    if (eocd < 0) throw new Error(`找不到 ZIP 目录（文件大小：${bytes.length} 字节）。`);
    // ZIP directory offsets are relative to the ZIP payload. A browser file
    // provider may prepend a wrapper header, so locate the first local-file
    // header and account for that prefix when resolving all offsets.
    let zipStart = -1;
    for (let i = 0; i + 4 <= bytes.length; i++) {
      if (u32(i) === 0x04034b50) { zipStart = i; break; }
    }
    if (zipStart < 0) {
      const rawDirOffset = u32(eocd + 16);
      const rawDirSize = u32(eocd + 12);
      const candidate = eocd - rawDirSize - rawDirOffset;
      zipStart = candidate >= 0 ? candidate : 0;
    }
    const entries = u16(eocd + 10);
    const directoryOffset = zipStart + u32(eocd + 16);
    const records = new Map();
    let cursor = directoryOffset;
    for (let i = 0; i < entries && cursor + 46 <= bytes.length; i++) {
      if (u32(cursor) !== 0x02014b50) break;
      const method = u16(cursor + 10);
      const compressedSize = u32(cursor + 20);
      const uncompressedSize = u32(cursor + 24);
      const nameLength = u16(cursor + 28);
      const extraLength = u16(cursor + 30);
      const commentLength = u16(cursor + 32);
      const localOffset = zipStart + u32(cursor + 42);
      const name = new TextDecoder().decode(bytes.slice(cursor + 46, cursor + 46 + nameLength));
      if (wanted.has(name)) records.set(name, { method, compressedSize, uncompressedSize, localOffset });
      cursor += 46 + nameLength + extraLength + commentLength;
    }
    const result = new Map();
    for (const name of wanted) {
      const record = records.get(name);
      if (!record) continue;
      const nameLength = u16(record.localOffset + 26);
      const extraLength = u16(record.localOffset + 28);
      const start = record.localOffset + 30 + nameLength + extraLength;
      const compressed = bytes.slice(start, start + record.compressedSize);
      let content;
      if (record.method === 0) {
        content = compressed;
      } else if (record.method === 8 && typeof DecompressionStream === "function") {
        const stream = new Blob([compressed]).stream().pipeThrough(new DecompressionStream("deflate-raw"));
        content = new Uint8Array(await new Response(stream).arrayBuffer());
      } else {
        throw new Error(`暂不支持 ZIP 压缩方式：${record.method}`);
      }
      if (record.uncompressedSize && content.byteLength !== record.uncompressedSize) {
        throw new Error(`ZIP 文件内容不完整：${name}`);
      }
      result.set(name, new TextDecoder("utf-8").decode(content));
    }
    return result;
  }

  async function loadZip(buffer) {
    if (!window.JSZip) return null;
    try { return await window.JSZip.loadAsync(buffer); } catch (firstError) {
      // Some WPS exports have a non-ZIP prefix. Try each local-file signature
      // so the browser can recover the valid XLSX payload without uploading it.
      const bytes = new Uint8Array(buffer);
      for (let i = 0; i < bytes.length - 3; i++) {
        if (bytes[i] === 0x50 && bytes[i + 1] === 0x4b && bytes[i + 2] === 0x03 && bytes[i + 3] === 0x04) {
          try { return await window.JSZip.loadAsync(bytes.slice(i).buffer); } catch (_) { /* keep scanning */ }
        }
      }
      return null;
    }
  }

  async function parseRawXlsx(buffer) {
    const wanted = new Set(["xl/sharedStrings.xml", "xl/worksheets/sheet1.xml"]);
    let zip = await loadZip(buffer);
    let sharedXml = "";
    let sheetXml = "";
    if (zip) {
      const sharedEntry = zip.file("xl/sharedStrings.xml");
      const sheetEntry = zip.file("xl/worksheets/sheet1.xml");
      if (sheetEntry) {
        sharedXml = sharedEntry ? await sharedEntry.async("text") : "";
        sheetXml = await sheetEntry.async("text");
      }
    }
    if (!sheetXml) {
      const native = await readZipEntriesNative(buffer, wanted);
      sharedXml = native.get("xl/sharedStrings.xml") || "";
      sheetXml = native.get("xl/worksheets/sheet1.xml") || "";
    }
    if (!sheetXml) throw new Error("找不到第一个工作表。");
    const sharedDoc = sharedXml ? new DOMParser().parseFromString(sharedXml, "application/xml") : null;
    const shared = sharedDoc ? Array.from(sharedDoc.getElementsByTagName("si")).map((si) => Array.from(si.getElementsByTagName("t")).map((t) => t.textContent || "").join("")) : [];
    const doc = new DOMParser().parseFromString(sheetXml, "application/xml");
    return Array.from(doc.getElementsByTagName("row")).map((row) => {
        const cell = Array.from(row.children).find((c) => /^A\d+$/i.test(c.getAttribute("r") || ""));
        if (!cell) return "";
        const v = cell.getElementsByTagName("v")[0]?.textContent || "";
        const type = cell.getAttribute("t") || "";
        if (type === "s") return shared[Number(v)] ?? "";
        if (type === "inlineStr") return Array.from(cell.getElementsByTagName("t")).map((t) => t.textContent || "").join("");
        return v === "" ? "" : Number(v);
      });
  }

  async function readBufferRows(buffer) {
    const signature = new Uint8Array(buffer);
    let values = [];
    try {
      if (window.XLSX) {
        const wb = window.XLSX.read(buffer, { type: "array", cellDates: false, WTF: true });
        const ws = wb.Sheets[wb.SheetNames[0]];
        values = window.XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: "" }).map((r) => r?.[0] ?? "");
      }
    } catch (_) { values = []; }
    if (values.some((v) => idPattern.test(text(v)))) return values;
    // Do not reject a file solely because its first bytes are not PK. Some
    // browser file providers prepend a wrapper header even though the XLSX
    // ZIP payload is complete. The ZIP/raw parser below can locate and read
    // that payload. OLE .xls remains unsupported by the raw parser but is
    // still passed to the SheetJS reader above.
    return parseRawXlsx(buffer);
  }

  function localHelperAllowed() {
    // The helper is an optional local-only fallback. A shared/hosted page
    // must remain a pure browser app and must never contact a user's machine.
    if (!window.location || !/^https?:$/.test(window.location.protocol)) return false;
    const host = String(window.location.hostname || "").toLowerCase();
    const port = String(window.location.port || (window.location.protocol === "https:" ? "443" : "80"));
    return (host === "127.0.0.1" || host === "localhost") && (port === "8765" || port === "8767");
  }

  async function readLocalHelper(fileName, lastModified) {
    if (!localHelperAllowed()) return null;
    // Keep the reader on the same computer. It is never used by a remotely
    // shared page, where the browser alone must handle the workbook.
    const origin = "http://127.0.0.1:8767";
    try {
      const response = await fetch(`${origin}/api/read-file`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ name: fileName, lastModified })
      });
      if (!response.ok) return null;
      return await response.arrayBuffer();
    } catch (_) {
      return null;
    }
  }

  async function readFileRows(file) {
    let firstError;
    try {
      return await readBufferRows(await file.arrayBuffer());
    } catch (err) {
      firstError = err;
    }
    // When a browser/WPS file provider exposes an encrypted virtual view,
    // ask the optional local helper to read the original workbook from the
    // user's Desktop/Downloads/Documents folder instead.
    const fallback = await readLocalHelper(file.name, file.lastModified);
    if (fallback) return readBufferRows(fallback);
    throw firstError;
  }

  function render() {
    $("count").textContent = rows.length;
    $("titleCount").textContent = rows.filter((r) => r[0]).length;
    $("missing").textContent = rows.reduce((n, r) => n + r.slice(3).filter((v) => v == null).length, 0);
    const escape = (v) => String(v).replace(/[&<>'"]/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", "'": "&#39;", '"': "&quot;" }[c]));
    $("body").innerHTML = rows.slice(0, 100).map((r) => `<tr>${r.map((v, i) => `<td class="${i > 2 ? "num" : ""}">${v == null ? "" : escape(v)}</td>`).join("")}</tr>`).join("") || '<tr><td colspan="5" style="text-align:center;color:#8492a3;padding:30px">没有识别到商品数据</td></tr>';
    $("export").disabled = !rows.length;
  }

  function setStatus(message, ok) {
    $("status").className = "status " + (ok ? "ok" : "error");
    $("status").textContent = message;
  }

  function clearSelectedFile() {
    selected = null;
    $("file").value = "";
    $("fileName").textContent = "尚未选择文件";
    if ($("systemPick")) $("systemPick").hidden = true;
  }

  async function chooseWithSystemPicker() {
    if (typeof window.showOpenFilePicker !== "function") {
      throw new Error("SYSTEM_PICKER_UNAVAILABLE");
    }
    const handles = await window.showOpenFilePicker({
      multiple: false,
      types: [{
        description: "Excel 文件",
        accept: {
          "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet": [".xlsx"],
          "application/vnd.ms-excel": [".xls"]
        }
      }]
    });
    if (!handles?.[0]) return null;
    return handles[0].getFile();
  }

  // Reset the browser file control immediately before opening the chooser.
  // This guarantees that selecting the same workbook again after 清空输入
  // produces a fresh change event and never reuses the previous File object.
  $("file").addEventListener("click", () => {
    selected = null;
    $("file").value = "";
  });

  $("file").addEventListener("change", (e) => {
    selected = e.target.files[0] || null;
    if (selected) $("paste").value = "";
    $("fileName").textContent = selected ? selected.name : "尚未选择文件";
    if ($("systemPick")) $("systemPick").hidden = true;
  });
  $("drop").addEventListener("drop", (e) => {
    selected = e.dataTransfer.files[0] || null;
    if (selected) $("paste").value = "";
    $("fileName").textContent = selected ? selected.name : "尚未选择文件";
    if ($("systemPick")) $("systemPick").hidden = true;
  });
  $("paste").addEventListener("input", (e) => {
    if (text(e.target.value).trim()) clearSelectedFile();
  });
  $("clear").addEventListener("click", clearSelectedFile);
  if ($("systemPick")) {
    $("systemPick").addEventListener("click", async () => {
      try {
        const file = await chooseWithSystemPicker();
        if (!file) return;
        selected = file;
        $("paste").value = "";
        $("fileName").textContent = file.name;
        $("systemPick").hidden = true;
        setStatus("已用系统方式重新选择文件，正在读取…", true);
        $("clean").click();
      } catch (err) {
        if (err?.name !== "AbortError") setStatus("系统文件读取不可用，请改用普通本地目录中的 .xlsx 文件。", false);
      }
    });
  }

  document.addEventListener("click", async (e) => {
    if (e.target?.id === "clean") {
      e.preventDefault(); e.stopImmediatePropagation();
      try {
        let values;
        const pasted = $("paste").value.replace(/\r/g, "");
        // Pasted content takes precedence, so a previously selected file
        // cannot accidentally trigger a second file read.
        if (pasted.trim()) values = pasted.split("\n").flatMap((line) => line.split("\t"));
        else if (selected) values = await readFileRows(selected);
        else values = [];
        rows = clean(values); render();
        if ($("systemPick")) $("systemPick").hidden = true;
        setStatus(rows.length ? `清洗完成：识别到 ${rows.length} 条商品，可导出结果。` : "没有识别到符合规则的 ASIN/SKU，请检查输入内容。", rows.length > 0);
      } catch (err) {
        const detail = err?.message || String(err);
        const localOnly = localHelperAllowed();
        const message = /INVALID_XLSX_SIGNATURE/.test(detail)
          ? (localOnly
            ? "读取失败：浏览器返回的文件内容没有 ZIP/XLSX 数据；本机辅助读取也未找到原文件。请确认文件在桌面、下载或文档目录后重试。"
            : "读取失败：浏览器返回的文件内容没有可解析的 ZIP/XLSX 数据。请确认选择的是标准 .xlsx 文件，或改用粘贴网页数据。")
          : /ZIP 目录|ZIP 文件内容不完整|找不到第一个工作表/.test(detail)
          ? (localOnly
            ? "读取失败：浏览器返回的内容不是可解析的 ZIP/XLSX 数据；本机辅助读取也未找到原文件。请改用普通本地 .xlsx 文件或粘贴网页数据。"
            : "读取失败：浏览器返回的内容不是可解析的 ZIP/XLSX 数据。请确认选择的是标准 .xlsx 文件，或改用粘贴网页数据。")
          : "读取失败：" + detail;
        if ($("systemPick")) {
          $("systemPick").hidden = typeof window.showOpenFilePicker !== "function";
        }
        setStatus(message, false);
      }
      return;
    }
    if (e.target?.id === "export") {
      e.preventDefault(); e.stopImmediatePropagation();
      if (!rows.length || !window.XLSX) return;
      const ws = window.XLSX.utils.aoa_to_sheet([["标题", "ASIN", "SKU", "原价", "活动价最大值"], ...rows]);
      ws["!cols"] = [{ wch: 70 }, { wch: 16 }, { wch: 38 }, { wch: 14 }, { wch: 16 }];
      const wb = window.XLSX.utils.book_new();
      window.XLSX.utils.book_append_sheet(wb, ws, "清洗结果");
      window.XLSX.writeFile(wb, "Amazon_清洗结果.xlsx");
    }
  }, true);
})();
