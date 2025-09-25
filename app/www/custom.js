/* www/custom.js
 * Signals shinylive presence, confirms SheetJS, enables browser-side parse and export.
 */
(function () {
  "use strict";

  function whenShinyReady(cb, tries = 0) {
    if (window.Shiny && Shiny.fileInputBinding) { cb(); return; }
    if (tries > 200) { console.warn("Shiny not detected; server mode or too early."); return; }
    setTimeout(function () { whenShinyReady(cb, tries + 1); }, 50);
  }

  function signalFlags() {
    try {
      var isShinylive = !!(window.webR || window.webr);
      Shiny.setInputValue("client_is_shinylive", isShinylive, { priority: "event" });
      var hasSheetJS = !!(window.XLSX && XLSX.utils && XLSX.writeFile);
      Shiny.setInputValue("client_has_sheetjs", hasSheetJS, { priority: "event" });
    } catch (e) {}
  }

  // Allow R to request (re)loading SheetJS dynamically (fallback if local failed)
  function ensureSheetJS() {
    if (window.XLSX && XLSX.utils && XLSX.writeFile) { signalFlags(); return; }
    var s = document.createElement("script");
    // Try local first (already referenced by app.R, but retry just in case of load order)
    s.src = "xlsx.full.min.js";
    s.onload = signalFlags;
    s.onerror = function () {
      // Fallback to CDN
      var s2 = document.createElement("script");
      s2.src = "https://cdn.sheetjs.com/xlsx-0.20.3/package/dist/xlsx.full.min.js";
      s2.onload = signalFlags;
      s2.onerror = signalFlags;
      document.head.appendChild(s2);
    };
    document.head.appendChild(s);
  }

  // Client-side parse (optional; helps in Shinylive)
  function parseFile(file) {
    return new Promise(function (resolve, reject) {
      var reader = new FileReader();
      reader.onerror = function () { reject(reader.error); };
      reader.onload = function (e) {
        try {
          var data = new Uint8Array(e.target.result);
          var wb = XLSX.read(data, { type: "array", cellDates: false });
          var sheets = wb.SheetNames.map(function (name) {
            var ws = wb.Sheets[name];
            var aoa = XLSX.utils.sheet_to_json(ws, { header: 1, raw: false, defval: "" });
            return { name: name, data: aoa.slice(0, Math.min(aoa.length, 10000)) };
          });
          resolve({ name: file.name, sheets: sheets });
        } catch (err) { reject(err); }
      };
      reader.readAsArrayBuffer(file);
    });
  }

  whenShinyReady(function () {
    // Initial flags + a short delayed refresh
    signalFlags();
    setTimeout(signalFlags, 400);

    // Let R request a (re)load of SheetJS
    Shiny.addCustomMessageHandler("load_sheetjs", function () { ensureSheetJS(); });

    // Patch file input to parse Excel in-browser when possible (Shinylive)
    try {
      var fib = Shiny.fileInputBinding && Shiny.fileInputBinding.prototype;
      if (fib && typeof fib.onChange === "function") {
        var origOnChange = fib.onChange;
        fib.onChange = function () {
          var res = origOnChange.apply(this, arguments);
          try {
            var ev = arguments[0];
            var el = ev && ev.target ? ev.target : null;
            if (!el || el.id !== "legacy_files") return res;

            var files = el.files ? Array.prototype.slice.call(el.files) : [];
            if (!files.length) return res;

            if (!(window.XLSX && XLSX.utils)) return res; // skip if SheetJS not loaded
            (async function () {
              try {
                var results = [];
                for (var i = 0; i < files.length; i++) {
                  results.push(await parseFile(files[i]));
                }
                Shiny.setInputValue("excel_parsed", { files: results }, { priority: "event" });
              } catch (err) {
                console.error(err); alert("Error parsing file(s) in browser: " + err);
              }
            })();
          } catch (e) {}
          return res;
        };
      }
    } catch (e) {
      console.warn("Could not patch fileInputBinding.onChange:", e);
    }

    // Browser-side Excel download (called from R)
    Shiny.addCustomMessageHandler("download_xlsx", function (payloadJSON) {
      try {
        if (!(window.XLSX && XLSX.utils && XLSX.writeFile)) {
          alert("Browser export unavailable: SheetJS not loaded.");
          return;
        }
        var payload = JSON.parse(payloadJSON);
        var wb = XLSX.utils.book_new();
        var sheetNames = Object.keys(payload.sheets || {});
        for (var s = 0; s < sheetNames.length; s++) {
          var sheetName = sheetNames[s];
          var df = payload.sheets[sheetName];
          if (!df || !df.length) continue;
          var cols = Object.keys(df[0]);
          var ws = XLSX.utils.json_to_sheet(df, { header: cols, skipHeader: false });
          XLSX.utils.book_append_sheet(wb, ws, sheetName);
        }
        XLSX.writeFile(wb, payload.filename || "SEND_abbrev.xlsx");
      } catch (e) {
        console.error(e); alert("Export failed in browser: " + e);
      }
    });
  });
})();
