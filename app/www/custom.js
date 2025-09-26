/* www/custom.js */
(function () {
  "use strict";

  function whenShinyReady(cb, tries = 0) {
    if (window.Shiny && Shiny.fileInputBinding) { cb(); return; }
    if (tries > 200) { console.warn("Shiny not detected; too early?"); return; }
    setTimeout(function(){ whenShinyReady(cb, tries+1); }, 50);
  }

  // Optional: parse Excel client-side (useful in Shinylive)
  function parseFile(file) {
    return new Promise(function (resolve, reject) {
      var reader = new FileReader();
      reader.onerror = function(){ reject(reader.error); };
      reader.onload = function(e){
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
    console.log("[custom.js] loaded; SheetJS:", !!(window.XLSX && XLSX.utils && XLSX.writeFile));

    // Patch file input to parse in browser (Shinylive)
    try {
      var fib = Shiny.fileInputBinding && Shiny.fileInputBinding.prototype;
      if (fib && typeof fib.onChange === "function") {
        var origOnChange = fib.onChange;
        fib.onChange = function () {
          var res = origOnChange.apply(this, arguments);
          try {
            var ev = arguments[0], el = ev && ev.target ? ev.target : null;
            if (!el || el.id !== "legacy_files") return res;
            var files = el.files ? Array.prototype.slice.call(el.files) : [];
            if (!files.length || !(window.XLSX && XLSX.utils)) return res;
            (async function(){
              try {
                var results = [];
                for (var i=0; i<files.length; i++) results.push(await parseFile(files[i]));
                Shiny.setInputValue("excel_parsed", { files: results }, { priority: "event" });
              } catch (err) { console.error(err); alert("Error parsing file(s) in browser: " + err); }
            })();
          } catch (e) {}
          return res;
        };
      }
    } catch (e) { console.warn("Could not patch fileInputBinding.onChange:", e); }

    // Download: build workbook in browser and save as .xlsx
    Shiny.addCustomMessageHandler("download_xlsx", function (payloadJSON) {
      try {
        if (!(window.XLSX && XLSX.utils && XLSX.writeFile)) {
          alert("Export unavailable: SheetJS not loaded.");
          return;
        }
        console.log("[custom.js] download_xlsx message received");
        var payload = JSON.parse(payloadJSON);
        var wb = XLSX.utils.book_new();
        var sheetNames = Object.keys(payload.sheets || {});
        for (var s=0; s<sheetNames.length; s++) {
          var sheetName = sheetNames[s];
          var df = payload.sheets[sheetName];
          if (!df || !df.length) continue;
          var cols = Object.keys(df[0]);
          var ws = XLSX.utils.json_to_sheet(df, { header: cols, skipHeader: false });
          XLSX.utils.book_append_sheet(wb, ws, sheetName);
        }
        XLSX.writeFile(wb, payload.filename || "SEND_abbrev.xlsx"); // triggers browser download
      } catch (e) {
        console.error(e); alert("Export failed in browser: " + e);
      }
    });
  });
})();