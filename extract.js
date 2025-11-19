// extract.js (robust extractor that maps embedded images to rows)
// Usage: node extract.js
import fs from "fs-extra";
import JSZip from "jszip";
import { XMLParser } from "fast-xml-parser";
const EXCEL_FILE = "realdata.xlsx";
const OUTPUT_JSON = "data.json";

async function main() {
  try {
    console.log("Reading XLSX:", EXCEL_FILE);
    const buf = await fs.readFile(EXCEL_FILE);
    const zip = await JSZip.loadAsync(buf);
    const parser = new XMLParser({
      ignoreAttributes: false,
      attributeNamePrefix: "@_",
      removeNSPrefix: true
    });

    // --- 1) load sharedStrings (if present) ---
    let sharedStrings = [];
    if (zip.file("xl/sharedStrings.xml")) {
      const sstXml = await zip.file("xl/sharedStrings.xml").async("string");
      const sst = parser.parse(sstXml);
      if (sst?.sst?.si) {
        const sis = Array.isArray(sst.sst.si) ? sst.sst.si : [sst.sst.si];
        sharedStrings = sis.map(si => {
          // si.t may be a string or object
          if (si.t === undefined && si.r) {
            // rich text: join all t pieces
            const parts = Array.isArray(si.r) ? si.r : [si.r];
            return parts.map(p => (typeof p.t === "object" && p.t["#text"]) ? p.t["#text"] : p.t || "").join("");
          }
          if (typeof si.t === "object" && si.t["#text"]) return si.t["#text"];
          return si.t ?? "";
        });
      }
    }

    // --- 2) gather media files (xl/media/) into base64 map ---
    const mediaBase64 = {};
    for (const name of Object.keys(zip.files)) {
      if (name.startsWith("xl/media/")) {
        const ext = name.split(".").pop().toLowerCase();
        const base64 = await zip.file(name).async("base64");
        mediaBase64[name] = `data:image/${ext};base64,${base64}`;
      }
    }
    console.log(`Found ${Object.keys(mediaBase64).length} media files.`);

    // --- 3) parse drawings and drawing rels to map embed rIds -> media paths ---
    // We'll build maps per drawing file:
    // drawingRelsMap: 'xl/drawings/drawingN.xml' -> { rId: 'xl/media/imageX.png' }
    // drawingAnchorsMap: 'xl/drawings/drawingN.xml' -> [ { row: 5, relId: 'rId1' }, ... ]
    const drawingRelsMap = {};
    const drawingAnchorsMap = {};

    const drawingFiles = Object.keys(zip.files).filter(p => p.startsWith("xl/drawings/") && p.endsWith(".xml"));
    for (const drawingFile of drawingFiles) {
      // parse rels for this drawing
      const relsPath = drawingFile.replace("xl/drawings/", "xl/drawings/_rels/") + ".rels";
      const relMap = {};
      if (zip.file(relsPath)) {
        try {
          const relsXml = await zip.file(relsPath).async("string");
          const relsObj = parser.parse(relsXml);
          const relList = relsObj.Relationships?.Relationship
            ? (Array.isArray(relsObj.Relationships.Relationship) ? relsObj.Relationships.Relationship : [relsObj.Relationships.Relationship])
            : [];
          for (const r of relList) {
            const id = r["@_Id"] || r.Id;
            const target = r["@_Target"] || r.Target;
            if (id && target) {
              // normalize: rel target might be ../media/image1.png or media/image1.png
              let t = target;
              if (t.startsWith("../")) t = t.replace("../", "xl/");
              else if (!t.startsWith("xl/")) t = "xl/" + t;
              relMap[id] = t;
            }
          }
        } catch (e) {
          console.warn("Could not parse rels for", drawingFile, e.message);
        }
      }
      drawingRelsMap[drawingFile] = relMap;

      // parse anchors (positions) in drawing file
      try {
        const drawingXml = await zip.file(drawingFile).async("string");
        const dObj = parser.parse(drawingXml);
        // Try several namespace variants
        const wsDr = dObj.wsDr || dObj["xdr:wsDr"] || dObj["xdr:wsDr"] || dObj["xdr:wsDr"] || dObj;
        // collect both oneCellAnchor and twoCellAnchor
        let anchors = [];
        if (wsDr.oneCellAnchor) anchors = anchors.concat(Array.isArray(wsDr.oneCellAnchor) ? wsDr.oneCellAnchor : [wsDr.oneCellAnchor]);
        if (wsDr.twoCellAnchor) anchors = anchors.concat(Array.isArray(wsDr.twoCellAnchor) ? wsDr.twoCellAnchor : [wsDr.twoCellAnchor]);
        if (wsDr["xdr:oneCellAnchor"]) anchors = anchors.concat(Array.isArray(wsDr["xdr:oneCellAnchor"]) ? wsDr["xdr:oneCellAnchor"] : [wsDr["xdr:oneCellAnchor"]]);
        if (wsDr["xdr:twoCellAnchor"]) anchors = anchors.concat(Array.isArray(wsDr["xdr:twoCellAnchor"]) ? wsDr["xdr:twoCellAnchor"] : [wsDr["xdr:twoCellAnchor"]]);

        const anchorList = [];
        for (const a of anchors) {
          // anchor may have 'from' or 'xdr:from'
          const from = a.from || a["xdr:from"] || a["xdr:from"] || a;
          const pic = a.pic || a["xdr:pic"] || a;
          // get row (the row element might be a number string or object)
          let rowIdx = null;
          try {
            const rowVal = from?.row ?? from?.["xdr:row"] ?? from?.["xdr:row"] ?? from?.r;
            if (rowVal !== undefined && rowVal !== null) {
              rowIdx = parseInt(rowVal, 10) + 1; // convert 0-based -> 1-based
            } else {
              // try nested structure
              const nested = from?.["xdr:row"] || from?.row;
              if (nested !== undefined) rowIdx = parseInt(nested, 10) + 1;
            }
          } catch (e) {
            rowIdx = null;
          }

          // locate r:embed
          let relId = null;
          try {
            // possible paths to blip embed
            const blip = (pic && (pic.blipFill?.blip || pic["xdr:blipFill"]?.["a:blip"] || pic.blipFill?.["a:blip"])) || (a.pic?.blipFill?.blip);
            if (blip) {
              relId = blip["@_r:embed"] || blip["@_embed"] || blip["r:embed"] || blip.embed;
            }
          } catch (e) {
            relId = null;
          }

          if (rowIdx && relId) anchorList.push({ row: rowIdx, rId: relId });
        }

        drawingAnchorsMap[drawingFile] = anchorList;
        console.log(`Parsed ${anchorList.length} anchors from ${drawingFile}`);
      } catch (err) {
        console.warn("Failed parsing drawing file", drawingFile, err.message);
        drawingAnchorsMap[drawingFile] = [];
      }
    } // end for each drawing file

    // --- 4) Build a global row -> media dataURL map ---
    const rowToImage = {}; // 1-based row -> dataURL

    for (const drawingFile of Object.keys(drawingAnchorsMap)) {
      const anchors = drawingAnchorsMap[drawingFile];
      const relMap = drawingRelsMap[drawingFile] || {};
      for (const a of anchors) {
        const rId = a.rId;
        const mediaTarget = relMap[rId]; // e.g. 'xl/media/image1.png' or 'media/image1.png'
        if (!mediaTarget) continue;
        // normalize key to actual zip file key
        let mediaKey = mediaTarget;
        if (!mediaKey.startsWith("xl/")) mediaKey = "xl/" + mediaKey;
        if (!zip.file(mediaKey)) {
          // sometimes rel target points to ../media/image1.png
          const alt = mediaKey.replace("xl/drawings/../", "xl/");
          mediaKey = alt;
        }
        const dataUrl = mediaBase64[mediaKey];
        if (dataUrl) {
          // If multiple images map to same row, keep first (or you could collect an array)
          if (!rowToImage[a.row]) rowToImage[a.row] = dataUrl;
          else {
            // already exists — keep existing but log
            console.log(`Multiple images for row ${a.row}; keeping first.`);
          }
        } else {
          // mediaKey not found in mediaBase64; list available media keys for debugging
          //console.warn("No media data for", mediaKey);
        }
      }
    }

    console.log("Mapped images to rows for", Object.keys(rowToImage).length, "rows.");

    // --- 5) parse sheet1 rows and compose final JSON ---
    const sheetXml = await zip.file("xl/worksheets/sheet1.xml").async("string");
    const sheetObj = parser.parse(sheetXml);
    const sheetRows = sheetObj.worksheet?.sheetData?.row;
    const rowsArray = Array.isArray(sheetRows) ? sheetRows : (sheetRows ? [sheetRows] : []);

    const out = [];
    for (const r of rowsArray) {
      const rowNum = parseInt(r["@_r"], 10);
      if (rowNum === 1) continue; // skip header

      // create blank row object
      const obj = { name: "", type: "", website: "", description: "", image: "" };

      const cells = Array.isArray(r.c) ? r.c : (r.c ? [r.c] : []);
      for (const c of cells) {
        const cellRef = c["@_r"]; // like "A2"
        const col = cellRef.replace(/[0-9]/g, "");
        let val = "";
        if (c.v !== undefined) {
          // if type is shared string
          if (c["@_t"] === "s") {
            val = sharedStrings[ parseInt(c.v, 10) ] || "";
          } else {
            val = (typeof c.v === "object" && c.v["#text"]) ? c.v["#text"] : c.v;
          }
        } else if (c.is && c.is.t) {
          val = c.is.t;
        }
        val = String(val || "").replace(/&#10;|\\n/g, " ").trim();

        if (col === "A") obj.name = val;
        if (col === "B") obj.type = val;
        if (col === "C") obj.website = val;
        if (col === "D") obj.description = val;
      }

      // attach image if mapped
      if (rowToImage[rowNum]) obj.image = rowToImage[rowNum];
      else obj.image = "";

      out.push(obj);
    }

    // write JSON
    await fs.writeJSON(OUTPUT_JSON, out, { spaces: 2 });
    console.log("Wrote", out.length, "rows to", OUTPUT_JSON);
  } catch (err) {
    console.error("ERROR:", err && err.stack ? err.stack : err);
    process.exit(1);
  }
}

main();
