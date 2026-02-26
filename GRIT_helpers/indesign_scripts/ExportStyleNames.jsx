#target "InDesign"

(function () {
  function toJsonString(value, pretty) {
    var indentUnit = pretty ? "  " : "";

    function repeat(str, count) {
      var out = "";
      for (var i = 0; i < count; i++) out += str;
      return out;
    }

    function escapeString(str) {
      return String(str)
        .replace(/\\/g, "\\\\")
        .replace(/\"/g, "\\\"")
        .replace(/\r/g, "\\r")
        .replace(/\n/g, "\\n")
        .replace(/\t/g, "\\t");
    }

    function serialize(v, level) {
      if (v === null) return "null";

      var t = typeof v;
      if (t === "string") return '"' + escapeString(v) + '"';
      if (t === "number") return isFinite(v) ? String(v) : "null";
      if (t === "boolean") return v ? "true" : "false";

      if (v instanceof Array) {
        if (v.length === 0) return "[]";
        var arrParts = [];
        for (var i = 0; i < v.length; i++) {
          var arrVal = serialize(v[i], level + 1);
          if (indentUnit) {
            arrParts.push("\n" + repeat(indentUnit, level + 1) + arrVal);
          } else {
            arrParts.push(arrVal);
          }
        }
        return indentUnit
          ? "[" + arrParts.join(",") + "\n" + repeat(indentUnit, level) + "]"
          : "[" + arrParts.join(",") + "]";
      }

      if (t === "object") {
        var keys = [];
        for (var k in v) {
          if (v.hasOwnProperty(k)) keys.push(k);
        }
        if (keys.length === 0) return "{}";

        var objParts = [];
        for (var j = 0; j < keys.length; j++) {
          var key = keys[j];
          var val = serialize(v[key], level + 1);
          var keyPart = '"' + escapeString(key) + '":' + (indentUnit ? " " : "");
          if (indentUnit) {
            objParts.push("\n" + repeat(indentUnit, level + 1) + keyPart + val);
          } else {
            objParts.push(keyPart + val);
          }
        }
        return indentUnit
          ? "{" + objParts.join(",") + "\n" + repeat(indentUnit, level) + "}"
          : "{" + objParts.join(",") + "}";
      }

      return "null";
    }

    if (typeof JSON !== "undefined" && JSON && typeof JSON.stringify === "function") {
      return JSON.stringify(value, null, pretty ? 2 : 0);
    }
    return serialize(value, 0);
  }

  if (app.documents.length === 0) {
    alert("Open a document first.");
    return;
  }

  var doc = app.activeDocument;
  var out = {
    document: doc.name,
    paragraphStyles: [],
    characterStyles: []
  };

  var i;
  for (i = 0; i < doc.allParagraphStyles.length; i++) {
    out.paragraphStyles.push(doc.allParagraphStyles[i].name);
  }

  for (i = 0; i < doc.allCharacterStyles.length; i++) {
    out.characterStyles.push(doc.allCharacterStyles[i].name);
  }

  var file = File.saveDialog("Save styles JSON", "*.json");
  if (!file) return;

  file.encoding = "UTF-8";
  file.open("w");
  file.write(toJsonString(out, true));
  file.close();

  alert("Exported style names.");
})();
