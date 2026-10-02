// Strip trailing encoded newline characters from InDesign hyperlink URLs
// Removes %0D (carriage return) and %0A (line feed) from END of URLs only.

(function () {
    if (app.documents.length === 0) {
        alert("No document is open.");
        return;
    }

    var doc = app.activeDocument;
    var fixed = 0;
    var changes = [];

    for (var i = 0; i < doc.hyperlinkURLDestinations.length; i++) {
        var dest = doc.hyperlinkURLDestinations[i];

        try {
            var oldURL = dest.destinationURL;

            // Strip one or more trailing encoded CR/LF characters.
            var newURL = oldURL.replace(/(%0D|%0A)+$/gi, "");

            if (newURL !== oldURL) {
                dest.destinationURL = newURL;
                fixed++;

                changes.push(
                    oldURL + "\r→ " + newURL
                );
            }
        } catch (e) {
            // Skip invalid/inaccessible destinations
        }
    }

    if (fixed === 0) {
        alert("Done. No hyperlinks ending in %0D or %0A were found.");
    } else {
        alert(
            "Done. Fixed " + fixed + " hyperlink" +
            (fixed === 1 ? "." : "s.") +
            "\r\r" +
            changes.join("\r\r")
        );
    }
})();
