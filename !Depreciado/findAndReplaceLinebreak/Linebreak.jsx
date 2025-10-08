
// Ensure a document is open
if (app.documents.length === 0) {
    alert("No document open.");
} else {
    var doc = app.activeDocument;

    // Save current find/change preferences
    var findPrefs = app.findTextPreferences.properties;
    var changePrefs = app.changeTextPreferences.properties;

    try {
        // Clear previous find/change preferences
        app.findTextPreferences = NothingEnum.NOTHING;
        app.changeTextPreferences = NothingEnum.NOTHING;

        // Set find what: literal "\n"
        app.findTextPreferences.findWhat = "\\n";

        // Set change to: forced line break (SpecialCharacter.FORCED_LINE_BREAK)
        app.changeTextPreferences.changeTo = SpecialCharacters.FORCED_LINE_BREAK;

        // Run find/change
        var results = doc.changeText();

        alert("Replaced " + results.length + " occurrences of \\n with forced line break.");
    } catch (e) {
        alert("Error: " + e);
    } finally {
        // Restore previous preferences
        app.findTextPreferences.properties = findPrefs;
        app.changeTextPreferences.properties = changePrefs;
    }
}