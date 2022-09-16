const path = require("path")
const fs = require("fs")
const admZip = require("adm-zip");
const xmlEditor = require("xml_editor");

const get_Bool = (val, def) => (typeof val === "boolean" ? val : def);
const get_Str = (val, def) => (typeof val === "string" ? val : def);

const defaultOptions = {
    // path to where the "unzip" dir will be placed
    outputDir: __dirname
};

const relevantEntries = {
    // all items are xml files, created when docx unzipped, that need altering
    "app.xml": null,
    "core.xml": null,
    "comments.xml": null,
    "commentsExtended.xml": null,
    "commentsExtensible.xml": null,
    "commentsIds.xml": null,
    "document.xml": null,
    "settings.xml": null
}

const zip = new admZip();
module.exports = function(/**String */ docxStr, /**object */ options) {
    /**
     * Unzips a docx file to a directory called unzip.
     * @param docxStr docx file full path of entry
     * @options 
     */
    /**
     * Access contents, xml files
     * Save xml files to variables
     * Update xml files through xml_editor
     * Zip updated xml files to create new docx
     * Save path to new docx
     */

    // create object based default options, allowing them to be overwritten
    const opts = Object.assign(Object.create(null), defaultOptions);

    // test docxStr
    try {
        fs.existsSync(docxStr)
    } catch (err) {
        console.error(err);
        return
    }
    // assign options
    const _zip = new admZip(docxStr);

    /**
     * 
     * @param entryID String of xml entry file (ei comments.xml, not full path)
     * @returns Buffer or Null in case of error
     */
    function getEntry(/**String */ entryID) {
        if (entryID) {
            item = _zip.getEntries();
        }
        return null
    }
    /**
     * Save all entries of the docx to a list, save ones of importance
     */
    function getEntries() {
        for (let i = 0; i < relevantEntries.length(); i++) {
            console.log(relevantEntries[i]);
        }
        return null;
    }

    return {
        /**
         * @param outputDir Full path to extract directory
         * @param overwrite If the file already exists at the target path, the file will be overwritten if this is true
         *                  Default is False
         */
        extractArchive: function(/**String */ outputDir, /**Bool */ overwrite) {
            overwrite = get_Bool(overwrite, false);
            _zip.extractAllTo(outputDir, overwrite);
        },
        findFileComponent: function(/**String */ fileStr, /**String */ fileID) {

        },
        /**
         * 
         * @returns Full path to new docx file or Null in case of error
         */
        getNewDocx: function() {
            return _zip.newDocx || null;
        }
    }
}

exports._getNewDocx = function() {

}