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
    "docProps/app.xml": null,
    "docProps/core.xml": null,
    "word/comments.xml": null,
    "word/commentsExtended.xml": null,
    "word/commentsExtensible.xml": null,
    "word/commentsIds.xml": null,
    "word/document.xml": null,
    "word/settings.xml": null
}

/**
 * Unzips a docx file to a directory called unzip.
 * @param {String} docxStr docx file full path of entry
 * @options 
 */
/**
 * Access contents, xml files
 * Save xml files to variables
 * Update xml files through xml_editor
 * Zip updated xml files to create new docx
 * Save path to new docx
 */

module.exports = function(/**String */ docxStr, /**object */ options) {
    let inBuffer = null;
    // create object based default options, allowing them to be overwritten
    const opts = Object.assign(Object.create(null), defaultOptions);

    // test docxStr
    if (docxStr && "string" === typeof docxStr) {
        try {
            fs.existsSync(docxStr);
            inBuffer = docxStr;
        } catch (err) {
            console.error(err);
            return
        }
    } else {console.error("docx is not passed in to docx_zip");}
    
    // assign options
    const _zip = new admZip(docxStr);

    /**
     * Save all entries of the docx to a list, save ones of importance
     */
    function getEntries() {
        for (let [xml, value] of Object.entries(relevantEntries)) {
            let entry = _zip.getEntry(xml);
            relevantEntries[xml] = entry.toJSON();
        }
        return relevantEntries;
    }

    return {
        /**
         * 
         * @param {String} extractDir Full path to the docx file
         * @param {String} outputFile Full path to the file to unzip the file to
         */
        createZipArchive: function(/**String */ extractDir, /**String */ outputFile) {
            _zip.addLocalFolder(extractDir);
            _zip.writeZip(outputFile);
        },
        /**
         * @param {String} outputDir Full path to extract directory
         * @param {bool} overwrite If the file already exists at the target path, the file will be overwritten if this is true
         *                  Default is False
         */
        extractArchive: function(/**String */ outputDir, /**Bool */ overwrite) {
            overwrite = get_Bool(overwrite, false);
            _zip.extractAllTo(outputDir, overwrite);
        },
        /**
         * 
         * @param {String} fileStr  
         * @returns 
         */
        findFileComponents: function(/**String */ fileStr, /**String */ fileID) {
            return getEntries();
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