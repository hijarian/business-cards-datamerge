/**
 * Скрипт автоматического конвертирования всех открытых файлов в кривые и сохранения их в EPS и PDF
 */
function main() {
    if (app.documents.length == 0) {
        alert("Не открыто ни одного файла - прекращаем работу.");
        return;
    }
    while (app.documents.length != 0) {
        var doc = app.documents[0];
        var file = doc.fullName;
        var filename = file.path + '/' + file.name;
        saveDocAsPDF(doc, filename);
        convertTextToOutlines(doc);
        saveDocAsEPS (doc, filename + ' кривые');
        doc.close( SaveOptions.DONOTSAVECHANGES );
    }
}

function convertTextToOutlines(doc) {
	while (doc.textFrames.length != 0) {
        doc.textFrames[0].createOutline();
    }
}

function saveDocAsEPS(doc, filename) {
    var epsFile = new File(filename + '.eps');
    var epsOptions = new EPSSaveOptions();
    doc.saveAs(epsFile, epsOptions);
}

function saveDocAsPDF(doc, filename) {
    var pdfFile = new File(filename + '.pdf');
    var pdfOptions = getPDFOptions ();
    doc.saveAs(pdfFile, pdfOptions);
}

// Создаём второй аргумент для метода saveAs.
function getPDFOptions()
{
	var pdfSaveOpts = new PDFSaveOptions();
	pdfSaveOpts.acrobatLayers = true;
	pdfSaveOpts.colorBars = false;
	pdfSaveOpts.compressArt = true; //default
	pdfSaveOpts.embedICCProfile = true;
	pdfSaveOpts.enablePlainText = true;
	pdfSaveOpts.generateThumbnails = true; // default
	pdfSaveOpts.optimization = true;
	pdfSaveOpts.pageInformation = false;
	pdfSaveOpts.trimMarks = true; // MUST HAVE, CLIENT REQUIREMENT
    pdfSaveOpts.offset = 12; // MUST HAVE, CLIENT REQUIREMENT. CUSTOM CRAFTED VALUE ACCORDING TO ~4mm!
    return pdfSaveOpts;
}

main();
