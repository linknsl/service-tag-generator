// Установка пути к каталогу с файлами
var folderPath = $.fileName.split('\\').slice(0, -1).join('\\');
var pathgroup = "double";

// Получение списка файлов
var files = Folder(folderPath + "/../../output/" + pathgroup).getFiles("0_2_Green_Red_0000.pdf");
var outputFolder = new Folder(folderPath + "/../../output/" + pathgroup);
if (!outputFolder.exists) {
    outputFolder.create();
}
// Define the output directory path and log file path
var logFilePath = outputFolder + "/log_generate.txt";

var logFile = new File(logFilePath);
logFile.open("w");
logFile.close();

writeToLog("Start script");
for (var f = 0; f < files.length; f++) {
    var count = 1;
    writeToLog("Reading.... files:" + files.length);
    var file = files[f];
    writeToLog("Read file :" + file.toString());
    // Открытие файла в Illustrator
    var doc = app.open(file);
    var doc = activeDocument


    // Сохранение файла в PDF-формате
    for (outdoc = 1; outdoc < 10; outdoc++) {
        if (outdoc > 0) {
            replaceTextAllfiles(outdoc);
        }
        var numStr = String(1000 * outdoc);
        while (numStr.length < 4) {
            numStr = "0" + numStr;
        }
        writeToLog("Replace text in all documents: " + numStr);
        var pdfFile = new File(outputFolder.fullName + "/" + file.name.replace(/0000.pdf$/, "") + numStr + ".pdf");
        writeToLog("save file:" + pdfFile.toString());
        doc.saveAs(pdfFile, PDFSaveOptions);
    }
    doc.close();
    writeToLog("Finish:");
}

function replaceTextAllfiles(number) {
    for (var i = 0; i < doc.textFrames.length; i++) {
        var textFrame = doc.textFrames[i];
        var textRange = textFrame.textRange;
        var contents = textRange.contents;

        var replacedContents = contents.replace(/\b\d{4}\b/g, function (match) {
            var num = parseInt(match, 10);
            return String((num % 1000) + 1000 * number);
        });
        if (contents !== replacedContents) {
            textRange.contents = replacedContents;
        }
    }
}

function writeToLog(text) {
    logFile.open("a");
    logFile.writeln(new Date() + " - " + text);
    logFile.close();
}
