// Установка пути к каталогу с файлами
var folderPath = $.fileName.split('\\').slice(0, -1).join('\\');
var setQuantity = 19;
var countTemplate = 1;
var column = 12, quantity = Math.ceil(setQuantity / 6), line = Math.ceil(quantity / column), spacing = 2.5,
    stop = quantity;

// Получение списка файлов
var files = Folder(folderPath + "/../../resources/template/Color").getFiles("*.ai");
var outputFolder = new Folder(files[0].parent.fullName + "/../../../output/color"); // Путь к папке output, находящейся рядом с папкой source
if (!outputFolder.exists) {
    outputFolder.create(); // Создание папки output, если она не существует
} else {
    var outfiles = outputFolder.getFiles();
    for (var r = 0; r < outfiles.length; r++) {
        outfiles[r].remove();
    }
}
// Define the output directory path and log file path
var logFilePath = outputFolder + "/log.txt";

var logFile = new File(logFilePath);
logFile.open("w");
logFile.close();

writeToLog("Script path: " + folderPath);
// Обработка каждого файла
writeToLog("Start script");
writeToLog("quantity: " + quantity);
writeToLog("line: " + line);
writeToLog("column: " + column);
writeToLog("stop: " + column);
for (var f = 0; f < files.length; f++) {
    var count = 1;
    writeToLog("Reading.... files:" + files.length);
    var file = files[f];
    writeToLog("Read file :" + file.toString());
    // Открытие файла в Illustrator
    var doc = app.open(file);
    var doc = activeDocument
    curAbIdx = doc.artboards.getActiveArtboardIndex();

    selection = null;

// Copy Artwork
    writeToLog("Start Copy Double");
    doc.selectObjectsOnActiveArtboard();
    var abItems = selection;
    for (var i = 0; i < line; i++) {
        for (var j = 0; j < column; j++) {
            if (count <= (stop - countTemplate)) {
                writeToLog("added number list: " + count);
                suffix = " copy" + fillZero((i + 1), line.toString().length);
                duplicateArtboard(curAbIdx, abItems, spacing, suffix, j, i);
                count++
            }
        }
    }

    doc.selection = null;

    writeToLog("Start renumber");
    for (var montageAreaIndex = 0; montageAreaIndex < doc.artboards.length; montageAreaIndex++) {
        writeToLog("renumber list: " + montageAreaIndex);
        replaceTextInTextFrames(montageAreaIndex, doc.artboards.length - (montageAreaIndex + 1))
    }
    doc.selection = null;

    // Сохранение файла в PDF-формате
    var pdfFile = new File(outputFolder.fullName + "/" + file.name.replace(/\.ai$/, "") + ".pdf");
    writeToLog("save file:" + pdfFile.toString());
    doc.saveAs(pdfFile, PDFSaveOptions);
    doc.close();
}

function replaceTextInTextFrames(montageAreaIndex, numberAdd) {
    // Выбираем нужную монтажную область
    doc.artboards.setActiveArtboardIndex(montageAreaIndex);

    // Выбираем все объекты на активной монтажной области
    doc.selectObjectsOnActiveArtboard();

    // Получаем выделенные текстовые фреймы
    var textFrames = [];
    for (var i = 0; i < doc.selection.length; i++) {
        var item = doc.selection[i];
        if (item.typename === "TextFrame") {
            textFrames.push(item);
        }
    }
    // Изменяем текст только в выделенных текстовых фреймах
    for (var j = 0; j < textFrames.length; j++) {
        var textFrame = textFrames[j];
        var textRange = textFrame.textRange;
        var contents = textRange.contents;

        var replacedContents = contents.replace(/\b0{0,3}[0-9]{1,3}\b/g, function (match) {
            var num = parseInt(match, 10);
            var numStr = String(num + numberAdd);
            while (numStr.length < 3) {
                numStr = "0" + numStr;
            }
            return numStr;
        });

        if (contents !== replacedContents) {
            textRange.contents = replacedContents;
        }
    }
    // Снимаем выделение со всех объектов
    doc.selection = null;
}

function fillZero(number, size) {
    var str = '000000000' + number;
    return str.slice(str.length - size);
}

function duplicateArtboard(thisAbIdx, items, spacing, suffix, counter, row) {
    var doc = activeDocument,
        thisAb = doc.artboards[thisAbIdx],
        thisAbRect = thisAb.artboardRect,
        idx = doc.artboards.length - 1,
        lastAb = doc.artboards[idx],
        lastAbRect = lastAb.artboardRect,
        abWidth = thisAbRect[2] - thisAbRect[0] + spacing,
        abHeight = thisAbRect[3] - thisAbRect[1] - spacing;


    var newAb = doc.artboards.add(thisAbRect);

    if (counter % column === 0) {
        newAb.artboardRect = [
            thisAbRect[0] + abWidth,
            thisAbRect[1] + abHeight * row,
            thisAbRect[2] + abWidth,
            thisAbRect[3] + abHeight * row
        ];
    } else {
        newAb.artboardRect = [
            lastAbRect[2] + spacing,
            lastAbRect[1],
            lastAbRect[2] + abWidth,
            lastAbRect[3]
        ];
    }
    newAb.name = thisAb.name + suffix;

    var docCoordSystem = CoordinateSystem.DOCUMENTCOORDINATESYSTEM,
        abCoordSystem = CoordinateSystem.ARTBOARDCOORDINATESYSTEM,
        isDocCoords = app.coordinateSystem == docCoordSystem,
        dupArr = getDuplicates(items);

    // Move copied items to the new artboard
    for (var i = 0; i < dupArr.length; i++) {
        var pos = isDocCoords ? dupArr[i].position : doc.convertCoordinate(dupArr[i].position, docCoordSystem, abCoordSystem);
        dupArr[i].position = [pos[0] + abWidth * (counter + 1), pos[1] + abHeight * row];
    }
}

function getDuplicates(collection) {
    var arr = [];
    for (var i = 0, len = collection.length; i < len; i++) {
        arr.push(collection[i].duplicate());
    }
    return arr;
}

function writeToLog(text) {
    logFile.open("a");
    logFile.writeln(new Date() + " - " + text);
    logFile.close();
}