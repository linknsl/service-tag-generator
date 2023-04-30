// Установка пути к каталогу с файлами
var folderPath = "D:\\design\\pdf_illustrator\\src\\sourceDouble";
var setQuantity = 1000;
var countTemplate = 2;
var column = 12, quantity = Math.ceil(setQuantity / 6), line = Math.ceil(quantity / column), spacing = 2.5,
    stop = quantity * 2;
var logFilePath = folderPath + "/log.txt";

var logFile = new File(logFilePath);
logFile.open("w");
logFile.close();
// Получение списка файлов
var files = Folder(folderPath).getFiles("*.ai");
var outputFolder = new Folder(files[0].parent.fullName + "/../outputDouble/"); // Путь к папке output, находящейся рядом с папкой source
if (!outputFolder.exists) {
    outputFolder.create(); // Создание папки output, если она не существует
} else {
    var outfiles = outputFolder.getFiles();
    for (var r = 0; r < outfiles.length; r++) {
        outfiles[r].remove();
    }
}
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

// Copy Artwork
    for (var i = 0; i < line; i++) {
        for (var j = 0; j < column; j++) {
            suffix = fillZero((i + 1), line.toString().length);
            for (var template = 0; template < 2; template++) {
                if (count <= (stop - countTemplate)) {
                    writeToLog("added number sheet: " + count);
                    doc.artboards.setActiveArtboardIndex(template);
                    doc.selectObjectsOnActiveArtboard();
                    duplicateArtboard(template, selection, spacing, suffix, j, i);
                    count++
                }
            }
        }
    }

    doc.selection = null;
    writeToLog("Start renumber");
    for (j = 0; j < 2; j++) {
        var k = 0
        for (var montageAreaIndex = j; montageAreaIndex < doc.artboards.length; montageAreaIndex = montageAreaIndex + 2) {
            writeToLog("renumber list: " + montageAreaIndex);
            doc.artboards.setActiveArtboardIndex(montageAreaIndex);
            doc.selectObjectsOnActiveArtboard();
            replaceTextInTextFrames(montageAreaIndex, doc.artboards.length / 2 - 1 - k++)
        }
    }
    doc.selection = null;

    // Сохранение файла в PDF-формате
    for (outdoc = 0; outdoc < 10; outdoc++) {
        if (outdoc > 0) {
            replaceTextAllfiles(outdoc);
        }
        var numStr = String(1000 * outdoc);
        while (numStr.length < 4) {
            numStr = "0" + numStr;
        }
        writeToLog("Replace text in all documents: " + numStr);
        var pdfFile = new File(outputFolder.fullName + "/" + file.name.replace(/\.ai$/, "") + "_" + numStr + ".pdf");
        writeToLog("save file:" + pdfFile.toString());
        doc.saveAs(pdfFile, PDFSaveOptions);
    }
    doc.close();
    writeToLog("Finish:");
}

function replaceTextAllfiles(number) {
    var doc = app.activeDocument;
    for (var i = 0; i < doc.textFrames.length; i++) {
        var textFrame = doc.textFrames[i];
        var textRange = textFrame.textRange;
        var contents = textRange.contents;

        var replacedContents = contents.replace(/\b0{0,3}[0-9]{1,3}\b/g, function (match) {
            var num = parseInt(match, 10);
            var numStr = String(num + 1000 * number);
            while (numStr.length < 4) {
                numStr = "0" + numStr;
            }
            return numStr;
        });

        if (contents !== replacedContents) {
            textRange.contents = replacedContents;
        }
    }
}

function replaceTextInTextFrames(montageAreaIndex, numberAdd) {
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
            while (numStr.length < 4) {
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
        thisAb2 = doc.artboards[thisAbIdx + 1],
        thisAb2Rect = thisAb2.artboardRect,
        idx = doc.artboards.length - 1,
        lastAb = doc.artboards[idx],
        lastAbRect = lastAb.artboardRect,
        shift = thisAb2Rect[2] - thisAb2Rect[0] + spacing
    abWidth = thisAbRect[2] - thisAbRect[0] + spacing + shift,
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
            lastAbRect[2] + (abWidth - shift),
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