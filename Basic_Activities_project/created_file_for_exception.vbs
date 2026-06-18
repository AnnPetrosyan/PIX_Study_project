'textData - строка для записи в файл
'rewriteFile - булево значение указывающее что создаваемый файл может быть перезаписан (если False и файл уже существовует, будет вызвано исключение)
'outLenString - количество записанных символов

Dim textData, rewriteFile, outLenString

textData = textData
rewriteFile = rewriteFile
outLenString = outLenString

With WScript.CreateObject("Scripting.FileSystemObject").CreateTextFile("file.txt", rewriteFile)
    .Write (textData)
    .Close
End With

outLenString = Len(textData)

Wscript.Echo outLenString