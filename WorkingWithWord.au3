#include <Word.au3>
Local $oWord = _Word_Create()
Local $sDocument = "C:\Users\vinh\Desktop\sample.doc"
Local $oDoc = _Word_DocOpen($oWord, $sDocument)
WinWaitActive("sample [Compatibility Mode] - Word")
sleep(1000)
_Word_DocFindReplace($oDoc, "ông UserName1 và bà UserName2", "ông tống văn vinh")
_Word_DocFindReplace($oDoc, "phuongName", "phường Văn Phú")
_Word_DocSaveAs($oDoc, "D:\Unity\lol.doc")
_Word_Quit($oWord)