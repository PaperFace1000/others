#title="パス起動"
// -----------------------------------------------------------------------------
// パス起動（仮）
// about:  カーソルがある行、もしくは選択されているテキストをパスとして起動します。
// create: 2018/02/05 
// update: 
// Copyright (c) PaperFace1000. All Rights Reserved.
//  http://paperface.hatenablog.com/
// -----------------------------------------------------------------------------

var fso = new ActiveXObject('Scripting.FileSystemObject');
var sh = new ActiveXObject( "WScript.Shell" );

var doc = Document;
var sel = doc.Selection;

var line;

if(!sel.isEmpty){
    line = sel.Text.replace("\n", "");
}else{
    line = doc.GetLine(sel.GetActivePointY(mePosLogical), meGetLineWithNewLines).replace("\n", "");
}

function deleteQuote(str)
{
    var result;
    var re1 = /^\"/g;
    var re2 = /\"$/g;
    result = str.replace(re1, "");
    result = result.replace(re2, "");
    
    return result;
}

function addQuote(str)
{
    return "\"" + str + "\"";
}

var path = deleteQuote(line);

if(fso.FolderExists(path) || fso.FileExists(path)){
    sh.Run(addQuote(path));
    Quit();
}else{
    Alert("パスが見つかりません。\npath: " + line);
}