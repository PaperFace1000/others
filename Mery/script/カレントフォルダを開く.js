#title="カレントフォルダを開く"
// -----------------------------------------------------------------------------
// カレントフォルダを開く
// about:  今開いているファイルのカレントフォルダを開く。
// create: 2018/02/07 PaperFace1000
// update: 
// Copyright (c) PaperFace1000. All Rights Reserved.
//  http://paperface.hatenablog.com/
// -----------------------------------------------------------------------------

var fso = new ActiveXObject('Scripting.FileSystemObject');
var sh = new ActiveXObject( "WScript.Shell" );

var name = Document.Name;
var dirPath = Document.Path;

// 一応、システムパスを取得して絶対パスで指定したexplorer.exeを使う
sh.CurrentDirectory = dirPath;
var cmd = fso.BuildPath(fso.GetSpecialFolder(0), "explorer.exe") + 
          " /select,\"" + name + "\"";
          
sh.Exec(cmd);
