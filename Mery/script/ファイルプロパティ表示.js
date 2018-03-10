#title="ファイルプロパティ表示"
// ----------------------------------------------------------------------------- 
// ファイルプロパティ表示 
// about:  今開いているファイルのプロパティを表示する。 
// create: 2018/02/05 PaperFace1000 
// update: 
// Copyright (c) PaperFace1000. All Rights Reserved. 
//  http://paperface.hatenablog.com/ 
// ----------------------------------------------------------------------------- 

// 起動オプション 
//var IS_OUTPUT_TO_OUTPUT_BAR = true; //アウトプットバーに表示する 
var IS_OUTPUT_TO_OUTPUT_BAR = false; //メッセージダイアログに表示する 

var IS_SET_TO_CLIP_BOARD = true; //クリップボードにプロパティ情報を埋め込む 
//var IS_SET_TO_CLIP_BOARD = false; //クリップポードにプロパティ情報を埋め込む 




var fso = new ActiveXObject('Scripting.FileSystemObject'); 

var fullPath = Document.FullName; 
if(fullPath == "") Quit(); 

var file = fso.GetFile(fullPath); 

// 日付フォーマット 
function formatDate(date) 
{ 
    var format = "yyyy/MM/dd HH:mm:ss"; 
    var obj = new Date(date); 
    format = format.replace(/yyyy/g, obj.getFullYear()); 
    format = format.replace(/MM/g, ('0' + (obj.getMonth() + 1)).slice(-2)); 
    format = format.replace(/dd/g, ('0' + obj.getDate()).slice(-2)); 
    format = format.replace(/HH/g, ('0' + obj.getHours()).slice(-2)); 
    format = format.replace(/mm/g, ('0' + obj.getMinutes()).slice(-2)); 
    format = format.replace(/ss/g, ('0' + obj.getSeconds()).slice(-2)); 
    return format; 
} 

// ファイル属性名取得 
// 全部テストしてないから全部検知できているかわからん... 
function getAttributesName(file) 
{ 
    var result = ""; 
    var attributes = file.Attributes; 
    
    if(attributes == 0) result += "Normal, ";               //標準ファイル 
    if(attributes & 1) result += "ReadOnly, ";              //読み取り専用 
    if(attributes & 2) result += "Hidden, ";                //隠しファイル 
    if(attributes & 4) result += "System, ";                //システムファイル 
    if(attributes & 8) result += "VolumeLabel, ";           //ディスクボリュームラベル 
    if(attributes & 16) result += "Directory(Folder), ";    //ディレクトリ（フォルダ） 
    if(attributes & 32) result += "Archive, ";              //アーカイブ 
    if(attributes & 64) result += "Alias(Shortcut), ";      //エイリアス（ショートカット） 
    if(attributes & 128) result += "Compressed, ";          //圧縮ファイル 
    
    if(result.length > 2){ 
        result = result.substring(0, result.length - 2); 
    }else{ 
        result = "Unknown"; 
    } 
    
    return result; 
} 

// スペース区切りのプロパティ情報作成 
function createMsgSpaceBlank(file) 
{ 
    var msg = 
        "ファイル名:       " + file.Name + "\n" + 
        "フォルダ:         " + file.ParentFolder + "\n" + 
        "フルパス:         " + file.Path + "\n" + 
        "サイズ:           " + file.Size.toString().replace(/(\d)(?=(\d\d\d)+$)/g, "$1,") + " byte \n" + 
        "属性:             " + getAttributesName(file) + "\n" + 
        "作成日時:         " + formatDate(file.DateCreated) + "\n" + 
        "更新日時:         " + formatDate(file.DateLastModified) + "\n" + 
        "最終アクセス日時: " + formatDate(file.DateLastAccessed); 
    return msg; 
} 

// タブ区切りのプロパティ情報 
function createMsgTabBlank(file) 
{ 
    var msg = 
        "ファイル名:\t\t" + file.Name + "\n" + 
        "フォルダ:\t\t" + file.ParentFolder + "\n" + 
        "フルパス:\t\t" + file.Path + "\n" + 
        "サイズ:\t\t" + file.Size.toString().replace(/(\d)(?=(\d\d\d)+$)/g, "$1,") + " byte \n" + 
        "属性:\t\t" + getAttributesName(file) + "\n" + 
        "作成日時:\t\t" + formatDate(file.DateCreated) + "\n" + 
        "更新日時:\t\t" + formatDate(file.DateLastModified) + "\n" + 
        "最終アクセス日時:\t" + formatDate(file.DateLastAccessed); 
    return msg; 
} 

if(IS_OUTPUT_TO_OUTPUT_BAR){ 
    // Meryのアウトプットバーへ出力する場合 
    var result = createMsgSpaceBlank(file); 
    
    OutputBar.Clear(); 
    OutputBar.Visible = true; 
    OutputBar.Writeln(result); 
}else{ 
    // メッセージダイアログとして出力する場合 
    var result = createMsgTabBlank(file); 
    
    var sh = new ActiveXObject( "WScript.Shell" ); 
    var icon = 64;  //(i)アイコン 
    var button = 0; // [OK]ボタン 
    sh.Popup(result, 0, file.Name + " のプロパティ", icon + button); 
} 

if(IS_SET_TO_CLIP_BOARD) 
{ 
    ClipboardData.SetData(createMsgSpaceBlank(file)); 
} 