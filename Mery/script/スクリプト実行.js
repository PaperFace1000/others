#title="スクリプト実行"
// -----------------------------------------------------------------------------
// 今編集中のドキュメントをスクリプトとして実行
// about:  今編集しているテキストをスクリプトとして実行します。
//         拡張子によってどの言語か判定します。
// create: 2018/02/20 PaperFace1000
// update: 
// ---------------------------------------------------------------------------

// スクリプトエンジンパス
var PYTHON_PATH = "\"C:\\ProgramData\\Anaconda3\\python.exe\"";
var PYTHONW_PATH = "\"C:\\ProgramData\\Anaconda3\\pythonw.exe\"";

// デフォルト拡張子
var DEFAULT_EX = ".py";

main();

// 日付フォーマット
function formatDate (date, format) 
{
  format = format.replace(/yyyy/g, date.getFullYear());
  format = format.replace(/MM/g, ('0' + (date.getMonth() + 1)).slice(-2));
  format = format.replace(/dd/g, ('0' + date.getDate()).slice(-2));
  format = format.replace(/HH/g, ('0' + date.getHours()).slice(-2));
  format = format.replace(/mm/g, ('0' + date.getMinutes()).slice(-2));
  format = format.replace(/ss/g, ('0' + date.getSeconds()).slice(-2));
  format = format.replace(/SSS/g, ('00' + date.getMilliseconds()).slice(-3));
  return format;
}

function main()
{
    var fso = new ActiveXObject('Scripting.FileSystemObject');
    var sh = new ActiveXObject( "WScript.Shell" );
    
    if(Document.Name == ""){
        var scriptDir = fso.BuildPath(fso.GetParentFolderName(ScriptFullName), "exec_scripts");
        var outputFileName = Prompt("このスクリプトは保存されていません。ファイル名を決めてください。", 
                                    fso.BuildPath(scriptDir, formatDate(new Date(), "yyyyMMdd_HHmmss_SSS" + DEFAULT_EX)));
        if(!fso.FolderExists(scriptDir)){
            fso.CreateFolder(scriptDir);
        }
        Document.Save(outputFileName);
    }
    if(!Document.Saved){
        Document.Save();
    }
    
    var workDir = fso.GetParentFolderName(Document.FullName);
    var fileName = Document.Name;
    
    sh.CurrentDirectory = workDir;
    var cmd;
    
    switch(fso.GetExtensionName(fileName).toLowerCase()){
        case "py":
            cmd = PYTHON_PATH + " \"" + fileName + "\"";
            Document.Mode = "Python";
            break;
        case "pyw":
            cmd = PYTHONW_PATH + " \"" + fileName + "\"";
            Document.Mode = "Python";
            break;
        default:
            Quit();
    }
    
    //Alert(cmd);
    
    var output = sh.Exec(cmd);
    var stdout = output.StdOut.ReadAll();
    var stderr = output.StdErr.ReadAll();
    
    OutputBar.Visible = true;
    OutputBar.Writeln(formatDate(new Date(), "yyyy/MM/dd HH:mm:ss.SSS") + ">" + cmd);
    if(stdout != ""){
        OutputBar.Writeln(stdout);
    }
    if(stderr != ""){
        OutputBar.Writeln(stderr);
    }
}