function showFiles(oFolder){
	var files = new Enumerator(oFolder.Files);
	var s = "";
	var fso = new ActiveXObject("Scripting.FileSystemObject");
	var oExcel = new ActiveXObject("excel.application");
    for (; !files.atEnd(); files.moveNext()) {
        var f= files.item();
		s += files.item();
	
		var ext=f.name.substring(f.name.lastIndexOf(".")+1);
		if(ext=="xls"||ext=="xlsx"){
		
			mnm=fso.GetParentFolderName(f.path)+"\\"+fso.getbasename(f.path);
			oExcel.Workbooks.Open(f.path);
			oExcel.ActiveWorkbook.ExportAsFixedFormat(0,mnm+ext+".pdf");
			oExcel.ActiveWorkbook.Close();
			
		}
        s += "\n";
    }
	oExcel.Quit();
	return s;
}
function showSubFolders(oFolder){
	var subFlds = new Enumerator(oFolder.SubFolders);
	var s="";
    for (; !subFlds.atEnd(); subFlds.moveNext()) {
		var subFld =subFlds.item()
		s+=saveAs(subFld);
    }
	return s;
}

function saveAs(Folder){
	var fso = new ActiveXObject("Scripting.FileSystemObject");
	var oFolder;
	oFolder=fso.GetFolder(Folder);
	var s="";
	s+=showFiles(oFolder);
	s+=showSubFolders(oFolder);
	return s;
}
s=saveAs("E:\\workspace\\files");
WScript.Echo(s);