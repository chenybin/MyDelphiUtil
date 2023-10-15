unit uExcelUtils;

interface
uses
    ExcelXP, StrUtils, SysUtils, Variants, DBGridEh, Dialogs, DBGridEhImpExp,  EhlibADO, Forms, Windows;

function ProcessText(AText: string): string;
function GetWorkSheetValue(AWS: TExcelWorksheet; ARow, ACol: Integer): string;overload;
function GetWorkSheetValue(AWS: TExcelWorksheet; ARow: Integer; AColumnName: string): string;overload;
function OpenExcelFile(AExcelApp: TExcelApplication;
    AWorkSheet: TExcelWorkSheet; AFileName: string): Boolean;
function  DBGridEhToExportFile(dbgrideh:TDBGridEh; filename:string='导出的文件';AIsSaveAll: Boolean = True):string ;
implementation

// 去掉文字中的回车换行等信息
function ProcessText(AText: string): string;
begin
    Result := stringreplace(AText, '''', '',
        [rfReplaceAll, rfIgnoreCase]);
    Result := stringreplace(Result, #13, ' ',
        [rfReplaceAll, rfIgnoreCase]);
    Result := stringreplace(Result, #10, ' ',
        [rfReplaceAll, rfIgnoreCase]);
    Result := stringreplace(Result, #9, ' ',
        [rfReplaceAll, rfIgnoreCase]);
end;

// 获取Excel对象中某一个单元格的值
function GetWorkSheetValue(AWS: TExcelWorksheet; ARow, ACol: Integer): string; overload;
begin
    Result := ProcessText(AWS.Cells.Item[ARow, ACol].value);
end;

// 获取Excel对象中某一个单元格的值
function GetWorkSheetValue(AWS: TExcelWorksheet; ARow: Integer; AColumnName: string): string;overload;
var
  ACol: Integer;
begin
    Result := ProcessText(AWS.Cells.Item[ARow, ACol].value);
end;

// 打开一个Excel文件
function OpenExcelFile(AExcelApp: TExcelApplication;
    AWorkSheet: TExcelWorkSheet; AFileName: string): Boolean;

begin
    Result := False;
    if not FileExists(AFileName) then Exit;

    AExcelApp.Workbooks.Open(AFileName, EmptyParam, EmptyParam, EmptyParam, EmptyParam, EmptyParam, EmptyParam
        , EmptyParam, EmptyParam, EmptyParam, EmptyParam, EmptyParam, EmptyParam, EmptyParam, EmptyParam, 0);
    (AExcelApp.Worksheets.Item[1] as _WorkSheet).Activate(0);
    AWorkSheet.ConnectTo(AExcelApp.ActiveSheet as _WorkSheet);

    Result := True;
end;

// 导出DBGridEh表格到Excel中，
function  DBGridEhToExportFile(dbgrideh:TDBGridEh; filename:string;AIsSaveAll: Boolean):string ;
var ExpClass:TDBGridEhExportClass;
Ext:String;
sd1:tsavedialog;
begin
result:='' ;
sd1:=tsavedialog.Create(nil);
   sd1.Filter:='Excel 文件(*.xlsx)|*.xlsx|Excel 文件(*.xls)|*.xls|分隔符格式(*.csv)|*.csv|Html文件(*.htm)|*.htm|WORD 文件(*.rtf)|*.rtf|文本文件(*.txt)|*.txt';
   sd1.FileName:=filename;
 if (dbgrideh is TDBGridEh) then
  if sd1.Execute then
  begin
   case sd1.FilterIndex of
   1: begin ExpClass := TDBGridEhExportAsXLSX; Ext := 'xlsx'; end;
    2: begin ExpClass := TDBGridEhExportAsXLS; Ext := 'xls'; end;
    3: begin ExpClass := TDBGridEhExportAsCSV; Ext := 'csv'; end;
    4: begin ExpClass := TDBGridEhExportAsHTML; Ext := 'htm'; end;
    5: begin ExpClass := TDBGridEhExportAsRTF; Ext := 'rtf'; end;
    6: begin ExpClass := TDBGridEhExportAsText; Ext := 'txt'; end;
   else
    ExpClass := nil; Ext := '';
   end;
   if ExpClass <> nil then
   begin
    if UpperCase(Copy(sd1.FileName,Length(sd1.FileName)-2,3)) <>
      UpperCase(Ext) then
     sd1.FileName := sd1.FileName + '.' + Ext;
   if FileExists( sd1.FileName) then
    begin
    if application.MessageBox('文件已存在,替换?','提示',mb_yesno+mb_defbutton1+mb_iconquestion+mb_systemmodal)=idyes then
     begin
      if  DeleteFile( PWideChar(sd1.FileName))=false then
      begin
        showmessage(filename+'文件正在使用,无法替换.'+chr(13)+chr(10)+'请关闭文件：'+sd1.FileName+'.在重新导入。');
      result:='';
        exit;
      end;
    end;
    end;
    SaveDBGridEhToExportFile(ExpClass,TDBGridEh(dbgrideh),  sd1.FileName,AIsSaveAll);  //改为false 只导出选择行
    result := sd1.FileName
   end;
  end
  else
  begin
   result:='';
  end;
 sd1.Free ;
end;

end.
