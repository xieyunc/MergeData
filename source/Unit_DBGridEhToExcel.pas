unit Unit_DBGridEhToExcel;

 

interface

uses

SysUtils, Variants, Classes, Graphics, Controls, Forms, Excel2000, ComObj,

Dialogs, DB, DBGridEh, windows,ComCtrls,ExtCtrls;

 

type

 

TDBGridEhToExcel = class(TComponent)

private

    FProgressForm: TForm;                                  {进度窗体}

    FtempGauge: TProgressBar;                           {进度条}

    FShowProgress: Boolean;                                {是否显示进度窗体}

    FShowOpenExcel:Boolean;                                {是否导出后打开Excel文件}

    FDBGridEh: TDBGridEh;

    FTitleName: TCaption;                                  {Excel文件标题}

    FUserName: TCaption;                                   {制表人}

    procedure SetShowProgress(const Value: Boolean);       {是否显示进度条}

    procedure SetShowOpenExcel(const Value: Boolean);      {是否打开生成的Excel文件}

    procedure SetDBGridEh(const Value: TDBGridEh);

    procedure SetTitleName(const Value: TCaption);         {标题名称}

    procedure SetUserName(const Value: TCaption);          {使用人名称}

    procedure CreateProcessForm(AOwner: TComponent);       {生成进度窗体}

public

    constructor Create(AOwner: TComponent); override;

    destructor Destroy; override;

    procedure ExportToExcel; {输出Excel文件}

published

    property DBGridEh: TDBGridEh read FDBGridEh write SetDBGridEh;

    property ShowProgress: Boolean read FShowProgress write SetShowProgress;    //是否显示进度条

    property ShowOpenExcel: Boolean read FShowOpenExcel write SetShowOpenExcel; //是否打开Excel

    property TitleName: TCaption read FTitleName write SetTitleName;

    property UserName: TCaption read FUserName write SetUserName;

end;

 

implementation

 

constructor TDBGridEhToExcel.Create(AOwner: TComponent);

begin

inherited Create(AOwner);

FShowProgress := True;

FShowOpenExcel:= True;

end;

 

procedure TDBGridEhToExcel.SetShowProgress(const Value: Boolean);

begin

FShowProgress := Value;

end;

 

procedure TDBGridEhToExcel.SetDBGridEh(const Value: TDBGridEh);

begin

FDBGridEh := Value;

end;

 

procedure TDBGridEhToExcel.SetTitleName(const Value: TCaption);

begin

FTitleName := Value;

end;

 

procedure TDBGridEhToExcel.SetUserName(const Value: TCaption);

begin

FUserName := Value;

end;

 

function IsFileInUse(fName: string ): boolean;

var

HFileRes: HFILE;

begin

Result :=false;

if not FileExists(fName) then exit;

HFileRes :=CreateFile(pchar(fName), GENERIC_READ

             or GENERIC_WRITE,0, nil,OPEN_EXISTING,FILE_ATTRIBUTE_NORMAL, 0);

Result :=(HFileRes=INVALID_HANDLE_VALUE);

if not Result then

    CloseHandle(HFileRes);

end;

 

procedure TDBGridEhToExcel.ExportToExcel;

var

XLApp: Variant;

Sheet: Variant;

s1, s2: string;

Caption,Msg: String;

Row, Col: integer;

iCount, jCount: Integer;

FBookMark: TBookmark;

FileName: String;

SaveDialog1: TSaveDialog;

begin

    //如果数据集为空或没有打开则退出

    if not DBGridEh.DataSource.DataSet.Active then Exit;

 

    SaveDialog1 := TSaveDialog.Create(Nil);

    SaveDialog1.FileName := TitleName + '_' + FormatDateTime('YYMMDDHHmmSS', now);

    SaveDialog1.Filter := 'Excel文件|*.xls';

    if SaveDialog1.Execute then

        FileName := SaveDialog1.FileName;

    SaveDialog1.Free;

    if FileName = '' then Exit;

 

    while IsFileInUse(FileName) do

    begin

      if Application.MessageBox('目标文件使用中，请退出目标文件后点击确定继续！',

        '注意', MB_OKCANCEL + MB_ICONWARNING) = IDOK then

      begin

 

      end

      else

      begin

        Exit;

      end;

    end;

 

    if FileExists(FileName) then

    begin

      Msg := '已存在文件（' + FileName + '），是否覆盖？';

      if Application.MessageBox(PChar(Msg), '提示', MB_YESNO + MB_ICONQUESTION + MB_DEFBUTTON2) = IDYES then

      begin

   //删除文件

        DeleteFile(PChar(FileName))

      end

      else

        exit;

    end;

    Application.ProcessMessages;

 

    Screen.Cursor := crHourGlass;

    //显示进度窗体

    if ShowProgress then

        CreateProcessForm(nil);

   

    if not VarIsEmpty(XLApp) then

    begin

        XLApp.DisplayAlerts := False;

        XLApp.Quit;

        VarClear(XLApp);

    end;

 

    //通过ole创建Excel对象

    try

        XLApp := CreateOleObject('Excel.Application');

    except

        MessageDlg('创建Excel对象失败，请检查你的系统是否正确安装了Excel软件！', mtError, [mbOk], 0);

        Screen.Cursor := crDefault;

        Exit;

    end;

 

    //生成工作页

    XLApp.WorkBooks.Add[XLWBatWorksheet];

    XLApp.WorkBooks[1].WorkSheets[1].Name := TitleName;

    Sheet := XLApp.Workbooks[1].WorkSheets[TitleName];

 

    //写标题

    sheet.cells[1, 1] := TitleName;

    sheet.range[sheet.cells[1, 1], sheet.cells[1, DBGridEh.Columns.Count]].Select; //选择该列

    XLApp.selection.HorizontalAlignment := $FFFFEFF4;                               //居中

    XLApp.selection.MergeCells := True;                                             //合并

 

    //写表头

    Row := 1;

    jCount := 3;

    for iCount := 0 to DBGridEh.Columns.Count - 1 do

    begin

        Col := 2;

        Row := iCount+1;

        Caption := DBGridEh.Columns[iCount].Title.Caption;

        while POS('|', Caption) > 0 do

        begin

            jCount := 4;

            s1 := Copy(Caption, 1, Pos('|',Caption)-1);

            if s2 = s1 then

            begin

                sheet.range[sheet.cells[Col, Row-1],sheet.cells[Col, Row]].Select;

                XLApp.selection.HorizontalAlignment := $FFFFEFF4;

                XLApp.selection.MergeCells := True;

            end

            else

                Sheet.cells[Col,Row] := Copy(Caption, 1, Pos('|',Caption)-1);

            Caption := Copy(Caption,Pos('|', Caption)+1, Length(Caption));

            Inc(Col);

            s2 := s1;

        end;

        Sheet.cells[Col, Row] := Caption;

        Inc(Row);

    end;

 

    //合并表头并居中

    if jCount = 4 then

        for iCount := 1 to DBGridEh.Columns.Count do

            if Sheet.cells[3, iCount].Value = '' then

            begin

                sheet.range[sheet.cells[2, iCount],sheet.cells[3, iCount]].Select;

                XLApp.selection.HorizontalAlignment := $FFFFEFF4;

                XLApp.selection.MergeCells := True;

            end

            else begin

                sheet.cells[3, iCount].Select;

                XLApp.selection.HorizontalAlignment := $FFFFEFF4;

            end;

 

    //读取数据

    DBGridEh.DataSource.DataSet.DisableControls;

    FBookMark := DBGridEh.DataSource.DataSet.GetBookmark;

    DBGridEh.DataSource.DataSet.First;

    while not DBGridEh.DataSource.DataSet.Eof do

    begin

 

        for iCount := 1 to DBGridEh.Columns.Count do

        begin

            //Sheet.cells[jCount, iCount] :=DBGridEh.Columns.Items[iCount-1].Field.AsString;

 

 

          case DBGridEh.DataSource.DataSet.FieldByName(DBGridEh.Columns.Items[iCount-1].FieldName).DataType of

            ftSmallint, ftInteger, ftWord, ftAutoInc, ftBytes:

              Sheet.cells[jCount, iCount] :=DBGridEh.Columns.Items[iCount-1].Field.asinteger;

            ftFloat, ftCurrency, ftBCD:

              Sheet.cells[jCount, iCount] :=DBGridEh.Columns.Items[iCount-1].Field.AsFloat;

          else

            if DBGridEh.DataSource.DataSet.FieldByName(DBGridEh.Columns.Items[iCount-1].FieldName) is TBlobfield then // 此类型的字段(图像等)暂无法读取显示

              Sheet.cells[jCount, iCount] :=DBGridEh.Columns.Items[iCount-1].Field.AsString

            else

              Sheet.cells[jCount, iCount] :=''''+DBGridEh.Columns.Items[iCount-1].Field.AsString;

          end;

         

        end;

        Inc(jCount);

 

        //显示进度条进度过程

        if ShowProgress then

        begin

            FtempGauge.Position := DBGridEh.DataSource.DataSet.RecNo;

            FtempGauge.Refresh;

        end;

 

        DBGridEh.DataSource.DataSet.Next;

    end;

    if DBGridEh.DataSource.DataSet.BookmarkValid(FBookMark) then

        DBGridEh.DataSource.DataSet.GotoBookmark(FBookMark);

    DBGridEh.DataSource.DataSet.EnableControls;

 

    //读取表脚

    if DBGridEh.FooterRowCount > 0 then

    begin

        for Row := 0 to DBGridEh.FooterRowCount-1 do

        begin

            for Col := 0 to DBGridEh.Columns.Count-1 do

                Sheet.cells[jCount, Col+1] := DBGridEh.GetFooterValue(Row,DBGridEh.Columns[Col]);

            Inc(jCount);

        end;

    end;

 

    //调整列宽

//    for iCount := 1 to DBGridEh.Columns.Count do

//        Sheet.Columns[iCount].EntireColumn.AutoFit;

 

    sheet.cells[1, 1].Select;

    XlApp.Workbooks[1].SaveAs(FileName);

 

    XlApp.Visible := True;

    XlApp := Unassigned;

 

    if ShowProgress then

        FreeAndNil(FProgressForm);

    Screen.Cursor := crDefault;

   

end;

 

destructor TDBGridEhToExcel.Destroy;

begin

inherited Destroy;

end;

 

procedure TDBGridEhToExcel.CreateProcessForm(AOwner: TComponent);

var

Panel: TPanel;

begin

if Assigned(FProgressForm) then

     exit;

 

FProgressForm := TForm.Create(AOwner);

with FProgressForm do

begin

    try

      Font.Name := '宋体';                                  {设置字体}

      Font.Size := 10;

      BorderStyle := bsNone;

      Width := 300;

      Height := 30;

      BorderWidth := 1;

      Color := clBlack;

      Position := poScreenCenter;

      Panel := TPanel.Create(FProgressForm);

      with Panel do

      begin

        Parent := FProgressForm;

        Align := alClient;

        Caption := '正在导出Excel，请稍候......';

        Color:=$00E9E5E0;

     end;

      FtempGauge:=TProgressBar.Create(Panel);

      with FtempGauge do

      begin

        Parent := Panel;

        Align:=alClient;

        Min := 0;

        Max:= DBGridEh.DataSource.DataSet.RecordCount;

        Position := 0;

      end;

    except

 

    end;

end;

FProgressForm.Show;

FProgressForm.Update;

end;

 

procedure TDBGridEhToExcel.SetShowOpenExcel(const Value: Boolean);

begin

   FShowOpenExcel:=Value;

end;

 

end.
