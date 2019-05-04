unit Unit_DBGridEhToExcel;

 

interface

uses

SysUtils, Variants, Classes, Graphics, Controls, Forms, Excel2000, ComObj,

Dialogs, DB, DBGridEh, windows,ComCtrls,ExtCtrls;

 

type

 

TDBGridEhToExcel = class(TComponent)

private

    FProgressForm: TForm;                                  {���ȴ���}

    FtempGauge: TProgressBar;                           {������}

    FShowProgress: Boolean;                                {�Ƿ���ʾ���ȴ���}

    FShowOpenExcel:Boolean;                                {�Ƿ񵼳����Excel�ļ�}

    FDBGridEh: TDBGridEh;

    FTitleName: TCaption;                                  {Excel�ļ�����}

    FUserName: TCaption;                                   {�Ʊ���}

    procedure SetShowProgress(const Value: Boolean);       {�Ƿ���ʾ������}

    procedure SetShowOpenExcel(const Value: Boolean);      {�Ƿ�����ɵ�Excel�ļ�}

    procedure SetDBGridEh(const Value: TDBGridEh);

    procedure SetTitleName(const Value: TCaption);         {��������}

    procedure SetUserName(const Value: TCaption);          {ʹ��������}

    procedure CreateProcessForm(AOwner: TComponent);       {���ɽ��ȴ���}

public

    constructor Create(AOwner: TComponent); override;

    destructor Destroy; override;

    procedure ExportToExcel; {���Excel�ļ�}

published

    property DBGridEh: TDBGridEh read FDBGridEh write SetDBGridEh;

    property ShowProgress: Boolean read FShowProgress write SetShowProgress;    //�Ƿ���ʾ������

    property ShowOpenExcel: Boolean read FShowOpenExcel write SetShowOpenExcel; //�Ƿ��Excel

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

    //������ݼ�Ϊ�ջ�û�д����˳�

    if not DBGridEh.DataSource.DataSet.Active then Exit;

 

    SaveDialog1 := TSaveDialog.Create(Nil);

    SaveDialog1.FileName := TitleName + '_' + FormatDateTime('YYMMDDHHmmSS', now);

    SaveDialog1.Filter := 'Excel�ļ�|*.xls';

    if SaveDialog1.Execute then

        FileName := SaveDialog1.FileName;

    SaveDialog1.Free;

    if FileName = '' then Exit;

 

    while IsFileInUse(FileName) do

    begin

      if Application.MessageBox('Ŀ���ļ�ʹ���У����˳�Ŀ���ļ�����ȷ��������',

        'ע��', MB_OKCANCEL + MB_ICONWARNING) = IDOK then

      begin

 

      end

      else

      begin

        Exit;

      end;

    end;

 

    if FileExists(FileName) then

    begin

      Msg := '�Ѵ����ļ���' + FileName + '�����Ƿ񸲸ǣ�';

      if Application.MessageBox(PChar(Msg), '��ʾ', MB_YESNO + MB_ICONQUESTION + MB_DEFBUTTON2) = IDYES then

      begin

   //ɾ���ļ�

        DeleteFile(PChar(FileName))

      end

      else

        exit;

    end;

    Application.ProcessMessages;

 

    Screen.Cursor := crHourGlass;

    //��ʾ���ȴ���

    if ShowProgress then

        CreateProcessForm(nil);

   

    if not VarIsEmpty(XLApp) then

    begin

        XLApp.DisplayAlerts := False;

        XLApp.Quit;

        VarClear(XLApp);

    end;

 

    //ͨ��ole����Excel����

    try

        XLApp := CreateOleObject('Excel.Application');

    except

        MessageDlg('����Excel����ʧ�ܣ��������ϵͳ�Ƿ���ȷ��װ��Excel�����', mtError, [mbOk], 0);

        Screen.Cursor := crDefault;

        Exit;

    end;

 

    //���ɹ���ҳ

    XLApp.WorkBooks.Add[XLWBatWorksheet];

    XLApp.WorkBooks[1].WorkSheets[1].Name := TitleName;

    Sheet := XLApp.Workbooks[1].WorkSheets[TitleName];

 

    //д����

    sheet.cells[1, 1] := TitleName;

    sheet.range[sheet.cells[1, 1], sheet.cells[1, DBGridEh.Columns.Count]].Select; //ѡ�����

    XLApp.selection.HorizontalAlignment := $FFFFEFF4;                               //����

    XLApp.selection.MergeCells := True;                                             //�ϲ�

 

    //д��ͷ

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

 

    //�ϲ���ͷ������

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

 

    //��ȡ����

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

            if DBGridEh.DataSource.DataSet.FieldByName(DBGridEh.Columns.Items[iCount-1].FieldName) is TBlobfield then // �����͵��ֶ�(ͼ���)���޷���ȡ��ʾ

              Sheet.cells[jCount, iCount] :=DBGridEh.Columns.Items[iCount-1].Field.AsString

            else

              Sheet.cells[jCount, iCount] :=''''+DBGridEh.Columns.Items[iCount-1].Field.AsString;

          end;

         

        end;

        Inc(jCount);

 

        //��ʾ���������ȹ���

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

 

    //��ȡ���

    if DBGridEh.FooterRowCount > 0 then

    begin

        for Row := 0 to DBGridEh.FooterRowCount-1 do

        begin

            for Col := 0 to DBGridEh.Columns.Count-1 do

                Sheet.cells[jCount, Col+1] := DBGridEh.GetFooterValue(Row,DBGridEh.Columns[Col]);

            Inc(jCount);

        end;

    end;

 

    //�����п�

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

      Font.Name := '����';                                  {��������}

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

        Caption := '���ڵ���Excel�����Ժ�......';

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
