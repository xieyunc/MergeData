unit uMain;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, DBGridEhImpExp, 
  Controls, Forms, Dialogs, StdCtrls, ExtCtrls, Grids, CnProgressFrm,
  DBGrids, ComCtrls, DB, ADODB, Buttons, DBGridEhGrouping, GridsEh, DBGridEh,
  Menus,EhLibADO;

type
  TMain = class(TForm)
    Conn_SQL: TADOConnection;
    qry_SQL: TADOQuery;
    Conn_XLS: TADOConnection;
    qry_XLS: TADOQuery;
    DS_SQL: TDataSource;
    PageControl1: TPageControl;
    TabSheet1: TTabSheet;
    DS_XLS: TDataSource;
    OpenDialog1: TOpenDialog;
    TabSheet2: TTabSheet;
    dbgrdh_SQL: TDBGridEh;
    Panel2: TPanel;
    Label3: TLabel;
    btn_StartMerge: TButton;
    cbb_d_checkFld: TComboBox;
    dbgrdh_XLS: TDBGridEh;
    Panel1: TPanel;
    Label1: TLabel;
    SpeedButton1: TSpeedButton;
    btn_s_Open: TButton;
    Edit1: TEdit;
    cbb_s_Table: TComboBox;
    Label4: TLabel;
    cbb_s_qymcFld: TComboBox;
    btn_D_Open: TButton;
    lbl1: TLabel;
    cbb_d_qymcFld: TComboBox;
    lbl2: TLabel;
    cbb_s_checkFld: TComboBox;
    pm1: TPopupMenu;
    mniExcel1: TMenuItem;
    lbl3: TLabel;
    edt_Conn: TEdit;
    procedure SpeedButton1Click(Sender: TObject);
    procedure btn_s_OpenClick(Sender: TObject);
    procedure btn_StartMergeClick(Sender: TObject);
    procedure qry_SQLAfterOpen(DataSet: TDataSet);
    procedure qry_XLSAfterOpen(DataSet: TDataSet);
    procedure btn_D_OpenClick(Sender: TObject);
    procedure cbb_s_TableChange(Sender: TObject);
    procedure cbb_s_qymcFldChange(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure mniExcel1Click(Sender: TObject);
    procedure cbb_d_checkFldChange(Sender: TObject);
  private
    { Private declarations }
    procedure OutToFile(IADO : TADOQuery; DgEh : TDBGridEh);
  public
    { Public declarations }
  end;

var
  Main: TMain;

implementation
uses Unit_DBGridEhToExcel;
{$R *.dfm}

procedure TMain.btn_s_OpenClick(Sender: TObject);
var
  fn,fld1,fld2,sqlstr:string;
begin
  fn := Edit1.Text;
  fld1 := cbb_s_qymcFld.Text;
  fld2 := cbb_s_checkFld.Text;
  Screen.Cursor := crHourGlass;
  try
    Conn_XLS.Close;
    Conn_XLS.ConnectionString := 'Provider=Microsoft.Jet.OLEDB.4.0;Data Source='+fn+';Extended Properties=EXCEL 8.0;Persist Security Info=False';
    Conn_XLS.Connected := True;
    with qry_XLS do
    begin
      Close;
      sqlstr := 'select trim('+fld1+') as '+fld1+','+fld2+' from ['+cbb_s_Table.Text+'] '+
                'group by trim('+fld1+'),'+fld2;
      SQL.Text := sqlstr;
      Open;
    end;
  finally
    Screen.Cursor := crDefault;
  end;
end;

procedure TMain.cbb_d_checkFldChange(Sender: TObject);
begin
  btn_StartMerge.Enabled := cbb_d_checkFld.Text<>'';
end;

procedure TMain.cbb_s_qymcFldChange(Sender: TObject);
begin
  btn_s_Open.Enabled := (cbb_s_Table.Text<>'') and (cbb_s_qymcFld.Text<>'');
end;

procedure TMain.cbb_s_TableChange(Sender: TObject);
begin
  cbb_s_qymcFld.Items.Clear;
  cbb_s_qymcFld.Text := '';
  cbb_s_checkFld.Items.Clear;
  cbb_s_checkFld.Text := '';
  if cbb_s_Table.Text<>'' then
  begin
    Conn_XLS.GetFieldNames(cbb_s_Table.Text,cbb_s_qymcFld.ITems);
    Conn_XLS.GetFieldNames(cbb_s_Table.Text,cbb_s_checkFld.ITems);
  end;
end;

procedure TMain.FormCreate(Sender: TObject);
begin
  PageControl1.ActivePageIndex := 0;
end;

procedure TMain.mniExcel1Click(Sender: TObject);
begin
  OutToFile(qry_SQL,dbgrdh_SQL);
end;
{
var
   GridtoExcel: TDBGridEhToExcel;
begin
   try
     GridtoExcel := TDBGridEhToExcel.Create(nil);
     GridtoExcel.DBGridEh := dbgrdh_SQL;
     GridtoExcel.TitleName := '处理结果';
     GridtoExcel.ShowProgress := true;
     GridtoExcel.ShowOpenExcel := true;
     GridtoExcel.ExportToExcel;
   finally
     GridtoExcel.Free;
   end;
end;
}
procedure TMain.OutToFile(IADO: TADOQuery; DgEh: TDBGridEh);
var
  ExpClass:TDBGridEhExportclass;
  Ext:String;
  FSaveDialog: TSaveDialog;
begin
  try
    if not IADO.IsEmpty then
    begin
      FSaveDialog := TSaveDialog.Create(Self);
      FSaveDialog.Filter:='Excel 文档 (*.xls)|*.XLS|Text files (*.txt)|*.TXT|Comma separated values (*.csv)|*.CSV|HTML file (*.htm)|*.HTM|Word 文档 (*.rtf)|*.RTF';
      if FSaveDialog.Execute and (trim(FSaveDialog.FileName)<>'') then
      begin
        case FSaveDialog.FilterIndex of
            1: begin ExpClass := TDBGridEhExportAsXLS; Ext := 'xls'; end;
            2: begin ExpClass := TDBGridEhExportAsText; Ext := 'txt'; end;
            3: begin ExpClass := TDBGridEhExportAsCSV; Ext := 'csv'; end;
            4: begin ExpClass := TDBGridEhExportAsHTML; Ext := 'htm'; end;
            5: begin ExpClass := TDBGridEhExportAsRTF; Ext := 'rtf'; end;
        end;
        if ExpClass <> nil then
        begin
          if UpperCase(Copy(FSaveDialog.FileName,Length(FSaveDialog.FileName)-2,3)) <> UpperCase(Ext) then
            FSaveDialog.FileName := FSaveDialog.FileName + '.' + Ext;
            if FileExists(FSaveDialog.FileName) then
            begin
              if application.MessageBox('文件名已存在，是否覆盖   ', '提示', MB_ICONASTERISK or MB_OKCANCEL)<>idok then
                exit;
            end;
           Screen.Cursor := crHourGlass;
           SaveDBGridEhToExportFile(ExpClass,DgEh,FSaveDialog.FileName,true);
           Screen.Cursor := crDefault;
           MessageBox(Handle, '导出成功  ', '提示', MB_OK +
             MB_ICONINFORMATION);
          end;
      end;
      FSaveDialog.Destroy;
    end;
  except
    on e: exception do
    begin
      Application.MessageBox(PChar(e.message), '错误', MB_OK + MB_ICONSTOP);
    end;
  end;                       
end;

procedure TMain.btn_D_OpenClick(Sender: TObject);
var
  fld1,fld2:string;
begin
  Screen.Cursor := crHourGlass;
  Conn_SQL.Close;
  Conn_SQL.ConnectionString := edt_Conn.Text;
  fld1 := cbb_d_qymcFld.Text;
  fld2 := cbb_d_checkFld.Text;
  try
    qry_SQL.Close;
    qry_SQL.SQL.Text := 'select * from nn';
    //qry_SQL.SQL.Text := 'select '+fld1+','+fld2+' from 处理结果';
    qry_SQL.Open;

    cbb_d_checkFld.Items.Clear;
    cbb_d_qymcFld.Items.Clear;
    Conn_SQL.GetFieldNames('nn',cbb_d_qymcFld.Items);
    cbb_d_qymcFld.Text := '组织机构代码';//'企业名称';
    Conn_SQL.GetFieldNames('nn',cbb_d_checkFld.Items);

  finally
    Screen.Cursor := crDefault;
  end;
end;

procedure TMain.btn_StartMergeClick(Sender: TObject);
var
  s_qymcFld,d_qymcFld,s_checkFld,d_checkFld:string;
  d_qymc,s_qymc,d_check,s_check:string;
  InsertCount,UpdateCount:Integer;
begin
  btn_StartMerge.Enabled := False;
  s_qymcFld := cbb_s_qymcFld.Text;
  d_qymcFld := cbb_d_qymcFld.Text;//'企业名称';
  s_checkFld := cbb_s_checkFld.Text;
  d_CheckFld := cbb_d_checkFld.Text;
  Screen.Cursor := crHourGlass;
  qry_SQL.DisableControls;
  qry_XLS.EnableControls;
  qry_XLS.First;
  Conn_SQL.BeginTrans;
  try
    try
      InsertCount := 0;
      UpdateCount := 0;
      ShowProgress('正在处理，请稍候……',qry_XLS.RecordCount);
      while not qry_XLS.Eof do
      begin
        UpdateProgress(qry_XLS.RecNo);
        Application.ProcessMessages;
        s_qymc := Trim(qry_XLS.FieldByName(s_qymcFld).AsString);
        s_check := Trim(qry_XLS.FieldByName(s_checkFld).AsString);
        if s_qymc='' then
        begin
          qry_XLS.Next;
          Continue;
        end;

        if not qry_SQL.Locate(d_qymcFld,s_qymc,[]) then
        begin
          qry_XLS.Next;
          Continue;
          //qry_SQL.Append;
          //Inc(InsertCount);
        end
        else
        begin
          qry_SQL.Edit;
          //Inc(UpdateCount);
        end;

        d_qymc := Trim(qry_SQL.FieldByName(d_qymcFld).AsString);
        d_check := Trim(qry_SQL.FieldByName(d_checkFld).AsString);
{
        if d_qymc<>s_qymc then
          qry_SQL.FieldByName(d_qymcFld).Value := s_qymc;
}
//        if Pos(','+s_check+',',','+d_check+',')=0 then
          if d_check='' then
          begin
            qry_SQL.FieldByName(d_checkFld).Value := s_check;
            Inc(UpdateCount);
          end;
//          else if s_check<>'' then
//            qry_SQL.FieldByName(d_checkFld).Value := d_check+','+s_check;


        qry_SQL.Post;

        qry_XLS.Next;
      end;
      Conn_SQL.CommitTrans;
      HideProgress;
      Application.MessageBox(pChar(Format('处理完成！本次共新记录：%d条，更新记录：%d条',[InsertCount,UpdateCount])),'系统提示',MB_OK+MB_ICONERROR);
    except
      on e:Exception do
      begin
        Conn_SQL.RollbackTrans;
        Application.MessageBox(pChar('记录号：'+IntToStr(qry_XLS.RecNo)+'处理失败！'+#13+'失败原因为：'+e.Message),'系统提示',MB_OK+MB_ICONERROR);
      end;
    end;
  finally
    btn_StartMerge.Enabled := True;
    HideProgress;
    qry_SQL.EnableControls;
    qry_XLS.EnableControls;
    Screen.Cursor := crDefault;
  end;
end;

procedure TMain.qry_SQLAfterOpen(DataSet: TDataSet);
var
  i: Integer;
begin
  for i := 0 to dbgrdh_SQL.Columns.Count-1 do
  begin
    dbgrdh_SQL.Columns[i].Title.TitleButton := True;
    if dbgrdh_SQL.Columns[i].Width>200 then
      dbgrdh_SQL.Columns[i].Width := 200
    else if dbgrdh_SQL.Columns[i].Width<50 then
      dbgrdh_SQL.Columns[i].Width := 50;
  end;
end;

procedure TMain.qry_XLSAfterOpen(DataSet: TDataSet);
var
  i: Integer;
begin
  for i := 0 to dbgrdh_XLS.Columns.Count-1 do
    if dbgrdh_XLS.Columns[i].Width>200 then
      dbgrdh_XLS.Columns[i].Width := 200;
end;

procedure TMain.SpeedButton1Click(Sender: TObject);
var
  fn:string;
begin
  if OpenDialog1.Execute then
  begin
    Edit1.Text := OpenDialog1.FileName;
    fn := Edit1.Text;
    Screen.Cursor := crHourGlass;
    try
      Conn_XLS.Close;
      Conn_XLS.ConnectionString := 'Provider=Microsoft.Jet.OLEDB.4.0;Data Source='+fn+';Extended Properties=EXCEL 8.0;Persist Security Info=False';
      Conn_XLS.Connected := True;
      Conn_XLS.GetTableNames(cbb_s_Table.Items);
    finally
      Screen.Cursor := crDefault;
    end;
  end;
end;

end.
