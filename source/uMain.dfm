object Main: TMain
  Left = 0
  Top = 0
  Caption = #25968#25454#21512#24182#22788#29702
  ClientHeight = 482
  ClientWidth = 652
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -12
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  Position = poScreenCenter
  OnCreate = FormCreate
  PixelsPerInch = 96
  TextHeight = 14
  object PageControl1: TPageControl
    Left = 0
    Top = 0
    Width = 652
    Height = 482
    ActivePage = TabSheet2
    Align = alClient
    TabOrder = 0
    object TabSheet1: TTabSheet
      Caption = 'Excel'#21407#22987#25968#25454
      object dbgrdh_XLS: TDBGridEh
        Left = 0
        Top = 0
        Width = 644
        Height = 391
        Align = alClient
        DataGrouping.GroupLevels = <>
        DataSource = DS_XLS
        Flat = False
        FooterColor = clWindow
        FooterFont.Charset = DEFAULT_CHARSET
        FooterFont.Color = clWindowText
        FooterFont.Height = -12
        FooterFont.Name = 'Tahoma'
        FooterFont.Style = []
        OptionsEh = [dghFixed3D, dghHighlightFocus, dghClearSelection, dghRowHighlight, dghDialogFind, dghShowRecNo, dghColumnResize, dghColumnMove]
        RowDetailPanel.Color = clBtnFace
        TabOrder = 0
        TitleFont.Charset = DEFAULT_CHARSET
        TitleFont.Color = clWindowText
        TitleFont.Height = -12
        TitleFont.Name = 'Tahoma'
        TitleFont.Style = []
        object RowDetailData: TRowDetailPanelControlEh
        end
      end
      object Panel1: TPanel
        Left = 0
        Top = 391
        Width = 644
        Height = 62
        Align = alBottom
        TabOrder = 1
        object Label1: TLabel
          Left = 8
          Top = 9
          Width = 88
          Height = 14
          Caption = #36873#25321'Excel'#25991#20214#65306
        end
        object SpeedButton1: TSpeedButton
          Left = 336
          Top = 7
          Width = 60
          Height = 22
          Caption = #36873#25321'Excel'
          OnClick = SpeedButton1Click
        end
        object Label4: TLabel
          Left = 9
          Top = 36
          Width = 84
          Height = 14
          Caption = #20027#20851#38190#23383#23383#27573#65306
        end
        object lbl2: TLabel
          Left = 340
          Top = 36
          Width = 60
          Height = 14
          Caption = #26356#26032#23383#27573#65306
        end
        object btn_s_Open: TButton
          Left = 560
          Top = 5
          Width = 75
          Height = 52
          Caption = #25171#24320#34920
          Enabled = False
          TabOrder = 0
          OnClick = btn_s_OpenClick
        end
        object Edit1: TEdit
          Left = 94
          Top = 7
          Width = 233
          Height = 22
          TabOrder = 1
        end
        object cbb_s_Table: TComboBox
          Left = 404
          Top = 7
          Width = 145
          Height = 22
          ItemHeight = 14
          TabOrder = 2
          OnChange = cbb_s_TableChange
        end
        object cbb_s_qymcFld: TComboBox
          Left = 94
          Top = 34
          Width = 233
          Height = 22
          ItemHeight = 14
          TabOrder = 3
          OnChange = cbb_s_qymcFldChange
        end
        object cbb_s_checkFld: TComboBox
          Left = 404
          Top = 34
          Width = 145
          Height = 22
          ItemHeight = 14
          TabOrder = 4
        end
      end
    end
    object TabSheet2: TTabSheet
      Caption = #22788#29702#32467#26524
      ImageIndex = 2
      object dbgrdh_SQL: TDBGridEh
        Left = 0
        Top = 0
        Width = 644
        Height = 376
        Align = alClient
        DataGrouping.GroupLevels = <>
        DataSource = DS_SQL
        Flat = False
        FooterColor = clWindow
        FooterFont.Charset = DEFAULT_CHARSET
        FooterFont.Color = clWindowText
        FooterFont.Height = -12
        FooterFont.Name = 'Tahoma'
        FooterFont.Style = []
        IndicatorTitle.ShowDropDownSign = True
        IndicatorTitle.TitleButton = True
        OptionsEh = [dghFixed3D, dghHighlightFocus, dghClearSelection, dghAutoSortMarking, dghMultiSortMarking, dghRowHighlight, dghDialogFind, dghShowRecNo, dghColumnResize, dghColumnMove]
        PopupMenu = pm1
        RowDetailPanel.Color = clBtnFace
        SortLocal = True
        TabOrder = 0
        TitleFont.Charset = DEFAULT_CHARSET
        TitleFont.Color = clWindowText
        TitleFont.Height = -12
        TitleFont.Name = 'Tahoma'
        TitleFont.Style = []
        object RowDetailData: TRowDetailPanelControlEh
        end
      end
      object Panel2: TPanel
        Left = 0
        Top = 376
        Width = 644
        Height = 77
        Align = alBottom
        TabOrder = 1
        DesignSize = (
          644
          77)
        object Label3: TLabel
          Left = 320
          Top = 47
          Width = 60
          Height = 14
          Caption = #26356#26032#23383#27573#65306
        end
        object lbl1: TLabel
          Left = 9
          Top = 46
          Width = 84
          Height = 14
          Caption = #20027#20851#38190#23383#23383#27573#65306
        end
        object lbl3: TLabel
          Left = 8
          Top = 18
          Width = 84
          Height = 14
          Caption = #25968#25454#24211#36830#25509#20018#65306
        end
        object btn_StartMerge: TButton
          Left = 553
          Top = 43
          Width = 75
          Height = 25
          Caption = #24320#22987#21512#24182
          Enabled = False
          TabOrder = 2
          OnClick = btn_StartMergeClick
        end
        object cbb_d_checkFld: TComboBox
          Left = 381
          Top = 45
          Width = 161
          Height = 22
          ItemHeight = 14
          TabOrder = 4
          OnChange = cbb_d_checkFldChange
        end
        object btn_D_Open: TButton
          Left = 553
          Top = 11
          Width = 75
          Height = 25
          Caption = #25171#24320#34920
          TabOrder = 0
          OnClick = btn_D_OpenClick
        end
        object cbb_d_qymcFld: TComboBox
          Left = 102
          Top = 45
          Width = 207
          Height = 22
          ItemHeight = 14
          TabOrder = 3
          Text = #20225#19994#21517#31216
        end
        object edt_Conn: TEdit
          Left = 102
          Top = 15
          Width = 440
          Height = 22
          Anchors = [akLeft, akTop, akRight]
          TabOrder = 1
          Text = 
            'Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security In' +
            'fo=False;Initial Catalog=mergedata;Data Source=127.0.0.1,9105;'
        end
      end
    end
  end
  object Conn_SQL: TADOConnection
    ConnectionString = 
      'Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security In' +
      'fo=False;Initial Catalog=EnterpriseAnalysis;Data Source=127.0.0.' +
      '1,8829'
    LoginPrompt = False
    Provider = 'SQLOLEDB.1'
    Left = 488
    Top = 64
  end
  object qry_SQL: TADOQuery
    CacheSize = 1000
    Connection = Conn_SQL
    CursorType = ctStatic
    AfterOpen = qry_SQLAfterOpen
    Parameters = <>
    SQL.Strings = (
      'select * from '#22788#29702#32467#26524)
    Left = 536
    Top = 64
  end
  object Conn_XLS: TADOConnection
    ConnectionString = 
      'Provider=Microsoft.Jet.OLEDB.4.0;Data Source=H:\'#21335#23665'\2014-3-3\3\20' +
      '13+'#28145#22323#28779#28844#32479#35745'.xls;Extended Properties=EXCEL 8.0;Persist Security Inf' +
      'o=False'
    LoginPrompt = False
    Provider = 'Microsoft.Jet.OLEDB.4.0'
    Left = 155
    Top = 384
  end
  object qry_XLS: TADOQuery
    Connection = Conn_XLS
    AfterOpen = qry_XLSAfterOpen
    Parameters = <>
    Left = 209
    Top = 384
  end
  object DS_SQL: TDataSource
    DataSet = qry_SQL
    Left = 584
    Top = 64
  end
  object DS_XLS: TDataSource
    DataSet = qry_XLS
    Left = 256
    Top = 384
  end
  object OpenDialog1: TOpenDialog
    DefaultExt = '.xls'
    Filter = 'Excel 97-2003'#25991#20214'(*.xls)|*.xls|'#25152#26377#25991#20214'(*.*)|*.*'
    Title = #25171#24320'Excel'#25991#20214
    Left = 123
    Top = 384
  end
  object pm1: TPopupMenu
    Left = 312
    Top = 176
    object mniExcel1: TMenuItem
      Caption = #23548#20986#21040'Excel'#25991#20214
      OnClick = mniExcel1Click
    end
  end
end
