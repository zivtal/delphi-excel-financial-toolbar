object CurrencyRateSet: TCurrencyRateSet
  Left = 0
  Top = 0
  BorderIcons = [biSystemMenu]
  BorderStyle = bsDialog
  Caption = 'Currency Rate Settings'
  ClientHeight = 404
  ClientWidth = 763
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  Position = poDesktopCenter
  OnCreate = FormCreate
  OnKeyDown = FormKeyDown
  PixelsPerInch = 96
  TextHeight = 13
  object Label1: TLabel
    Left = 8
    Top = 12
    Width = 56
    Height = 13
    Caption = 'Bank name:'
  end
  object Label2: TLabel
    Left = 8
    Top = 39
    Width = 50
    Height = 13
    Caption = 'URL Mask:'
  end
  object Label3: TLabel
    Left = 8
    Top = 66
    Width = 95
    Height = 13
    Caption = 'URL Currencies'#39' list:'
  end
  object BANKNAME: TComboBox
    Left = 112
    Top = 10
    Width = 237
    Height = 21
    TabOrder = 0
    OnChange = BANKNAMEChange
    OnKeyDown = FormKeyDown
  end
  object Save: TButton
    Left = 355
    Top = 8
    Width = 75
    Height = 25
    Caption = 'Add'
    Enabled = False
    TabOrder = 11
    OnClick = SaveClick
    OnKeyDown = FormKeyDown
  end
  object Remove: TButton
    Left = 436
    Top = 8
    Width = 75
    Height = 25
    Caption = 'Remove'
    Enabled = False
    TabOrder = 12
    OnClick = RemoveClick
    OnKeyDown = FormKeyDown
  end
  object Export: TButton
    Left = 517
    Top = 8
    Width = 75
    Height = 25
    Caption = 'Export'
    TabOrder = 13
    OnClick = ExportClick
    OnKeyDown = FormKeyDown
  end
  object Import: TButton
    Left = 598
    Top = 8
    Width = 75
    Height = 25
    Caption = 'Import'
    TabOrder = 14
    OnClick = ImportClick
    OnKeyDown = FormKeyDown
  end
  object MASKURL: TEdit
    Left = 112
    Top = 37
    Width = 642
    Height = 21
    TabOrder = 1
    OnChange = CheckChanges
    OnKeyDown = FormKeyDown
  end
  object LISTURL: TEdit
    Left = 112
    Top = 64
    Width = 642
    Height = 21
    TabOrder = 2
    OnChange = LISTURLChange
    OnKeyDown = FormKeyDown
  end
  object GroupBox1: TGroupBox
    Left = 8
    Top = 104
    Width = 369
    Height = 73
    Caption = 'Main header:'
    TabOrder = 3
    object Label4: TLabel
      Left = 12
      Top = 22
      Width = 23
      Height = 11
      Caption = 'Start:'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -9
      Font.Name = 'Tahoma'
      Font.Style = []
      ParentFont = False
    end
    object Label5: TLabel
      Left = 190
      Top = 22
      Width = 19
      Height = 11
      Caption = 'End:'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -9
      Font.Name = 'Tahoma'
      Font.Style = []
      ParentFont = False
    end
    object SOM: TEdit
      Left = 12
      Top = 35
      Width = 170
      Height = 21
      TabOrder = 0
      OnChange = CheckChanges
      OnKeyDown = FormKeyDown
    end
    object EOM: TEdit
      Left = 190
      Top = 35
      Width = 170
      Height = 21
      TabOrder = 1
      OnChange = CheckChanges
      OnKeyDown = FormKeyDown
    end
  end
  object GroupBox2: TGroupBox
    Left = 8
    Top = 183
    Width = 369
    Height = 73
    Caption = 'Sub header:'
    TabOrder = 4
    object Label6: TLabel
      Left = 12
      Top = 22
      Width = 23
      Height = 11
      Caption = 'Start:'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -9
      Font.Name = 'Tahoma'
      Font.Style = []
      ParentFont = False
    end
    object Label7: TLabel
      Left = 190
      Top = 22
      Width = 19
      Height = 11
      Caption = 'End:'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -9
      Font.Name = 'Tahoma'
      Font.Style = []
      ParentFont = False
    end
    object SOS: TEdit
      Left = 12
      Top = 35
      Width = 170
      Height = 21
      TabOrder = 0
      OnChange = CheckChanges
      OnKeyDown = FormKeyDown
    end
    object EOS: TEdit
      Left = 190
      Top = 35
      Width = 170
      Height = 21
      TabOrder = 1
      OnChange = CheckChanges
      OnKeyDown = FormKeyDown
    end
  end
  object GroupBox3: TGroupBox
    Left = 8
    Top = 262
    Width = 369
    Height = 73
    Caption = 'Currency code header:'
    TabOrder = 5
    object Label8: TLabel
      Left = 12
      Top = 22
      Width = 23
      Height = 11
      Caption = 'Start:'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -9
      Font.Name = 'Tahoma'
      Font.Style = []
      ParentFont = False
    end
    object Label9: TLabel
      Left = 190
      Top = 22
      Width = 19
      Height = 11
      Caption = 'End:'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -9
      Font.Name = 'Tahoma'
      Font.Style = []
      ParentFont = False
    end
    object SOC: TEdit
      Left = 12
      Top = 35
      Width = 170
      Height = 21
      TabOrder = 0
      OnChange = CheckChanges
      OnKeyDown = FormKeyDown
    end
    object EOC: TEdit
      Left = 190
      Top = 35
      Width = 170
      Height = 21
      TabOrder = 1
      OnChange = CheckChanges
      OnKeyDown = FormKeyDown
    end
  end
  object GroupBox4: TGroupBox
    Left = 383
    Top = 104
    Width = 371
    Height = 73
    Caption = 'Nominal header:'
    TabOrder = 6
    object Label10: TLabel
      Left = 12
      Top = 22
      Width = 23
      Height = 11
      Caption = 'Start:'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -9
      Font.Name = 'Tahoma'
      Font.Style = []
      ParentFont = False
    end
    object Label11: TLabel
      Left = 190
      Top = 22
      Width = 19
      Height = 11
      Caption = 'End:'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -9
      Font.Name = 'Tahoma'
      Font.Style = []
      ParentFont = False
    end
    object SON: TEdit
      Left = 12
      Top = 35
      Width = 170
      Height = 21
      TabOrder = 0
      OnChange = CheckChanges
      OnKeyDown = FormKeyDown
    end
    object EON: TEdit
      Left = 190
      Top = 35
      Width = 170
      Height = 21
      TabOrder = 1
      OnChange = CheckChanges
      OnKeyDown = FormKeyDown
    end
  end
  object GroupBox5: TGroupBox
    Left = 383
    Top = 183
    Width = 371
    Height = 73
    Caption = 'Rate header:'
    TabOrder = 7
    object Label12: TLabel
      Left = 12
      Top = 22
      Width = 23
      Height = 11
      Caption = 'Start:'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -9
      Font.Name = 'Tahoma'
      Font.Style = []
      ParentFont = False
    end
    object Label13: TLabel
      Left = 190
      Top = 22
      Width = 19
      Height = 11
      Caption = 'End:'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -9
      Font.Name = 'Tahoma'
      Font.Style = []
      ParentFont = False
    end
    object SOR: TEdit
      Left = 12
      Top = 35
      Width = 170
      Height = 21
      TabOrder = 0
      OnChange = CheckChanges
      OnKeyDown = FormKeyDown
    end
    object EOR: TEdit
      Left = 190
      Top = 35
      Width = 170
      Height = 21
      TabOrder = 1
      OnChange = CheckChanges
      OnKeyDown = FormKeyDown
    end
  end
  object GroupBox6: TGroupBox
    Left = 8
    Top = 341
    Width = 746
    Height = 52
    Caption = 'Currencies:'
    TabOrder = 9
    object Label14: TLabel
      Left = 12
      Top = 20
      Width = 35
      Height = 13
      Caption = 'Native:'
    end
    object Label15: TLabel
      Left = 107
      Top = 20
      Width = 42
      Height = 13
      Caption = 'Support:'
    end
    object DEFCODE: TEdit
      Left = 53
      Top = 18
      Width = 44
      Height = 21
      CharCase = ecUpperCase
      TabOrder = 0
      OnChange = CheckChanges
      OnKeyDown = FormKeyDown
    end
    object SUPPORT: TEdit
      Left = 155
      Top = 18
      Width = 500
      Height = 21
      CharCase = ecUpperCase
      TabOrder = 1
      OnChange = CheckChanges
      OnKeyDown = FormKeyDown
    end
    object Load: TButton
      Left = 662
      Top = 16
      Width = 75
      Height = 25
      Caption = 'Load'
      TabOrder = 2
      OnClick = LoadClick
      OnKeyDown = FormKeyDown
    end
  end
  object GroupBox7: TGroupBox
    Left = 8
    Top = 407
    Width = 747
    Height = 194
    Caption = 'Preview:'
    TabOrder = 10
    Visible = False
    object Label17: TLabel
      Left = 12
      Top = 23
      Width = 36
      Height = 13
      Caption = 'Target:'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'Tahoma'
      Font.Style = []
      ParentFont = False
    end
    object Label16: TLabel
      Left = 119
      Top = 23
      Width = 27
      Height = 13
      Caption = 'Date:'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'Tahoma'
      Font.Style = []
      ParentFont = False
    end
    object Label18: TLabel
      Left = 7
      Top = 58
      Width = 36
      Height = 13
      Caption = 'Target:'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'Tahoma'
      Font.Style = []
      ParentFont = False
    end
    object Label19: TLabel
      Left = 7
      Top = 92
      Width = 37
      Height = 13
      Caption = 'Source:'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'Tahoma'
      Font.Style = []
      ParentFont = False
    end
    object Label20: TLabel
      Left = 7
      Top = 126
      Width = 41
      Height = 13
      Caption = 'Nominal:'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'Tahoma'
      Font.Style = []
      ParentFont = False
    end
    object Label21: TLabel
      Left = 7
      Top = 160
      Width = 27
      Height = 13
      Caption = 'Rate:'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'Tahoma'
      Font.Style = []
      ParentFont = False
    end
    object TestDate: TDateTimePicker
      Left = 152
      Top = 21
      Width = 81
      Height = 21
      Date = 43830.000000000000000000
      Time = 0.666811400464212100
      TabOrder = 1
    end
    object TestTarget: TEdit
      Left = 54
      Top = 20
      Width = 44
      Height = 21
      CharCase = ecUpperCase
      TabOrder = 0
      Text = 'USD'
    end
    object PTarget: TEdit
      Left = 53
      Top = 56
      Width = 44
      Height = 21
      TabStop = False
      CharCase = ecUpperCase
      TabOrder = 3
    end
    object PSource: TEdit
      Left = 53
      Top = 90
      Width = 44
      Height = 21
      TabStop = False
      CharCase = ecUpperCase
      TabOrder = 4
    end
    object PNominal: TEdit
      Left = 54
      Top = 124
      Width = 44
      Height = 21
      TabStop = False
      CharCase = ecUpperCase
      TabOrder = 5
    end
    object PRate: TEdit
      Left = 54
      Top = 158
      Width = 44
      Height = 21
      TabStop = False
      CharCase = ecUpperCase
      TabOrder = 6
    end
    object Test: TButton
      Left = 662
      Top = 16
      Width = 75
      Height = 25
      Caption = 'Preview'
      TabOrder = 2
      OnClick = TestClick
    end
    object DATA: TMemo
      Left = 119
      Top = 48
      Width = 618
      Height = 129
      ReadOnly = True
      ScrollBars = ssVertical
      TabOrder = 7
    end
  end
  object GroupBox8: TGroupBox
    Left = 383
    Top = 262
    Width = 371
    Height = 73
    Caption = 'Aditional settings:'
    TabOrder = 8
    object INVERSE: TCheckBox
      Left = 10
      Top = 20
      Width = 127
      Height = 17
      Caption = 'Inverse currency rate'
      TabOrder = 0
      OnClick = CheckChanges
      OnKeyDown = FormKeyDown
    end
    object THIRDPART: TCheckBox
      Left = 10
      Top = 42
      Width = 151
      Height = 17
      Caption = 'Enable 3rd part convertion'
      TabOrder = 1
      OnClick = CheckChanges
      OnKeyDown = FormKeyDown
    end
  end
  object ClearCache: TButton
    Left = 679
    Top = 8
    Width = 75
    Height = 25
    Caption = 'Clear Cache'
    TabOrder = 15
    OnClick = ClearCacheClick
    OnKeyDown = FormKeyDown
  end
end
