object SheetHeader: TSheetHeader
  Left = 0
  Top = 0
  BorderIcons = [biSystemMenu]
  BorderStyle = bsDialog
  Caption = 'Sheet header'
  ClientHeight = 113
  ClientWidth = 725
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  Position = poScreenCenter
  PixelsPerInch = 96
  TextHeight = 13
  object Label1: TLabel
    Left = 16
    Top = 8
    Width = 31
    Height = 13
    Caption = 'Client:'
  end
  object Label2: TLabel
    Left = 16
    Top = 43
    Width = 56
    Height = 13
    Caption = 'Period-end:'
  end
  object Label3: TLabel
    Left = 227
    Top = 43
    Width = 71
    Height = 13
    Caption = 'Currency/Unit:'
  end
  object Label4: TLabel
    Left = 16
    Top = 80
    Width = 24
    Height = 13
    Caption = 'Title:'
  end
  object PeriodEnd: TDateTimePicker
    Left = 80
    Top = 43
    Width = 121
    Height = 21
    Date = 44264.000000000000000000
    Time = 0.885727465276431800
    TabOrder = 0
  end
  object CurrencyUnit: TEdit
    Left = 304
    Top = 43
    Width = 121
    Height = 21
    TabOrder = 1
  end
  object Title: TEdit
    Left = 80
    Top = 77
    Width = 544
    Height = 21
    TabOrder = 2
  end
  object Client: TComboBox
    Left = 80
    Top = 8
    Width = 345
    Height = 21
    TabOrder = 3
  end
  object Add: TButton
    Left = 630
    Top = 75
    Width = 75
    Height = 25
    Caption = 'Add'
    TabOrder = 4
    OnClick = AddClick
  end
end
