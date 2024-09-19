object Form1: TForm1
  Left = 283
  Top = 190
  Width = 1044
  Height = 823
  Caption = 'Form1'
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  OnActivate = FormActivate
  PixelsPerInch = 96
  TextHeight = 13
  object PageControl1: TPageControl
    Left = 0
    Top = 0
    Width = 1017
    Height = 753
    ActivePage = TabSheetProv
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -16
    Font.Name = 'MS Sans Serif'
    Font.Style = []
    MultiLine = True
    ParentFont = False
    TabIndex = 2
    TabOrder = 0
    TabPosition = tpLeft
    object TabSheetShablon: TTabSheet
      Caption = #1064#1072#1073#1083#1086#1085#1099
      OnShow = TabSheetShablonShow
      object DateTimePickerShablon: TDateTimePicker
        Left = 8
        Top = 0
        Width = 89
        Height = 32
        CalAlignment = dtaLeft
        Date = 45433.5829704514
        Format = 'yyyy'
        Time = 45433.5829704514
        DateFormat = dfShort
        DateMode = dmUpDown
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -19
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        Kind = dtkDate
        ParseInput = False
        ParentFont = False
        TabOrder = 0
      end
      object ButtonWordShablon: TButton
        Left = 880
        Top = 0
        Width = 97
        Height = 33
        Caption = #1057#1083#1080#1103#1085#1080#1077
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -19
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 1
      end
      object ComboBoxShablon: TComboBox
        Left = 96
        Top = 0
        Width = 249
        Height = 32
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -19
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ItemHeight = 0
        ParentFont = False
        TabOrder = 2
        Text = 'ComboBoxShablon'
      end
      object Edit1: TEdit
        Left = 344
        Top = 0
        Width = 121
        Height = 32
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -19
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 3
        Text = 'EditShablon'
      end
      object DBNavigatorShablon: TDBNavigator
        Left = 464
        Top = 0
        Width = 224
        Height = 33
        VisibleButtons = [nbFirst, nbPrior, nbNext, nbLast]
        TabOrder = 4
      end
      object ButtonSelectShablon: TButton
        Left = 784
        Top = 0
        Width = 97
        Height = 33
        Caption = #1064#1072#1073#1083#1086#1085
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -19
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 5
      end
      object ButtonRefreshShablon: TButton
        Left = 688
        Top = 0
        Width = 97
        Height = 33
        Caption = #1054#1073#1085#1086#1074#1080#1090#1100
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -19
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 6
      end
      object Memo1: TMemo
        Left = 8
        Top = 48
        Width = 969
        Height = 697
        Lines.Strings = (
          'Memo1')
        TabOrder = 7
      end
      object ProgressBarShablon: TProgressBar
        Left = 8
        Top = 32
        Width = 969
        Height = 17
        Min = 0
        Max = 100
        TabOrder = 8
      end
    end
    object TabSheet2: TTabSheet
      Caption = 'TabSheet2'
      ImageIndex = 1
    end
    object TabSheetProv: TTabSheet
      Caption = #1055#1088#1086#1074#1077#1088#1082#1072
      ImageIndex = 2
      OnShow = TabSheetProvShow
      object Button1: TButton
        Left = 472
        Top = 8
        Width = 137
        Height = 33
        Caption = #1057#1090#1072#1088#1090
        TabOrder = 0
        OnClick = Button1Click
      end
      object Button2: TButton
        Left = 832
        Top = 0
        Width = 137
        Height = 33
        Caption = #1055#1088#1086#1092#1080#1083#1100
        TabOrder = 1
        OnClick = Button2Click
      end
      object DateTimePickerProv: TDateTimePicker
        Left = 192
        Top = 16
        Width = 105
        Height = 32
        CalAlignment = dtaLeft
        Date = 45370.9239593634
        Format = 'yyyy'
        Time = 45370.9239593634
        DateFormat = dfShort
        DateMode = dmUpDown
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -19
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        Kind = dtkDate
        ParseInput = False
        ParentFont = False
        TabOrder = 2
      end
      object Button4: TButton
        Left = 672
        Top = 8
        Width = 121
        Height = 33
        Caption = #1055#1077#1095#1072#1090#1100
        TabOrder = 3
        OnClick = Button4Click
      end
      object PageControlProv: TPageControl
        Left = 16
        Top = 88
        Width = 953
        Height = 641
        ActivePage = TabSheet6
        TabIndex = 1
        TabOrder = 4
        object TabSheetProvZad: TTabSheet
          Caption = #1047#1072#1076#1072#1095#1080
          object CheckListBox1: TCheckListBox
            Left = 8
            Top = 8
            Width = 425
            Height = 393
            OnClickCheck = CheckListBox1ClickCheck
            ItemHeight = 20
            TabOrder = 0
          end
          object Button3: TButton
            Left = 504
            Top = 352
            Width = 137
            Height = 33
            Caption = #1060#1072#1081#1083' '#1087#1088#1086#1074#1077#1088#1082#1080
            TabOrder = 1
            OnClick = Button3Click
          end
        end
        object TabSheet6: TTabSheet
          Caption = #1056#1077#1079#1091#1083#1100#1090#1072#1090
          ImageIndex = 1
          object ProvRichEdit: TRichEdit
            Left = 16
            Top = 8
            Width = 921
            Height = 585
            Lines.Strings = (
              'ProvRichEdit')
            ReadOnly = True
            ScrollBars = ssBoth
            TabOrder = 0
          end
        end
      end
    end
    object TabSheetSelect: TTabSheet
      Caption = #1042#1099#1073#1086#1088#1082#1080
      ImageIndex = 3
      OnShow = TabSheetSelectShow
      object DateTimePickerSelect: TDateTimePicker
        Left = 8
        Top = 0
        Width = 105
        Height = 32
        CalAlignment = dtaLeft
        Date = 45407.5726256944
        Format = 'yyyy'
        Time = 45407.5726256944
        DateFormat = dfShort
        DateMode = dmUpDown
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -19
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        Kind = dtkDate
        ParseInput = False
        ParentFont = False
        TabOrder = 0
      end
      object ComboBoxSelect: TComboBox
        Left = 120
        Top = 0
        Width = 489
        Height = 32
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -19
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ItemHeight = 0
        ParentFont = False
        TabOrder = 1
        Text = #1053#1072#1080#1084#1077#1085#1086#1074#1072#1085#1080#1077' '#1074#1099#1073#1086#1088#1082#1080
      end
      object ButtonSelect: TButton
        Left = 616
        Top = 8
        Width = 75
        Height = 25
        Caption = #1042#1099#1073#1088#1072#1090#1100
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -19
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 2
        OnClick = ButtonSelectClick
      end
      object ButtonPrintSelect: TButton
        Left = 696
        Top = 8
        Width = 89
        Height = 25
        Caption = #1055#1077#1095#1072#1090#1100
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -19
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 3
        OnClick = ButtonPrintSelectClick
      end
      object ButtonSelectProfile: TButton
        Left = 896
        Top = 8
        Width = 75
        Height = 25
        Caption = #1055#1088#1086#1092#1080#1083#1100
        TabOrder = 4
        OnClick = ButtonSelectProfileClick
      end
      object RichEditSelect: TRichEdit
        Left = 8
        Top = 40
        Width = 969
        Height = 529
        Lines.Strings = (
          'RichEditSelect')
        ScrollBars = ssBoth
        TabOrder = 5
      end
    end
    object TabSheetSelectProfile: TTabSheet
      Caption = #1055#1088#1086#1092#1080#1083#1100
      ImageIndex = 4
      object ListView1: TListView
        Left = 80
        Top = 144
        Width = 250
        Height = 150
        Columns = <>
        TabOrder = 0
      end
      object Button5: TButton
        Left = 136
        Top = 56
        Width = 75
        Height = 25
        Caption = 'Button5'
        TabOrder = 1
        OnClick = Button5Click
      end
      object ListBox1: TListBox
        Left = 432
        Top = 160
        Width = 297
        Height = 297
        ItemHeight = 20
        MultiSelect = True
        ScrollWidth = 1000
        TabOrder = 2
      end
    end
  end
end
