object FormColumnSettings: TFormColumnSettings
  Left = 364
  Top = 252
  BorderIcons = [biSystemMenu]
  Caption = 'Column Settings'
  ClientHeight = 335
  ClientWidth = 287
  Color = clBtnFace
  Constraints.MinHeight = 370
  Constraints.MinWidth = 295
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -14
  Font.Name = 'Tahoma'
  Font.Style = []
  KeyPreview = True
  OldCreateOrder = False
  Position = poScreenCenter
  OnCreate = FormCreate
  OnKeyPress = FormKeyPress
  OnResize = FormResize
  DesignSize = (
    287
    335)
  PixelsPerInch = 120
  TextHeight = 17
  object Bevel1: TBevel
    Left = 21
    Top = 366
    Width = 329
    Height = 18
    Anchors = [akLeft, akTop, akRight, akBottom]
    Shape = bsBottomLine
  end
  object Label2: TLabel
    Left = 25
    Top = 322
    Width = 192
    Height = 17
    Alignment = taCenter
    Anchors = [akBottom]
    Caption = 'The selected column should be '
    OnClick = FormCreate
  end
  object Label3: TLabel
    Left = 276
    Top = 322
    Width = 64
    Height = 17
    Anchors = [akBottom]
    Caption = 'pixels wide'
  end
  object Panel1: TPanel
    Left = 0
    Top = 0
    Width = 287
    Height = 46
    Align = alTop
    BevelOuter = bvNone
    BorderWidth = 4
    TabOrder = 0
    object Label1: TLabel
      Left = 4
      Top = 4
      Width = 279
      Height = 38
      Align = alClient
      Alignment = taCenter
      AutoSize = False
      Caption = 
        'Check the columns you would like to make visible in this Folder.' +
        '  Drag and Drop to reorder the columns. '
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -15
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
      WordWrap = True
      ExplicitLeft = 5
      ExplicitTop = 5
      ExplicitWidth = 365
      ExplicitHeight = 36
    end
  end
  object CheckBoxLiveUpdate: TCheckBox
    Left = 24
    Top = 350
    Width = 105
    Height = 23
    Anchors = [akRight, akBottom]
    Caption = 'Live Update'
    TabOrder = 1
    OnClick = CheckBoxLiveUpdateClick
  end
  object ButtonOk: TButton
    Left = 136
    Top = 398
    Width = 98
    Height = 32
    Anchors = [akRight, akBottom]
    Caption = '&OK'
    ModalResult = 1
    TabOrder = 2
  end
  object ButtonCancel: TButton
    Left = 251
    Top = 398
    Width = 98
    Height = 32
    Anchors = [akRight, akBottom]
    Caption = '&Cancel'
    ModalResult = 2
    TabOrder = 3
  end
  object VSTColumnNames: TVirtualStringTree
    Left = 10
    Top = 61
    Width = 357
    Height = 244
    Anchors = [akLeft, akTop, akRight, akBottom]
    CheckImageKind = ckDarkCheck
    Header.AutoSizeIndex = 0
    Header.Font.Charset = DEFAULT_CHARSET
    Header.Font.Color = clWindowText
    Header.Font.Height = -11
    Header.Font.Name = 'MS Sans Serif'
    Header.Font.Style = []
    Header.MainColumn = -1
    Header.Options = [hoColumnResize, hoDrag]
    HintAnimation = hatNone
    TabOrder = 4
    TreeOptions.AutoOptions = [toAutoDropExpand, toAutoScroll, toAutoScrollOnExpand, toAutoTristateTracking]
    TreeOptions.MiscOptions = [toAcceptOLEDrop, toCheckSupport, toInitOnSave, toToggleOnDblClick]
    TreeOptions.PaintOptions = [toShowButtons, toShowRoot, toThemeAware, toUseBlendedImages]
    OnChecking = VSTColumnNamesChecking
    OnDragAllowed = VSTColumnNamesDragAllowed
    OnDragOver = VSTColumnNamesDragOver
    OnDragDrop = VSTColumnNamesDragDrop
    OnFocusChanging = VSTColumnNamesFocusChanging
    OnFreeNode = VSTColumnNamesFreeNode
    OnGetText = VSTColumnNamesGetText
    OnInitNode = VSTColumnNamesInitNode
    Columns = <>
  end
  object EditPixelWidth: TEdit
    Left = 225
    Top = 319
    Width = 46
    Height = 25
    Anchors = [akBottom]
    TabOrder = 5
    OnExit = EditPixelWidthExit
    OnKeyPress = EditPixelWidthKeyPress
  end
end
