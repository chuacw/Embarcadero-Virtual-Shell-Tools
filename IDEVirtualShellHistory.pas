unit IDEVirtualShellHistory;

// Version 1.3.0
//
// The contents of this file are subject to the Mozilla Public License
// Version 1.1 (the "License"); you may not use this file except in compliance
// with the License. You may obtain a copy of the License at http://www.mozilla.org/MPL/
//
// Alternatively, you may redistribute this library, use and/or modify it under the terms of the
// GNU Lesser General Public License as published by the Free Software Foundation;
// either version 2.1 of the License, or (at your option) any later version.
// You may obtain a copy of the LGPL at http://www.gnu.org/copyleft/.
//
// Software distributed under the License is distributed on an "AS IS" basis,
// WITHOUT WARRANTY OF ANY KIND, either express or implied. See the License for the
// specific language governing rights and limitations under the License.
//
//----------------------------------------------------------------------------

//The initial developer of this code is Robert Lee.


{

Requirements:
  - Mike Lischke's Virtual Treeview (VT)
    http://www.lischke-online.de/VirtualTreeview/VT.html
  - Jim Kuenaman's Virtual Shell Tools
    http://groups.yahoo.com/group/VirtualExplorerTree

Credits:
  Special thanks to Mike Lischke (VT) and Jim Kuenaman (VSTools) for the
  magnificent components they have made available to the Delphi community.

How to use:
  VirtualShellHistory is basically a History manager, it keeps an array of
  TNamespaces that the attached VET control has been navigating.
  To use it just attach a VirtualExplorerTree to the VET property and use the
  following methods or properties:
  - Back/Next: to move to the previous or next directory.
  - FillPopupMenu: fills a TPopupMenu to mimic the Explorer's Back and Next Buttons PopupMenu.
  - Add, Delete, Clear, Count, Items, ItemIndex: to manage the items.
  - MaxCount: maximum items capacity.

Todo:
  -
Known issues:
  -

History:
12 April 2004 - version 0.4
  - Fixed recursion bug, the itemchange was dispatched to the VET control for
    every item when the items were loaded from file.
  - Added unicode support for TBX items.
  - The component automatically deletes invalid namespaces when the itemindex is
    changed and generates an OnChanged event with hctInvalidSelected as the
    parameter.

5 September 2002 - version 0.3.1
  - Added an extra check to LoadFromRegistry and fixed a problem with setting
    the focus to the ActiveControl, thanks to Ebi.

16 July 2002 - version 0.3
  - Robert debugged my changes and I added a Most Recently Used layer to the class
  - Added Support for Toolbar 2000 and TBX Toolbar

14 July 2002 - version 0.2  (Jim Kueneman)
  - Finshed Component and added to VirtualShellTools package
  - Added more options such as using the entire path, images, etc in the popupmenu

13 July 2002 - version 0.1  (Robert Lee)
  - First release, with a lousy unicode support.

==============================================================================}

interface

{$include VSToolsAddIns.inc}
{$include Compilers.inc}

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms,
  Dialogs, Menus, IDEVirtualTrees, IDEVirtualExplorerTree, IDEVirtualShellUtilities,
  IDEVirtualWideStrings, ShlObj, ShellAPI, CommCtrl, ImgList, IDEVirtualUtilities,
  IDEVirtualUnicodeDefines, IDEVirtualSystemImageLists, IDEVirtualPIDLTools,
  Registry
  {$IFDEF TOOLBAR_2000}
  ,TB2Item
  {$ENDIF TOOLBAR_2000}
  {$IFDEF TBX}
  ,TB2Item, TBX, TBXThemes, TB2Common
  {$ENDIF TBX}
  ;


const
  REGISTRYDATASIZE = 'VirtualShellHistoryMRUSize';
  REGISTRYDATA = 'VirtualShellHistoryMRUData';

type
  TVSHChangeType = (
    hctAdded,           // A path was added to the history list
    hctDeleted,         // A path was deleted from the history list
    hctSelected,        // A path was selected
    hctInvalidSelected  // An invalid path was selected
  );

  TVSHMenuTextType = (
    mttName, // Simple in folder name used in menu
    mttPath  // Full Path used in menu
  );

  TBaseVSPState = (
    bvsChangeNotified,    // The component has been notified that a change happened in the associated VET or Combobox
    bvsChangeDispatching, // The component has been interacting with and is changing the associated VET or Combobox
    bvsChangeItemsLoading // The items are being loaded from stream
  );
  TBaseVSPStates = set of TBaseVSPState;

  TFillPopupDirection = (
    fpdNewestToOldest,   // Fill menu present to past
    fpdOldestToNewest    // Fill menu past to present
  );

type
  TCustomVirtualShellHistory = class;
  TBaseVirtualShellPersistent = class;

  TVSPChangeEvent = procedure(Sender: TBaseVirtualShellPersistent; ItemIndex: Integer; ChangeType: TVSHChangeType) of object;
  TVSPGetImageEvent = procedure(Sender: TBaseVirtualShellPersistent; NS: TNamespace; var ImageList: TImageList; var ImageIndex: Integer) of object;

  TVSHMenuOptions = class(TPersistent)
  private
    FEllipsisPlacement: TShortenStringEllipsis; // Where the ellipsis goes if the string must be shortened
    FImages: Boolean;            // Show Images in menu
    FImageBorder: Integer;       // Border between Image and the drawn frame
    FLargeImages: Boolean;       // Use large (32x32) images in the menu
    FTextType: TVSHMenuTextType; // What type of text is used in the menu
    FMaxWidth: Integer;          // Set a max width for the Menu, the text will be shortened

  public
    constructor Create;
    destructor Destroy; override;

    procedure Assign(Source: TPersistent); override;
    procedure AssignTo(Dest: TPersistent); override;
  published
    property EllipsisPlacement: TShortenStringEllipsis read FEllipsisPlacement write FEllipsisPlacement default sseFilePathMiddle;
    property Images: Boolean read FImages write FImages default False;
    property ImageBorder: Integer read FImageBorder write FImageBorder default 1;
    property LargeImages: Boolean read FLargeImages write FLargeImages default False;
    property TextType: TVSHMenuTextType read FTextType write FTextType default mttName;
    property MaxWidth: Integer read FMaxWidth write FMaxWidth default -1;
  end;

  TBaseVirtualShellPersistent = class(TComponent)
  private
    FOnChange: TVSPChangeEvent;
    {$IFDEF EXPLORERCOMBOBOX}
    FVirtualExplorerComboBox: TCustomVirtualExplorerCombobox;
    {$ENDIF}
    FVirtualExplorerTree: TCustomVirtualExplorerTree;
    FLevels: Integer;
    FOnGetImage: TVSPGetImageEvent;
    FMenuOptions: TVSHMenuOptions;
    FItemIndex: integer;
    FNamespaces: TList;
    FState: TBaseVSPStates;

    {$IFDEF EXPLORERCOMBOBOX}
    procedure SetVirtualExplorerComboBox(const Value: TCustomVirtualExplorerCombobox);
    {$ENDIF}
    procedure SetVirtualExplorerTree(const Value: TCustomVirtualExplorerTree);
    procedure SetLevels(const Value: Integer);
    procedure SetMenuOptions(const Value: TVSHMenuOptions);
    function GetLargeSysImages: TImageList;
    function GetSmallSysImages: TImageList;
    function GetItems(Index: integer): TNamespace;
    procedure SetItemIndex(Value: integer);
    function GetCount: integer;
    function GetHasBackItems: Boolean;
    function GetHasNextItems: Boolean;
  protected
    procedure ChangeLinkChanging(Server: TObject; NewPIDL: PItemIDList); dynamic;
    procedure ChangeLinkDispatch(PIDL: PItemIDList); virtual;
    procedure ChangeLinkFreeing(ChangeLink: IVETChangeLink); dynamic;
    function CreateMenuOptions: TVSHMenuOptions; dynamic;
    procedure ValidateLevels;
    procedure DoGetImage(NS: TNamespace; var ImageList: TImageList; var ImageIndex: Integer); virtual;
    procedure DoItemChange(ItemIndex: Integer; ChangeType: TVSHChangeType);
    procedure OnMenuItemClick(Sender: TObject); virtual;
    procedure OnMenuItemDraw(Sender: TObject; ACanvas: TCanvas; ARect: TRect; Selected: Boolean); virtual;
    procedure OnMenuItemMeasure(Sender: TObject; ACanvas: TCanvas; var Width, Height: Integer); virtual;

    property Count: integer read GetCount;
    property HasBackItems: Boolean read GetHasBackItems;
    property HasNextItems: Boolean read GetHasNextItems;
    property ItemIndex: integer read FItemIndex write SetItemIndex;
    property LargeSysImages: TImageList read GetLargeSysImages;
    property Levels: Integer read FLevels write SetLevels default 10;
    property MenuOptions: TVSHMenuOptions read FMenuOptions write SetMenuOptions;
    property Namespaces: TList read FNamespaces write FNamespaces;
    property OnChange: TVSPChangeEvent read FOnChange write FOnChange;
    property OnGetImage: TVSPGetImageEvent read FOnGetImage write FOnGetImage;
    property SmallSysImages: TImageList read GetSmallSysImages;
    property State: TBaseVSPStates read FState write FState;
    {$IFDEF EXPLORERCOMBOBOX}
    property VirtualExplorerComboBox: TCustomVirtualExplorerCombobox read FVirtualExplorerComboBox write SetVirtualExplorerComboBox;
    {$ENDIF}
    property VirtualExplorerTree: TCustomVirtualExplorerTree read FVirtualExplorerTree write SetVirtualExplorerTree;
  public
    constructor Create(AOwner: TComponent); override;
    destructor Destroy; override;

    function Add(Value: TNamespace; Release: Boolean = False): integer; virtual;
    procedure Clear; virtual;
    procedure Delete(Index: integer);
    procedure FillPopupMenu(Popupmenu: TPopupMenu; FillDirection: TFillPopupDirection;
      AddClearItem: Boolean = False; ClearItemText: WideString = ''); virtual;
    {$IFDEF TBX_OR_TB2K}
    procedure FillPopupMenu_TBX(PopupMenu: TTBCustomItem; FillDirection: TFillPopupDirection;
      AddClearItem: Boolean = False; ClearItemText: WideString = ''); virtual;
    {$ENDIF}
    property Items[Index: integer]: TNamespace read GetItems; default;
    procedure LoadFromFile(FileName: WideString);
    procedure LoadFromStream(S: TStream); virtual;
    procedure LoadFromRegistry(RootKey: DWORD; SubKey: string);
    procedure SaveToFile(FileName: WideString);
    procedure SaveToStream(S: TStream); virtual;
    procedure SaveToRegistry(RootKey: DWORD; SubKey: string);
  end;

  TCustomVirtualShellMRU = class(TBaseVirtualShellPersistent)
  public
    property ItemIndex;
    property LargeSysImages;
    property SmallSysImages;
  end;

  TVirtualShellMRU = class(TCustomVirtualShellMRU)
  published
    property Count;
    property Levels;
    property MenuOptions;
    property OnChange;
    property OnGetImage;
    {$IFDEF EXPLORERCOMBOBOX}
    property VirtualExplorerComboBox;
    {$ENDIF}
    property VirtualExplorerTree;
  end;

  TCustomVirtualShellHistory = class(TBaseVirtualShellPersistent)
  public
    function Add(Value: TNamespace; Release: Boolean = False): integer; override;
    procedure Back;
    procedure Next;

    property HasBackItems;
    property HasNextItems;
    property ItemIndex;
    property LargeSysImages;
    property SmallSysImages;
  end;

  TVirtualShellHistory = class(TCustomVirtualShellHistory)
  published
    property Count;
    property Levels;
    property MenuOptions;
    property OnChange;
    property OnGetImage;
    {$IFDEF EXPLORERCOMBOBOX}
    property VirtualExplorerComboBox;
    {$ENDIF}
    property VirtualExplorerTree;
  end;

implementation

//WMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWM
// TBX Unicode menu item
{$IFDEF TBX}
type
  TTBXItemAccess = class(TTBXCustomItem);
  TTBXItemViewerAccess = class(TTBXItemViewer);
  TTBViewAccess = class(TTBView);

  TTBXUnicodeItem = class(TTBXItem)
  private
    FCaption: WideString;
    procedure SetCaption(const Value: WideString);
    procedure ReadCaptionW(Reader: TReader);
    procedure WriteCaptionW(Writer: TWriter);
  protected
    procedure DefineProperties(Filer: TFiler); override;
    function GetItemViewerClass(AView: TTBView): TTBItemViewerClass; override;
  public
    constructor Create(AOwner: TComponent); override;
  published
    // Don't let the VCL store the WideString, it doesn't write it correctly.
    // Use DefineProperties instead
    property Caption: WideString read FCaption write SetCaption stored False; // Hides the inherited Caption
  end;

  TTBXUnicodeItemViewer = class(TTBXItemViewer)
  protected
    function  GetTextSize(Canvas: TCanvas; const Text: string; TextFlags: Cardinal;
      Rotated: Boolean; StateFlags: Integer): TSize; override;
    procedure Paint(const Canvas: TCanvas; const ClientAreaRect: TRect;
      IsHoverItem: Boolean; IsPushed: Boolean; UseDisabledShadow: Boolean); override;
  end;

procedure FillItemInfo(ItemViewer: TTBItemViewer; var ItemInfo: TTBXItemInfo);
const
  CToolbarStyle: array [Boolean] of Integer = (0, IO_TOOLBARSTYLE);
  CCombo: array [Boolean] of Integer = (0, IO_COMBO);
  CSubmenuItem: array [Boolean] of Integer = (0, IO_SUBMENUITEM);
  CDesigning: array [Boolean] of Integer = (0, IO_DESIGNING);
  CAppActive: array [Boolean] of Integer = (0, IO_APPACTIVE);
var
  Item: TTBXCustomItem;
  View: TTBViewAccess;

  ClientAreaRect: TRect;
  IsHoverItem, IsOpen, IsPushed: Boolean;
  ToolbarStyle: Boolean;
  HasArrow: Boolean;
  IsSplit: Boolean;
  ImageIsShown: Boolean;
  ImgSize: TSize;
  IsComboPushed: Boolean;
begin
  Item := TTBXCustomItem(ItemViewer.Item);
  View := TTBViewAccess(ItemViewer.View);

  ClientAreaRect := ItemViewer.BoundsRect;
  OffsetRect(ClientAreaRect, -ClientAreaRect.Left, -ClientAreaRect.Top);

  IsOpen := ItemViewer = View.OpenViewer;
  IsHoverItem := ItemViewer = View.Selected;
  IsPushed := IsHoverItem and (IsOpen or (View.MouseOverSelected and View.Capture));

  ToolbarStyle := ItemViewer.IsToolbarStyle;
  IsSplit := tbisCombo in TTBXItemAccess(Item).ItemStyle;
  IsComboPushed := IsSplit and IsPushed and not View.Capture;
  if IsComboPushed then IsPushed := False;

  if TTBXItemViewerAccess(ItemViewer).GetImageShown then
  begin
    ImgSize := TTBXItemViewerAccess(ItemViewer).GetImageSize;
    with ImgSize do if (CX <= 0) or (CY <= 0) then
    begin
      CX := 0;
      CY := 0;
      ImageIsShown := False;
    end
    else ImageIsShown := True;
  end
  else
  begin
    ImgSize.CX := 0;
    ImgSize.CY := 0;
    ImageIsShown := False;
  end;
  IsSplit := tbisCombo in TTBXItemAccess(Item).ItemStyle;

  FillChar(ItemInfo, SizeOf(ItemInfo), 0);
  ItemInfo.ViewType := GetViewType(View);
  ItemInfo.ItemOptions := CToolbarStyle[ToolbarStyle] or CCombo[IsSplit] or
    CDesigning[csDesigning in Item.ComponentState] or CSubmenuItem[tbisSubmenu in TTBXItemAccess(Item).ItemStyle] or
    CAppActive[Application.Active];
  ItemInfo.Enabled := Item.Enabled or View.Customizing;
  ItemInfo.Pushed := IsPushed;
  ItemInfo.Selected := Item.Checked;
  ItemInfo.ImageShown := ImageIsShown;
  ItemInfo.ImageWidth := ImgSize.CX;
  ItemInfo.ImageHeight := ImgSize.CY;
  if IsHoverItem then
  begin
    if not ItemInfo.Enabled and not View.MouseOverSelected then ItemInfo.HoverKind := hkKeyboardHover
    else if ItemInfo.Enabled then ItemInfo.HoverKind := hkMouseHover;
  end
  else ItemInfo.HoverKind := hkNone;
  ItemInfo.IsPopupParent := ToolbarStyle and
    (((vsModal in View.State) and Assigned(View.OpenViewer)) or (tbisSubmenu in TTBXItemAccess(Item).ItemStyle)) and
    ((IsSplit and IsComboPushed) or (not IsSplit and IsPushed));
  ItemInfo.IsVertical := (View.Orientation = tbvoVertical) and not IsSplit;
//  if not ToolbarStyle then ItemInfo.PopupMargin := GetPopupMargin(Self);
  ItemInfo.PopupMargin := GetPopupMargin(ItemViewer);

  HasArrow := (tbisSubmenu in TTBXItemAccess(Item).ItemStyle) and
    ((tbisCombo in TTBXItemAccess(Item).ItemStyle) or (tboDropdownArrow in Item.EffectiveOptions));

  if ToolbarStyle then
  begin
    if HasArrow then
      ItemInfo.ComboPart := cpCombo;
    if IsSplit then
      ItemInfo.ComboPart := cpSplitLeft;
  end;
end;

function DrawXPText(ACanvas: TCanvas; Caption: WideString; var ARect: TRect; Flags: Cardinal): integer;
var
  S: string;
  BS: TBrushStyle;
begin
  BS := ACanvas.Brush.Style;
  ACanvas.Brush.Style := bsClear;
  if Win32Platform = VER_PLATFORM_WIN32_WINDOWS then begin
    S := Caption;
    Result := Windows.DrawText(ACanvas.Handle, PChar(S), -1, ARect, Flags);
  end
  else
    Result := Windows.DrawTextW(ACanvas.Handle, PWideChar(Caption), -1, ARect, Flags);
  ACanvas.Brush.Style := BS;
end;

{ TSpTBXCustomItem }

constructor TTBXUnicodeItem.Create(AOwner: TComponent);
begin
  inherited;
  FCaption := '';
end;

procedure TTBXUnicodeItem.DefineProperties(Filer: TFiler);
begin
  inherited;
  // Don't let the VCL store the WideString, it doesn't write it correctly.
  // Use DefineProperties instead, with a new name for the property
  Filer.DefineProperty('CaptionW', ReadCaptionW, WriteCaptionW, FCaption <> '');
end;

procedure TTBXUnicodeItem.ReadCaptionW(Reader: TReader);
begin
  case Reader.NextValue of
    vaLString, vaString:
      SetCaption(Reader.ReadString);
  else
    SetCaption(Reader.ReadWideString);
  end;
end;

procedure TTBXUnicodeItem.WriteCaptionW(Writer: TWriter);
begin
  Writer.WriteWideString(FCaption);
end;

function TTBXUnicodeItem.GetItemViewerClass(AView: TTBView): TTBItemViewerClass;
begin
  Result := TTBXUnicodeItemViewer;
end;

procedure TTBXUnicodeItem.SetCaption(const Value: WideString);
begin
  if FCaption <> Value then begin
    FCaption := Value;
    inherited Caption := Value;
  end;
end;

{ TTBXUnicodeItemViewer }

function TTBXUnicodeItemViewer.GetTextSize(Canvas: TCanvas;
  const Text: string; TextFlags: Cardinal; Rotated: Boolean;
  StateFlags: Integer): TSize;
var
  R: TRect;
  I: TTBXUnicodeItem;
begin
  Result := inherited GetTextSize(Canvas, Text, TextFlags, Rotated, StateFlags);
  if (Result.cx > 0) and (Result.cy > 0) then begin
    I := TTBXUnicodeItem(Item);
    R := Rect(0, 1, 0, 0);
    DrawXPText(Canvas, I.Caption, R, DT_CALCRECT);

    Result.cx := Result.cx + (R.Right - Result.cx);
    Result.cy := Result.cy + (R.Bottom - Result.cy);
  end;
end;

procedure TTBXUnicodeItemViewer.Paint(const Canvas: TCanvas; const ClientAreaRect: TRect;
  IsHoverItem, IsPushed, UseDisabledShadow: Boolean);
var
  Item: TTBXItemAccess;
  View: TTBViewAccess;
  ItemInfo: TTBXItemInfo;

  R: TRect;
  ComboRect: TRect;
  CaptionRect: TRect;
  ImageRect: TRect;
  C: TColor;

  ToolbarStyle: Boolean;
  HasArrow: Boolean;
  IsSplit: Boolean;
  ImageIsShown: Boolean;
  ImageOrCheckShown: Boolean;
  ImgAndArrowWidth: Integer;
  ImgSize: TSize;
  IsComboPushed: Boolean;
  IsCaptionShown: Boolean;
  IsTextRotated: Boolean;
  ItemLayout: TTBXItemLayout;
  S: string;
  StateFlags: Integer;
  IsSpecialDropDown: Boolean;
  TextFlags: Cardinal;
  TextMetrics: TTextMetric;
  TextSize: TSize;
  Margins: TTBXMargins;
begin
  Item := TTBXItemAccess(Self.Item);
  View := TTBViewAccess(Self.View);
  FillItemInfo(Self, ItemInfo);

  ToolbarStyle := IsToolbarStyle;
  IsSplit := tbisCombo in TTBXItemAccess(Item).ItemStyle;
  IsComboPushed := IsSplit and IsPushed and not View.Capture;
  if IsComboPushed then IsPushed := False;

  ItemLayout := Item.Layout;
  if ItemLayout = tbxlAuto then
  begin
    if tboImageAboveCaption in Item.EffectiveOptions then ItemLayout := tbxlGlyphTop
    else if View.Orientation <> tbvoVertical then ItemLayout := tbxlGlyphLeft
    else ItemLayout := tbxlGlyphTop;
  end;

  HasArrow := (tbisSubmenu in TTBXItemAccess(Item).ItemStyle) and
    ((tbisCombo in TTBXItemAccess(Item).ItemStyle) or (tboDropdownArrow in Item.EffectiveOptions));

  if GetImageShown then
  begin
    ImgSize := GetImageSize;
    with ImgSize do if (CX <= 0) or (CY <= 0) then
    begin
      CX := 0;
      CY := 0;
      ImageIsShown := False;
    end
    else ImageIsShown := True;
  end
  else
  begin
    ImgSize.CX := 0;
    ImgSize.CY := 0;
    ImageIsShown := False;
  end;
  ImageOrCheckShown := ImageIsShown or (not ToolbarStyle and Item.Checked);

  StateFlags := GetStateFlags(ItemInfo);

  Canvas.Font := TTBViewAccess(View).GetFont;
  Canvas.Font.Color := CurrentTheme.GetItemTextColor(ItemInfo);
  DoAdjustFont(Canvas.Font, StateFlags);
  C := Canvas.Font.Color;

  { Setup font }
  TextFlags := GetTextFlags;
  IsCaptionShown := CaptionShown;
  IsTextRotated := (View.Orientation = tbvoVertical) and ToolbarStyle;
  if IsCaptionShown then
  begin
    S := GetCaptionText;
    if (Item.Layout <> tbxlAuto) or (tboImageAboveCaption in Item.EffectiveOptions) then
      IsTextRotated := False;
    if IsTextRotated or not ToolbarStyle then TextFlags := TextFlags or DT_SINGLELINE;
    TextSize := GetTextSize(Canvas, S, TextFlags, IsTextRotated, StateFlags);
  end
  else
  begin
    StateFlags := 0;
    SetLength(S, 0);
    IsTextRotated := False;
    TextSize.CX := 0;
    TextSize.CY := 0;
  end;

  IsSpecialDropDown := HasArrow and not IsSplit and ToolbarStyle and
    ((Item.Layout = tbxlGlyphTop) or (Item.Layout = tbxlAuto) and (tboImageAboveCaption in Item.EffectiveOptions)) and
    (ImgSize.CX > 0) and not (IsTextRotated) and (TextSize.CX > 0);

  { Border & Arrows }
  R := ClientAreaRect;
  with CurrentTheme do if ToolbarStyle then
  begin
    GetMargins(MID_TOOLBARITEM, Margins);
    if HasArrow then with R do
    begin
      ItemInfo.ComboPart := cpCombo;
      if IsSplit then
      begin
        ItemInfo.ComboPart := cpSplitLeft;
        ComboRect := R;
        Dec(Right, SplitBtnArrowWidth);
        ComboRect.Left := Right;
      end
      else if not IsSpecialDropDown then
      begin
        if View.Orientation <> tbvoVertical then
          ComboRect := Rect(Right - DropdownArrowWidth - DropdownArrowMargin, 0,
            Right - DropdownArrowMargin, Bottom)
        else
          ComboRect := Rect(0, Bottom - DropdownArrowWidth - DropdownArrowMargin,
            Right, Bottom - DropdownArrowMargin);
      end
      else
      begin
        ImgAndArrowWidth := ImgSize.CX + DropdownArrowWidth + 2;
        ComboRect.Right := (R.Left + R.Right + ImgAndArrowWidth + 2) div 2;
        ComboRect.Left := ComboRect.Right - DropdownArrowWidth;
        ComboRect.Top := (R.Top + R.Bottom - ImgSize.CY - 2 - TextSize.CY) div 2;
        ComboRect.Bottom := ComboRect.Top + ImgSize.CY;
      end;
    end
    else SetRectEmpty(ComboRect);

    if not IsSplit then
    begin
      CurrentTheme.PaintButton(Canvas, R, ItemInfo);

      if HasArrow then
      begin
        PaintDropDownArrow(Canvas, ComboRect, ItemInfo);
        if not IsSpecialDropDown then
        begin
          if View.Orientation <> tbvoVertical then Dec(R.Right, DropdownArrowWidth)
          else Dec(R.Bottom, DropdownArrowWidth);
        end;
      end;
    end
    else // IsSplit
    begin
      CurrentTheme.PaintButton(Canvas, R, ItemInfo);
      ItemInfo.Pushed := IsComboPushed;
      ItemInfo.Selected := False;
      ItemInfo.ComboPart := cpSplitRight;

      CurrentTheme.PaintButton(Canvas, ComboRect, ItemInfo);
      ItemInfo.ComboPart := cpSplitLeft;
      ItemInfo.Pushed := IsPushed;
      ItemInfo.Selected := Item.Checked;
    end;

    InflateRect(R, -2, -2);
  end
  else
  begin
    GetMargins(MID_MENUITEM, Margins);
    CurrentTheme.PaintMenuItem(Canvas, R, ItemInfo);
    Inc(R.Left, Margins.LeftWidth);
    Dec(R.Right, Margins.RightWidth);
    Inc(R.Top, Margins.TopHeight);
    Dec(R.Bottom, Margins.BottomHeight);
  end;

  { Caption }
  if IsCaptionShown then
  begin
    if ToolbarStyle then
    begin
      TextFlags := TextFlags or DT_CENTER or DT_VCENTER;
      CaptionRect := R;

      if ImageIsShown then
        Case ItemLayout of
          tbxlGlyphLeft:
            begin
              Inc(CaptionRect.Left, ImgSize.CX + 3);
              TextFlags := TextFlags and not DT_CENTER;
            end;
          tbxlGlyphTop:
            begin
              Inc(CaptionRect.Top, ImgSize.CY + 1);
              if IsTextRotated then Inc(CaptionRect.Top, 3);
              TextFlags := TextFlags and not DT_VCENTER;
            end;
        end;

      CaptionRect.Left := CaptionRect.Left + 3;
      CaptionRect.Top := (CaptionRect.Top + CaptionRect.Bottom - TextSize.CY) div 2;
      CaptionRect.Right := CaptionRect.Left + TextSize.CX;
      CaptionRect.Bottom := CaptionRect.Top + TextSize.CY;
    end
    else with CurrentTheme do
    begin
      TextFlags := DT_LEFT or DT_VCENTER or TextFlags;
      TextSize := GetTextSize(Canvas, S, TextFlags, False, StateFlags); { TODO : Check if this line is required }
      GetTextMetrics(Canvas.Handle, TextMetrics);

      CaptionRect := R;
      Inc(CaptionRect.Left, ItemInfo.PopupMargin + MenuImageTextSpace + MenuLeftCaptionMargin);
      with TextMetrics, CaptionRect do
        if (Bottom - Top) - (tmHeight + tmExternalLeading) = Margins.BottomHeight then Dec(Bottom);
      Inc(CaptionRect.Top, TextMetrics.tmExternalLeading);
      CaptionRect.Right := CaptionRect.Left + TextSize.CX;
    end;

    Canvas.Font.Color := C;
    DrawXPText(Canvas, TTBXUnicodeItem(Item).Caption, CaptionRect, TextFlags);
  end;

  { Shortcut and/or submenu arrow (menus only) }
  if not ToolbarStyle then
  begin
    S := Item.GetShortCutText;
    if Length(S) > 0 then
    begin
      CaptionRect := R;
      with CaptionRect, TextMetrics do
      begin
        Left := Right - (Bottom - Top) - GetTextWidth(Canvas.Handle, S, True);
        if (Bottom - Top) - (tmHeight + tmExternalLeading) = Margins.BottomHeight then Dec(Bottom);
        Inc(Top, TextMetrics.tmExternalLeading);
      end;
      Canvas.Font.Color := C;
      CurrentTheme.PaintCaption(Canvas, CaptionRect, ItemInfo, S, TextFlags, False);
    end;
  end;

  { Image, or check box }
  if ImageOrCheckShown then
  begin
    ImageRect := R;

    if ToolBarStyle then
    begin
      if IsSpecialDropDown then OffsetRect(ImageRect, (-CurrentTheme.DropdownArrowWidth + 1) div 2, 0);
      if ItemLayout = tbxlGlyphLeft then ImageRect.Right := ImageRect.Left + ImgSize.CX + 2
      else
      begin
        ImageRect.Top := (ImageRect.Top + ImageRect.Bottom - ImgSize.cy - 2 - TextSize.cy) div 2;
        ImageRect.Bottom := ImageRect.Top + ImgSize.CY;
      end;
    end
    else ImageRect.Right := ImageRect.Left + ClientAreaRect.Bottom - ClientAreaRect.Top;

    if ImageIsShown then with ImageRect, ImgSize do
    begin
      Left := Left + ((Right - Left) - CX) div 2;
      ImageRect.Top := Top + ((Bottom - Top) - CY) div 2;
      Right := Left + CX;
      Bottom := Top + CY;
      DrawItemImage(Canvas, ImageRect, ItemInfo);
    end
    else if not ToolbarStyle and Item.Checked then
      CurrentTheme.PaintCheckMark(Canvas, ImageRect, ItemInfo);
  end;
end;

(*
procedure TTBXUnicodeItemViewer.Paint(const Canvas: TCanvas; const ClientAreaRect: TRect;
  IsHoverItem, IsPushed, UseDisabledShadow: Boolean);
var
  Item: TTBXCustomItem;
  View: TTBViewAccess;
  ItemInfo: TTBXItemInfo;

  R: TRect;
  ComboRect: TRect;
  CaptionRect: TRect;
  ImageRect: TRect;
  C: TColor;

  ToolbarStyle: Boolean;
  HasArrow: Boolean;
  IsSplit: Boolean;
  ImageIsShown: Boolean;
  ImageOrCheckShown: Boolean;
  ImgAndArrowWidth: Integer;
  ImgSize: TSize;
  IsComboPushed: Boolean;
  IsCaptionShown: Boolean;
  IsTextRotated: Boolean;
  ItemLayout: TTBXItemLayout;
  PaintDefault: Boolean;
  S: string;
  StateFlags: Integer;
  IsSpecialDropDown: Boolean;
  TextFlags: Cardinal;
  TextMetrics: TTextMetric;
  TextSize: TSize;
  TextAlignment: TAlignment;
  Margins: TTBXMargins;
begin
  Item := TTBXCustomItem(Self.Item);
  View := TTBViewAccess(Self.View);
  FillItemInfo(Self, ItemInfo);

  ToolbarStyle := IsToolbarStyle;
  IsSplit := tbisCombo in TTBXItemAccess(Item).ItemStyle;
  IsComboPushed := IsSplit and IsPushed and not View.Capture;
  if IsComboPushed then IsPushed := False;

  ItemLayout := Item.Layout;
  if ItemLayout = tbxlAuto then
  begin
    if tboImageAboveCaption in Item.EffectiveOptions then ItemLayout := tbxlGlyphTop
    else if View.Orientation <> tbvoVertical then ItemLayout := tbxlGlyphLeft
    else ItemLayout := tbxlGlyphTop;
  end;

  HasArrow := (tbisSubmenu in TTBXItemAccess(Item).ItemStyle) and
    ((tbisCombo in TTBXItemAccess(Item).ItemStyle) or (tboDropdownArrow in Item.EffectiveOptions));

  if GetImageShown then
  begin
    ImgSize := GetImageSize;
    with ImgSize do if (CX <= 0) or (CY <= 0) then
    begin
      CX := 0;
      CY := 0;
      ImageIsShown := False;
    end
    else ImageIsShown := True;
  end
  else
  begin
    ImgSize.CX := 0;
    ImgSize.CY := 0;
    ImageIsShown := False;
  end;
  ImageOrCheckShown := ImageIsShown or (not ToolbarStyle and Item.Checked);

  StateFlags := GetStateFlags(ItemInfo);

  Canvas.Font := TTBViewAccess(View).GetFont;
  Canvas.Font.Color := CurrentTheme.GetItemTextColor(ItemInfo);
  DoAdjustFont(Canvas.Font, StateFlags);
  C := Canvas.Font.Color;

  { Setup font }
  TextFlags := GetTextFlags;
  IsCaptionShown := CaptionShown;
  IsTextRotated := (View.Orientation = tbvoVertical) and ToolbarStyle;
  if IsCaptionShown then
  begin
    S := GetCaptionText;
    if (Item.Layout <> tbxlAuto) or (tboImageAboveCaption in Item.EffectiveOptions) then
      IsTextRotated := False;
    if IsTextRotated or not ToolbarStyle then TextFlags := TextFlags or DT_SINGLELINE;
    TextSize := GetTextSize(Canvas, S, TextFlags, IsTextRotated, StateFlags);
  end
  else
  begin
    StateFlags := 0;
    SetLength(S, 0);
    IsTextRotated := False;
    TextSize.CX := 0;
    TextSize.CY := 0;
  end;

  IsSpecialDropDown := HasArrow and not IsSplit and ToolbarStyle and
    ((Item.Layout = tbxlGlyphTop) or (Item.Layout = tbxlAuto) and (tboImageAboveCaption in Item.EffectiveOptions)) and
    (ImgSize.CX > 0) and not (IsTextRotated) and (TextSize.CX > 0);

  { Border & Arrows }
  R := ClientAreaRect;
  with CurrentTheme do
  begin
    GetMargins(MID_MENUITEM, Margins);
    CaptionRect := R;

    CurrentTheme.PaintMenuItem(Canvas, CaptionRect, ItemInfo);

    Inc(R.Left, Margins.LeftWidth);
    Dec(R.Right, Margins.RightWidth);
    Inc(R.Top, Margins.TopHeight);
    Dec(R.Bottom, Margins.BottomHeight);
  end;

  { Caption }
  if IsCaptionShown then
  begin
    with CurrentTheme do
    begin
      TextFlags := DT_LEFT or DT_VCENTER or TextFlags;
      TextSize := GetTextSize(Canvas, S, TextFlags, False, StateFlags); { TODO : Check if this line is required }
      GetTextMetrics(Canvas.Handle, TextMetrics);

      CaptionRect := R;
      Inc(CaptionRect.Left, ItemInfo.PopupMargin + MenuImageTextSpace + MenuLeftCaptionMargin);
      with TextMetrics, CaptionRect do
        if (Bottom - Top) - (tmHeight + tmExternalLeading) = Margins.BottomHeight then Dec(Bottom);
      Inc(CaptionRect.Top, TextMetrics.tmExternalLeading);

      CaptionRect.Right := CaptionRect.Left + TextSize.CX;
    end;

    Canvas.Font.Color := C;
    DrawXPText(Canvas, TTBXUnicodeItem(Item).Caption, CaptionRect, TextFlags);
  end;

  { Shortcut and/or submenu arrow (menus only) }
  if not ToolbarStyle then
  begin
    S := Item.GetShortCutText;
    if Length(S) > 0 then
    begin
      CaptionRect := R;
      with CaptionRect, TextMetrics do
      begin
        Left := Right - (Bottom - Top) - GetTextWidth(Canvas.Handle, S, True);
        if (Bottom - Top) - (tmHeight + tmExternalLeading) = Margins.BottomHeight then Dec(Bottom);
        Inc(Top, TextMetrics.tmExternalLeading);
      end;
      Canvas.Font.Color := C;
      DrawXPText(Canvas, TTBXUnicodeItem(Item).Caption, CaptionRect, TextFlags);
    end;
  end;

  { Image, or check box }
  if ImageOrCheckShown then
  begin
    ImageRect := R;

    if ImageIsShown then with ImageRect, ImgSize do
    begin
      Left := Left + ((Right - Left) - CX) div 2;
      ImageRect.Top := Top + ((Bottom - Top) - CY) div 2;
      Right := Left + CX;
      Bottom := Top + CY;
      DrawItemImage(Canvas, ImageRect, ItemInfo);
    end
    else if not ToolbarStyle and Item.Checked then
      CurrentTheme.PaintCheckMark(Canvas, ImageRect, ItemInfo);
  end;
end;

*)
{$ENDIF TBX}


//WMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWM
{ TCustomVirtualShellHistory }

function TCustomVirtualShellHistory.Add(Value: TNamespace; Release: Boolean = False): integer;
begin
  //If ItemIndex is NOT the LastItem then delete all Namespaces between
  //ItemIndex and LastItem (delete all the "Next branch").
  if (Count > 0) and (FItemIndex < Count-1) then
    while Count-1 > FItemIndex do
      Delete(Count-1);
  Result := inherited Add(Value);
end;

procedure TCustomVirtualShellHistory.Back;
begin
  SetItemIndex(ItemIndex - 1);
end;

procedure TCustomVirtualShellHistory.Next;
begin
  SetItemIndex(ItemIndex + 1);
end;

//WMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWM
{ TVSHMenuOptions }

procedure TVSHMenuOptions.Assign(Source: TPersistent);
begin
  if Source is TVSHMenuOptions then
  begin
    Images := TVSHMenuOptions(Source).Images;
    ImageBorder := TVSHMenuOptions(Source).ImageBorder;
    LargeImages := TVSHMenuOptions(Source).LargeImages;
    TextType := TVSHMenuOptions(Source).TextType;
    MaxWidth := TVSHMenuOptions(Source).MaxWidth ;
  end else
    inherited
end;

procedure TVSHMenuOptions.AssignTo(Dest: TPersistent);
begin
  if Dest is TVSHMenuOptions then
  begin
    TVSHMenuOptions(Dest).Images := Images;
    TVSHMenuOptions(Dest).ImageBorder := ImageBorder;
    TVSHMenuOptions(Dest).LargeImages := LargeImages;
    TVSHMenuOptions(Dest).TextType := TextType;
    TVSHMenuOptions(Dest).MaxWidth := MaxWidth;
  end else
    inherited
end;

constructor TVSHMenuOptions.Create;
begin
  FMaxWidth := -1;
  FImageBorder := 1;
  EllipsisPlacement := sseFilePathMiddle
end;

destructor TVSHMenuOptions.Destroy;
begin
  inherited;
end;

//WMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWM
{ TBaseVirtualShellPersistent }

function TBaseVirtualShellPersistent.Add(Value: TNamespace; Release: Boolean = False): integer;
var
  NS: TNamespace;
begin
  Result := -1; // if no item is added it should return -1
  //Check if the LastItem is the same as Value
  if (Count = 0) or (not ILIsEqual(Items[Count-1].AbsolutePIDL, Value.AbsolutePIDL)) then
  begin
    ValidateLevels;  //validate first
    if Release then
      NS := Value
    else
      NS := Value.Clone(True);  //Clone it, we should free it explicitly
    Result := Namespaces.Add(NS);
    ItemIndex := Result;
    DoItemChange(FItemIndex, hctAdded);
  end;
end;

procedure TBaseVirtualShellPersistent.ChangeLinkChanging(Server: TObject; NewPIDL: PItemIDList);
var
  NS: TNamespace;
begin
  //VET informs us that it has changed its Directory
  // Don't add the new change if
  //  1) We initiated the change in the target VET or Combobox
  //  2) We are in a recursive ChangeNotified call
  //  3) The PIDL is not valid
  //  4) The Object that sent the notification was not a VET or a Combobox
  {$IFDEF EXPLORERCOMBOBOX}
   if (not (bvsChangeDispatching in FState)) and (not (bvsChangeNotified in FState)) and
    Assigned(NewPIDL) and ((Server = FVirtualExplorerTree) or (Server = FVirtualExplorerComboBox)) then
  {$ELSE}
   if (not (bvsChangeDispatching in FState)) and (not (bvsChangeNotified in FState)) and
    Assigned(NewPIDL) and ((Server = FVirtualExplorerTree)) then
  {$ENDIF}
  begin
    Include(FState, bvsChangeNotified);
    NS := TNamespace.Create(NewPIDL, nil); //create a temp namespace based on the PIDL
    try
      NS.FreePIDLOnDestroy := False;
      Add(NS); //Add will clone the namespace and will take care of everything
    finally
      NS.Free; //we don't use it anymore
      Exclude(FState, bvsChangeNotified);
    end
  end;
end;

procedure TBaseVirtualShellPersistent.ChangeLinkDispatch(PIDL: PItemIDList);
begin
  //Informs all the Client VETs that we have selected a new Directory
  if not (bvsChangeNotified in FState) then
  begin
    Include(FState, bvsChangeDispatching);
    VETChangeDispatch.DispatchChange(Self, PIDL);
    Exclude(FState, bvsChangeDispatching);
  end
end;

procedure TBaseVirtualShellPersistent.ChangeLinkFreeing(ChangeLink: IVETChangeLink);
begin
  {$IFDEF EXPLORERCOMBOBOX}
  if ChangeLink.ChangeLinkClient = Self then
  begin
    if ChangeLink.ChangeLinkServer = FVirtualExplorerTree then
      FVirtualExplorerTree := nil
    else
    if ChangeLink.ChangeLinkServer = FVirtualExplorerComboBox then
      FVirtualExplorerComboBox := nil
  end
  {$ELSE}
  if ChangeLink.ChangeLinkClient = Self then
  begin
    if ChangeLink.ChangeLinkServer = FVirtualExplorerTree then
      FVirtualExplorerTree := nil
  end
  {$ENDIF}
end;

procedure TBaseVirtualShellPersistent.Clear;
begin
  while Count > 0 do
    Delete(0); // Make sure we fire all events
end;

constructor TBaseVirtualShellPersistent.Create(AOwner: TComponent);
begin
  inherited;
  FNamespaces := TList.Create;
  FMenuOptions := CreateMenuOptions;
  FLevels := 10;
end;

function TBaseVirtualShellPersistent.CreateMenuOptions: TVSHMenuOptions;
begin
  Result := TVSHMenuOptions.Create;
end;

procedure TBaseVirtualShellPersistent.Delete(Index: integer);
var
  Temp: TNamespace;
  Changed: Boolean;
begin
  if (Index > -1) and (Index < Count) then
  begin
    Temp := TNamespace( Namespaces[Index]);
    Namespaces.Delete(Index);
    Temp.Free;
    DoItemChange(Index, hctDeleted);
    // The deleted item is the currently selected item need to change it
    if Index <= ItemIndex then
    begin
      // If the ItemIndex items is the one deleted then it is effectivly selecting
      // the next lower item
      Changed := Index = ItemIndex;
      // If a lower than selected item or the selected item is deleted the rest
      // will be shifted down so the index effectively is lowered by 1 (unless it
      // is the first item) although the actual selected item is the same.
      if Index > 0 then
        Dec(FItemIndex);

      if Changed then
        DoItemChange(FItemIndex, hctSelected);
    end
  end;
  if Count = 0 then
    SetItemIndex(-1);
end;

destructor TBaseVirtualShellPersistent.Destroy;
begin
  VETChangeDispatch.UnRegisterChangeLink(Self, Self, utAll);
  Clear;
  FNamespaces.Free;
  FreeAndNil(FMenuOptions); 
  inherited;
end;

procedure TBaseVirtualShellPersistent.DoGetImage(NS: TNamespace;
  var ImageList: TImageList; var ImageIndex: Integer);
begin
  if Assigned(OnGetImage) then
    OnGetImage(Self, NS, ImageList, ImageIndex);
end;

procedure TBaseVirtualShellPersistent.DoItemChange(ItemIndex: integer; ChangeType: TVSHChangeType);
begin
  if Assigned(OnChange) and not(csDestroying in ComponentState) then
    OnChange(Self, ItemIndex, ChangeType);
end;

procedure TBaseVirtualShellPersistent.FillPopupMenu(Popupmenu: TPopupMenu;
  FillDirection: TFillPopupDirection; AddClearItem: Boolean = False; ClearItemText: WideString = '');
//Fills a TPopupMenu to mimic the Explorer's Back and Next Buttons PopupMenu.
//Depending on the value of the BackPopup boolean the PopupMenu is filled with
//the corresponding Back or Next Namespaces folder names.
//When UnicodeEnabled is true the PopupMenu is OwnerDrawed to draw the widestrings.

  procedure AddToPopup(AIndex: integer);
  var
    M: TMenuItem;
  begin
    M := TMenuItem.Create(PopupMenu);
    M.Caption := Items[AIndex].NameInFolder;
    M.Tag := AIndex; //this represents the real MenuItem index (the back popupmenu is upside down)
    M.OnClick := OnMenuItemClick;
    M.OnDrawItem := OnMenuItemDraw;
    M.OnMeasureItem := OnMenuItemMeasure;
    Popupmenu.Items.Add(M);
  end;

  procedure AddClear;
  var
    M: TMenuItem;
  begin
    // Draw a divider line and let VCL draw it
    M := TMenuItem.Create(PopupMenu);
    M.Caption := '-';
    M.Tag := -1; //this represents the real MenuItem index (the back popupmenu is upside down)
    Popupmenu.Items.Add(M);

    M := TMenuItem.Create(PopupMenu);
    M.Caption := ClearItemText;
    M.Tag := -1; //this represents the real MenuItem index (the back popupmenu is upside down)
    M.OnClick := OnMenuItemClick;
    M.OnDrawItem := OnMenuItemDraw;
    M.OnMeasureItem := OnMenuItemMeasure;
    Popupmenu.Items.Add(M);
  end;

var
  i: integer;
begin
  {$IFNDEF DELPHI_5_UP}
  ClearMenuItems(Popupmenu);
  {$ELSE}
  Popupmenu.Items.Clear;
  {$ENDIF DELPHI_5_UP} 
  // Don't use the currently selected item in the MRU list
  if Count > 1 then
  begin
    if FillDirection = fpdOldestToNewest then
    begin
      if Self is TCustomVirtualShellHistory then
      begin
        for i := FItemIndex + 1 to Count - 1 do
          AddToPopup(i); //upside down
      end else
      if Self is TBaseVirtualShellPersistent then // This is true for ShellHistory too so must be last
      begin
        for i := 0 to Count - 2 do
          AddToPopup(i);
      end
    end else
    begin
      if Self is TCustomVirtualShellHistory then
      begin
        for i := FItemIndex - 1 downto 0 do
          AddToPopup(i); //upside down
      end else
      if Self is TBaseVirtualShellPersistent then // This is true for ShellHistory too so must be last
      begin
        for i := Count - 2 downto 0 do
        AddToPopup(i);
      end
    end;
    if AddClearItem then
      AddClear;

    PopupMenu.OwnerDraw := true;
  end;
end;

{$IFDEF TBX_OR_TB2K}
procedure TBaseVirtualShellPersistent.FillPopupMenu_TBX(PopupMenu: TTBCustomItem;
  FillDirection: TFillPopupDirection; AddClearItem: Boolean; ClearItemText: WideString);
// Fills a TTBXCustomItem to mimic the Explorer's Back and Next Buttons TBX Item.
// Depending on the value of the PopupType the PopupMenu is filled with
// the corresponding Back, Next or All Namespaces folder names.

  procedure AddToPopup(AIndex: integer);
  var
    M: TTBCustomItem;
    NS: TNamespace;
  begin
    NS := Items[AIndex];
   {$IFDEF TBX}
    M := TTBXUnicodeItem.Create(PopupMenu);
    if MenuOptions.TextType = mttName then
      TTBXUnicodeItem(M).Caption := NS.NameInFolder
    else
      TTBXUnicodeItem(M).Caption := NS.NameParseAddress;
   {$ELSE}
    M := TTBItem.Create(PopupMenu);
    if MenuOptions.TextType = mttName then
      M.Caption := NS.NameInFolder
    else
      M.Caption := NS.NameParseAddress;
   {$ENDIF}
    M.Tag := AIndex; //this represents the real MenuItem index (the back popupmenu is upside down)
    M.ImageIndex := NS.GetIconIndex(False, icSmall);
    M.Images := SmallSysImages;
    M.OnClick := OnMenuItemClick;

    Popupmenu.Add(M);
  end;


  procedure AddClear;
  var
    M: TTBXCustomItem;
    SI: TTBXSeparatorItem;
  begin
   {$IFDEF TBX}
    SI := TTBXSeparatorItem.Create(PopupMenu);
    M := TTBXItem.Create(PopupMenu);
   {$ELSE}
    SI := TTBSeparatorItem.Create(PopupMenu);
    M := TTBItem.Create(PopupMenu);
   {$ENDIF}

    // Draw a divider line and let VCL draw it
    SI.Tag := -1; //this represents the real MenuItem index (the back popupmenu is upside down)
    PopupMenu.Add(SI);

    M.Caption := ClearItemText;
    M.Tag := -1; //this represents the real MenuItem index (the back popupmenu is upside down)
    M.OnClick := OnMenuItemClick;
    M.Images := SmallSysImages;
    PopupMenu.Add(M);
  end;

var
  i: integer;
begin
  PopupMenu.Clear;
  // Don't use the currently selected item in the MRU list
  if Count > 1 then
  begin
    if FillDirection = fpdOldestToNewest then
    begin
      if Self is TCustomVirtualShellHistory then
      begin
        for i := FItemIndex + 1 to Count - 1 do
          AddToPopup(i); //upside down
      end else
      if Self is TBaseVirtualShellPersistent then // This is true for ShellHistory too so must be last
      begin
        for i := 0 to Count - 2 do
          AddToPopup(i);
      end
    end else
    begin
      if Self is TCustomVirtualShellHistory then
      begin
        for i := FItemIndex - 1 downto 0 do
          AddToPopup(i); //upside down
      end else
      if Self is TBaseVirtualShellPersistent then // This is true for ShellHistory too so must be last
      begin
        for i := Count - 2 downto 0 do
        AddToPopup(i);
      end
    end;
    if AddClearItem then
      AddClear;
  end;
end;
{$ENDIF TBX_OR_TB2K}

function TBaseVirtualShellPersistent.GetCount: integer;
begin
  Result := Namespaces.Count
end;

function TBaseVirtualShellPersistent.GetHasBackItems: Boolean;
begin
  Result := ItemIndex > 0
end;

function TBaseVirtualShellPersistent.GetHasNextItems: Boolean;
begin
  Result := ItemIndex < Count - 1;
end;

function TBaseVirtualShellPersistent.GetItems(Index: integer): TNamespace;
begin
  if (Index > -1) and (Index < Count) then
    Result := TNamespace(Namespaces[Index])
  else
    Result := nil
end;

function TBaseVirtualShellPersistent.GetLargeSysImages: TImageList;
begin
  Result := IDEVirtualSystemImageLists.LargeSysImages;
end;

function TBaseVirtualShellPersistent.GetSmallSysImages: TImageList;
begin
  Result := IDEVirtualSystemImageLists.SmallSysImages;
end;

procedure TBaseVirtualShellPersistent.LoadFromFile(FileName: WideString);
var
  S: TWideFileStream;
begin
  S := TWideFileStream.Create(FileName, fmOpenRead or fmShareExclusive);
  try
    LoadFromStream(S)
  finally
    S.Free
  end
end;

procedure TBaseVirtualShellPersistent.LoadFromRegistry(RootKey: DWORD; SubKey: string);
var
  Reg: TRegistry;
  Stream: TMemoryStream;
begin
  Reg := TRegistry.Create;
  Stream := TMemoryStream.Create;
  try
    Reg.RootKey := RootKey;
    if Reg.OpenKey(SubKey, False) and Reg.ValueExists(REGISTRYDATASIZE) then
    begin
      Stream.Size := Reg.ReadInteger(REGISTRYDATASIZE);
      Reg.ReadBinaryData(REGISTRYDATA, Stream.Memory^, Stream.Size);
      LoadFromStream(Stream);
    end
  finally
    Reg.CloseKey;
    Reg.Free;
    Stream.Free;
    inherited;
  end
end;

procedure TBaseVirtualShellPersistent.LoadFromStream(S: TStream);
var
  C, I: integer;
begin
  Include(FState, bvsChangeItemsLoading);
  try
    Clear;
    S.ReadBuffer(C, SizeOf(C));
    for I := 0 to C - 1 do begin
      // Dispatch the item change only for the last item
      if I = C - 1 then
        Exclude(FState, bvsChangeItemsLoading);
      Add(TNamespace.Create(PIDLMgr.LoadFromStream(S), nil), True);
    end;
  finally
    Exclude(FState, bvsChangeItemsLoading);
  end;
end;

procedure TBaseVirtualShellPersistent.OnMenuItemClick(Sender: TObject);
var
  M: TComponent;
  OldFocus: TWinControl;
begin
  OldFocus := Screen.ActiveForm.ActiveControl;

  M := nil;

  if (Sender is TMenuItem) then
    M := TMenuItem(Sender)
  else begin
    {$IFDEF TOOLBAR_2000}
    if not Assigned(M) and (Sender is TTBCustomItem) then
      M := TTBCustomItem(Sender);
    {$ENDIF}

    {$IFDEF TBX}
    if not Assigned(M) and (Sender is TTBXCustomItem) then
      M := TTBXCustomItem(Sender);
    {$ENDIF}
  end;

  if Assigned(M) then
  begin
    M := TMenuItem(Sender);
    if M.Tag > -1 then
      SetItemIndex(M.Tag)
    else
      Clear;
    if Assigned(OldFocus) then
      Screen.ActiveForm.SetFocusedControl(OldFocus)
  end;
end;

procedure TBaseVirtualShellPersistent.OnMenuItemDraw(Sender: TObject;
  ACanvas: TCanvas; ARect: TRect; Selected: Boolean);
var
  WS: WideString;
  S: string;
  i, Border: integer;
  ImageRect: TRect;
  TargetImageIndex: Integer;
  TargetImageList: TImageList;
  RTL: Boolean; // Left to Right reading
  OldMode: Longint;
begin
  if Sender is TMenuItem then
  begin
    RTL := Application.UseRightToLeftReading;
    i := TMenuItem(Sender).Tag;

    if IsWinXP and Selected then
    begin
      ACanvas.Brush.Color := clHighlight;
      ACanvas.Font.Color := clBtnHighlight;
      ACanvas.FillRect(ARect);
    end else
    begin
      ACanvas.Brush.Color := clMenu;
      ACanvas.Font.Color := clMenuText;
      ACanvas.FillRect(ARect);
    end;

    if i > -1 then
    begin
      if MenuOptions.TextType = mttName then
        WS := Items[i].NameInFolder
      else
        if MenuOptions.TextType = mttPath then
          WS := Items[i].NameParseAddress;
    end
    else
      WS := TMenuItem(Sender).Caption;

    if MenuOptions.Images then
    begin
      TargetImageIndex := -1;
      Border := MenuOptions.ImageBorder;
      if MenuOptions.LargeImages then
      begin
        TargetImageList := LargeSysImages;
        if i > -1 then
          TargetImageIndex := Items[i].GetIconIndex(False, icLarge)
      end else
      begin
        TargetImageList := SmallSysImages;
        if i > -1 then
          TargetImageIndex := Items[i].GetIconIndex(False, icSmall);
      end;
      // Allow custom icons
      if i > -1 then
        DoGetImage(Namespaces[i], TargetImageList, TargetImageIndex);


      if RTL then
        ImageRect := Rect(ARect.Right - (TargetImageList.Width + 2 * Border), ARect.Top, ARect.Right, ARect.Bottom)
      else
        ImageRect := Rect(ARect.Left, ARect.Top, ARect.Left + TargetImageList.Width + 2 * Border, ARect.Bottom);

      if Selected and not IsWinXP then
        DrawEdge(ACanvas.Handle, ImageRect, BDR_RAISEDINNER, BF_RECT);

      OffsetRect(ImageRect, Border, ((ImageRect.Bottom - ImageRect.Top) - TargetImageList.Height) div 2);

      ImageList_Draw(TargetImageList.Handle, TargetImageIndex, ACanvas.Handle,
        ImageRect.Left, ImageRect.Top, ILD_TRANSPARENT);

      if RTL then
        ARect.Right := ARect.Right - (TargetImageList.Width + (2 * Border) + 1)
      else
        ARect.Left := ARect.Left + TargetImageList.Width + (2 * Border) + 1
    end;

    if Selected and not IsWinXP then
    begin
      ACanvas.Brush.Color := clHighlight;
      ACanvas.Font.Color := clBtnHighlight;
      ACanvas.FillRect(ARect);
    end;

    Inc(ARect.Left, 2);
    if TextExtentW(WS, ACanvas).cx > ARect.Right-ARect.Left then
      WS := ShortenStringEx(ACanvas.Handle, WS, ARect.Right-ARect.Left, RTL, MenuOptions.EllipsisPlacement);
    OldMode := SetBkMode(ACanvas.Handle, TRANSPARENT);
    // Remove the & chars
    i := Pos('&', WS);
    System.Delete(WS, i, 1);
    if IsUnicode then
      DrawTextW_VST(ACanvas.handle, PWideChar(WS), StrLenW(PWideChar(WS)), ARect, DT_SINGLELINE or DT_VCENTER)
    else begin
      S := WS;
      DrawText(ACanvas.handle, PChar(S), Length(S), ARect, DT_SINGLELINE or DT_VCENTER)
    end;
    SetBkMode(ACanvas.Handle, OldMode);
    //Note: it seems that DrawTextW doesn't draw the prefix.
  end;
end;

procedure TBaseVirtualShellPersistent.OnMenuItemMeasure(Sender: TObject;
  ACanvas: TCanvas; var Width, Height: Integer);
var
  WS: WideString;
  i: integer;
  Border: Integer;
begin
  if Sender is TMenuItem then
  begin
    i := TMenuItem(Sender).Tag;

    if i > -1 then
    begin
      if MenuOptions.TextType = mttName then
        WS := Items[i].NameInFolder
      else
        WS := Items[i].NameParseAddress;
      Width := TextExtentW(WS, ACanvas).cx;
    end else
      Width := TextExtentW(TMenuItem(Sender).Caption, ACanvas).cx;

    if MenuOptions.Images then
    begin
      Border := 2 * MenuOptions.ImageBorder;
      if MenuOptions.LargeImages then
      begin
        Inc(Width, LargeSysImages.Width + Border);
        if LargeSysImages.Height + Border > Height then
          Height := LargeSysImages.Height + Border
      end else
      begin
        Inc(Width, SmallSysImages.Width + Border);
        if SmallSysImages.Height + Border > Height then
          Height := SmallSysImages.Height + Border
      end
    end;
  end;

  if MenuOptions.MaxWidth > 0 then
  begin
    if Width > MenuOptions.MaxWidth then
      Width := MenuOptions.MaxWidth;
  end;

  if Width > Screen.Width then
    Width := Screen.Width - 12  // Purely imperical value seen on XP for the unaccessable borders
end;

procedure TBaseVirtualShellPersistent.SaveToFile(FileName: WideString);
var
  S: TWideFileStream;
begin
  S := TWideFileStream.Create(FileName, fmCreate or fmShareExclusive);
  try
    SaveToStream(S)
  finally
    S.Free
  end
end;

procedure TBaseVirtualShellPersistent.SaveToRegistry(RootKey: DWORD; SubKey: string);
var
  Reg: TRegistry;
  Stream: TMemoryStream;
begin
  Reg := TRegistry.Create;
  Stream := TMemoryStream.Create;
  try
    Reg.RootKey := RootKey;
    if Reg.OpenKey(SubKey, True) then
    begin
      SaveToStream(Stream);
      Reg.WriteInteger(REGISTRYDATASIZE, Stream.Size);
      Reg.WriteBinaryData(REGISTRYDATA, Stream.Memory^, Stream.Size)
    end
  finally
    Reg.CloseKey;
    Reg.Free;
    Stream.Free;
    inherited;
  end
end;

procedure TBaseVirtualShellPersistent.SaveToStream(S: TStream);
var
  i: integer;
begin
  S.WriteBuffer(Namespaces.Count, SizeOf(Count));
  for i := 0 to Count - 1 do
    PIDLMgr.SaveToStream(S, TNamespace(Namespaces[i]).AbsolutePIDL);
end;

procedure TBaseVirtualShellPersistent.SetItemIndex(Value: integer);
var
  PrevItemIndex: integer;
  NS: TNamespace;
begin
  if Value < 0 then Value := 0
  else
    if Value > Count - 1 then Value := Count - 1;

  if FItemIndex <> Value then
  begin
    PrevItemIndex := FItemIndex;
    if Count = 0 then
      FItemIndex := -1
    else begin
      FItemIndex := Value;
      //Inform VET that we have selected a new Directory
      if not (bvsChangeItemsLoading in FState) then begin
        NS := Items[FItemIndex];
        // Delete the item if it's invalid
        if Assigned(NS) and NS.FileSystem and (NS.Extension <> '.zip') and not DirExistsW(NS.NameForParsing) then begin
          Delete(FItemIndex);
          if PrevItemIndex > Count - 1 then PrevItemIndex := Count - 1;
          FItemIndex := PrevItemIndex;
          DoItemChange(FItemIndex, hctInvalidSelected);
          Exit;
        end
        else begin
          ChangeLinkDispatch(NS.AbsolutePIDL);
          DoItemChange(FItemIndex, hctSelected);
        end;
      end;
    end;
  end;
end;

procedure TBaseVirtualShellPersistent.SetLevels(const Value: Integer);
begin
  if FLevels <> Value then
  begin
    FLevels := Value;
    ValidateLevels;
  end;
end;

procedure TBaseVirtualShellPersistent.SetMenuOptions(const Value: TVSHMenuOptions);
begin
  if Assigned(FMenuOptions) then
    FMenuOptions.Free;
  FMenuOptions := Value;
end;

{$IFDEF EXPLORERCOMBOBOX}
procedure TBaseVirtualShellPersistent.SetVirtualExplorerComboBox(const Value: TCustomVirtualExplorerCombobox);
begin
  if FVirtualExplorerComboBox <> Value then
  begin
    VirtualExplorerTree := nil;
    if Assigned(FVirtualExplorerComboBox) then
    begin
      VETChangeDispatch.UnRegisterChangeLink(FVirtualExplorerComboBox, Self, utLink);
      VETChangeDispatch.UnRegisterChangeLink(Self, FVirtualExplorerComboBox, utLink);
    end;
    FVirtualExplorerComboBox := Value;
    if Assigned(FVirtualExplorerComboBox) then
    begin
      //two way dispaching
      VETChangeDispatch.RegisterChangeLink(FVirtualExplorerComboBox, Self, ChangeLinkChanging, ChangeLinkFreeing);
      VETChangeDispatch.RegisterChangeLink(Self, FVirtualExplorerComboBox, FVirtualExplorerComboBox.ChangeLinkChanging, nil);
    end;
  end;
end;
{$ENDIF}

procedure TBaseVirtualShellPersistent.SetVirtualExplorerTree(const Value: TCustomVirtualExplorerTree);
begin
  if FVirtualExplorerTree <> Value then
  begin
    {$IFDEF EXPLORERCOMBOBOX}
    VirtualExplorerComboBox := nil;
    {$ENDIF}
    if Assigned(FVirtualExplorerTree) then
    begin
      VETChangeDispatch.UnRegisterChangeLink(FVirtualExplorerTree, Self, utLink);
      VETChangeDispatch.UnRegisterChangeLink(Self, FVirtualExplorerTree, utLink);
    end;
    FVirtualExplorerTree := Value;
    if Assigned(FVirtualExplorerTree) then
    begin
      //two way dispaching
      VETChangeDispatch.RegisterChangeLink(FVirtualExplorerTree, Self, ChangeLinkChanging, ChangeLinkFreeing);
      VETChangeDispatch.RegisterChangeLink(Self, FVirtualExplorerTree, FVirtualExplorerTree.ChangeLinkChanging, nil);
    end;
  end;
end;

procedure TBaseVirtualShellPersistent.ValidateLevels;
begin
  while Count >= Levels do
    Delete(0); //delete will fire the change event, free the namespace properly and check the itemindex
end;

end.
