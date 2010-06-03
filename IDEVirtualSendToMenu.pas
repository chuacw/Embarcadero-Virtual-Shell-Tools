unit IDEVirtualSendToMenu;

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
// The initial developer of this code is Jim Kueneman <jimdk@mindspring.com>
//
//----------------------------------------------------------------------------

interface

{$include Compilers.inc}
{$include VSToolsAddIns.inc}

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms,
  Menus, Registry, ShlObj, ShellAPI, ActiveX, IDEVirtualShellUtilities, ImgList,
  CommCtrl, IDEVirtualResources, IDEVirtualWideStrings, IDEVirtualSystemImageLists,
  IDEVirtualShellContainers, IDEVirtualUtilities,
  {$IFDEF TOOLBAR_2000} TB2Item, {$ENDIF TOOLBAR_2000}
  {$IFDEF TBX} TB2Item, TBX, {$ENDIF TBX}
  IDEVirtualUnicodeDefines;

type
  TVirtualSendToMenu = class;  // Forward

  TSendToMenuOptions = class(TPersistent)
  private
    FImages: Boolean;
    FLargeImages: Boolean;
    FImageBorder: Integer;
    FMaxWidth: Integer;
    FEllipsisPlacement: TShortenStringEllipsis;
  public
    constructor Create;
  published
    property EllipsisPlacement: TShortenStringEllipsis read FEllipsisPlacement write FEllipsisPlacement default sseMiddle;
    property Images: Boolean read FImages write FImages default True;
    property LargeImages: Boolean read FLargeImages write FLargeImages default False;
    property ImageBorder: Integer read FImageBorder write FImageBorder default 1;
    property MaxWidth: Integer read FMaxWidth write FMaxWidth default -1;
  end;

  TVirtualSendToEvent = procedure(Sender: TVirtualSendToMenu;
    SendToTarget: TNamespace; var SourceData: IDataObject) of object;
  TVirtualSendToGetImageEvent = procedure(Sender: TVirtualSendToMenu;
    NS: TNamespace; var ImageList: TImageList; var ImageIndex: Integer) of object;

  TVirtualSendToMenuItem = class(TMenuItem)
  private
    FNamespace: TNamespace;
  protected
    property Namespace: TNamespace read FNamespace write FNamespace;
  public
    destructor Destroy; override;
    procedure Click; override;
  end;

  {$IFDEF TBX_OR_TB2K}
  {$IFDEF TBX}
  TVirtualSendToMenuItem_TB2000 = class(TTBXItem)
  {$ELSE}
  TVirtualSendToMenuItem_TB2000 = class(TTBItem)
  {$ENDIF TBX}
  private
    FNamespace: TNamespace;
  protected
    property Namespace: TNamespace read FNamespace write FNamespace;
  public
    destructor Destroy; override;
    procedure Click; override;
  end;
  {$ENDIF TBX_OR_TB2K}

  TVirtualSendToMenu = class(TPopupMenu)
  private
    FSendToItems: TVirtualNameSpaceList;
    FSendToEvent: TVirtualSendToEvent;
    FOptions: TSendToMenuOptions;
    FOnGetImage: TVirtualSendToGetImageEvent;
  protected
    procedure DoGetImage(NS: TNamespace; var ImageList: TImageList; var ImageIndex: Integer);
    procedure DoSendTo(SendToTarget: TNamespace; var SourceData: IDataObject); virtual;
    function EnumSendToCallback(APIDL: PItemIDList; AParent: TNamespace;
      Data: Pointer; var Terminate: Boolean): Boolean;
    procedure OnMenuItemDraw(Sender: TObject; ACanvas: TCanvas; ARect: TRect; Selected: Boolean); virtual;
    procedure OnMenuItemMeasure(Sender: TObject; ACanvas: TCanvas; var Width, Height: Integer); virtual;
  public
    constructor Create(AOwner: TComponent); override;
    destructor Destroy; override;
    procedure Populate(MenuItem: TMenuItem); virtual;
    {$IFDEF TBX_OR_TB2K}
    procedure Populate_TB2000(MenuItem: TTBCustomItem); virtual;
    {$ENDIF TBX_OR_TB2K}
    procedure Popup(X, Y: Integer); override;
    property SendToItems: TVirtualNameSpaceList read FSendToItems;
  published
    property SendToEvent: TVirtualSendToEvent read FSendToEvent write FSendToEvent;
    property OnGetImage: TVirtualSendToGetImageEvent read FOnGetImage write FOnGetImage;
    property Options: TSendToMenuOptions read FOptions write FOptions;
  end;

implementation

function SendToMenuSort(Item1, Item2: Pointer): Integer;
begin
  if Assigned(Item1) and Assigned(Item2) then
    Result := TNamespace(Item2).ComparePIDL(TNamespace(Item1).RelativePIDL, False)
  else
    Result := 0
end;

{ TVirtualSendToMenuItem }

procedure TVirtualSendToMenuItem.Click;
var
  Menu: TVirtualSendToMenu;
  DataObject: IDataObject;
  DropTarget: IDropTarget;
  DropEffect: Longint;
begin
  inherited;
  Menu := Owner as TVirtualSendToMenu;
  Menu.DoSendTo(Namespace, DataObject);
  if Assigned(DataObject) then
  begin
    DropEffect := DROPEFFECT_COPY or DROPEFFECT_MOVE or DROPEFFECT_LINK;
    DropTarget := Namespace.DropTargetInterface;
    if Assigned(DropTarget) then
      if Succeeded(DropTarget.DragEnter(DataObject, MK_LBUTTON, Point(0, 0), DropEffect)) then
        (DropTarget.Drop(DataObject, MK_LBUTTON, Point(0, 0), DropEffect))
  end
end;

destructor TVirtualSendToMenuItem.Destroy;
begin
  Namespace.Free;
  inherited;
end;

{$IFDEF TBX_OR_TB2K}
{ TVirtualSendToMenuItem_TB2000 }

procedure TVirtualSendToMenuItem_TB2000.Click;
var
  Menu: TVirtualSendToMenu;
  DataObject: IDataObject;
  DropTarget: IDropTarget;
  DropEffect: Longint;
begin
  inherited;
  Menu := Owner as TVirtualSendToMenu;
  Menu.DoSendTo(Namespace, DataObject);
  if Assigned(DataObject) then
  begin
    DropEffect := DROPEFFECT_COPY or DROPEFFECT_MOVE or DROPEFFECT_LINK;
    DropTarget := Namespace.DropTargetInterface;
    if Assigned(DropTarget) then
      if Succeeded(DropTarget.DragEnter(DataObject, MK_LBUTTON, Point(0, 0), DropEffect)) then
        (DropTarget.Drop(DataObject, MK_LBUTTON, Point(0, 0), DropEffect))
  end
end;

destructor TVirtualSendToMenuItem_TB2000.Destroy;
begin
  Namespace.Free;
  inherited;
end;

{$ENDIF TBX_OR_TB2K}

{ TVirtualSendToMenu }

constructor TVirtualSendToMenu.Create(AOwner: TComponent);
begin
  inherited;
  FSendToItems := TVirtualNameSpaceList.Create(False);
  FOptions := TSendToMenuOptions.Create;
end;

destructor TVirtualSendToMenu.Destroy;
begin
  SendToItems.Free;
  FOptions.Free;
  inherited;
end;

procedure TVirtualSendToMenu.DoGetImage(NS: TNamespace;
  var ImageList: TImageList; var ImageIndex: Integer);
begin
  if Assigned(OnGetImage) then
    OnGetImage(Self, NS, ImageList, ImageIndex);
end;

procedure TVirtualSendToMenu.DoSendTo(SendToTarget: TNamespace;
   var SourceData: IDataObject);
begin
  SourceData := nil;
  if Assigned(SendToEvent) then
    SendToEvent(Self, SendToTarget, SourceData);
end;

function TVirtualSendToMenu.EnumSendToCallback(APIDL: PItemIDList;
  AParent: TNamespace; Data: Pointer; var Terminate: Boolean): Boolean;
var
  NS: TNamespace;
begin
  if AParent.IsMyComputer then
  begin
    NS := TNamespace.Create(PIDLMgr.AppendPIDL(AParent.AbsolutePIDL, APIDL), nil);
    if NS.Removable then
      TVirtualNameSpaceList(Data).Add(NS)
    else
      NS.Free
  end else
    TVirtualNameSpaceList(Data).Add(TNamespace.Create(PIDLMgr.AppendPIDL(AParent.AbsolutePIDL, APIDL), nil));
  Result := True
end;

procedure TVirtualSendToMenu.OnMenuItemDraw(Sender: TObject;
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
  MenuItem: TVirtualSendToMenuItem;
begin
  if Sender is TMenuItem then
  begin
    RTL := Application.UseRightToLeftReading;
    i := TMenuItem(Sender).MenuIndex;

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

    MenuItem := TVirtualSendToMenuItem( Items[i]);

    if i > -1 then
      WS := MenuItem.Namespace.NameInFolder
    else
      WS := TMenuItem(Sender).Caption;

    if Options.Images then
    begin
      TargetImageIndex := -1;
      Border := Options.ImageBorder;
      if Options.LargeImages then
      begin
        TargetImageList := LargeSysImages;
        if i > -1 then
          TargetImageIndex := MenuItem.Namespace.GetIconIndex(False, icLarge)
      end else
      begin
        TargetImageList := SmallSysImages;
        if i > -1 then
          TargetImageIndex := MenuItem.Namespace.GetIconIndex(False, icSmall);
      end;
      // Allow custom icons
      if i > -1 then
        DoGetImage(MenuItem.Namespace, TargetImageList, TargetImageIndex);


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
      WS := ShortenStringEx(ACanvas.Handle, WS, ARect.Right-ARect.Left, RTL, Options.EllipsisPlacement);
    OldMode := SetBkMode(ACanvas.Handle, TRANSPARENT);
    // The are WideString compatible
    //Note: it seems that DrawTextW doesn't draw the prefix.
    i := Pos('&', WS);
    System.Delete(WS, i, 1);
    if IsUnicode then
      DrawTextW_VST(ACanvas.handle, PWideChar(WS), StrLenW(PWideChar(WS)), ARect, DT_SINGLELINE or DT_VCENTER)
    else begin
      S := WS;
      DrawText(ACanvas.handle, PChar(S), Length(S), ARect, DT_SINGLELINE or DT_VCENTER)
    end;
    SetBkMode(ACanvas.Handle, OldMode);
  end;
end;

procedure TVirtualSendToMenu.OnMenuItemMeasure(Sender: TObject;
  ACanvas: TCanvas; var Width, Height: Integer);
var
  WS: WideString;
  i: integer;
  Border: Integer;
begin
  if Sender is TMenuItem then
  begin
    i := TMenuItem(Sender).MenuIndex;

    if i > -1 then
    begin
      WS := TVirtualSendToMenuItem( Items[i]).Namespace.NameInFolder;
      Width := TextExtentW(WS, ACanvas).cx;
    end else
      Width := TextExtentW(TMenuItem(Sender).Caption, ACanvas).cx;

    if Options.Images then
    begin
      Border := 2 * Options.ImageBorder;
      if Options.LargeImages then
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

  if Options.MaxWidth > 0 then
  begin
    if Width > Options.MaxWidth then
      Width := Options.MaxWidth;
  end;

  if Width > Screen.Width then
    Width := Screen.Width - 12  // Purely imperical value seen on XP for the unaccessable borders
end;

procedure TVirtualSendToMenu.Populate(MenuItem: TMenuItem);
var
  NS: TNamespace;
  M: TVirtualSendToMenuItem;
  i: integer;
  OldErrorMode: integer;
begin
  OldErrorMode := SetErrorMode(SEM_FAILCRITICALERRORS or SEM_NOOPENFILEERRORBOX);
  try
    SendToItems.Clear;
    NS := CreateSpecialNamespace(CSIDL_SENDTO);
    NS.EnumerateFolder(False, True, False, EnumSendToCallback, SendToItems);
    SendToItems.Sort(SendToMenuSort);
    for i := 0 to SendToItems.Count - 1 do
    begin
      M := TVirtualSendToMenuItem.Create(Self);
      M.Namespace := SendToItems[i];
      M.Caption := M.Namespace.NameNormal;
      M.ImageIndex := M.Namespace.GetIconIndex(False, icSmall);
      MenuItem.Add(M);
    end;
    SendToItems.Clear;
    NS.Free;

    DrivesFolder.EnumerateFolder(False, False, False, EnumSendToCallback, SendToItems);
    SendToItems.Sort(SendToMenuSort);
    for i := 0 to SendToItems.Count - 1 do
    begin
      M := TVirtualSendToMenuItem.Create(Self);
      M.Namespace := SendToItems[i];
      M.Caption := M.Namespace.NameNormal;
      M.ImageIndex := M.Namespace.GetIconIndex(False, icSmall);
      MenuItem.Add(M);
    end;
    SendToItems.Clear;
  finally
    SetErrorMode(OldErrorMode);
  end;
end;

{$IFDEF TBX_OR_TB2K}
procedure TVirtualSendToMenu.Populate_TB2000(MenuItem: TTBCustomItem);
var
  NS: TNamespace;
  M: TVirtualSendToMenuItem_TB2000;
  i: integer;
begin
  {$IFDEF TBX}
  if MenuItem is TTBXSubmenuItem then
    TTBXSubmenuItem(MenuItem).SubMenuImages := SmallSysImages;
  {$ELSE}
  if MenuItem is TTBSubmenuItem then
    TTBSubmenuItem(MenuItem).SubMenuImages := SmallSysImages;
  {$ENDIF TBX}
  MenuItem.Clear;

  SendToItems.Clear;
  NS := CreateSpecialNamespace(CSIDL_SENDTO);
  NS.EnumerateFolder(False, True, False, EnumSendToCallback, SendToItems);
  SendToItems.Sort(SendToMenuSort);
  for i := 0 to SendToItems.Count - 1 do
  begin
    M := TVirtualSendToMenuItem_TB2000.Create(Self);
    M.Namespace := SendToItems[i];
    M.Caption := M.Namespace.NameNormal;
    M.ImageIndex := M.Namespace.GetIconIndex(False, icSmall);
    MenuItem.Add(M);
  end;
  SendToItems.Clear;
  NS.Free;

  DrivesFolder.EnumerateFolder(False, False, False, EnumSendToCallback, SendToItems);
  SendToItems.Sort(SendToMenuSort);
  for i := 0 to SendToItems.Count - 1 do
  begin
    M := TVirtualSendToMenuItem_TB2000.Create(Self);
    M.Namespace := SendToItems[i];
    M.Caption := M.Namespace.NameNormal;
    M.ImageIndex := M.Namespace.GetIconIndex(False, icSmall);
    MenuItem.Add(M);
  end;
  SendToItems.Clear;
end;
{$ENDIF TBX_OR_TB2K}

procedure TVirtualSendToMenu.Popup(X, Y: Integer);
begin
  Images := SmallSysImages;
  {$IFNDEF DELPHI_5_UP}
  ClearMenuItems(Self);
  {$ELSE}
  Items.Clear;
  {$ENDIF DELPHI_5_UP}
  Populate(Items);
  inherited;
end;

{ TSendToMenuOptions }

constructor TSendToMenuOptions.Create;
begin
  Images := True;
  LargeImages := False;
  ImageBorder := 1;
  FEllipsisPlacement := sseMiddle;
  MaxWidth := -1;
end;

end.
