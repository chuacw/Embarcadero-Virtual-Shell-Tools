unit IDEVirtualShellContainers;

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
//
// Classes to help with storing TNamespaces

interface

{$include Compilers.inc}
{$include VSToolsAddIns.inc}

uses
  Windows, Messages, SysUtils, Classes, IDEVirtualShellUtilities;

type
  TVirtualNameSpaceList = class;  // Forward


  TObjectList = class(TList)
  private
    FOwnsObjects: Boolean;
  protected
    function GetItem(Index: Integer): TObject;
    procedure SetItem(Index: Integer; AObject: TObject);
  public
    constructor Create; overload;
    constructor Create(AOwnsObjects: Boolean); overload;

    function Add(AObject: TObject): Integer;
    {$IFDEF DELPHI_5_UP}
    function Extract(Item: TObject): TObject;
    {$ENDIF}
    function Remove(AObject: TObject): Integer;
    function IndexOf(AObject: TObject): Integer;
    function FindInstanceOf(AClass: TClass; AExact: Boolean = True; AStartAt: Integer = 0): Integer;
    procedure Insert(Index: Integer; AObject: TObject);
    function First: TObject;
    function Last: TObject;
    property OwnsObjects: Boolean read FOwnsObjects write FOwnsObjects;
    property Items[Index: Integer]: TObject read GetItem write SetItem; default;
  end;



  {$IFDEF DELPHI_5_UP}
  TVirtualNamespaceListNotifyEvent = procedure(Sender: TVirtualNameSpaceList; Namespace: TNamespace;
    Action: TListNotification);
  {$ELSE}
  TVirtualNamespaceListNotifyEvent = procedure(Sender: TVirtualNameSpaceList; Namespace: TNamespace);
  {$ENDIF}

  TVirtualNameSpaceList  = class(TObjectList)
    FOnChanged : TVirtualNamespaceListNotifyEvent;
  protected
    {$IFDEF DELPHI_5_UP}
    procedure Notify(Ptr: Pointer; Action: TListNotification); override;
    {$ENDIF}
    function GetItems(Index: Integer): TNameSpace;
    procedure SetItems(Index: Integer; ANameSpace: TNameSpace);
  public
    function Add(ANameSpace: TNamespace): Integer;
    {$IFDEF DELPHI_5_UP}
    function Extract(Item: TNameSpace): TNameSpace;
    {$ENDIF DELPHI_5_UP}
    procedure FillArray(var NamespaceArray: TNamespaceArray);
    function First: TNameSpace;
    procedure FreeNamespaces;
    function IndexOf(ANameSpace: TNameSpace): Integer;
    procedure Insert(Index: Integer; ANameSpace: TNameSpace);
    function Last: TNameSpace;
    function Remove(ANameSpace: TNameSpace): Integer;

    property Items[Index: Integer]: TNamespace read GetItems write SetItems; default;
    property OnChanged: TVirtualNamespaceListNotifyEvent  read FOnChanged write FOnChanged;
  end;

implementation

{ TObjectList }

function TObjectList.Add(AObject: TObject): Integer;
begin
  Result := inherited Add(AObject);
end;

constructor TObjectList.Create;
begin
  inherited Create;
  FOwnsObjects := True;
end;

constructor TObjectList.Create(AOwnsObjects: Boolean);
begin
  inherited Create;
  FOwnsObjects := AOwnsObjects;
end;

{$IFDEF DELPHI_5_UP}
function TObjectList.Extract(Item: TObject): TObject;
begin
  Result := TObject(inherited Extract(Item));
end;
{$ENDIF DELPHI_5_UP}

function TObjectList.FindInstanceOf(AClass: TClass; AExact: Boolean;
  AStartAt: Integer): Integer;
var
  I: Integer;
begin
  Result := -1;
  for I := AStartAt to Count - 1 do
    if (AExact and
        (Items[I].ClassType = AClass)) or
       (not AExact and
        Items[I].InheritsFrom(AClass)) then
    begin
      Result := I;
      break;
    end;
end;

function TObjectList.First: TObject;
begin
  Result := TObject(inherited First);
end;

function TObjectList.GetItem(Index: Integer): TObject;
begin
  Result := inherited Items[Index];
end;

function TObjectList.IndexOf(AObject: TObject): Integer;
begin
  Result := inherited IndexOf(AObject);
end;

procedure TObjectList.Insert(Index: Integer; AObject: TObject);
begin
  inherited Insert(Index, AObject);
end;

function TObjectList.Last: TObject;
begin
  Result := TObject(inherited Last);
end;

function TObjectList.Remove(AObject: TObject): Integer;
begin
  Result := inherited Remove(AObject);
end;

procedure TObjectList.SetItem(Index: Integer; AObject: TObject);
begin
  inherited Items[Index] := AObject;
end;

{ TVirtualNameSpaceList }

function TVirtualNameSpaceList.Add(ANameSpace: TNameSpace): Integer;
begin
  Result := inherited Add(ANameSpace);
end;

{$IFDEF DELPHI_5_UP}
function TVirtualNameSpaceList.Extract(Item: TNameSpace): TNameSpace;
begin
  Result := TNamespace( inherited Extract(Item))
end;
{$ENDIF DELPHI_5_UP}

procedure TVirtualNameSpaceList.FillArray(var NamespaceArray: TNamespaceArray);
begin
  SetLength(NamespaceArray, Count);
  MoveMemory(@NamespaceArray[0], List, SizeOf(TNamespace)*Count);
end;

function TVirtualNameSpaceList.First: TNameSpace;
begin
  Result := TNamespace( inherited First)
end;

procedure TVirtualNameSpaceList.FreeNamespaces;
var
  i: integer;
begin
  for i := 0 to Count - 1 do
  begin
    TObject(Items[i]).Free;
    Items[i] := nil
  end;
end;

function TVirtualNameSpaceList.GetItems (Index: Integer): TNameSpace;
begin
  Result := TNameSpace(inherited Items[Index]);
end;

function  TVirtualNameSpaceList.IndexOf (ANameSpace: TNameSpace): Integer;
begin
  Result := inherited IndexOf(ANameSpace);
end;

procedure TVirtualNameSpaceList.Insert (Index: Integer; ANameSpace: TNameSpace);
begin
  inherited Insert(Index, ANameSpace);
end;

function TVirtualNameSpaceList.Last: TNameSpace;
begin
  Result := TNamespace( inherited Last)
end;

{$IFDEF DELPHI_5_UP}
procedure TVirtualNameSpaceList.Notify(Ptr: Pointer;
  Action: TListNotification);
begin
  if Assigned(FOnChanged) then
    FOnChanged(Self, TNamespace(Ptr), Action);
  inherited;
end;
{$ENDIF}

function  TVirtualNameSpaceList.Remove (ANameSpace: TNameSpace): Integer;
begin
  Result := inherited Remove(ANameSpace);
end;

procedure TVirtualNameSpaceList.SetItems (Index: Integer; ANameSpace: TNameSpace);
begin
  inherited Items[Index] := ANameSpace;
end;

end.
