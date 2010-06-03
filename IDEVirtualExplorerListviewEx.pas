unit IDEVirtualExplorerListviewEx;

{==============================================================================
Version 1.4.6
(VirtualShellTools release 1.1.x)

Software is distributed on an "AS IS" basis, WITHOUT WARRANTY OF ANY KIND,
either express or implied.
The initial developer of this code is Robert Lee.

Requirements:
  - Mike Lischke's Virtual Treeview (VT)
    http://www.lischke-online.de/VirtualTreeview/VT.html
  - Jim Kuenaman's Virtual Shell Tools (VSTools)
    http://groups.yahoo.com/group/VirtualExplorerTree
  - Mike Lischke's Theme Manager 1.9.0 or above (ONLY for Delphi 5 or Delphi 6)
  - Comctl32.dll v.5.81 (Internet Explorer 5.00) or later is required for the VCL
    Listview to work as expected in Thumbnail viewstyle.

Credits:
  Special thanks to Mike Lischke (VT) and Jim Kuenaman (VSTools) for the
  magnificent components they made available to the Delphi community.
  Thanks to (in alphabetical order):
    Aaron, Adem Baba (ImageMagick conversion), Nils Haeck (ImageMagick wrapper),
    Gerald K�der (bugs hunter), Werner Lehmann (Thumbs Cache),
    Bill Miller (HyperVirtualExplorer), Renate Schaaf (Graphics guru),
    Boris Tadjikov (bugs hunter), Milan Vandrovec (CBuilder port),
    Philip Wand, Troy Wolbrink (Unicode support).

Known issues:
  - If the Anchors property is changed at runtime and the ViewStyle is not
    vsxReport SyncOptions must be called.
  - If the node selection is changed at runtime and the ViewStyle is not
    vsxReport SyncSelected must be called.
  - If ComCtrl 6 is used the scrollbars are not invalidated correctly when
    the BevelKind <> bkNone, and the ViewStyle <> vsxReport.
    This is a VCL bug.
  - If ComCtrl 6 is used the Listview standard hints are not correctly painted
    (only the first line is showed), this happens only when ViewStyle <> vsxReport.
    This is a VCL bug.

Development notes:
  - Don't use PaintTo, WinXP doesn't support it.
  - Don't use DrawTextW, Win9x doesn't support it, instead use:
    if Win32Platform = VER_PLATFORM_WIN32_WINDOWS then Windows.DrawText(...)
    else Windows.DrawTextW(...);
  - Don't use Item.DisplayRect, use TUnicodeOwnerDataListView.GetLVItemRect
    instead, ComCtrl 6 returns the item rect adding the item space.
  - Bug in Delphi 5 TGraphic class: when creating a bitmap by using its class
    type (TGraphicClass) it doesn't calls the TBitmap constructor.
    That's because in Delphi 5 the TGraphic.Create is protected, and
    TBitmap.Create is public, so TGraphicClass(ABitmap).Create will NOT call
    TBitmap.Create because it's not visible by TGraphicClass.
    To fix this we need to make it visible by creating a TGraphicClass cracker.
    Fixed in LoadGraphic helper.
    More info on:
    http://groups.google.com/groups?hl=en&lr=&ie=UTF-8&selm=VA.00000331.0164178b%40xmission.com

VCL TListview bugs fixed by VELVEx:
  - TListview bug: setting ShowHint to False doesn't disable the hints.
    Fixed in TSyncListView.CMShowHintChanged.
  - TListview bug: if the ClientHeight is too small to fit 2 items the
    PageUp/PageDown key buttons won't work.
    To fix this I had to fake these keys to Up/Down in TSyncListView.KeyDown.
  - TListview bug only in Delphi 7: the Listview invalidates the canvas when
    scrolling, the problem is in ComCtrls.pas: TCustomListView.WMVScroll.
    Fixed in TSyncListView.WMVScroll, I had to call the TWinControl WM_VSCROLL
    handler, sort of inherited-inherited.
    Contact Mike Lischke, this was introduced by D7 ThemeManager engine.
  - TListview bug only in Win2K and WinXP: the Listview Edit control is not
    showed correctly when the font size is large (Height > 15), this is visible
    in vsReport, vsList and vsSmallIcon viewstyles. We must reshow the Edit
    control.
    Fixed in TUnicodeOwnerDataListView.CNNotify with LVN_BEGINLABELEDIT and
    LVN_BEGINLABELEDITW messages.
  - TListview bug only on WinXP using ComCtl 6: a thick black border around
    the edit box is showed when editing an item in the list.
    Fixed in TUnicodeOwnerDataListView.WMCtlColorEdit.
  - OwnerData TListview bug on WinXP using ComCtl 6: it has to invalidate its
    Canvas on resizing.
    Fixed in TUnicodeOwnerDataListView.WMWindowPosChanging.
    This is fixed on Delphi 2005:
    http://qc.borland.com/qc/wc/wc.exe/details?ReportID=5920
  - OwnerData TListview bug on WinXP using ComCtl 6: Item.DisplayRect(drBounds)
    returns an incorrect Rect, the R.Left and R.Right includes the item spacing,
    this problem is related with the selectable space between icons issue.
    This is corrected in TUnicodeOwnerDataListView.GetLVItemIconRect.
  - OwnerData TListview bug on WinXP using ComCtl 6: the white space between
    icons (vsxIcon or vsxThumbs) becomes selectable, but not the space between
    the captions.
    This is related to previous bug.
    Fixed in TUnicodeOwnerDataListView.WndProc.
  - OwnerData TListview bug: when the Listview is small enough to not fit 2 fully
    visible items the PageUp/PageDown buttons don't work.
    Fixed in TSyncListView.KeyDown.
  - OwnerData TListview bug: when Shift-Selecting an item, it just selects all
    the items from the last selected to the current selected, index wise, it
    should box-select the items.
    Fixed in TSyncListView.OwnerDataStateChange, KeyMultiSelectEx and
    MouseMultiSelectEx.
  - OwnerData TListview bug: the OnAdvancedCustomDrawItem event doesn't catch
    cdPostPaint paint stage.
    Fixed in TSyncListView.CNNotify.
  - OwnerData TListview bug: when the Listview is unfocused and a previously
    selected item caption is clicked it enters in editing mode. This is an
    incorrect TListview behavior.
    To fix this issue I set a flag: FPrevEditing in TSyncListView.CMEnter and
    deactivate it in TSyncListView.CanEdit and when the selection changes in
    TSyncListView.CNNotify.
  - OwnerData TListview bug: the virtual TListView raises an AV in vsReport mode.
    Fixed in TSyncListView.LVMInsertColumn and TSyncListView.LVMSetColumn.
  - OwnerData TListview bug: when the icon arrangement is iaLeft the arrow keys
    are scrambled.
    Fixed in TSyncListView.KeyDown.

To Do
  -

History:

6 November 2005
  - Added support for threaded enumeration of objects - Jim K

28 September 2005 - version 1.4.6
  - Made GetChildByIndex function more robust.
  - Fixed 0000017 Mantis entry: AV when deleting file.
  - Fixed 0000023 Mantis entry: incorrect mouse click events when
    HideCaptions or ComCtl6 was used.
  - Fixed problem in CM_FONTCHANGED and Begin/End Update causing
    TListview to access window handle during streaming, thanks to
    Jim for the fix.
  - Fixed incorrect selection painting, now it uses the
    Colors.FocusedSelectionColor and Colors.UnFocusedSelectionColor
    properties.
  - Fixed incorrect focus painting when syncronizing the selection.
  - Added new border styles, tbWindowed and tbFit.
  - Added new SingleLineCaptions property to the ThumbsOptions.
  - Added Ctrl-I shortcut to invert the selection.

31 May 2005 - version 1.4.5
  - Fixed AV when creating large Metafiles thumbnails, thanks to
    Polo for the fix.
  - Fixed incorrect drag image when DragImageKind was diNoImage, thanks
    to Thomas Bauer for reporting this.
  - Denied Theme Manager subclassing.
  - Added new property: ThumbsOptions.CacheOptions.AdjustIconSize to allow
    dynamic thumbnails resizing.

21 December 2004 - version 1.4.4
  - Added EditFile method, to easily browse for a file to select it and begin
    to edit it.
  - Fixed incorrect checkboxes sync.
  - Fixed incorrect mouse button click handling when ComCtrls 6 is used,
    thanks to Gabriel Cristescu for reporting this.

27 August 2004 - version 1.4.3
  - Added checkbox support for vsxThumbs viewstyle.
  - Added partial background image support to non vsxReport ViewStyles, the
    background should be loaded in the listview using the ListView_SetBkImage
    API:
    var
      BK: TLVBKImage;
    begin
      // LVBKIF_SOURCE_HBITMAP flag is not supported by ListView_SetBkImage
      // TLVBKImage.hbm is not supported by ListView_SetBkImage
      Fillchar(BK, SizeOf(BK), 0);
      BK.ulFlags := LVBKIF_SOURCE_URL or LVBKIF_STYLE_TILE;
      BK.pszImage := PChar(Edit1.text);
      ListView_SetBkImage(LV.ChildListview.Handle, @BK);
    end;

23 May 2004 - version 1.4.2
  - Added support for toHideOverlay and toRestoreTopNodeOnRefresh properties.
  - Added new property: ThumbsOptions.CacheOptions.CacheProcessing to allow
    ascending or descending sorting of the thread cache processing list.
    When this property is tcpAscending the top files in the listview are
    thumbnailed first.
  - Fixed incorrect selection painting, it now uses the values in
    Colors.FocusedSelectionColor and Colors.UnFocusedSelectionColor.
  - Fixed an incorrect call to OnEditCancelled.
  - Reworked the internal cache events.

5 March 2004 - version 1.4.1
  - Compatible with VSTools 1.1.15
  - Fixed drag and drop synchronization, it now correctly fires OnDragDrop
    event when the ViewStyle <> vsxReport.
  - Fixed incorrect icon spacing when the handle is recreated.

==============================================================================}

interface

{$include Compilers.inc}
{$include VSToolsAddIns.inc}

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, ComCtrls, ImgList,
  Dialogs, Forms, IDEVirtualTrees, IDEVirtualExplorerTree, IDEVirtualShellutilities,
  ActiveX, Commctrl, IDEVirtualUtilities, IDEVirtualUnicodeDefines, IDEVirtualWideStrings,
  ShlObj, IDEVirtualResources, IDEVirtualShellTypes, IDEVirtualSystemImageLists,
  IDEVirtualIconThread, // Got to have it for the Thumbnail thread
  {$IFDEF USEGRAPHICEX} GraphicEx, {$ELSE}
    {$IFDEF USEIMAGEEN} ImageEnIo, {$ELSE}
    {$IFDEF USEIMAGEMAGICK} MagickImage, ImageMagickAPI, {$ENDIF}
    {$ENDIF}
  {$ENDIF}
  IDEVirtualUnicodeControls, JPeg, ComObj;

const
  SCROLL_DELAY_TIMER = 100;
  SCROLL_TIMER = 101;
  WM_VLVEXTHUMBTHREAD = WM_SHELLNOTIFY + 100;

  CM_DENYSUBCLASSING = CM_BASE + 2000; // Prevent Theme Manager subclassing

  {
  // TListView Extended Styles
  LVS_EX_LABELTIP     = $00004000;  // Partially hidden captions hint
  LVS_EX_BORDERSELECT = $00008000;  // Highlight the border when selecting items

  // TListView Extended Styles for ComCtl32.dll version 6
  // LVS_EX_DOUBLEBUFFER = $00010000;  // Enables alpha blended selection, defined in VirtualShellTypes
  LVS_EX_HIDELABELS   = $00020000;  // Hide the captions, like Shift clicking Thumbnails viewstyle option in Explorer
  LVS_EX_SINGLEROW    = $00040000;  // Like thumb stripes mode
  LVS_EX_SNAPTOGRID   = $00080000;  // Snaps the icons to grid
  LVS_EX_SIMPLESELECT = $00100000;  // ?
  }

type
  TCustomVirtualExplorerListviewEx = class;
  TThumbsCacheItem = class;
  TThumbsCache = class;

  TIconSpacing = record
    X: Word;
    Y: Word;
  end;

  TIconAttributes = record
    Size: TPoint;
    Spacing: TIconSpacing;
  end;

  TThumbnailState = (
    tsEmpty,           //empty data
    tsProcessing,      //processing thumbnail
    tsValid,           //valid image
    tsInvalid);        //not an image

  TThumbnailBorder = (tbNone, tbRaised, tbDoubleRaised, tbSunken,
    tbDoubleSunken, tbBumped, tbEtched, tbFramed, tbWindowed, tbFit);

  TThumbnailHighlight = (thNone, thSingleColor, thMultipleColors);

  TThumbnailImageLibrary = (timNone, timGraphicEx, timImageEn, timEnvision, timImageMagick);

  TViewStyleEx = (vsxIcon, vsxSmallIcon, vsxList, vsxReport, vsxThumbs);

  TThumbsCacheStorage = (tcsCentral, tcsPerFolder);

  TThumbsCacheProcessing = (tcpAscending, tcpDescending);

  PThumbnailData = ^TThumbnailData;
  TThumbnailData = packed record
    CachePos: integer;
    Reloading: Boolean;
    State: TThumbnailState;
  end;

  PThumbnailThreadData = ^TThumbnailThreadData;
  TThumbnailThreadData = packed record
    ImageWidth: integer;
    ImageHeight: integer;
    State: TThumbnailState;
    MemStream: TMemoryStream;
    CompressedStream: Boolean;
  end;

  TOLEListviewButtonState = (
    LButtonDown,
    RButtonDown,
    MButtonDown
  );
  TOLEListviewButtonStates = set of TOLEListviewButtonState;

  TRectArray = array of TRect;

  TLVThumbsDraw = procedure(Sender: TCustomVirtualExplorerListviewEx; ACanvas: TCanvas;
    ListItem: TListItem; ThumbData: PThumbnailData; AImageRect, ADetailsRect: TRect;
    var DefaultDraw: Boolean) of object;

  TLVThumbsDrawHint = procedure(Sender: TCustomVirtualExplorerListviewEx; HintBitmap: TBitmap;
    Node: PVirtualNode; var DefaultDraw: Boolean) of object;

  TLVThumbsGetDetails = procedure(Sender: TCustomVirtualExplorerListviewEx; Node: PVirtualNode;
    HintDetails: Boolean; var Details: WideString) of object;

  TLVThumbsCacheItemReadEvent = procedure(Sender: TCustomVirtualExplorerListviewEx;
    Filename: WideString; Thumbnail: TBitmap; var DoDefault: Boolean) of object;

  TLVThumbsCacheItemLoadEvent = procedure(Sender: TCustomVirtualExplorerListviewEx;
    NS: TNamespace; CacheItem: TThumbsCacheItem; var DoDefault: Boolean) of object;

  TLVThumbsCacheItemProcessingEvent = procedure(Sender: TCustomVirtualExplorerListviewEx;
    NS: TNamespace; Thumbnail: TBitmap; var ImageWidth, ImageHeight: Integer; var DoDefault: Boolean) of object;

  TLVThumbsCacheEvent = procedure(Sender: TThumbsCache; CacheFilePath: WideString; Comments: TWideStringList; var DoDefault: Boolean) of object;

  TExtensionsList = class(TWideStringList)
  private
    function GetColors(Index: integer): TColor;
    procedure SetColors(Index: integer; const Value: TColor);
  public
    constructor Create; virtual;
    function Add(const Extension: WideString; HighlightColor: TColor): Integer; reintroduce;
    function AddObject(const S: WideString; AObject: TObject): Integer; override;
    function IndexOf(const S: WideString): Integer; override;
    function DeleteString(const S: WideString): Boolean; virtual;
    property Colors[Index: integer]: TColor read GetColors write SetColors;
  end;

  TBitmapHint = class(THintWindow)
  private
    FHintBitmap: TBitmap;
    FActivating: Boolean;
    procedure WMEraseBkgnd(var Message: TWMEraseBkgnd); message WM_ERASEBKGND;
    procedure CMTextChanged(var Message: TMessage); message CM_TEXTCHANGED;
  protected
    procedure Paint; override;
  public
    property Activating: Boolean read FActivating;
    procedure ActivateHint(Rect: TRect; const AHint: string); override;
    procedure ActivateHintData(Rect: TRect; const AHint: string; AData: Pointer); override;
  end;

  TThumbsCacheItem = class
  private
    FFileDateTime: TDateTime;
    FImageWidth: Integer;
    FImageHeight: Integer;
    FExif: WideString;
    FComment: WideString;
    FCompressed: Boolean;
    FThumbImageStream: TMemoryStream;
    FStreamSignature: WideString;
  protected
    FFilename: WideString;
    procedure Changed; virtual;
    function DefaultStreamSignature: WideString; virtual;

    property CompressedThumbImageStream: Boolean read FCompressed write FCompressed;
    property ThumbImageStream: TMemoryStream read FThumbImageStream;
  public
    constructor Create(AFilename: WideString); virtual;
    constructor CreateFromStream(ST: TStream); virtual;
    destructor Destroy; override;
    procedure Assign(CI: TThumbsCacheItem); virtual;
    procedure Fill(AFileDateTime: TDateTime; AExif, AComment: WideString;
      AImageWidth, AImageHeight: Integer; ACompressed: Boolean;
      AThumbImageStream: TMemoryStream);
    function LoadFromStream(ST: TStream): Boolean; virtual;
    procedure SaveToStream(ST: TStream); virtual;

    function ReadBitmap(OutBitmap: TBitmap): Boolean;
    procedure WriteBitmap(ABitmap: TBitmap; CompressIt: Boolean);

    property Comment: WideString read FComment write FComment;
    property Exif: WideString read FExif write FExif;
    property Filename: WideString read FFilename;
    property FileDateTime: TDateTime read FFileDateTime write FFileDateTime;
    property ImageWidth: Integer read FImageWidth write FImageWidth;
    property ImageHeight: Integer read FImageHeight write FImageHeight;
    property StreamSignature: WideString read FStreamSignature;
  end;

  TThumbsCacheItemClass = class of TThumbsCacheItem;

  TThumbsCache = class
  private
    FDirectory: WideString;
    FLoadedFromFile: Boolean;
    FStreamVersion: Integer;
    FSize: Integer;
    FInvalidCount: Integer;
    FThumbWidth: Integer;
    FThumbHeight: Integer;
    FComments: TWideStringList;
    function GetCount: integer;
  protected
    FHeaderFilelist: TWideStringList; //List of filenames encoded like: "filename*comment", it owns TMemoryStreams
    FScreenBuffer: TWideStringList; //List of filenames, owns TBitmaps, it's a cache to speed screen rendering
    function DefaultStreamVersion: integer; virtual;
  public
    constructor Create; virtual;
    destructor Destroy; override;
    procedure Clear;
    function IndexOf(Filename: WideString): integer;
    function Add(Filename: WideString; CI: TThumbsCacheItem): Integer; overload;
    function Add(Filename, AExif, AComment: WideString; AFileDateTime: TDateTime;
      AImageWidth, AImageHeight: Integer; ACompressIt: Boolean;
      AThumbImage: TBitmap): Integer; overload;
    function Add(Filename, AExif, AComment: WideString; AFileDateTime: TDateTime;
      AImageWidth, AImageHeight: Integer; ACompressed: Boolean;
      AThumbImageStream: TMemoryStream): Integer; overload;
    procedure Assign(AThumbsCache: TThumbsCache);
    function Delete(Filename: WideString): boolean;
    function Read(Index: integer; var OutCacheItem: TThumbsCacheItem): Boolean; overload;
    function Read(Index: integer; OutBitmap: TBitmap): Boolean; overload;
    procedure LoadFromFile(const Filename: WideString); overload;
    procedure LoadFromFile(const Filename: WideString; InvalidFiles: TWideStringList); overload;
    procedure SaveToFile(const Filename: WideString);
    property Directory: WideString read FDirectory write FDirectory;
    property ThumbWidth: integer read FThumbWidth write FThumbWidth;
    property ThumbHeight: integer read FThumbHeight write FThumbHeight;
    property Comments: TWideStringList read FComments;
    property Count: integer read GetCount;  //Count includes the deleted thumbs, ValidCount = Count - InvalidCount
    property InvalidCount: integer read FInvalidCount;
    property LoadedFromFile: boolean read FLoadedFromFile;
    property StreamVersion: integer read FStreamVersion;
    property Size: integer read FSize;
  end;

  TCacheList = class(TWideStringList)
  private
    FCentralFolder: WideString;
    FDefaultFilename: WideString;
    procedure SetCentralFolder(const Value: WideString);
  public
    constructor Create; virtual;
    procedure DeleteAllFiles;
    procedure DeleteInvalidFiles;
    function GetCacheFileToLoad(Dir: Widestring): WideString;
    function GetCacheFileToSave(Dir: Widestring): WideString;
    procedure SaveToFile; reintroduce;
    procedure LoadFromFile; reintroduce;
    property CentralFolder: WideString read FCentralFolder write SetCentralFolder;
    property DefaultFilename: WideString read FDefaultFilename write FDefaultFilename;
  end;

  TThumbsCacheOptions = class(TPersistent)
  private
    FOwner: TCustomVirtualExplorerListviewEx;
    FAdjustIconSize: boolean;
    FAutoSave: boolean;
    FAutoLoad: boolean;
    FCompressed: boolean;
    FStorageType: TThumbsCacheStorage;
    FDefaultFilename: WideString;
    FBrowsingFolder: WideString;
    FCacheProcessing: TThumbsCacheProcessing;
    function GetCentralFolder: WideString;
    function GetSize: integer;
    function GetThumbsCount: integer;
    procedure SetBrowsingFolder(const Value: WideString);
    procedure SetCacheProcessing(const Value: TThumbsCacheProcessing);
    procedure SetCentralFolder(const Value: WideString);
    procedure SetCompressed(const Value: boolean);
  protected
    FThumbsCache: TThumbsCache; //Cache of the browsing folder
    FCacheList: TCacheList; //List of cache files
  public
    constructor Create(AOwner: TCustomVirtualExplorerListviewEx); virtual;
    destructor Destroy; override;
    procedure Assign(Source: TPersistent); override;
    procedure ClearCache(DeleteAllFiles: Boolean = False);
    function GetCacheFileFromCentralFolder(Dir: WideString): WideString;
    function RenameCacheFileFromCentralFolder(Dir, NewDirName: WideString; NewCacheFilename: WideString = ''): Boolean;
    procedure Reload(Node: PVirtualNode); overload;
    procedure Reload(Filename: WideString); overload;
    procedure Load(Force: Boolean = false);
    procedure Save;
    function Read(Node: PVirtualNode; var OutCacheItem: TThumbsCacheItem): Boolean; overload;
    function Read(Filename: WideString; var OutCacheItem: TThumbsCacheItem): Boolean; overload;
    function Read(Node: PVirtualNode; OutBitmap: TBitmap): Boolean; overload;
    function Read(Filename: WideString; OutBitmap: TBitmap): Boolean; overload;

    property Size: integer read GetSize;
    property ThumbsCount: integer read GetThumbsCount;
    property Owner: TCustomVirtualExplorerListviewEx read FOwner;
    property BrowsingFolder: WideString read FBrowsingFolder write SetBrowsingFolder;
  published
    property AdjustIconSize: boolean read FAdjustIconSize write FAdjustIconSize default False;
    property AutoLoad: boolean read FAutoLoad write FAutoLoad default False;
    property AutoSave: boolean read FAutoSave write FAutoSave default False;
    property DefaultFilename: WideString read FDefaultFilename write FDefaultFilename;
    property StorageType: TThumbsCacheStorage read FStorageType write FStorageType default tcsCentral;
    property CacheProcessing: TThumbsCacheProcessing read FCacheProcessing write SetCacheProcessing default tcpDescending;
    property CentralFolder: WideString read GetCentralFolder write SetCentralFolder;
    property Compressed: boolean read FCompressed write SetCompressed default False;
  end;

  TThumbsOptions = class(TPersistent)
  private
    FOwner: TCustomVirtualExplorerListviewEx;
    FCacheOptions: TThumbsCacheOptions;
    FThumbsIconAtt: TIconAttributes;
    FDetails: Boolean;
    FBorderOnFiles: Boolean;
    FDetailsHeight: Integer;
    FHighlightColor: TColor;
    FBorder: TThumbnailBorder;
    FHighlight: TThumbnailHighlight;
    FLoadAllAtOnce: Boolean;
    FBorderSize: Integer;
    FUseShellExtraction: Boolean;
    FShowSmallIcon: Boolean;
    FShowXLIcons: Boolean;
    FStretch: Boolean;
    FUseSubsampling: Boolean;
    function GetHeight: Integer;
    function GetSpaceHeight: Word;
    function GetSpaceWidth: Word;
    function GetWidth: Integer;
    procedure SetBorder(const Value: TThumbnailBorder);
    procedure SetBorderSize(const Value: Integer);
    procedure SetBorderOnFiles(const Value: Boolean);
    procedure SetDetailedHints(const Value: Boolean);
    procedure SetDetails(const Value: Boolean);
    procedure SetDetailsHeight(const Value: Integer);
    procedure SetHeight(const Value: Integer);
    procedure SetHighlight(const Value: TThumbnailHighlight);
    procedure SetHighlightColor(const Value: TColor);
    procedure SetSpaceHeight(const Value: Word);
    procedure SetSpaceWidth(const Value: Word);
    procedure SetWidth(const Value: Integer);
    procedure SetUseShellExtraction(const Value: Boolean);
    procedure SetShowSmallIcon(const Value: Boolean);
    procedure SetShowXLIcons(const Value: Boolean);
    procedure SetStretch(const Value: Boolean);
    procedure SetUseSubsampling(const Value: Boolean);
    function GetDetailedHints: Boolean;
    function GetHideCaptions: Boolean;
    procedure SetHideCaptions(const Value: Boolean);
    function GetSingleLineCaptions: Boolean;
    procedure SetSingleLineCaptions(const Value: Boolean);
  public
    constructor Create(AOwner: TCustomVirtualExplorerListviewEx); virtual;
    destructor Destroy; override;
    property Owner: TCustomVirtualExplorerListviewEx read FOwner;
  published
    property Border: TThumbnailBorder read FBorder write SetBorder default tbFramed;
    property BorderSize: Integer read FBorderSize write SetBorderSize default 4;
    property BorderOnFiles: Boolean read FBorderOnFiles write SetBorderOnFiles default False;
    property Width: Integer read GetWidth write SetWidth default 120;
    property Height: Integer read GetHeight write SetHeight default 120;
    property SpaceWidth: Word read GetSpaceWidth write SetSpaceWidth default 40;
    property SpaceHeight: Word read GetSpaceHeight write SetSpaceHeight default 40;
    property DetailedHints: Boolean read GetDetailedHints write SetDetailedHints default False;
    property Details: Boolean read FDetails write SetDetails default False;
    property DetailsHeight: Integer read FDetailsHeight write SetDetailsHeight default 40;
    property HideCaptions: Boolean read GetHideCaptions write SetHideCaptions default False;
    property SingleLineCaptions: Boolean read GetSingleLineCaptions write SetSingleLineCaptions default False;
    property Highlight: TThumbnailHighlight read FHighlight write SetHighlight default thMultipleColors;
    property HighlightColor: TColor read FHighlightColor write SetHighlightColor default $EFD3D3;
    property LoadAllAtOnce: Boolean read FLoadAllAtOnce write FLoadAllAtOnce default False;
    property ShowSmallIcon: Boolean read FShowSmallIcon write SetShowSmallIcon default True;
    property ShowXLIcons: Boolean read FShowXLIcons write SetShowXLIcons default True;
    property UseShellExtraction: Boolean read FUseShellExtraction write SetUseShellExtraction default True;
    property UseSubsampling: Boolean read FUseSubsampling write SetUseSubsampling default True;
    property Stretch: Boolean read FStretch write SetStretch default False;
    property CacheOptions: TThumbsCacheOptions read FCacheOptions write FCacheOptions;
  end;

  TThumbThread = class(TVirtualImageThread)
  private
    FOwner: TCustomVirtualExplorerListviewEx;
    FThumbThreadData: TThumbnailThreadData;
    FThumbCompression: boolean;
    FThumbWidth: integer;
    FThumbHeight: integer;
    FTransparentColor: TColor;
    FThumbStretch: boolean;
    FThumbSubsampling: boolean;
  protected
    procedure ExtractInfo(PIDL: PItemIDList; Info: PVirtualThreadIconInfo); override;
    procedure ExtractedInfoLoad(Info: PVirtualThreadIconInfo); override; // Load Info before being sent to Control(s)
    procedure InvalidateExtraction; override;
    procedure ReleaseItem(Item: PVirtualThreadIconInfo; const Malloc: IMalloc); override;
    function CreateThumbnail(Filename: WideString; var Thumbnail: TBitmap;
      var ImageWidth, ImageHeight: integer; var CompressIt: boolean): Boolean; virtual;
  public
    constructor Create(AOwner: TCustomVirtualExplorerListviewEx); virtual;
    procedure ResetThumbOptions; virtual;
    property ThumbCompression: boolean read FThumbCompression;
    property ThumbWidth: integer read FThumbWidth;
    property ThumbHeight: integer read FThumbHeight;
    property ThumbStretch: boolean read FThumbStretch;
    property ThumbSubsampling: boolean read FThumbSubsampling;
    property TransparentColor: TColor read FTransparentColor;
    property Owner: TCustomVirtualExplorerListviewEx read FOwner;
  end;

  TThumbThreadClass = class of TThumbThread;

  TThumbThreadClassEvent = procedure(Sender: TCustomVirtualExplorerListviewEx; var ThreadClass: TThumbThreadClass) of object;

  //TUnicodeOwnerDataListView adds Unicode support to OWNER-DATA-ONLY TListview
  TUnicodeOwnerDataListView = class(TListView)
  private
    FIsComCtl6: Boolean;
    FHideCaptions: Boolean;
    FSingleLineCaptions: Boolean;
    FEditingItemIndex: Integer;
    FSingleLineMaxChars: Integer;
    procedure SetHideCaptions(const Value: Boolean);
    procedure SetSingleLineCaptions(const Value: Boolean);
    procedure CMDenySubclassing(var Message: TMessage); message CM_DENYSUBCLASSING;
    procedure CMFontchanged(var Message: TMessage); message CM_FONTCHANGED;
    procedure CNNotify(var Message: TWMNotify); message CN_NOTIFY;
    procedure WMCtlColorEdit(var Message: TWMCtlColorEdit); message WM_CTLCOLOREDIT;
    procedure WMWindowPosChanging(var Message: TWMWindowPosChanging); message WM_WINDOWPOSCHANGING;
  protected
    Win32PlatformIsUnicode: boolean;
    PWideFindString: PWideChar;
    CurrentDispInfo: PLVDispInfoW;
    OriginalDispInfoMask: Cardinal;
    procedure WndProc(var Msg: TMessage); override;
    procedure CreateWnd; override;
    procedure CreateWindowHandle(const Params: TCreateParams); override;
    function GetItem(Value: TLVItemW): TListItem;
    function GetItemCaption(Index: Integer): WideString; virtual; abstract;
    property EditingItemIndex: Integer read FEditingItemIndex;
  public
    constructor Create(AOwner: TComponent); override;
    function GetFilmstripHeight: Integer;
    function GetLVItemRect(Index: integer; DisplayCode: TDisplayCode): TRect;
    procedure UpdateSingleLineMaxChars;
    property HideCaptions: Boolean read FHideCaptions write SetHideCaptions default False;
    property SingleLineCaptions: Boolean read FSingleLineCaptions write SetSingleLineCaptions default False;
    property IsComCtl6: Boolean read FIsComCtl6;
    property SingleLineMaxChars: Integer read FSingleLineMaxChars;
  end;

  //TSyncListView is used to sync VCL LV with VELV
  TSyncListView = class(TUnicodeOwnerDataListView)
  private
    FVETController: TCustomVirtualExplorerListviewEx;
    FSavedPopupNamespace: TNamespace;
    FFirstShiftClicked: Integer;
    FOwnerDataPause: Boolean;
    FSelectionPause: Boolean;
    FInPaintCycle: Boolean;
    FPrevEditing: Boolean;
    FDetailedHints: Boolean;
    FThumbnailHintBitmap: TBitmap;
    FDefaultTooltipsHandle: THandle;
    procedure ContextMenuCmdCallback(Namespace: TNamespace; Verb: WideString; MenuItemID: Integer; var Handled: Boolean);
    procedure ContextMenuShowCallback(Namespace: TNamespace; Menu: hMenu; var Allow: Boolean);
    procedure ContextMenuAfterCmdCallback(Namespace: TNamespace; Verb: WideString; MenuItemID: Integer; Successful: Boolean);
    procedure SetDetailedHints(const Value: Boolean);
    procedure UpdateHintHandle;
    procedure CNNotify(var Message: TWMNotify); message CN_NOTIFY;
    procedure CMEnter(var Message: TCMEnter); message CM_ENTER;
    procedure CMExit(var Message: TCMExit); message CM_EXIT;
    procedure CMMouseWheel(var Message: TCMMouseWheel); message CM_MOUSEWHEEL;
    procedure CMShowHintChanged(var Message: TMessage); message CM_SHOWHINTCHANGED;
    procedure LVMSetColumn(var Message: TMessage); message LVM_SETCOLUMN;
    procedure LVMInsertColumn(var Message: TMessage); message LVM_INSERTCOLUMN;
    procedure WMPaint(var Message: TWMPaint); message WM_PAINT;
    procedure WMEraseBkgnd(var Message: TWMEraseBkgnd); message WM_ERASEBKGND;
  protected
    procedure WndProc(var Msg: TMessage); override;
    procedure WMVScroll(var Message: TWMVScroll); message WM_VSCROLL;
    procedure WMHScroll(var Message: TWMHScroll); message WM_HSCROLL;
    procedure WMLButtonDown(var Message: TWMLButtonDown); message WM_LBUTTONDOWN;
    procedure CMHintShow(var Message: TCMHintShow); message CM_HINTSHOW;

    function OwnerDataFetch(Item: TListItem; Request: TItemRequest): Boolean; override;
    function OwnerDataHint(StartIndex: Integer; EndIndex: Integer): Boolean; override;
    function OwnerDataFind(Find: TItemFind; const FindString: string;
      const FindPosition: TPoint; FindData: Pointer; StartIndex: Integer;
      Direction: TSearchDirection; Wrap: Boolean): Integer; override;
    function OwnerDataStateChange(StartIndex, EndIndex: Integer; OldState,
      NewState: TItemStates): Boolean; override;

    procedure DoContextPopup(MousePos: TPoint; var Handled: Boolean); override;
    procedure Edit(const Item: TLVItem); override;
    function CanEdit(Item: TListItem): Boolean; override;
    function GetItemCaption(Index: Integer): WideString; override;
    function IsBackgroundValid: Boolean;

    procedure KeyDown(var Key: Word; Shift: TShiftState); override;
    procedure KeyPress(var Key: Char); override;
    procedure KeyUp(var Key: Word; Shift: TShiftState); override;
    procedure MouseDown(Button: TMouseButton; Shift: TShiftState;
      X: Integer; Y: Integer); override;
    procedure MouseMove(Shift: TShiftState; X: Integer; Y: Integer);
      override;
    procedure MouseUp(Button: TMouseButton; Shift: TShiftState; X: Integer;
      Y: Integer); override;
    procedure DblClick; override;
  public
    constructor Create(AOwner: TComponent); override;
    destructor Destroy; override;
    procedure CreateHandle; override;
    function IsItemGhosted(Item: TListitem): Boolean;
    procedure FetchThumbs(StartIndex, EndIndex: integer);
    procedure UpdateArrangement;
    property DetailedHints: Boolean read FDetailedHints write SetDetailedHints default False;
    property InPaintCycle: Boolean read FInPaintCycle;
    property OwnerDataPause: boolean read FOwnerDataPause write FOwnerDataPause;
    property VETController: TCustomVirtualExplorerListviewEx read FVETController write FVETController;
  end;

  //TOLEListview adds drag & drop support to TSyncListView
  TOLEListview = class(TSyncListView, IDropTarget, IDropSource)
  private
    FDropTargetHelper: IDropTargetHelper;
    FDragDataObject: IDataObject;
    FCurrentDropIndex: integer;
    FDragging: Boolean;
    FMouseButtonState: TOLEListviewButtonStates;
    FAutoScrolling: Boolean;
    FDropped: Boolean;
    FDragItemIndex: integer;

    //Scrolling support
    FScrollDelayTimer: THandle;
    FScrollTimer: THandle;
    FAutoScrollTimerStub: Pointer;    // Stub for timer callback function
    procedure AutoScrollTimerCallback(Window: hWnd; Msg, idEvent: integer; dwTime: Longword); stdcall;
  protected
    procedure ClearTimers;
    procedure CreateDragImage(TotalDragRect: TRect; RectArray: TRectArray; var Bitmap: TBitmap);
    procedure CreateWnd; override;
    procedure DestroyWnd; override;
    function DragEnter(const dataObj: IDataObject; grfKeyState: Longint; pt: TPoint; var dwEffect: Longint): HResult; virtual; stdcall;
    function IDropTarget.DragOver = DragOverOLE; // Naming Clash
    procedure DoContextPopup(MousePos: TPoint; var Handled: Boolean); override;
    function DragOverOLE(grfKeyState: Longint; pt: TPoint; var dwEffect: Longint): HResult; virtual; stdcall;
    function Drop(const dataObj: IDataObject; grfKeyState: Longint; pt: TPoint; var dwEffect: Longint): HResult; virtual; stdcall;
    function DragLeave: HResult; virtual; stdcall;
    function GiveFeedback(dwEffect: Longint): HResult; virtual; stdcall;
    function ListIndexToNamespace(ItemIndex: integer): TNamespace;
    function ListItemToNamespace(Item: TListItem; BackGndIfNIL: Boolean): TNamespace;
    function QueryContinueDrag(fEscapePressed: BOOL; grfKeyState: Longint): HResult; virtual; stdcall;

    procedure WMLButtonDown(var Message: TWMLButtonDown); message WM_LBUTTONDOWN;
    procedure WMLButtonUp(var Message: TWMLButtonUp); message WM_LBUTTONUP;
    procedure WMRButtonDown(var Message: TWMRButtonDown); message WM_RBUTTONDOWN;
    procedure WMRButtonUp(var Message: TWMRButtonUp); message WM_RBUTTONUP;
    procedure WMMouseMove(var Message: TWMMouseMove); message WM_MOUSEMOVE;
    procedure WMTimer(var Message: TWMTimer); message WM_TIMER;

    property CurrentDropIndex: integer read FCurrentDropIndex write FCurrentDropIndex default -2;
    property DragDataObject: IDataObject read FDragDataObject write FDragDataObject;
    property DropTargetHelper: IDropTargetHelper read FDropTargetHelper;
    property MouseButtonState: TOLEListviewButtonStates read FMouseButtonState write FMouseButtonState;
    property AutoScrolling: Boolean read FAutoScrolling;
  public
    constructor Create(AOwner: TComponent); override;
    destructor Destroy; override;

    function Dragging: Boolean; reintroduce;
  end;

  TCustomVirtualExplorerListviewEx = class(TVirtualExplorerListview)
  private
    FVisible: boolean;
    FAccumulatedChanging: boolean;
    FViewStyle: TViewStyleEx;
    FThumbThread: TThumbThread;
    FThumbsThreadPause: Boolean;
    FImageLibrary: TThumbnailImageLibrary;
    FExtensionsList: TExtensionsList;
    FShellExtractExtensionsList: TExtensionsList;
    FExtensionsExclusionList: TExtensionsList;
    FOnThumbsDrawBefore, FOnThumbsDrawAfter: TLVThumbsDraw;
    FOnThumbsDrawHint: TLVThumbsDrawHint;
    FOnThumbsGetDetails: TLVThumbsGetDetails;
    FOnThumbsCacheItemAdd: TLVThumbsCacheItemProcessingEvent;
    FOnThumbsCacheItemRead: TLVThumbsCacheItemReadEvent;
    FOnThumbsCacheItemLoad: TLVThumbsCacheItemLoadEvent;
    FOnThumbsCacheItemProcessing: TLVThumbsCacheItemProcessingEvent;
    FOnThumbsCacheLoad: TLVThumbsCacheEvent;
    FOnThumbsCacheSave: TLVThumbsCacheEvent;
    FInternalDataOffset: Cardinal; // offset to the internal data of the ExplorerListviewEx
    FThumbsOptions: TThumbsOptions;
    FOnThumbThreadClass: TThumbThreadClassEvent;

    function GetThumbThread: TThumbThread;
    procedure SetThumbThreadClassEvent(const Value: TThumbThreadClassEvent);
    procedure SetViewStyle(const Value: TViewStyleEx);
    procedure SetVisible(const Value: boolean);
    procedure LVOnAdvancedCustomDrawItem(Sender: TCustomListView;
      Item: TListItem; State: TCustomDrawState; Stage: TCustomDrawStage;
      var DefaultDraw: Boolean);
    procedure CMBorderChanged(var Message: TMessage); message CM_BORDERCHANGED;
    procedure CMShowHintChanged(var Message: TMessage); message CM_SHOWHINTCHANGED;
  protected
    FListview: TOLEListview;
    FDummyIL: TImageList;
    FSearchCache: PVirtualNode;
    FLastDeletedNode: PVirtualNode;
    
    procedure CreateWnd; override;
    procedure RequestAlign; override;
    procedure SetParent(AParent: TWinControl); override;
    procedure SetZOrder(TopMost: Boolean); override;
    function GetAnimateWndParent: TWinControl; override;
    function GetClientRect: TRect; override;
    procedure Notification(AComponent: TComponent; Operation: TOperation); override;
    procedure CMBidimodechanged(var Message: TMessage); message CM_BIDIMODECHANGED;
    procedure CMCtl3DChanged(var Message: TMessage); message CM_CTL3DCHANGED;
    procedure CMColorChanged(var Message: TMessage); message CM_COLORCHANGED;
    procedure CMCursorChanged(var Message: TMessage); message CM_CURSORCHANGED;
    procedure CMEnabledchanged(var Message: TMessage); message CM_ENABLEDCHANGED;
    procedure CMFontChanged(var Message: TMessage); message CM_FONTCHANGED;
    procedure WMNCDestroy(var Message: TWMNCDestroy); message WM_NCDESTROY;
    {$IFDEF THREADEDICONS}
    procedure WMVTSetIconIndex(var Msg: TWMVTSetIconIndex); message WM_VTSETICONINDEX;
    {$ENDIF}
    procedure WMVLVExThumbThread(var Message: TMessage); message WM_VLVEXTHUMBTHREAD;
    procedure WMEnumThreadFinished(var Msg: TMessage); message WM_ENUMTHREADFINISHED;
    procedure WMEnumThreadStart(var Msg: TMessage); message WM_ENUMTHREADSTART;

    //VirtualTree methods
    procedure DoInitNode(Parent, Node: PVirtualNode; var InitStates: TVirtualNodeInitStates); override;
    procedure DoFreeNode(Node: PVirtualNode); override;
    procedure RebuildRootNamespace; override;
    procedure DoRootChanging(const NewRoot: TRootFolder; Namespace: TNamespace; var Allow: Boolean); override;
    procedure DoStructureChange(Node: PVirtualNode; Reason: TChangeReason); override;
    procedure ReReadAndRefreshNode(Node: PVirtualNode; SortNode: Boolean); override;
    procedure DoBeforeCellPaint(Canvas: TCanvas; Node: PVirtualNode; Column: TColumnIndex; CellRect: TRect); override;
    function InternalData(Node: PVirtualNode): Pointer;
    function IsAnyEditing: Boolean; override;
    function GetNodeByIndex(Index: Cardinal): PVirtualNode;
    procedure FlushSearchCache;

    //Thumbnails processing
    function IsValidChildListview: boolean;
    function GetThumbThreadClass: TThumbThreadClass; virtual;

    //Thumbnails cache
    function DoThumbsCacheItemAdd(NS: TNamespace; Thumbnail: TBitmap; ImageWidth, ImageHeight: Integer): Boolean; virtual; // A Thumbnail is about to be added to the cache
    function DoThumbsCacheItemLoad(NS: TNamespace; var CacheItem: TThumbsCacheItem): Boolean; virtual; // A Thumbnail was loaded from the cache
    function DoThumbsCacheItemRead(Filename: WideString; Thumbnail: TBitmap): Boolean; virtual; // A Thumbnail must be read from the cache
    function DoThumbsCacheItemProcessing(NS: TNamespace; Thumbnail: TBitmap; var ImageWidth, ImageHeight: Integer): Boolean; virtual; // A Thumbnail must be created
    function DoThumbsCacheLoad(Sender: TThumbsCache; CacheFilePath: WideString; Comments: TWideStringList): Boolean; virtual; // The cache was loaded from file
    function DoThumbsCacheSave(Sender: TThumbsCache; CacheFilePath: WideString; Comments: TWideStringList): Boolean; virtual; // The cache is about to be saved

    procedure ResetThumbImageList(ResetSpacing: boolean = True);
    procedure ResetThumbSpacing;
    procedure ResetThumbThread;

    //Thumbnails drawing
    function GetDetailsString(Node: PVirtualNode; ThumbFormatting: Boolean = True): WideString;
    function DoThumbsGetDetails(Node: PVirtualNode; HintDetails: Boolean): WideString;
    procedure DoThumbsDrawBefore(ACanvas: TCanvas; ListItem: TListItem; ThumbData: PThumbnailData;
      AImageRect, ADetailsRect: TRect; var DefaultDraw: Boolean); virtual;
    procedure DoThumbsDrawAfter(ACanvas: TCanvas; ListItem: TListItem; ThumbData: PThumbnailData;
      AImageRect, ADetailsRect: TRect; var DefaultDraw: Boolean); virtual;
    function DoThumbsDrawHint(HintBitmap: TBitmap; Node: PVirtualNode): Boolean;
    procedure DrawThumbBG(ACanvas: TCanvas; Item: TListItem; ThumbData: PThumbnailData; R: TRect);
    procedure DrawThumbFocus(ACanvas: TCanvas; Item: TListItem; ThumbData: PThumbnailData; R: TRect);
    procedure DrawIcon(ACanvas: TCanvas; Item: TListItem; ThumbData: PThumbnailData; RThumb, RDetails: TRect; var RDestination: TRect);

    property ViewStyle: TViewStyleEx read FViewStyle write SetViewStyle;
    property ThumbsOptions: TThumbsOptions read FThumbsOptions write FThumbsOptions;
    property ThumbThread: TThumbThread read GetThumbThread;
    property ThumbsThreadPause: Boolean read FThumbsThreadPause write FThumbsThreadPause;
    property OnThumbThreadClass: TThumbThreadClassEvent read FOnThumbThreadClass write SetThumbThreadClassEvent;
    property OnThumbsCacheItemAdd: TLVThumbsCacheItemProcessingEvent read FOnThumbsCacheItemAdd write FOnThumbsCacheItemAdd;
    property OnThumbsCacheItemLoad: TLVThumbsCacheItemLoadEvent read FOnThumbsCacheItemLoad write FOnThumbsCacheItemLoad;
    property OnThumbsCacheItemRead: TLVThumbsCacheItemReadEvent read FOnThumbsCacheItemRead write FOnThumbsCacheItemRead;
    property OnThumbsCacheItemProcessing: TLVThumbsCacheItemProcessingEvent read FOnThumbsCacheItemProcessing write FOnThumbsCacheItemProcessing;
    property OnThumbsCacheLoad: TLVThumbsCacheEvent read FOnThumbsCacheLoad write FOnThumbsCacheLoad;
    property OnThumbsCacheSave: TLVThumbsCacheEvent read FOnThumbsCacheSave write FOnThumbsCacheSave;
    property OnThumbsDrawBefore: TLVThumbsDraw read FOnThumbsDrawBefore write FOnThumbsDrawBefore;
    property OnThumbsDrawAfter: TLVThumbsDraw read FOnThumbsDrawAfter write FOnThumbsDrawAfter;
    property OnThumbsDrawHint: TLVThumbsDrawHint read FOnThumbsDrawHint write FOnThumbsDrawHint;
    property OnThumbsGetDetails: TLVThumbsGetDetails read FOnThumbsGetDetails write FOnThumbsGetDetails;

    function DoKeyAction(var CharCode: Word; var Shift: TShiftState): Boolean; override;
  public
    constructor Create(AOwner: TComponent); override;
    destructor Destroy; override;
    procedure Loaded; override;

    //VT methods
    procedure SetBounds(ALeft, ATop, AWidth, AHeight: Integer); override;
    function Focused: Boolean; override;
    procedure SetFocus; override;
    function EditNode(Node: PVirtualNode; Column: TColumnIndex): Boolean; override;
    function EditFile(APath: WideString): Boolean;
    procedure Clear; override;
    procedure CopyToClipBoard; override;
    procedure CutToClipBoard; override;
    function PasteFromClipboard: Boolean; override;
    function InvalidateNode(Node: PVirtualNode): TRect; override;

    //VET methods
    function BrowseToByPIDL(APIDL: PItemIDList; ExpandTarget, SelectTarget, SetFocusToVET,
      CollapseAllFirst: Boolean; ShowAllSiblings: Boolean = True): Boolean; override;
    procedure SelectedFilesDelete; override;
    procedure SelectedFilesPaste(AllowMultipleTargets: Boolean); override;
    procedure SelectedFilesShowProperties; override;

    //Synchronization methods
    procedure SyncInvalidate;
    procedure SyncItemsCount; virtual;
    procedure SyncOptions; virtual;
    procedure SyncSelectedItems(UpdateChildListview: Boolean = True);

    //Thumbnails public methods
    procedure FillExtensionsList(FillColors: Boolean = true); virtual;
    function GetThumbDrawingBounds(IncludeThumbDetails, IncludeBorderSize: Boolean): TRect;
    function IsImageFileIndex(FileName: WideString): integer;
    function IsImageFile(FileName: WideString): Boolean; overload;
    function IsImageFile(Node: PVirtualNode): TNamespace; overload;
    function ValidateThumbnail(Node: PVirtualNode; var ThumbData: PThumbnailData): Boolean;
    function ValidateListItem(Node: PVirtualNode; var ListItem: TListItem): Boolean;

    //Public properties
    property ImageLibrary: TThumbnailImageLibrary read FImageLibrary;
    property ChildListview: TOLEListview read FListview;
    property ExtensionsList: TExtensionsList read FExtensionsList;
    property ShellExtractExtensionsList: TExtensionsList read FShellExtractExtensionsList;
    property ExtensionsExclusionList: TExtensionsList read FExtensionsExclusionList;
  published
    property Visible: boolean read FVisible write SetVisible default true;
  end;

  TVirtualExplorerListviewEx = class(TCustomVirtualExplorerListviewEx)
  published
    property ViewStyle;
    property ThumbsOptions;
    property OnThumbThreadClass;
    property OnThumbsCacheItemAdd;
    property OnThumbsCacheItemLoad;
    property OnThumbsCacheItemRead;
    property OnThumbsCacheItemProcessing;
    property OnThumbsCacheLoad;
    property OnThumbsCacheSave;
    property OnThumbsDrawBefore;
    property OnThumbsDrawAfter;
    property OnThumbsDrawHint;
    property OnThumbsGetDetails;
  end;

// Misc helpers
function SpMakeObjectInstance(Method: TWndMethod): Pointer;
procedure SpFreeObjectInstance(ObjectInstance: Pointer);
function SpCompareText(W1, W2: WideString): Boolean;
function SpPathCompactPath(S: WideString; Max: Integer): WideString;

//Node manipulation helpers
function GetChildByIndex(ParentNode: PVirtualNode; ChildIndex: Cardinal; var SearchCache: PVirtualNode): PVirtualNode;
function IsThumbnailActive(ThumbnailState: TThumbnailState): Boolean;
function SupportsShellExtract(NS: TNamespace): boolean;

//Image manipulation helpers
procedure InitBitmap(OutB: TBitmap; W, H: integer; BackgroundColor: TColor);
procedure SpStretchDraw(G: TGraphic; OutBitmap: TBitmap; DestR: TRect; UseSubsampling: Boolean);
function RectAspectRatio(ImageW, ImageH, ThumbW, ThumbH: integer; Center, Stretch: boolean): TRect;
procedure DrawThumbBorder(ACanvas: TCanvas; ThumbBorder: TThumbnailBorder; R: TRect; Selected: Boolean);
function IsIncompleteJPGError(E: Exception): boolean;
function IsDelphiSupportedImageFile(FileName: WideString): Boolean;
function GetGraphicClass(Filename: WideString): TGraphicClass;
function LoadImageEnGraphic(Filename: WideString; outBitmap: TBitmap): Boolean;
function LoadGraphic(Filename: WideString; outP: TPicture): boolean;
function LoadImage(Filename: WideString; OutP: TPicture): boolean;
function MakeThumbFromFile(Filename: WideString; OutBitmap: TBitmap; ThumbW, ThumbH: integer;
  Center: boolean; BgColor: TColor; Stretch, Subsampling: Boolean; var ImageWidth, ImageHeight: integer): Boolean;

//Stream helpers
function ReadDateTimeFromStream(ST: TStream): TDateTime;
procedure WriteDateTimeToStream(ST: TStream; D: TDateTime);
function ReadIntegerFromStream(ST: TStream): Integer;
procedure WriteIntegerToStream(ST: TStream; I: Integer);
function ReadWideStringFromStream(ST: TStream): WideString;
procedure WriteWideStringToStream(ST: TStream; WS: WideString);
function ReadMemoryStreamFromStream(ST: TStream; MS: TMemoryStream): Boolean;
procedure WriteMemoryStreamToStream(ST: TStream; MS: TMemoryStream);
function ReadBitmapFromStream(ST: TStream; B: TBitmap): Boolean;
procedure WriteBitmapToStream(ST: TStream; B: TBitmap);
procedure ConvertBitmapStreamToJPGStream(MS: TMemoryStream; CompressionQuality: TJPEGQualityRange);
procedure ConvertJPGStreamToBitmapStream(MS: TMemoryStream);
procedure ConvertJPGStreamToBitmap(MS: TMemoryStream; OutBitmap: TBitmap);

implementation

uses
  ShellApi, Math;

//WMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWM
{ Helpers }

function SpMakeObjectInstance(Method: TWndMethod): Pointer;
begin
{$IFDEF COMPILER_6_UP}
  Result := Classes.MakeObjectInstance(Method);
{$ELSE}
  Result := Forms.MakeObjectInstance(Method);
{$ENDIF}
end;

procedure SpFreeObjectInstance(ObjectInstance: Pointer);
begin
{$IFDEF COMPILER_6_UP}
  Classes.FreeObjectInstance(ObjectInstance);
{$ELSE}
  Forms.FreeObjectInstance(ObjectInstance);
{$ENDIF}
end;

function SpCompareText(W1, W2: WideString): Boolean;
begin
  Result := False;
  if Win32Platform = VER_PLATFORM_WIN32_NT then begin
    if lstrcmpiW_VST(PWideChar(W1), PWideChar(W2)) = 0 then
      Result := True;
  end else
    if AnsiCompareText(W1, W2) = 0 then
      Result := True;
end;

function SpPathCompactPath(S: WideString; Max: Integer): WideString;
var
  L: Integer;
begin
  L := Length(S);
  if (L > 3) and (Max > 0) and (L > Max) then begin
    SetLength(S, Max - 3);
    S := S + '...';
  end;
  Result := S;
end;

function GetChildByIndex(ParentNode: PVirtualNode; ChildIndex: Cardinal; var SearchCache: PVirtualNode): PVirtualNode;
var
  N: PVirtualNode;
  Count: Cardinal;
begin
  Result := nil;
  if not Assigned(ParentNode) then Exit;

  Count := ParentNode.ChildCount;
  if ChildIndex >= Count then Exit;

  // This speeds up the search drastically
  if Assigned(SearchCache) and Assigned(SearchCache.Parent) and (SearchCache.Parent = ParentNode) then
  begin
    if ChildIndex >= SearchCache.Index then
    begin
      N := SearchCache;
      while Assigned(N) do
        if N.Index = ChildIndex then begin
          Result := N;
          Break;
        end
        else
          N := N.NextSibling;
    end else
    begin
      N := SearchCache;
      while Assigned(N) do
        if N.Index = ChildIndex then begin
          Result := N;
          Break;
        end
        else
          N := N.PrevSibling;
    end;
  end else
    if ChildIndex <= Count div 2 then begin
      N := ParentNode.FirstChild;
      while Assigned(N) do
        if N.Index = ChildIndex then begin
          Result := N;
          Break;
        end
        else
          N := N.NextSibling;
    end
    else begin
      N := ParentNode.LastChild;
      while Assigned(N) do
        if N.Index = ChildIndex then begin
          Result := N;
          Break;
        end
        else
          N := N.PrevSibling;
    end;

  SearchCache := Result;
end;

function IsThumbnailActive(ThumbnailState: TThumbnailState): Boolean;
begin
  Result := (ThumbnailState = tsValid) or (ThumbnailState = tsProcessing);
end;

function SupportsShellExtract(NS: TNamespace): boolean;
begin
  Result := Assigned(NS.ExtractImage.ExtractImageInterface);
end;

//WMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWM
{ Drawing helpers }

function RectAspectRatio(ImageW, ImageH, ThumbW, ThumbH: integer; Center, Stretch: boolean): TRect;
begin
  Result := Rect(0, 0, 0, 0);
  if (ImageW < 1) or (ImageH < 1) then Exit;

  if (ImageW <= ThumbW) and (ImageH <= ThumbH) and (not Stretch) then begin
    Result.Right := ImageW;
    Result.Bottom := ImageH;
  end
  else begin
    Result.Right := (ThumbH * ImageW) div ImageH;
    if (Result.Right <= ThumbW) then Result.Bottom := ThumbH
    else begin
      Result.Right := ThumbW;
      Result.Bottom := (ThumbW * ImageH) div ImageW;
    end;
  end;

  if Center then
    with Result do begin
      left := (ThumbW - right) div 2;
      top := (ThumbH - bottom) div 2;
      right := left + right;
      bottom := top + bottom;
    end;

  // TJPEGImage doesn't accepts images with 1 pixel width or height
  if Result.Right <= 1 then Result.Right := 2;
  if Result.Bottom <= 1 then Result.Bottom := 2;
end;

procedure InitBitmap(OutB: TBitmap; W, H: integer; BackgroundColor: TColor);
begin
//  OutB.PixelFormat := pf24bit; //do this first!
  OutB.Width := W;
  OutB.Height := H;
  OutB.Canvas.Brush.Color := BackgroundColor;
  OutB.Canvas.Fillrect(Rect(0, 0, W, H));
end;

procedure SpStretchDraw(G: TGraphic; OutBitmap: TBitmap; DestR: TRect; UseSubsampling: Boolean);
// Canvas.StretchDraw is NOT THREADSAFE!!!
// Use StretchBlt instead, we have to use a worker bitmap to do so
var
  Work: TBitmap;
begin
  Work := TBitmap.Create;
  Work.Canvas.Lock;
  try
    // Paint the Picture in Work
    if (G is TJpegImage) or (G is TBitmap) then
      Work.Assign(G) //assign works in this case
    else begin
      Work.Width := G.Width;
      Work.Height := G.Height;
      Work.Canvas.Draw(0, 0, G);
    end;

    if UseSubsampling then
      SetStretchBltMode(OutBitmap.Canvas.Handle, STRETCH_HALFTONE)
    else
      SetStretchBltMode(OutBitmap.Canvas.Handle, STRETCH_DELETESCANS);

    StretchBlt(OutBitmap.Canvas.Handle,
      DestR.Left, DestR.Top, DestR.Right - DestR.Left, DestR.Bottom - DestR.Top,
      Work.Canvas.Handle, 0, 0, G.Width, G.Height, SRCCopy);
  finally
    Work.Canvas.Unlock;
    Work.Free;
  end;
end;

procedure DrawThumbBorder(ACanvas: TCanvas; ThumbBorder: TThumbnailBorder; R: TRect; Selected: Boolean);
const
  Edge: array [TThumbnailBorder] of Cardinal = (0, BDR_RAISEDINNER, EDGE_RAISED,
    BDR_SUNKENOUTER, EDGE_SUNKEN, EDGE_BUMP, EDGE_ETCHED, 0, 0, 0);
begin
  if ThumbBorder <> tbNone then begin
    Case ThumbBorder of
      tbNone: ;
      tbFramed:
        begin
          ACanvas.Brush.Color := clBtnFace;
          ACanvas.FrameRect(R);
        end;
      tbWindowed:
        if Selected then begin
          ACanvas.Brush.Color := clHighlight;
          InflateRect(R, -2, -2);
          ACanvas.FrameRect(R);
          InflateRect(R, -1, -1);
          ACanvas.FrameRect(R);
        end
        else begin
          ACanvas.Brush.Color := clBtnFace;
          InflateRect(R, -2, -2);
          ACanvas.FrameRect(R);
        end;
      tbFit:
        if Selected then begin
          ACanvas.Brush.Color := clHighlight;
          ACanvas.FrameRect(R);
          InflateRect(R, -1, -1);
          ACanvas.FrameRect(R);
          ACanvas.Brush.Color := clWhite;
          InflateRect(R, -1, -1);
          ACanvas.FrameRect(R);
        end;
    else
      DrawEdge(ACanvas.Handle, R, Edge[ThumbBorder], BF_RECT);
    end;
  end;
end;

function IsIncompleteJPGError(E: Exception): boolean;
var
  S: string;
begin
  S := E.Message;
  Result := (S = 'JPEG error #68') or
            (S = 'JPEG error #67') or
            (S = 'JPEG error #60') or
            (S = 'JPEG error #57');
end;

function IsDelphiSupportedImageFile(FileName: WideString): Boolean;
var
  Ext: WideString;
begin
  Ext := WideLowerCase(ExtractFileExtW(Filename));
  Result := (Ext = '.jpg') or (Ext = '.jpeg') or (Ext = '.jif') or
    (Ext = '.bmp') or (Ext = '.wmf') or (Ext = '.emf') or (Ext = '.ico');
end;

function GetGraphicClass(Filename: WideString): TGraphicClass;
var
  Ext: WideString;
begin
  Ext := WideLowerCase(ExtractFileExtW(Filename));
  Delete(Ext, 1, 1);

  {$IFDEF USEGRAPHICEX}
  Result := GraphicEx.FileFormatList.GraphicFromExtension(Ext);
  {$ELSE}
    Result := nil;
    if (Ext = 'jpg') or (Ext = 'jpeg') or (Ext = 'jif') then Result := TJpegImage
    else if Ext = 'bmp' then Result := TBitmap
    else if (Ext = 'wmf') or (Ext = 'emf') then Result := TMetafile
    else if Ext = 'ico' then Result := TIcon;
    {$IFDEF USEIMAGEMAGICK}
    if Result = nil then
      Result := MagickImage.MagickFileFormatList.GraphicFromExtension(Ext);
    {$ENDIF}
  {$ENDIF}
end;

function LoadImageEnGraphic(Filename: WideString; outBitmap: TBitmap): Boolean;
// Loads an image file with a unicode name using ImageEn graphic library
{$IFDEF USEIMAGEEN}
var
  B: TBitmap;
  ImageEnIO: TImageEnIO;
  F: TWideFileStream;
{$ENDIF}
begin
  Result := false;
  {$IFDEF USEIMAGEEN}
  ImageEnIO := TImageEnIO.Create(nil);
  try
    ImageEnIO.AttachedBitmap := outBitmap;
    F := TWideFileStream.Create(Filename, fmOpenRead or fmShareDenyNone);
    try
      ImageEnIO.LoadFromStream(F);
      Result := true;
    finally
      F.Free;
    end;
  finally
    ImageEnIO.Free;
  end;
  {$ENDIF}
end;

// Bug in Delphi 5:
// When creating a bitmap by using its class type (TGraphicClass) it doesn't calls
// the TBitmap constructor.
// That's because in Delphi 5 the TGraphic.Create is protected, and TBitmap.Create
// is public, so TGraphicClass(ABitmap).Create will NOT call TBitmap.Create because
// it's not visible by TGraphicClass.
// To fix this we need to make it visible by creating a TGraphicClass cracker.
// More info on:
// http://groups.google.com/groups?hl=en&lr=&ie=UTF-8&selm=VA.00000331.0164178b%40xmission.com
type
  TFixedGraphic = class(TGraphic);
  TFixedGraphicClass = class of TFixedGraphic;

function LoadGraphic(Filename: WideString; outP: TPicture): boolean;
// Loads an image file with a unicode name
var
  NewGraphic: TGraphic;
  GraphicClass: TGraphicClass;
  F: TWideFileStream;
  {$IFDEF USEIMAGEEN}
  B: TBitmap;
  {$ENDIF}
begin
  Result := false;

  {$IFDEF USEIMAGEEN}
  B := TBitmap.Create;
  try
    Result := LoadImageEnGraphic(Filename, B);
  finally
    B.Free;
  end;
  {$ELSE}
  GraphicClass := GetGraphicClass(Filename);
  if GraphicClass = nil then
    outP.LoadFromFile(Filename) //try the default loader
  else begin
    F := TWideFileStream.Create(Filename, fmOpenRead or fmShareDenyNone);
    try
      NewGraphic := TFixedGraphicClass(GraphicClass).Create;
      try
        NewGraphic.LoadFromStream(F);
        outP.Graphic := NewGraphic;
        Result := true;
      finally
        NewGraphic.Free;
      end;
    finally
      F.Free;
    end;
  end;
  {$ENDIF}
end;

function LoadImage(Filename: WideString; OutP: TPicture): boolean;
var
  J: TJpegImage;
begin
  Result := false;
  if (Filename = '') or not Assigned(outP) then Exit;

  try
    LoadGraphic(Filename, OutP);
    if OutP.Graphic is TJpegImage then begin
      J := TJpegImage(OutP.Graphic);
      J.DIBNeeded; //load the JPG
    end;
  except
    on E:Exception do
      if not IsIncompleteJPGError(E) then
        raise;
  end;

  Result := true;
end;

function MakeThumbFromFile(Filename: WideString; OutBitmap: TBitmap; ThumbW, ThumbH: integer;
  Center: boolean; BgColor: TColor; Stretch, SubSampling: Boolean; var ImageWidth, ImageHeight: integer): Boolean;
var
  P: TPicture;
  J: TJpegImage;
  WMFScale: Single;
  DestR: TRect;
  Ext: string;
begin
  Result := false;
  if not Assigned(OutBitmap) then exit;
  Ext := Lowercase(ExtractFileExtW(Filename));

  P := TPicture.create;
  try
    LoadGraphic(Filename, P);
    if P.Graphic <> nil then begin
      if (Ext = '.jpg') or (Ext = '.jpeg') or (Ext = '.jif') then begin
        // 5x faster loading jpegs, try to load just the minimum possible jpg
        // From Danny Thorpe: http://groups.google.com/groups?hl=en&frame=right&th=69a64eafb3ee2b12&seekm=01bdee71%24e5a5ded0%247e018f0a%40agamemnon#link6
        try
          J := TJpegImage(P.graphic);
          J.Performance := jpBestSpeed;
          J.Scale := jsFullSize;
          while ((J.width > ThumbW) or (J.height > ThumbH)) and (J.Scale < jsEighth) do
            J.Scale := Succ(J.Scale);
          if J.Scale <> jsFullSize then
            J.Scale := Pred(J.Scale);

          J.DibNeeded; //now load the JPG
        except
          on E:Exception do
            if not IsIncompleteJPGError(E) then
              Raise;
        end;
      end
      else
        // We need to scale down the metafile images
        if (Ext = '.wmf') or (Ext = '.emf') then begin
          WMFScale := Min(1, Min(ThumbW/P.Graphic.Width, ThumbH/P.Graphic.Height));
          P.Graphic.Width := Round(P.Graphic.Width * WMFScale);
          P.Graphic.Height := Round(P.Graphic.Height * WMFScale);
        end;

      ImageWidth := P.Graphic.Width;
      ImageHeight := P.Graphic.Height;
      // Resize the thumb
      // Need to lock/unlock the canvas here
      OutBitmap.Canvas.Lock;
      try
        // init OutBitmap
        DestR := RectAspectRatio(ImageWidth, ImageHeight, ThumbW, ThumbH, Center, Stretch);

        if Center then
          InitBitmap(OutBitmap, ThumbW, ThumbH, BgColor)
        else
          InitBitmap(OutBitmap, DestR.Right, DestR.Bottom, BgColor);

        // StretchDraw is NOT THREADSAFE!!! Use SpStretchDraw instead
        SpStretchDraw(P.Graphic, OutBitmap, DestR, Subsampling);

        Result := True;
      finally
        OutBitmap.Canvas.UnLock;
      end;
    end;
  finally
    P.free;
  end;
end;

//WMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWM
{ Stream helpers }

function ReadDateTimeFromStream(ST: TStream): TDateTime;
begin
  ST.ReadBuffer(Result, SizeOf(Result));
end;

procedure WriteDateTimeToStream(ST: TStream; D: TDateTime);
begin
  ST.WriteBuffer(D, SizeOf(D));
end;

function ReadIntegerFromStream(ST: TStream): Integer;
begin
  ST.ReadBuffer(Result, SizeOf(Result));
end;

procedure WriteIntegerToStream(ST: TStream; I: Integer);
begin
  ST.WriteBuffer(I, SizeOf(I));
end;

function ReadWideStringFromStream(ST: TStream): WideString;
var
  L: integer;
  WS: WideString;
begin
  Result := '';
  ST.ReadBuffer(L, SizeOf(L));
  SetLength(WS, L);
  ST.ReadBuffer(PWideChar(WS)^, 2 * L);
  Result := WS;
end;

procedure WriteWideStringToStream(ST: TStream; WS: WideString);
var
  L: integer;
begin
  L := Length(WS);
  ST.WriteBuffer(L, SizeOf(L));
  ST.WriteBuffer(PWideChar(WS)^, 2 * L);
end;

function ReadMemoryStreamFromStream(ST: TStream; MS: TMemoryStream): Boolean;
var
  L: integer;
begin
  Result := false;
  ST.ReadBuffer(L, SizeOf(L));
  if L > 0 then begin
    MS.Size := L;
    ST.ReadBuffer(MS.Memory^, L);
    Result := true;
  end;
end;

procedure WriteMemoryStreamToStream(ST: TStream; MS: TMemoryStream);
var
  L: integer;
begin
  L := MS.Size;
  ST.WriteBuffer(L, SizeOf(L));
  ST.WriteBuffer(MS.Memory^, L);
end;

function ReadBitmapFromStream(ST: TStream; B: TBitmap): Boolean;
var
  MS: TMemoryStream;
begin
  Result := false;
  MS := TMemoryStream.Create;
  try
    if ReadMemoryStreamFromStream(ST, MS) then
      if Assigned(B) then begin
        B.LoadFromStream(MS);
        Result := true;
      end;
  finally
    MS.Free;
  end;
end;

procedure WriteBitmapToStream(ST: TStream; B: TBitmap);
var
  L: integer;
  MS: TMemoryStream;
begin
  if Assigned(B) then begin
    MS := TMemoryStream.Create;
    try
      B.SaveToStream(MS);
      WriteMemoryStreamToStream(ST, MS);
    finally
      MS.Free;
    end;
  end
  else begin
    L := 0;
    ST.WriteBuffer(L, SizeOf(L));
  end;
end;

procedure ConvertBitmapStreamToJPGStream(MS: TMemoryStream; CompressionQuality: TJPEGQualityRange);
var
  B: TBitmap;
  J: TJPEGImage;
begin
  B := TBitmap.Create;
  J := TJPEGImage.Create;
  try
    MS.Position := 0;
    B.LoadFromStream(MS);
    //WARNING, first set the JPEG options
    J.CompressionQuality := CompressionQuality; //90 is the default, 60 is the best setting
    //Now assign the Bitmap
    J.Assign(B);
    J.Compress;
    MS.Clear;
    J.SaveToStream(MS);
    MS.Position := 0;
  finally
    B.Free;
    J.Free;
  end;
end;

procedure ConvertJPGStreamToBitmapStream(MS: TMemoryStream);
var
  B: TBitmap;
  J: TJPEGImage;
begin
  B := TBitmap.Create;
  J := TJPEGImage.Create;
  try
    MS.Position := 0;
    J.LoadFromStream(MS);
    B.Assign(J);
    MS.Clear;
    B.SaveToStream(MS);
    MS.Position := 0;
  finally
    B.Free;
    J.Free;
  end;
end;

procedure ConvertJPGStreamToBitmap(MS: TMemoryStream; OutBitmap: TBitmap);
var
  B: TMemoryStream;
begin
  B := TMemoryStream.Create;
  try
    MS.Position := 0;
    B.LoadFromStream(MS);
    MS.Position := 0;
    ConvertJPGStreamToBitmapStream(B);
    OutBitmap.LoadFromStream(B);
  finally
    B.Free;
  end;
end;

//WMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWM
{ TExtensionsList }

constructor TExtensionsList.Create;
begin
  inherited;
  Sorted := True;
end;

function TExtensionsList.Add(const Extension: WideString; HighlightColor: TColor): Integer;
begin
  Result := inherited Add(Extension);
  if Result > -1 then
    Colors[Result] := HighlightColor;
end;

function TExtensionsList.AddObject(const S: WideString; AObject: TObject): Integer;
var
  Aux: WideString;
begin
  Aux := WideLowerCase(S);
  //Add the '.' part of the extension
  if (Length(Aux) > 0) and (Aux[1] <> '.') then
    Aux := '.' + Aux;
  Result := inherited AddObject(Aux, AObject);
end;

function TExtensionsList.IndexOf(const S: WideString): Integer;
var
  Aux: WideString;
begin
  Aux := WideLowerCase(S);
  //Add the '.' part of the extension
  if (Length(Aux) > 0) and (Aux[1] <> '.') then
    Aux := '.' + Aux;
  Result := inherited IndexOf(Aux);
end;

function TExtensionsList.DeleteString(const S: WideString): Boolean;
var
  I: integer;
begin
  I := IndexOf(S);
  if I > -1 then begin
    Delete(I);
    Result := True;
  end
  else
    Result := False;
end;

function TExtensionsList.GetColors(Index: integer): TColor;
begin
  if Assigned(Objects[Index]) then
    Result := TColor(Objects[Index])
  else
    Result := clNone;
end;

procedure TExtensionsList.SetColors(Index: integer; const Value: TColor);
begin
  Objects[Index] := Pointer(Value);
end;

//WMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWM
{ TBitmapHint }

procedure TBitmapHint.ActivateHint(Rect: TRect; const AHint: string);
begin
  FActivating := True;
  try
    inherited ActivateHint(Rect, AHint);
  finally
    FActivating := False;
  end;
end;

procedure TBitmapHint.ActivateHintData(Rect: TRect; const AHint: string; AData: Pointer);
begin
  //The AData parameter is a bitmap
  FHintBitmap := TBitmap(AData);
  Rect.Right := Rect.Left + FHintBitmap.Width - 2;
  Rect.Bottom := Rect.Top + FHintBitmap.Height - 2;
  inherited ActivateHintData(Rect, AHint, AData);
end;

procedure TBitmapHint.CMTextChanged(var Message: TMessage);
begin
  Message.Result := 1;
end;

procedure TBitmapHint.Paint;
begin
  Canvas.Draw(0, 0, FHintBitmap);
end;

procedure TBitmapHint.WMEraseBkgnd(var Message: TWMEraseBkgnd);
begin
  //Don't erase the background as this causes flickering
  Paint;
  Message.Result := 1;
end;

//WMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWM
{ TThumbsCache }

constructor TThumbsCache.Create;
begin
  FHeaderFilelist := TWideStringList.Create;
  FHeaderFilelist.OwnsObjects := true;
  FScreenBuffer := TWideStringList.Create;
  FScreenBuffer.OwnsObjects := true;
  FScreenBuffer.Sorted := true;
  FComments := TWideStringList.Create;
  Clear;
end;

destructor TThumbsCache.Destroy;
begin
  Clear;
  FHeaderFilelist.Free;
  FScreenBuffer.Free;
  FComments.Free;
  inherited;
end;

procedure TThumbsCache.Clear;
begin
  FHeaderFilelist.Clear;
  FScreenBuffer.Clear;
  FComments.Clear;

  FStreamVersion := DefaultStreamVersion;
  FLoadedFromFile := false;
  FDirectory := '';
  FSize := 0;
  FInvalidCount := 0;
  FThumbWidth := 0;
  FThumbHeight := 0;
end;

function TThumbsCache.IndexOf(Filename: WideString): Integer;
begin
  Result := FHeaderFilelist.IndexOf(Filename);
end;

function TThumbsCache.Add(Filename: WideString; CI: TThumbsCacheItem): Integer;
begin
  Result := FHeaderFilelist.AddObject(Filename, CI);
  if Result > -1 then
    FSize := FSize + CI.ThumbImageStream.Size;
end;

function TThumbsCache.Add(Filename, AExif, AComment: WideString;
  AFileDateTime: TDateTime; AImageWidth, AImageHeight: Integer;
  ACompressIt: Boolean; AThumbImage: TBitmap): Integer;
var
  CI: TThumbsCacheItem;
begin
  Result := -1;
  CI := TThumbsCacheItem.Create(Filename);
  try
    CI.WriteBitmap(AThumbImage, ACompressIt);
    CI.Fill(AFileDateTime, AExif, AComment, AImageWidth, AImageHeight,
      ACompressIt, nil);
    Result := Add(Filename, CI);
  except
    CI.Free;
  end;
end;

function TThumbsCache.Add(Filename, AExif, AComment: WideString;
  AFileDateTime: TDateTime; AImageWidth, AImageHeight: Integer;
  ACompressed: Boolean; AThumbImageStream: TMemoryStream): Integer;
var
  CI: TThumbsCacheItem;
begin
  Result := -1;
  CI := TThumbsCacheItem.Create(Filename);
  try
    CI.Fill(AFileDateTime, AExif, AComment, AImageWidth, AImageHeight,
      ACompressed, AThumbImageStream);
    Result := Add(Filename, CI);
  except
    CI.Free;
  end;
end;

function TThumbsCache.Delete(Filename: WideString): Boolean;
var
  I, J: integer;
  M: TMemoryStream;
begin
  I := IndexOf(Filename);
  Result := I > -1;
  if Result then begin
    M := TThumbsCacheItem(FHeaderFilelist.Objects[I]).ThumbImageStream;
    FSize := FSize - M.Size;

    // Don't delete it from the HeaderFilelist, instead free the ThumbsCacheItem and clear the string
    TThumbsCacheItem(FHeaderFilelist.Objects[I]).Free;
    FHeaderFilelist.Objects[I] := nil;
    FHeaderFilelist[I] := '';
    Inc(FInvalidCount);

    // Delete the ScreenBuffer item
    J := FScreenBuffer.IndexOf(Filename);
    if J > -1 then
      FScreenBuffer.Delete(J);

    Result := true;
  end;
end;

function TThumbsCache.Read(Index: integer; var OutCacheItem: TThumbsCacheItem): Boolean;
begin
  if (Index > -1) and (Index < FHeaderFilelist.Count) then
    OutCacheItem := FHeaderFilelist.Objects[Index] as TThumbsCacheItem
  else
    OutCacheItem := nil;
  Result := Assigned(OutCacheItem);
end;

function TThumbsCache.Read(Index: integer; OutBitmap: TBitmap): Boolean;
var
  CI: TThumbsCacheItem;
  I: integer;
  B: TBitmap;
begin
  Result := false;
  if Read(Index, CI) and Assigned(CI.ThumbImageStream) and Assigned(FScreenBuffer) then begin
    // Retrieve the bitmap from the ScreenBuffer
    I := FScreenBuffer.IndexOf(CI.Filename);
    if I > -1 then begin
      OutBitmap.Assign(TBitmap(FScreenBuffer.Objects[I]));
      Result := true;
    end
    else begin
      CI.ReadBitmap(OutBitmap);

      // Add it to the ScreenBuffer
      B := TBitmap.Create;
      try
        B.Assign(OutBitmap);
        FScreenBuffer.AddObject(CI.Filename, B); // The ScreenBuffer owns the bitmaps
      except
        B.Free;
      end;
      // Set capacity to 50
      while FScreenBuffer.Count > 50 do
        FScreenBuffer.Delete(0); // FIFO list

      Result := true;
    end;
  end;
end;

function TThumbsCache.DefaultStreamVersion: integer;
begin
  // Version 11, has FileDateTime defined as TDateTime, uses TThumbsCacheItem.SaveToStream,
  // has per file compression
  // and stores ThumbWidth/ThumbHeight
  Result := 11;
end;

procedure TThumbsCache.LoadFromFile(const Filename: Widestring);
var
  FileStream: TStream;
  CI: TThumbsCacheItem;
  I, FileCount: integer;
begin
  Clear;
  FileStream := TWideFileStream.Create(FileName, fmOpenRead or fmShareDenyWrite);
  try
    // Read the header
    FDirectory := ReadWideStringFromStream(FileStream);
    FileCount := ReadIntegerFromStream(FileStream);
    FStreamVersion := ReadIntegerFromStream(FileStream);
    FThumbWidth := ReadIntegerFromStream(FileStream);
    FThumbHeight := ReadIntegerFromStream(FileStream);

    if FStreamVersion = DefaultStreamVersion then begin
      // Read the file list
      for I := 0 to FileCount - 1 do begin
        CI := TThumbsCacheItem.CreateFromStream(FileStream);
        try
          if FileExistsW(CI.Filename) then
            Add(CI.Filename, CI)
          else begin
            // The file is not valid, don't add it
            CI.Free;
          end;
        except
          Inc(FInvalidCount);
          CI.Free;
        end;
      end;
      // Read the comments or extra options
      FComments.LoadFromStream(FileStream);

      FLoadedFromFile := true;
    end;

  finally
    FileStream.Free;
  end;
end;

procedure TThumbsCache.LoadFromFile(const Filename: Widestring; InvalidFiles: TWideStringList);
var
  AuxThumbsCache: TThumbsCache;
  CI, CICopy: TThumbsCacheItem;
  I: integer;
begin
  if not Assigned(InvalidFiles) or (InvalidFiles.Count = 0) then
    LoadFromFile(Filename)
  else begin
    // Tidy the list
    for I := 0 to InvalidFiles.Count - 1 do
      InvalidFiles[I] := WideLowerCase(InvalidFiles[I]);
    InvalidFiles.Sort;
    // Load a cache file and ATTACH only the streams that are NOT in InvalidFiles list
    AuxThumbsCache := TThumbsCache.Create;
    try
      AuxThumbsCache.LoadFromFile(Filename);
      for I := 0 to AuxThumbsCache.Count - 1 do begin
        if AuxThumbsCache.Read(I, CI) and (InvalidFiles.IndexOf(CI.Filename) < 0) then begin
          CICopy := TThumbsCacheItem.Create(CI.Filename);
          try
            CICopy.Assign(CI);
            Self.Add(CI.Filename, CICopy);
          except
            CICopy.Free;
          end;
        end;
      end;

      FLoadedFromFile := true;
    finally
      AuxThumbsCache.Free;
    end;
  end;
end;

procedure TThumbsCache.SaveToFile(const Filename: Widestring);
var
  FileStream: TStream;
  CI: TThumbsCacheItem;
  I: integer;
begin
  FileStream := TWideFileStream.Create(Filename, fmCreate);
  try
    // Write the header
    WriteWideStringToStream(FileStream, Directory);
    WriteIntegerToStream(FileStream, Count - InvalidCount);
    WriteIntegerToStream(FileStream, DefaultStreamVersion);
    WriteIntegerToStream(FileStream, FThumbWidth);
    WriteIntegerToStream(FileStream, FThumbHeight);
    // Write the file list
    FHeaderFilelist.Sort;
    for I := 0 to Count - 1 do
      if FHeaderFilelist[I] <> '' then begin
        Read(I, CI);
        CI.SaveToStream(FileStream);
      end;
    // Save comments or extra options
    FComments.SaveToStream(FileStream);
  finally
    FileStream.Free;
  end;
end;

procedure TThumbsCache.Assign(AThumbsCache: TThumbsCache);
var
  I: integer;
  CI, CICopy: TThumbsCacheItem;
begin
  if Assigned(AThumbsCache) then begin
    Clear;
    Directory := AThumbsCache.Directory;
    FLoadedFromFile := AThumbsCache.LoadedFromFile;
    FStreamVersion := AThumbsCache.StreamVersion;
    FThumbWidth := AThumbsCache.ThumbWidth;
    FThumbHeight := AThumbsCache.ThumbHeight;
    FComments.Assign(AThumbsCache.Comments);

    //Clone
    for I := 0 to AThumbsCache.Count - 1 do begin
      AThumbsCache.Read(I, CI);
      if CI.Filename <> '' then begin
        CICopy := TThumbsCacheItem.Create(CI.Filename);
        try
          CICopy.Assign(CI);
          Self.Add(CI.Filename, CICopy);
        except
          CICopy.Free;
        end;
      end;
    end;
  end;
end;

function TThumbsCache.GetCount: integer;
begin
  Result := FHeaderFilelist.Count;
end;

//WMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWM
{ TCacheList }

function PosRight(SubStr: WideChar; S: WideString): Integer;
// Like Pos function but starting from the last char
var
  I, L: integer;
begin
  Result := -1;
  L := Length(S);
  if L > 0 then
    for I := L downto 1 do
      if S[I] = SubStr then begin
        Result := I;
        Break;
      end;
end;

function NameFromRight(L: TWideStringList; const Index: Integer): WideString;
// Like TWideStringList.Names[] but starting from the last char
var
  P: Integer;
begin
  Result := L[Index];
  P := PosRight('\', Result);
  if P > 0 then SetLength(Result, P)
           else Result := '';
end;

function ValueFromIndexRight(L: TWideStringList; const Index: Integer): WideString;
// Like TWideStringList.ValueFromIndex but starting from the last char
var
  S: WideString;
begin
  Result := '';
  if Index > -1 then begin
    // get name
    S := NameFromRight(L, Index);
    // copy the result
    if S <> '' then
      Result := Copy(L[Index], Length(S) + 2, MaxInt);
  end;
end;

function IndexOfNameRight(L: TWideStringList; const Name: WideString): Integer;
// Like TWideStringList.IndexOfName but starting from the last char
var
  I: integer;
  S: WideString;
begin
  Result := -1;
  for I := 0 to L.Count - 1 do begin
    S := NameFromRight(L, I);
    if S <> '' then
      if SpCompareText(S, Name) then begin
        Result := I;
        Break;
      end;
  end;
end;

function IndexOfValueRight(L: TWideStringList; const Value: WideString): Integer;
// Like TWideStringList.IndexOfValue but starting from the last char
var
  P: Integer;
  I: integer;
begin
  Result := -1;
  for I := 0 to L.Count - 1 do begin
    P := PosRight('\', L[I]);
    if P > 0 then
      if SpCompareText(ValueFromIndexRight(L, I), Value) then begin
        Result := I;
        Break;
      end;
  end;
end;

constructor TCacheList.Create;
begin
  inherited;
  FDefaultFilename := 'CacheList.txt';
  FCentralFolder := '';
end;

function TCacheList.GetCacheFileToSave(Dir: Widestring): WideString;
var
  WCH: PWideChar;
  F, WS: WideString;
  I: integer;
begin
  Result := '';
  if FCentralFolder = '' then exit;

  Dir := IncludeTrailingBackslashW(Dir);
  I := IndexOfNameRight(Self, Dir);
  if I <> -1 then
    Result := FCentralFolder + ValueFromIndexRight(Self, I)
  else begin
    //find a unique file name to store the cache
    if IsDriveW(Dir) then
      WS := Dir[1]
    else begin
      // NameForParsingInFolder
      WS := '';
      WCH := StrRScanW(PWideChar(StripTrailingBackslashW(Dir)), WideChar('\'));
      if Assigned(WCH) then begin
        WCH := WCH + 1; //get rid of the first '\'
        WS := WCH;
      end;
      // Delete all the '=' chars
      WCH := StrScanW(PWideChar(WS), WideChar('='));
      if Assigned(WCH) then begin
        I := Length(WS) - Integer(StrLenW(WCH));
        if I > 0 then
          Dec(I);
        SetLength(WS, I);
        if WS = '' then
          WS := '0';
      end;
    end;
    F := WS + '.cache';

    //find a unique file name
    I := -1;
    while IndexOfValueRight(Self, F) <> -1 do begin
      inc(I);
      F := WS + '.' + inttostr(I) + '.cache';
    end;

    Add(Dir + '=' + F);
    Sort;
    Result := FCentralFolder + F;
  end;
end;

function TCacheList.GetCacheFileToLoad(Dir: Widestring): WideString;
var
  I: integer;
begin
  Result := '';
  if FCentralFolder <> '' then begin
    I := IndexOfNameRight(Self, IncludeTrailingBackslashW(Dir));
    if I <> -1 then
      Result := FCentralFolder + ValueFromIndexRight(Self, I);
  end;
end;

procedure TCacheList.LoadFromFile;
begin
  if FileExistsW(FCentralFolder + DefaultFileName) then
    inherited LoadFromFile(FCentralFolder + DefaultFileName)
  else
    Clear;
end;

procedure TCacheList.SaveToFile;
begin
  DeleteInvalidFiles;
  inherited SaveToFile(FCentralFolder + DefaultFileName);
end;

procedure TCacheList.DeleteAllFiles;
var
  I: integer;
  F: WideString;
begin
  if FCentralFolder = '' then Exit;

  // Delete all the cache files
  for I := 0 to Count - 1 do begin
    F := FCentralFolder + ValueFromIndexRight(Self, I);
    if (F <> '') and FileExistsW(F) then
      DeleteFileW(PWideChar(F));
  end;
  // Delete the cache list file
  F := FCentralFolder + FDefaultFileName;
  if (F <> '') and FileExistsW(F) then
    DeleteFileW(PWideChar(F));
  // Clear the cache list
  Clear;
end;

procedure TCacheList.DeleteInvalidFiles;
var
  I: integer;
  F: WideString;
begin
  if FCentralFolder = '' then Exit;

  // Delete the invalid cache files
  for I := Count - 1 downto 0 do begin
    // Delete the entry if the cache file doesn't exist
    F := FCentralFolder + ValueFromIndexRight(Self, I);
    if (F = '') or not FileExistsW(F) then
      Delete(I)
    else
      // Delete the cache file if the Dir doesn't exist
      if not DirExistsW(NameFromRight(Self, I)) then
        if (F <> '') and FileExistsW(F) then begin
          DeleteFileW(PWideChar(F));
          Delete(I);
        end;
  end;
end;

procedure TCacheList.SetCentralFolder(const Value: WideString);
begin
  if FCentralFolder <> Value then
    if Value = '' then
      FCentralFolder := ''
    else
      FCentralFolder := IncludeTrailingBackslashW(Value);
end;

//WMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWM
{ TThumbsCacheOptions }

constructor TThumbsCacheOptions.Create(AOwner: TCustomVirtualExplorerListviewEx);
begin
  FOwner := AOwner;
  FThumbsCache := TThumbsCache.Create;
  FCacheList := TCacheList.Create;
  FAutoLoad := False;
  FAutoSave := False;
  FCompressed := False;
  FDefaultFilename := 'Thumbnails.Cache';
  FStorageType := tcsCentral;
  FCacheProcessing := tcpDescending;
end;

destructor TThumbsCacheOptions.Destroy;
begin
  if AutoSave then
    Save; //Save the cache on exit
  FThumbsCache.Free;
  FCacheList.Free;
  inherited;
end;

procedure TThumbsCacheOptions.ClearCache(DeleteAllFiles: Boolean = False);
var
  F: WideString;
begin
  if Assigned(Owner.RootFolderNamespace) and
    SpCompareText(Owner.RootFolderNamespace.NameForParsing, BrowsingFolder) then
  begin
    if DeleteAllFiles then
      case FStorageType of
        tcsPerFolder:
          begin
            F := IncludeTrailingBackslashW(FBrowsingFolder) + FDefaultFilename;
            if (F <> '') and FileExistsW(F) then
              DeleteFileW(PWideChar(F));
          end;
        tcsCentral:
          FCacheList.DeleteAllFiles;
      end;
    Owner.ResetThumbThread; //reset and reload the thread
  end;
end;

procedure TThumbsCacheOptions.Reload(Node: PVirtualNode);
var
  TD: PThumbnailData;
begin
  if Assigned(Node) and (vsInitialized in Node.States) then
    if Owner.ValidateThumbnail(Node, TD) and (TD.State = tsValid) then begin
      // Reset the TD
      TD.Reloading := True;
      TD.State := tsEmpty;
      // TD.CachePos := -1; don't reset the CachePos
      if Owner.ViewStyle = vsxThumbs then
        Owner.FListview.UpdateItems(Node.Index, Node.Index); // invalidate canvas
    end;
end;

procedure TThumbsCacheOptions.Reload(Filename: WideString);
begin
  Reload(Owner.FindNode(Filename));
end;

procedure TThumbsCacheOptions.Load(Force: boolean);
var
  InvalidFiles: TWideStringList;
  N: PVirtualNode;
  NS: TNamespace;
  TD: PThumbnailData;
  F: WideString;
begin
  F := '';
  case FStorageType of
    tcsPerFolder:
      if FDefaultFilename <> '' then
        F := IncludeTrailingBackslashW(FBrowsingFolder) + FDefaultFilename;
    tcsCentral:
        F := FCacheList.GetCacheFileToLoad(FBrowsingFolder);
  end;

  if not Owner.DoThumbsCacheLoad(FThumbsCache, F, FThumbsCache.Comments) then
    FThumbsCache.FLoadedFromFile := True
  else
    if FileExistsW(F) then begin
      // Load only the bitmap streams that are not already loaded
      InvalidFiles := TWideStringList.Create;
      try
        // Iterate through the nodes and populate the InvalidFiles with the images filenames.
        if not Force then begin
          N := Owner.RootNode.FirstChild;
          while Assigned(N) do begin
            if (vsInitialized in N.States) and Owner.ValidateNamespace(N, NS) and Owner.ValidateThumbnail(N, TD) then
              if IsThumbnailActive(TD.State) then
                InvalidFiles.Add(WideLowerCase(NS.NameForParsing));
            N := N.NextSibling;
          end;
        end;
        // Load only the bitmap streams that are not already loaded
        FThumbsCache.LoadFromFile(F, InvalidFiles);
      finally
        InvalidFiles.Free;
      end;
    end;
end;

procedure TThumbsCacheOptions.Save;
var
  F: WideString;
begin
  if (FThumbsCache.Count = 0) or not DirExistsW(FBrowsingFolder) then Exit;

  case FStorageType of
    tcsPerFolder:
      if FDefaultFilename <> '' then begin
        F := IncludeTrailingBackslashW(FBrowsingFolder) + FDefaultFilename;
        if Owner.DoThumbsCacheSave(FThumbsCache, F, FThumbsCache.Comments) then
          FThumbsCache.SaveToFile(F);
      end;
    tcsCentral:
      if CentralFolder <> '' then
        if DirExistsW(CentralFolder) or CreateDirW(CentralFolder) then begin
          F := FCacheList.GetCacheFileToSave(FBrowsingFolder);
          if F <> '' then begin
            if Owner.DoThumbsCacheSave(FThumbsCache, F, FThumbsCache.Comments) then
              FThumbsCache.SaveToFile(F);
          end;
          FCacheList.SaveToFile;
        end;
  end;
end;

function TThumbsCacheOptions.Read(Node: PVirtualNode; var OutCacheItem: TThumbsCacheItem): Boolean;
var
  TD: PThumbnailData;
begin
  Result := False;
  if Assigned(FOwner) and FOwner.ValidateThumbnail(Node, TD) then
    if (TD.State = tsValid) and (TD.CachePos > -1) then
      Result := FThumbsCache.Read(TD.CachePos, OutCacheItem);
end;

function TThumbsCacheOptions.Read(Filename: WideString; var OutCacheItem: TThumbsCacheItem): Boolean;
var
  I: integer;
begin
  Result := False;
  I := FThumbsCache.IndexOf(Filename);
  if I > -1 then
    Result := FThumbsCache.Read(I, OutCacheItem);
end;

function TThumbsCacheOptions.Read(Node: PVirtualNode; OutBitmap: TBitmap): Boolean;
var
  TD: PThumbnailData;
begin
  Result := False;
  if Assigned(FOwner) and FOwner.ValidateThumbnail(Node, TD) then
    if (TD.State = tsValid) and (TD.CachePos > -1) then
      Result := FThumbsCache.Read(TD.CachePos, OutBitmap);
end;

function TThumbsCacheOptions.Read(Filename: WideString; OutBitmap: TBitmap): Boolean;
var
  I: integer;
begin
  Result := False;
  I := FThumbsCache.IndexOf(Filename);
  if I > -1 then
    Result := FThumbsCache.Read(I, OutBitmap);
end;

procedure TThumbsCacheOptions.Assign(Source: TPersistent);
begin
  if Source is TThumbsCacheOptions then begin
    AutoLoad := TThumbsCacheOptions(Source).AutoLoad;
    AutoSave := TThumbsCacheOptions(Source).AutoSave;
    DefaultFilename := TThumbsCacheOptions(Source).DefaultFilename;
    StorageType := TThumbsCacheOptions(Source).StorageType;
    CentralFolder := TThumbsCacheOptions(Source).CentralFolder;
  end
  else
    inherited;
end;

function TThumbsCacheOptions.GetCacheFileFromCentralFolder(Dir: WideString): WideString;
begin
  // Gets the corresponding cache file from the central folder
  Result := FCacheList.GetCacheFileToLoad(Dir);
end;

function TThumbsCacheOptions.RenameCacheFileFromCentralFolder(Dir, NewDirName: WideString; NewCacheFilename: WideString = ''): Boolean;
var
  I: integer;
  D, F, PrevF: WideString;
  NS: TNamespace;
begin
  Result := False;

  if Dir <> '' then
    Dir := IncludeTrailingBackslashW(Dir);
  if NewDirName <> '' then
    NewDirName := IncludeTrailingBackslashW(NewDirName);

  I := IndexOfNameRight(FCacheList, Dir);
  if I > -1 then begin
    // Delete the entry
    PrevF := ValueFromIndexRight(FCacheList, I);
    FCacheList.Delete(I);
    // Rename the previous Dir entry for the new one
    if NewDirName <> '' then
      D := NewDirName
    else
      D := Dir;
    // Rename the previous Cache File for the new one
    if NewCacheFilename <> '' then begin
      F := NewCacheFilename;
      FCacheList.Add(D + '=' + F);
    end
    else
      F := ExtractFileNameW(FCacheList.GetCacheFileToSave(D));
    // Update the BrowsingFolder property
    if SpCompareText(IncludeTrailingBackslashW(BrowsingFolder), Dir) then
      BrowsingFolder := D;
    // Rename the file
    if PrevF <> '' then begin
      NS := TNamespace.CreateFromFileName(FCacheList.CentralFolder + PrevF);
      try
        NS.SetNameOf(F);
      finally
        NS.Free;
      end;
    end;

    Result := True;
  end;
end;

function TThumbsCacheOptions.GetSize: integer;
begin
  Result := FThumbsCache.Size;
end;

function TThumbsCacheOptions.GetThumbsCount: integer;
begin
  Result := FThumbsCache.Count - FThumbsCache.InvalidCount;
end;

procedure TThumbsCacheOptions.SetBrowsingFolder(const Value: WideString);
begin
  if FBrowsingFolder <> Value then begin
    FThumbsCache.Directory := Value;
    FBrowsingFolder := Value;
  end;
end;

procedure TThumbsCacheOptions.SetCacheProcessing(const Value: TThumbsCacheProcessing);
begin
  if FCacheProcessing <> Value then begin
    FCacheProcessing := Value;
  end;
end;

function TThumbsCacheOptions.GetCentralFolder: WideString;
begin
  Result := FCacheList.CentralFolder;
end;

procedure TThumbsCacheOptions.SetCentralFolder(const Value: WideString);
begin
  if FCacheList.CentralFolder <> Value then begin
    FCacheList.CentralFolder := Value;
    FCacheList.LoadFromFile;
  end;
end;

procedure TThumbsCacheOptions.SetCompressed(const Value: boolean);
begin
  if FCompressed <> Value then begin
    FCompressed := Value;
    ClearCache(False);
  end;
end;

//WMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWM
{ TThumbsOptions }

constructor TThumbsOptions.Create(AOwner: TCustomVirtualExplorerListviewEx);
begin
  FOwner := AOwner;
  FCacheOptions := TThumbsCacheOptions.Create(AOwner);

  FThumbsIconAtt.Size.X := 120;
  FThumbsIconAtt.Size.Y := 120;
  FThumbsIconAtt.Spacing.X := 40;
  FThumbsIconAtt.Spacing.Y := 40;

  FBorder := tbFramed;
  FBorderSize := 4;
  FDetailsHeight := 40;
  FHighlight := thMultipleColors;
  FHighlightColor := $EFD3D3;
  FShowSmallIcon := True;
  FShowXLIcons := True;
  FUseShellExtraction := True;
  FUseSubSampling := True;
end;

destructor TThumbsOptions.Destroy;
begin
  FCacheOptions.Free;
  inherited;
end;

function TThumbsOptions.GetWidth: Integer;
begin
  Result := FThumbsIconAtt.Size.X;
end;

function TThumbsOptions.GetHeight: Integer;
begin
  Result := FThumbsIconAtt.Size.Y;
end;

function TThumbsOptions.GetSpaceWidth: Word;
begin
  Result := FThumbsIconAtt.Spacing.X;
end;

function TThumbsOptions.GetSpaceHeight: Word;
begin
  Result := FThumbsIconAtt.Spacing.Y;
end;

procedure TThumbsOptions.SetBorder(const Value: TThumbnailBorder);
begin
  if FBorder <> Value then begin
    FBorder := Value;
    if Owner.ViewStyle = vsxThumbs then Owner.SyncInvalidate;
  end;
end;

procedure TThumbsOptions.SetBorderSize(const Value: Integer);
begin
  if FBorderSize <> Value then begin
    FBorderSize := Value;
    if FBorderSize < 0 then FBorderSize := 0;
    Owner.ResetThumbImageList;
  end;
end;

procedure TThumbsOptions.SetBorderOnFiles(const Value: Boolean);
begin
  if FBorderOnFiles <> Value then begin
    FBorderOnFiles := Value;
    if Owner.ViewStyle = vsxThumbs then Owner.SyncInvalidate;
  end;
end;

function TThumbsOptions.GetDetailedHints: Boolean;
begin
  Result := Owner.FListview.DetailedHints;
end;

procedure TThumbsOptions.SetDetailedHints(const Value: Boolean);
begin
  Owner.FListview.DetailedHints := Value;
end;

procedure TThumbsOptions.SetDetails(const Value: Boolean);
begin
  if FDetails <> Value then begin
    FDetails := Value;
    Owner.ResetThumbImageList;
  end;
end;

procedure TThumbsOptions.SetDetailsHeight(const Value: Integer);
begin
  if FDetailsHeight <> Value then begin
    FDetailsHeight := Value;
    Owner.ResetThumbImageList;
  end;
end;

procedure TThumbsOptions.SetWidth(const Value: Integer);
begin
  if Value <> FThumbsIconAtt.Size.X then begin
    FThumbsIconAtt.Size.X := Value;
    Owner.ResetThumbImageList;
    Owner.ResetThumbThread; //reset and reload the thread
    Owner.ChildListview.UpdateSingleLineMaxChars;
  end;
end;

procedure TThumbsOptions.SetHeight(const Value: Integer);
begin
  if Value <> FThumbsIconAtt.Size.Y then begin
    FThumbsIconAtt.Size.Y := Value;
    Owner.ResetThumbImageList;
    Owner.ResetThumbThread; //reset and reload the thread
  end;
end;

procedure TThumbsOptions.SetSpaceWidth(const Value: Word);
begin
  if Value <> FThumbsIconAtt.Spacing.X then begin
    FThumbsIconAtt.Spacing.X := Value;
    if Owner.ViewStyle = vsxThumbs then begin
      Owner.ResetThumbSpacing;
      Owner.SyncInvalidate;
    end;
  end;
end;

procedure TThumbsOptions.SetSpaceHeight(const Value: Word);
begin
  if Value <> FThumbsIconAtt.Spacing.Y then begin
    FThumbsIconAtt.Spacing.Y := Value;
    if Owner.ViewStyle = vsxThumbs then begin
      Owner.ResetThumbSpacing;
      Owner.SyncInvalidate;
    end;
  end;
end;

function TThumbsOptions.GetHideCaptions: Boolean;
begin
  Result := Owner.ChildListview.HideCaptions;
end;

procedure TThumbsOptions.SetHideCaptions(const Value: Boolean);
begin
  Owner.ChildListview.HideCaptions := Value;
end;

function TThumbsOptions.GetSingleLineCaptions: Boolean;
begin
  Result := Owner.ChildListview.SingleLineCaptions;
end;

procedure TThumbsOptions.SetSingleLineCaptions(const Value: Boolean);
begin
  Owner.ChildListview.SingleLineCaptions := Value;
end;

procedure TThumbsOptions.SetHighlight(const Value: TThumbnailHighlight);
begin
  if FHighlight <> Value then begin
    FHighlight := Value;
    Owner.SyncInvalidate;
  end;
end;

procedure TThumbsOptions.SetHighlightColor(const Value: TColor);
begin
  if FHighlightColor <> Value then begin
    FHighlightColor := Value;
    if FHighlight = thSingleColor then
      Owner.SyncInvalidate;
  end;
end;

procedure TThumbsOptions.SetUseShellExtraction(const Value: Boolean);
begin
  if FUseShellExtraction <> Value then begin
    FUseShellExtraction := Value;
  end;
end;

procedure TThumbsOptions.SetShowSmallIcon(const Value: Boolean);
begin
  if FShowSmallIcon <> Value then begin
    FShowSmallIcon := Value;
    if Owner.ViewStyle = vsxThumbs then Owner.SyncInvalidate;
  end;
end;

procedure TThumbsOptions.SetShowXLIcons(const Value: Boolean);
begin
  if FShowXLIcons <> Value then begin
    FShowXLIcons := Value;
    if Owner.ViewStyle = vsxThumbs then Owner.SyncInvalidate;
  end;
end;

procedure TThumbsOptions.SetStretch(const Value: Boolean);
begin
  if FStretch <> Value then begin
    FStretch := Value;
    Owner.ResetThumbThread; //reset and reload the thread
  end;
end;

procedure TThumbsOptions.SetUseSubsampling(const Value: Boolean);
begin
  if FUseSubsampling <> Value then begin
    FUseSubsampling := Value;
    Owner.ResetThumbThread;
  end;
end;

//WMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWM
{ TThumbThread }

constructor TThumbThread.Create(AOwner: TCustomVirtualExplorerListviewEx);
begin
  inherited Create(False);
  FOwner := AOwner;
  {$IFDEF DELPHI_7_UP}
  Priority := tpIdle;
  {$ENDIF}
  ResetThumbOptions;
end;

procedure TThumbThread.ExtractedInfoLoad(Info: PVirtualThreadIconInfo);
begin
  inherited;
  // Use the UserData2 Pointer to pass the PThumbnailData to the Control(s)
  GetMem(Info.UserData2, SizeOf(TThumbnailThreadData));
  FillChar(Info.UserData2^, SizeOf(TThumbnailThreadData), #0);

  //Clone the FThumbThreadData
  with PThumbnailThreadData(Info.UserData2)^ do begin
    ImageWidth := FThumbThreadData.ImageWidth;
    ImageHeight := FThumbThreadData.ImageHeight;
    CompressedStream := FThumbThreadData.CompressedStream;
    State := FThumbThreadData.State;
    if Assigned(FThumbThreadData.MemStream) then begin
      MemStream := TMemoryStream.Create;
      MemStream.LoadFromStream(FThumbThreadData.MemStream);
    end;
  end;

  //Clear the FThumbThreadData
  FreeAndNil(FThumbThreadData.MemStream);
  FillChar(FThumbThreadData, SizeOf(FThumbThreadData), #0);
end;

function TThumbThread.CreateThumbnail(Filename: WideString;
  var Thumbnail: TBitmap; var ImageWidth, ImageHeight: integer; var CompressIt: boolean): Boolean;
{$IFDEF USEIMAGEEN}
var
  B: TBitmap;
  DestR: TRect;
{$ENDIF}
begin
  CompressIt := false;
  Result := False;
  try
    {$IFDEF USEIMAGEEN}
    if IsDelphiSupportedImageFile(Filename) then //common image, do default
      Result := MakeThumbFromFile(Filename, Thumbnail, FThumbWidth, FThumbHeight,
        false, FTransparentColor, FThumbStretch, FThumbSubsampling, ImageWidth, ImageHeight)
    else begin
      B := TBitmap.Create;
      try
        if LoadImageEnGraphic(Filename, B) then begin
          ImageWidth := B.Width;
          ImageHeight := B.Height;
          DestR := RectAspectRatio(ImageWidth, ImageHeight, FThumbWidth, FThumbHeight, false, FThumbStretch);
          InitBitmap(Thumbnail, DestR.Right, DestR.Bottom, FTransparentColor);
          // StretchDraw is NOT THREADSAFE!!! Use SpStretchDraw instead
          SpStretchDraw(B, Thumbnail, DestR, true);
          Result := true;
        end;
      finally
        B.Free;
      end;
    end;
    {$ELSE}
    Result := MakeThumbFromFile(Filename, Thumbnail, FThumbWidth, FThumbHeight,
      false, FTransparentColor, FThumbStretch, FThumbSubsampling, ImageWidth, ImageHeight);
    {$ENDIF}
  except
    //don't raise any image errors, just ignore them, the state will be tsInvalid
    FThumbThreadData.State := tsInvalid;
  end;

  CompressIt := Result and (ImageWidth > 200) and (ImageHeight > 200);
end;

procedure TThumbThread.ExtractInfo(PIDL: PItemIDList; Info: PVirtualThreadIconInfo);
var
  Folder, Desktop: IShellFolder;
  OldCB: Word;
  OldPIDL: PItemIDList;
  B: TBitmap;
  AllsOk: Boolean;

  function ExtractImageFromPIDL(OutBitmap: TBitmap): Boolean;
  var
    StrRet: TSTRRET;
    WS, Ext: WideString;
    ExtractImage: IExtractImage;
    Buffer: array[0..MAX_PATH] of WideChar;
    Priority, Flags: Longword;
    Size: TSize;
    Bits: HBitmap;
    ByShellExtract: Boolean;
  begin
    Result := False;
    if Succeeded(Folder.GetDisplayNameOf(OldPIDL, SHGDN_FORADDRESSBAR or SHGDN_FORPARSING, StrRet)) then
    begin
      WS := StrRetToStr(StrRet, OldPIDL);
      Ext := WideLowerCase(ExtractFileExtW(WS));

      // ByShellExtract is hardcoded to LargeIcon, take a look at TSyncListView.OwnerDataHint
      ByShellExtract := Info.LargeIcon;

      if not ByShellExtract then
        Result := CreateThumbnail(WS, OutBitmap, FThumbThreadData.ImageWidth, FThumbThreadData.ImageHeight, FThumbThreadData.CompressedStream)
      else
        // Load it with the Shell IExtract Interface
        // If ThumbsOptions.UseShellExtract = False and the file extension is in
        // ShellExtractExtensionsList then ByShellExtract will be True.
        // Suppose you want just images + html files, then you should turn off
        // UseShellExtract and add '.html' to the ShellExtractExtensionsList.
        if Succeeded(Folder.GetUIObjectOf(0, 1, OldPIDL, IID_IExtractImage, nil, ExtractImage)) then
        begin
          Priority := IEI_PRIORITY_NORMAL;
          Flags := IEIFLAG_ASPECT or IEIFLAG_OFFLINE;
          Size.cx := ThumbWidth;
          Size.cy := ThumbHeight;
          if Succeeded(ExtractImage.GetLocation(Buffer, Length(Buffer), Priority, Size, 32, Flags))
            and Succeeded(ExtractImage.Extract(Bits)) and (Bits <> 0) then
          begin
            OutBitmap.Handle := Bits;
            Result := True;
          end;
        end;
    end;
  end;


begin
  StripLastID(PIDL, OldCB, OldPIDL);
  try
    SHGetDesktopFolder(Desktop);
    {JIM // Special case for the desktop children, plus memory leak fix}
    if PIDL <> OldPIDL then
      AllsOk := Succeeded(Desktop.BindToObject(PIDL, nil, IID_IShellFolder, Folder))
    else begin
      Folder := Desktop;
      AllsOk := Assigned(Folder)
    end;

    if AllsOk then
    begin
      OldPIDL.mkid.cb := OldCB;

      FThumbThreadData.ImageWidth := 0;
      FThumbThreadData.ImageHeight := 0;
      FThumbThreadData.CompressedStream := false;
      FThumbThreadData.State := tsInvalid;
      FThumbThreadData.MemStream := nil;

      B := TBitmap.Create;
      try
        B.Canvas.Lock;
        if ExtractImageFromPIDL(B) then
        begin
          FThumbThreadData.MemStream := TMemoryStream.Create;
          try
            B.SaveToStream(FThumbThreadData.MemStream);
            FThumbThreadData.CompressedStream := FThumbCompression and FThumbThreadData.CompressedStream;
            if FThumbThreadData.CompressedStream then //JPEG compressed
              ConvertBitmapStreamToJPGStream(FThumbThreadData.MemStream, 60);
            FThumbThreadData.State := tsValid;
          except
            FreeAndNil(FThumbThreadData.MemStream);
          end;
        end;
      finally
        B.Canvas.Unlock;
        FreeAndNil(B);
      end;
    end;
  finally
    if Assigned(OldPIDL) then
      OldPIDL.mkid.cb := OldCB
  end
end;

procedure TThumbThread.InvalidateExtraction;
begin
  inherited;
  if Assigned(FThumbThreadData.MemStream) then
  begin
    FreeAndNil(FThumbThreadData.MemStream);
    FillChar(FThumbThreadData, SizeOf(FThumbThreadData), #0)
  end
end;

procedure TThumbThread.ReleaseItem(Item: PVirtualThreadIconInfo; const Malloc: IMalloc);
begin
  if Assigned(Item) then
  begin
    // Will only be valid if the image has been extracted already
    if Assigned(PThumbnailThreadData(Item.UserData2)) then
    begin
      FreeAndNil(PThumbnailThreadData(Item.UserData2).MemStream);
      FreeMem(PThumbnailThreadData(Item.UserData2));
    end;
  end;
  inherited;
end;

procedure TThumbThread.ResetThumbOptions;
begin
  QueryList.LockList;
  try
    FThumbWidth := FOwner.ThumbsOptions.Width;
    FThumbHeight := FOwner.ThumbsOptions.Height;
    FThumbStretch := FOwner.ThumbsOptions.Stretch;
    FThumbSubsampling := FOwner.ThumbsOptions.UseSubsampling;
    FThumbCompression := FOwner.ThumbsOptions.CacheOptions.Compressed;
    FTransparentColor := FOwner.Color;
  finally
    QueryList.UnlockList;
  end;
end;

//WMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWM
{ TUnicodeOwnerDataListView }

constructor TUnicodeOwnerDataListView.Create(AOwner: TComponent);
begin
  inherited;
  FEditingItemIndex := -1;
  OwnerData := true;
  Win32PlatformIsUnicode := (Win32Platform = VER_PLATFORM_WIN32_NT);
  FIsComCtl6 := GetComCtlVersion >= $00060000;
end;

procedure TUnicodeOwnerDataListView.CreateWnd;
begin
  inherited;
  // Enables alpha blended selection if ComCtrl 6 is used
  // Unfortunately the flicker is unbearable
  // if IsComCtl6 then
  //   ListView_SetExtendedListViewStyle(Handle, LVS_EX_DOUBLEBUFFER);
end;

procedure TUnicodeOwnerDataListView.CreateWindowHandle(const Params: TCreateParams);
var
  Column: TLVColumn;
begin
  // Do not call inherited, we need to create the unicode handle
  CreateUnicodeHandle(Self, Params, WC_LISTVIEW);
  if (Win32PlatformIsUnicode) then begin
    ListView_SetUnicodeFormat(Handle, True);
    // the only way I could get editing to work is after a column had been inserted
    Column.mask := 0;
    ListView_InsertColumn(Handle, 0, Column);
    ListView_DeleteColumn(Handle, 0);
  end;
end;

procedure TUnicodeOwnerDataListView.WndProc(var Msg: TMessage);
var
  Item: TListitem;
  P: TPoint;
  R: TRect;
  SimulateClick: Boolean;
begin
  // OwnerData TListview bug on WinXP using ComCtl 6: the white space between
  // icons (vsxIcon or vsxThumbs) becomes selectable, but not the space between
  // the captions.
  SimulateClick := False;
  Case Msg.Msg of
    WM_LBUTTONDOWN..WM_MBUTTONDBLCLK:
      if HideCaptions or FIsComCtl6 then begin
        P := Point(Msg.LParamLo, Msg.LParamHi);
        Item := GetItemAt(P.X, P.Y);
        if Assigned(Item) then begin
          // Disallow the click on the caption if HideCaption is true
          if HideCaptions then begin
            R := GetLVItemRect(Item.Index, drLabel);
            if PtInRect(R, P) then
              SimulateClick := True;
          end
          else
            // Disallow the click on the left and right side of the Icon when using ComCtrl 6
            if FIsComCtl6 then begin
              R := GetLVItemRect(Item.Index, drBounds);
              if not PtInRect(R, P) then begin
                Selected := nil;
                SimulateClick := True;
              end;
            end;
        end;
      end;
  end;

  if SimulateClick then begin
    Case Msg.Msg of
      WM_LBUTTONDOWN: MouseDown(mbLeft, [], P.X, P.Y);
      WM_LBUTTONUP: MouseUp(mbLeft, [], P.X, P.Y);
      WM_MBUTTONDOWN: MouseDown(mbMiddle, [], P.X, P.Y);
      WM_MBUTTONUP: MouseUp(mbMiddle, [], P.X, P.Y);
      WM_RBUTTONDOWN: MouseDown(mbRight, [], P.X, P.Y);
      WM_RBUTTONUP:
        begin
          MouseUp(mbRight, [], P.X, P.Y);
          DoContextPopup(P, SimulateClick);
        end;
    end;
  end
  else
    inherited;
end;

procedure TUnicodeOwnerDataListView.CMDenySubclassing(var Message: TMessage);
begin
  // Prevent Theme Manager subclassing
  // If a Windows XP Theme Manager component is used in the application it will
  // try to subclass all controls which do not explicitly deny this.
  Message.Result := 1;
end;

procedure TUnicodeOwnerDataListView.CMFontchanged(var Message: TMessage);
begin
  inherited;
  UpdateSingleLineMaxChars;
end;

procedure TUnicodeOwnerDataListView.CNNotify(var Message: TWMNotify);
var
  Item: TListItem;
  S: string;
  LVEditHandle: THandle;
begin
  if (not Win32PlatformIsUnicode) then
    inherited
  else begin
    with Message do
    begin
      Case NMHdr^.code of
        HDN_TRACKW:
          begin
            NMHdr^.code := HDN_TRACKA;
            try
              inherited;
            finally
              NMHdr^.code := HDN_TRACKW;
            end;
          end;
        LVN_GETDISPINFOW:
          begin
            // call inherited without the LVIF_TEXT flag
            CurrentDispInfo := PLVDispInfoW(NMHdr);
            try
              OriginalDispInfoMask := PLVDispInfoW(NMHdr)^.item.mask;
              PLVDispInfoW(NMHdr)^.item.mask := PLVDispInfoW(NMHdr)^.item.mask and (not LVIF_TEXT);
              try
                NMHdr^.code := LVN_GETDISPINFOA;
                try
                  inherited;
                finally
                  NMHdr^.code := LVN_GETDISPINFOW;
                end;
              finally
                PLVDispInfoW(NMHdr)^.item.mask := OriginalDispInfoMask;
              end;
            finally
              CurrentDispInfo := nil;
            end;

            // handle any text info
            with PLVDispInfoW(NMHdr)^.item do
              if ((mask and LVIF_TEXT) <> 0) and (iSubItem = 0) then
                if HideCaptions and (FEditingItemIndex <> iItem) then
                  pszText[0] := #0
                else
                  StrLCopyW(pszText, PWideChar(GetItemCaption(iItem)), cchTextMax - 1);
          end;
        LVN_ODFINDITEMW:
          with PNMLVFindItem(NMHdr)^ do
          begin
            if ((lvfi.flags and LVFI_PARTIAL) <> 0) or ((lvfi.flags and LVFI_STRING) <> 0) then
              PWideFindString := TLVFindInfoW(lvfi).psz
            else
              PWideFindString := nil;
            lvfi.psz := nil;
            NMHdr^.code := LVN_ODFINDITEMA;
            try
              inherited; {will result in call to OwnerDataFind}
            finally
              TLVFindInfoW(lvfi).psz := PWideFindString;
              NMHdr^.code := LVN_ODFINDITEMW;
              PWideFindString := nil;
            end;
          end;
        LVN_BEGINLABELEDITW:
          begin
            Item := GetItem(PLVDispInfoW(NMHdr)^.item);
            if not CanEdit(Item) then Result := 1;
          end;
        LVN_ENDLABELEDITW:
          with PLVDispInfoW(NMHdr)^ do
            if (item.pszText <> nil) and (item.IItem <> -1) then
              Edit(TLVItem(item));
        LVN_GETINFOTIPW:
          begin
            NMHdr^.code := LVN_GETINFOTIPA;
            try
              inherited;
            finally
              NMHdr^.code := LVN_GETINFOTIPW;
            end;
          end;
        else
          inherited;
      end;
    end;
  end;

  // Handle the edit control:
  // The Edit control is not showed correctly when the font size
  // is large (Height > 15), this is visible in vsReport, vsList and vsSmallIcon
  // viewstyles. We must reshow the Edit control.
  Case Message.NMHdr^.code of
    LVN_GETDISPINFO:  // handle HideCaptions for Ansi text
      if not Win32PlatformIsUnicode and HideCaptions then
          with PLVDispInfoA(Message.NMHdr)^.item do
            if ((mask and LVIF_TEXT) <> 0) and (iSubItem = 0) and (FEditingItemIndex = iItem) then
            pszText[0] := #0;
    LVN_BEGINLABELEDITA, LVN_BEGINLABELEDITW:
      if Message.Result = 0 then begin
        LVEditHandle := ListView_GetEditControl(Handle);
        if LVEditHandle <> 0 then begin
          SetWindowPos(LVEditHandle, 0, 0, 0, 500, 200, SWP_SHOWWINDOW + SWP_NOMOVE);
          if Win32PlatformIsUnicode then begin
            FEditingItemIndex := PLVDispInfoW(Message.NMHdr)^.item.iItem;
            if HideCaptions or SingleLineCaptions then
              SendMessageW(LVEditHandle, WM_SETTEXT, 0, Longint(PWideChar(GetItemCaption(FEditingItemIndex))));
          end
          else begin
            FEditingItemIndex := PLVDispInfoA(Message.NMHdr)^.item.iItem;
            if HideCaptions then begin
              S := GetItemCaption(FEditingItemIndex);
              SendMessage(LVEditHandle, WM_SETTEXT, 0, Longint(PChar(S)));
            end;
          end;
        end;
      end;
    LVN_ENDLABELEDITA, LVN_ENDLABELEDITW:
      FEditingItemIndex := -1;
  end;
end;

procedure TUnicodeOwnerDataListView.WMCtlColorEdit(var Message: TWMCtlColorEdit);
begin
  Message.Result := DefWindowProc(Handle, Message.Msg, Message.ChildDC, Message.ChildWnd);
end;

procedure TUnicodeOwnerDataListView.WMWindowPosChanging(var Message: TWMWindowPosChanging);
begin
  inherited;
  // Bug in ComCtrl 6 in WinXP, does not redraw on resize
  if FIsComCtl6 then
    InvalidateRect(Handle, nil, False);
end;

function TUnicodeOwnerDataListView.GetItem(Value: TLVItemW): TListItem;
begin
  Result := Items[Value.iItem];
end;

procedure TUnicodeOwnerDataListView.SetHideCaptions(const Value: Boolean);
begin
  if Value <> FHideCaptions then begin
    FHideCaptions := Value;
    Invalidate;
  end;
end;

procedure TUnicodeOwnerDataListView.SetSingleLineCaptions(const Value: Boolean);
begin
  if Value <> FSingleLineCaptions then begin
    FSingleLineCaptions := Value;
    IconOptions.WrapText := not Value;
  end;
end;

function TUnicodeOwnerDataListView.GetFilmstripHeight: Integer;
var
  Size: TSize;
  I: Integer;
begin
  Result := 0;
  if Assigned(LargeImages) then begin
    I := GetSystemMetrics(SM_CYHSCROLL);
    if not HideCaptions then begin
      Size := IDEVirtualWideStrings.TextExtentW('W', Font);
      if SingleLineCaptions then
        I := I + (Size.cy * 2)
      else
        I := I + (Size.cy * 4);
    end;
    Result := LargeImages.Height + I + 10;
  end;
end;

function TUnicodeOwnerDataListView.GetLVItemRect(Index: Integer; DisplayCode: TDisplayCode): TRect;
var
  RealWidthHalf: integer;
  Half: integer;
const
  IconSizePlus = 16;
  Codes: array[TDisplayCode] of Longint = (LVIR_BOUNDS, LVIR_ICON, LVIR_LABEL,
    LVIR_SELECTBOUNDS);
begin
  ListView_GetItemRect(Handle, Index, Result, Codes[DisplayCode]);
  if FIsComCtl6 and (DisplayCode <> drLabel) and (ViewStyle = vsIcon) and Assigned(LargeImages) then begin
    // Another WinXP painting issue...
    // Item.displayrect() returns a bigger TRect
    // We need to adjust it here
    RealWidthHalf := (LargeImages.Width + IconSizePlus) div 2;
    Half := (Result.Right - Result.Left) div 2 + Result.Left;
    Result.Left := Half - RealWidthHalf;
    Result.Right := Half + RealWidthHalf;
  end;
end;

procedure TUnicodeOwnerDataListView.UpdateSingleLineMaxChars;
var
  Size: TSize;
begin
  if Assigned(LargeImages) then begin
    Size := IDEVirtualWideStrings.TextExtentW('-', Font);
    if Size.cx <= 0 then Size.cx := 2;
    FSingleLineMaxChars := (LargeImages.Width div Size.cx) - 2;
  end
  else
    FSingleLineMaxChars := 0;
end;

//WMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWM
{ TSyncListView }

constructor TSyncListView.Create(AOwner: TComponent);
begin
  inherited;
  FDefaultTooltipsHandle := 0;

  Visible := false;
  ControlStyle := ControlStyle + [csNoDesignVisible];
  HideSelection := false;
  IconOptions.AutoArrange := true;
  FFirstShiftClicked := -1;
  FThumbnailHintBitmap := TBitmap.Create;
end;

destructor TSyncListView.Destroy;
begin
  FThumbnailHintBitmap.Free;
  inherited;
end;

procedure TSyncListView.CreateHandle;
var
  H: THandle;
begin
  inherited;
  H := ListView_GetToolTips(Handle);
  if H <> 0 then
    FDefaultTooltipsHandle := H;
  UpdateHintHandle;

  // Boris: When the handle is recreated the icon spacing is not updated
  if (HandleAllocated) and (VETController.ViewStyle = vsxThumbs) then
    VETController.ResetThumbSpacing;
end;

function TSyncListView.IsItemGhosted(Item: TListitem): Boolean;
var
  Node: PVirtualNode;
begin
  Result := false;
  if Assigned(Item) then
    if Item.Cut then
      Result := true //Hidden file
    else begin
      Node := PVirtualNode(Item.Data);
      if Assigned(Node) and (vsCutOrCopy in Node.States) then
        Result := true; //Marked as Cut or Copy
    end;
end;

procedure TSyncListView.FetchThumbs(StartIndex, EndIndex: integer);
begin
  //Fire the thread to create ALL the thumbnails at once
  if (Items.Count > 0) and (StartIndex > -1) and (StartIndex < Items.Count) and
  (EndIndex > -1) and (EndIndex < Items.Count) then
     OwnerDataHint(StartIndex, EndIndex);
end;

procedure TSyncListView.UpdateArrangement;
var
  i, Max, Temp: integer;
begin
  if Visible then begin
    //Update column size
    if (ViewStyle in [vsSmallIcon, vsList]) and (Items.Count > 0) then begin
      Max := 0;
      for i := 0 to Items.Count - 1 do
      begin
        Temp := TextExtentW(Items[I].Caption, Canvas).cx;
        if Temp > Max then
          Max := Temp;
      end;
      ListView_SetColumnWidth(Handle, 0, Max + SmallImages.Width + 10);
    end;

    //Update arrangement and scrollbars
    Parent.DisableAlign;
    try
      Width := Width + 1;
      Width := Width - 1;
    finally
      Parent.EnableAlign;
    end;
  end;
end;

procedure TSyncListView.CNNotify(var Message: TWMNotify);
var
  TmpItem: TLVItem;
  ChangeEventCalled, DummyB: boolean;
  Node: PVirtualNode;
begin
  // Update Selection and Focused Item in VET
  if (Message.NMHdr^.code = LVN_ITEMCHANGED) and Assigned(VETController) and (not FSelectionPause) then
    with PNMListView(Message.NMHdr)^ do
      if uChanged = LVIF_STATE then
        if not Assigned(Items[iItem]) then begin
          FPrevEditing := false;  // see CMEnter
          FFirstShiftClicked := -1;
          VETController.ClearSelection;
        end
        else begin
          Node := PVirtualNode(Items[iItem].Data);
          if Assigned(Node) then begin
            ChangeEventCalled := false;
            if ((uOldState and LVIS_SELECTED <> 0) and (uNewState and LVIS_SELECTED = 0)) or
               ((uOldState and LVIS_SELECTED = 0) and (uNewState and LVIS_SELECTED <> 0)) then begin
              VETController.Selected[Node] := (uNewState and LVIS_SELECTED <> 0);
              ChangeEventCalled := true; // changing the selection will call the OnChange event
            end;
            if (uOldState and LVIS_FOCUSED = 0) and (uNewState and LVIS_FOCUSED <> 0) then begin
              FPrevEditing := false;  // see CMEnter
              VETController.FocusedNode := Node;
              if not ChangeEventCalled and Assigned(VETController.OnChange) then
                VETController.OnChange(VETController, Node); // call VET event
            end;
          end;
        end;

  // The VCL totally destroys the speed of this message we will take care of it
  if Message.NMHdr^.code <> NM_CUSTOMDRAW then
    inherited;

  // Handle cdPostPaint, Delphi doesn't do it :(
  if (Message.NMHdr^.code = NM_CUSTOMDRAW) and Assigned(Canvas) then
    with PNMCustomDraw(Message.NMHdr)^ do begin
      case PNMCustomDraw(Message.NMHdr)^.dwDrawStage of
        CDDS_PREPAINT:
          if not FInPaintCycle then
            Message.Result := CDRF_SKIPDEFAULT
          else
            Message.Result := CDRF_NOTIFYITEMDRAW;
        CDDS_ITEMPREPAINT:
          begin
            Message.Result := CDRF_NOTIFYPOSTPAINT;
            Canvas.Lock;
            try
              // We are drawing an item, NOT a subitem, in a postpaint stage
              FillChar(TmpItem, SizeOf(TmpItem), 0);
              TmpItem.iItem := dwItemSpec;
              DummyB := True;
              Canvas.Handle := hdc;

              // Assign the font and brush
              Canvas.Font.Assign(Font);
              Canvas.Brush.Assign(Brush);

              if Assigned(OnAdvancedCustomDrawItem) then
                OnAdvancedCustomDrawItem(Self, Items[TmpItem.iItem], TCustomDrawState(Word(uItemState)), cdPrePaint , DummyB);

              // Set the font and brush colors
              if IsBackgroundValid then begin
                SetBkMode(hdc, TRANSPARENT);
                with PNMLVCustomDraw(Message.NMHdr)^ do begin
                  clrText := CLR_NONE;
                  clrTextBk := CLR_NONE;
                end;
              end
              else
                with PNMLVCustomDraw(Message.NMHdr)^ do begin
                  clrText := ColorToRGB(Canvas.Font.Color);
                  clrTextBk := ColorToRGB(Canvas.Brush.Color);
                end;

              Canvas.Handle := 0;
            finally
              Canvas.Unlock;
            end;
          end;
        CDDS_ITEMPOSTPAINT:
          begin
            Message.Result := CDRF_DODEFAULT;
            Canvas.Lock;
            try
             //We are drawing an item, NOT a subitem, in a postpaint stage
              FillChar(TmpItem, SizeOf(TmpItem), 0);
              TmpItem.iItem := dwItemSpec;
              DummyB := true;
              Canvas.Handle := hdc;
              if Assigned(OnAdvancedCustomDrawItem) then
                OnAdvancedCustomDrawItem(Self, Items[TmpItem.iItem], TCustomDrawState(Word(uItemState)), cdPostPaint, DummyB);
              Canvas.Handle := 0;
            finally
              Canvas.Unlock;
            end;
          end;
      end;
    end;

  // Fire OnEditCancelled
  if (Message.NMHdr^.code = LVN_ENDLABELEDITW) and Assigned(VETController.OnEditCancelled) then
    if Win32PlatformIsUnicode then begin
       with PLVDispInfoW(Message.NMHdr)^ do
         if (item.pszText = nil) or (item.IItem = -1) then
           VETController.OnEditCancelled(VETController, 0);
    end
    else begin
       with PLVDispInfo(Message.NMHdr)^ do
         if (item.pszText = nil) or (item.IItem = -1) then
           VETController.OnEditCancelled(VETController, 0);
    end;
end;

procedure TSyncListView.CMEnter(var Message: TCMEnter);
begin
  //When the Listview is unfocused and a previously selected item caption is
  //clicked it enters in editing mode. This is an incorrect TListview behavior.
  //Set a flag, and deactivate it in CanEdit and when the selection changes in CNNotify
  FPrevEditing := true;
  inherited;
  if Assigned(VETController) and Assigned(VETController.OnEnter) then VETController.OnEnter(VETController);
end;

procedure TSyncListView.CMExit(var Message: TCMExit);
begin
  inherited;
  if Assigned(VETController) and Assigned(VETController.OnExit) then VETController.OnExit(VETController);
end;

procedure TSyncListView.CMMouseWheel(var Message: TCMMouseWheel);
var
  I, dy: integer;
begin
  //Scroll by thumbs
  if (VETController.ViewStyle = vsxThumbs) and (Items.Count > 0) and Assigned(Items[0]) then begin
    if (IconOptions.Arrangement = iaTop) then begin
      I := VETController.ThumbsOptions.Height + VETController.ThumbsOptions.SpaceHeight + 8;
      dy := (Message.WheelDelta div WHEEL_DELTA) * I;
      Scroll(0, -dy);
      Message.Result := 1;
    end
    else begin
      I := VETController.ThumbsOptions.Width + VETController.ThumbsOptions.SpaceHeight + 8;
      dy := (Message.WheelDelta div WHEEL_DELTA) * I;
      Scroll(-dy, 0);
      Message.Result := 1;
    end;
  end
  else
    inherited;
end;

procedure TSyncListView.LVMInsertColumn(var Message: TMessage);
begin
  // Fix the VCL bug for XP
  with PLVColumn(Message.LParam)^ do
  begin
    // Fix TListView report mode bug.
    // But this screws up List style ... grrrr
    if (iImage = - 1) and (ViewStyle = vsReport) then
      Mask := Mask and not LVCF_IMAGE;
  end;
  inherited;
end;

procedure TSyncListView.LVMSetColumn(var Message: TMessage);
begin
  // Fix the VCL bug for XP
  with PLVColumn(Message.LParam)^ do
  begin
    // Fix TListView report mode bug.
    // But this screws up List style ... grrrr
    if (iImage = - 1) and (ViewStyle = vsReport) then
      Mask := Mask and not LVCF_IMAGE;
  end;
  inherited;
end;

procedure TSyncListView.WndProc(var Msg: TMessage);
var
  ContextResult: LRESULT;
begin
  inherited;
  if Assigned(VETController) then
    Case Msg.Msg of
      WM_INITMENUPOPUP, WM_DRAWITEM, WM_MEASUREITEM:
        if Assigned(FSavedPopupNamespace) then //show the Send To item in the contextmenu
          FSavedPopupNamespace.HandleContextMenuMsg(Msg.Msg, Msg.WParam, Msg.LParam, ContextResult);
    end;
end;

function TSyncListView.OwnerDataHint(StartIndex, EndIndex: Integer): Boolean;
var
  I, CacheIndex, W, H: integer;
  CI: TThumbsCacheItem;
  WSExt: WideString;
  Node: PVirtualNode;
  NS: TNamespace;
  Data: PThumbnailData;
  B: TBitmap;
  Cache: TThumbsCache;
  ByImageLibrary, ByShellExtract: boolean;
begin
  Result := inherited OwnerDataHint(StartIndex, EndIndex);

  if (csDesigning in ComponentState) or (not Assigned(VETController)) or (OwnerDataPause) or (Items.Count = 0) then
    exit;

  for I := StartIndex to EndIndex do begin
    Node := VETController.GetNodeByIndex(I); // This is fast enough
    if VETController.ValidateNamespace(Node, NS) then
    begin
      {$IFDEF THREADEDICONS}
      // Call the VET icon thread to pre-load the index
      if not NS.ThreadedIconLoaded and (VETController.ThreadedImagesEnabled) then
      begin
        if not NS.ThreadIconLoading then
        begin
          NS.ThreadIconLoading := True;
          if VETController.ViewStyle = vsxIcon then
            ImageThreadManager.AddNewItem(VETController, WM_VTSETICONINDEX, NS.AbsolutePIDL, True, Node, I)
          else
            ImageThreadManager.AddNewItem(VETController, WM_VTSETICONINDEX, NS.AbsolutePIDL, False, Node, I)
        end;
      end;
      {$ENDIF}

      // Call the thumb thread
      if (VETController.ViewStyle = vsxThumbs) and VETController.ValidateThumbnail(Node, Data) and
        (Data.State = tsEmpty) and (NS.FileSystem) and (not NS.Folder) then
      begin
        Data.State := tsInvalid;
        if not (vsInitialized in Node.States) then
          VETController.InitNode(Node);

        WSExt := ExtractFileExtW(NS.NameForParsing);
        if VETController.ExtensionsExclusionList.IndexOf(WSExt) > -1 then
        begin
          ByImageLibrary := False;
          ByShellExtract := False;
        end else
        begin
          ByImageLibrary := VETController.ExtensionsList.IndexOf(WSExt) > -1;
          ByShellExtract := False;
          if not ByImageLibrary then
            if VETController.ThumbsOptions.UseShellExtraction or (VETController.ShellExtractExtensionsList.IndexOf(WSExt) > -1) then
              ByShellExtract := SupportsShellExtract(NS);
        end;

        if ByImageLibrary or ByShellExtract then
        begin
          Cache := VETController.ThumbsOptions.CacheOptions.FThumbsCache;

          // If the cache was loaded from file we don't need to call the thread :)
          if Cache.LoadedFromFile and not Data.Reloading then begin
            CacheIndex := Cache.IndexOf(NS.NameForParsing);
            if (CacheIndex > -1) and Cache.Read(CacheIndex, CI) then begin
              Data.CachePos := CacheIndex;
              Data.State := tsValid;
              // Update the cache entry if the file was changed
              if not VETController.DoThumbsCacheItemLoad(NS, CI) then
                VETController.ThumbsOptions.CacheOptions.Reload(Node);
            end;
          end
          else begin
            //let the application decide
            if Assigned(VetController.OnThumbsCacheItemRead) then begin
              B := TBitmap.Create;
              try
                if not VETController.DoThumbsCacheItemRead(NS.NameForParsing, B) then begin
                  Data.CachePos := -1;
                  Data.State := tsValid;
                end;
              finally
                B.Free;
              end;
            end;
          end;

          if (Data.State <> tsValid) and Assigned(VETController.ThumbThread) then
          begin
            //Process the thumb outside the thread?
            if Assigned(VetController.OnThumbsCacheItemProcessing) then begin
              B := TBitmap.Create;
              try
                B.Width := VETController.ThumbsOptions.Width;
                B.Height := VETController.ThumbsOptions.Height;
                if not VETController.DoThumbsCacheItemProcessing(NS, B, W, H) then
                  if not Assigned(VETController.OnThumbsCacheItemAdd) or VETController.DoThumbsCacheItemAdd(NS, B, W, H) then begin
                    Data.CachePos := Cache.Add(NS.NameForParsing, '', '', NS.LastWriteDateTime, W, H, False, B);
                    Data.State := tsValid;
                  end;
              finally
                B.Free;
              end;
            end;
            //Call the thread
            if (Data.State <> tsValid) and not VETController.ThumbsThreadPause then begin
              if VETController.ThumbsOptions.LoadAllAtOnce or (VETController.ThumbsOptions.CacheOptions.CacheProcessing = tcpAscending) then
                VETController.ThumbThread.InsertNewItem(VETController, WM_VLVEXTHUMBTHREAD, NS.AbsolutePIDL, ByShellExtract, Node, I) //ByShellExtract hardcoded to LargeIcon parameter
              else
                VETController.ThumbThread.AddNewItem(VETController, WM_VLVEXTHUMBTHREAD, NS.AbsolutePIDL, ByShellExtract, Node, I); //ByShellExtract hardcoded to LargeIcon parameter
              Data.State := tsProcessing; //it should be tsProcessing at first
            end;
          end;
        end;
      end;
    end;
  end;
end;

function TSyncListView.OwnerDataFetch(Item: TListItem; Request: TItemRequest): Boolean;
var
  NS: TNamespace;
  Node: PVirtualNode;
  WS: WideString;
begin
  Result := True;
  if (csDesigning in ComponentState) or not Assigned(Item) or not Assigned(VETController) or
    (Items.Count = 0) or (Item.Index < 0) or (Item.Index >= Items.Count) then Exit;

  Node := VETController.GetNodeByIndex(Item.Index); // this is fast enough
  if VETController.ValidateNamespace(Node, NS) then begin
    if not (vsInitialized in Node.States) then
      VETController.InitNode(Node);
    Item.Data := Node; // Keep a reference of the Node

    // Fill the Item caption for W9x, the displayed caption is in unicode.
    WS := NS.NameInFolder;
    FVETController.DoGetVETText(0, Node, NS, WS);
    Item.Caption := WS;

    if irImage in Request then
    begin
      {$IFDEF THREADEDICONS}
      if not (csDesigning in ComponentState) and VETController.ThreadedImagesEnabled and not NS.ThreadedIconLoaded then
      begin
        // Show default images
        if NS.Folder and NS.FileSystem then Item.ImageIndex := VETController.UnknownFolderIconIndex
        else Item.ImageIndex := VETController.UnknownFileIconIndex
      end
      else begin
        if NS.ThreadIconLoading then begin
          if NS.Folder and NS.FileSystem then Item.ImageIndex := VETController.UnknownFolderIconIndex
          else Item.ImageIndex := VETController.UnknownFileIconIndex
        end else
          Item.ImageIndex := NS.GetIconIndex(false, icLarge);
      end;
      {$ELSE}
      Item.ImageIndex := NS.GetIconIndex(false, icLarge);
      {$ENDIF}
      Item.Cut := NS.Ghosted;

      if not (toHideOverlay in VETController.TreeOptions.VETImageOptions) and Assigned(NS.ShellIconOverlayInterface) then
        Item.OverlayIndex := NS.OverlayIndex - 1
      else
        if NS.Link then
          Item.OverlayIndex := 1
        else
          if NS.Share then Item.OverlayIndex := 0
          else Item.OverlayIndex := -1;
    end;
  end;
end;

function TSyncListView.OwnerDataFind(Find: TItemFind;
  const FindString: string; const FindPosition: TPoint;
  FindData: Pointer; StartIndex: Integer; Direction: TSearchDirection;
  Wrap: Boolean): Integer;
//OnDataFind gets called in response to calls to FindCaption, FindData,
//GetNearestItem, etc. It also gets called for each keystroke sent to the
//ListView (for incremental searching)
var
  I: Integer;
  Found: Boolean;
  Node: PVirtualNode;
  NS: TNamespace;
  WS: WideString;
begin
  Result := -1;
  if Assigned(PWideFindString) then
    WS := PWideFindString
  else
    WS := FindString;

  //search in VET
  if Assigned(VETController) and (not OwnerDataPause) then begin
    I := StartIndex;
    Found := false;
    if (Find = ifExactString) or (Find = ifPartialString) then
    begin
      WS := WideLowerCase(WS);
      repeat
        if I > Items.Count-1 then
          if Wrap then I := 0 else Exit;
        Node := PVirtualNode(Items[I].Data);
        if VETController.ValidateNamespace(Node, NS) then
          Found := Pos(WS, WideLowerCase(NS.NameInFolder)) = 1;
        Inc(I);
      until Found or (I = StartIndex);
      if Found then Result := I-1;
    end;
  end;

  // Fire OnDataFind, don't call inherited
  if Assigned(OnDataFind) then
    OnDataFind(Self, Find, WS, FindPosition, FindData, StartIndex, Direction, Wrap, Result);
end;

function TSyncListView.OwnerDataStateChange(StartIndex, EndIndex: Integer;
  OldState, NewState: TItemStates): Boolean;
begin
  // In OwnerDraw, the selections with the SHIFT and Mouse fire LVN_ODSTATECHANGED
  // and not the the CN_NOTIFY message
  Result := inherited OwnerDataStateChange(StartIndex, EndIndex, OldState, NewState);
  if Assigned(VETController) then begin
    VETController.SyncSelectedItems(False);
    if Assigned(VETController.OnChange) then
      VETController.OnChange(VETController, nil);
  end;
end;

function TSyncListView.GetItemCaption(Index: Integer): WideString;
var
  NS: TNamespace;
begin
  Result := '';
  if Assigned(VETController) then
    if FVETController.ValidateNamespace(PVirtualNode(Items[Index].Data), NS) then
    begin
      Result := NS.NameInFolder;
      FVETController.DoGetVETText(0, PVirtualNode(Items[Index].Data), NS, Result);
      if SingleLineCaptions and (EditingItemIndex <> Index) and (SingleLineMaxChars > 0) then
        if (ViewStyle = vsIcon) then
          Result := SpPathCompactPath(Result, SingleLineMaxChars);
    end;
end;

function TSyncListView.IsBackgroundValid: Boolean;
var
  BK: TLVBKImage;
begin
  Result := False;
  if Assigned(VETController) and (toShowBackground in VETController.TreeOptions.PaintOptions) then
  begin
    Fillchar(BK, SizeOf(BK), 0);
    ListView_GetBkImage(Handle, @BK);
    if BK.ulFlags > 0 then
      Result := True;
  end;
end;

procedure TSyncListView.Edit(const Item: TLVItem);
var
  WS: WideString;
  EditItem: TListItem;
  Node: PVirtualNode;
  NS: TNamespace;
  TD: PThumbnailData;
  ExtChanged: Boolean;
begin
  inherited;
  if (not Win32PlatformIsUnicode) then
    WS := Item.pszText
  else
    WS := TLVItemW(Item).pszText;
  EditItem := GetItem(TLVItemW(Item));
  if Assigned(VETController) and Assigned(EditItem) then begin
    Node := PVirtualNode(EditItem.Data);
    if VETController.ValidateNamespace(Node, NS) and VETController.ValidateThumbnail(Node, TD) then begin
      ExtChanged := not SpCompareText(NS.Extension, ExtractFileExtW(WS));
      if NS.SetNameOf(WS) then begin
        NS.InvalidateCache;
        if ExtChanged then begin
          if TD.State <> tsEmpty then
            TD.State := tsEmpty;
        end;
      end;
      if Assigned(VETController.OnEdited) then VETController.OnEdited(VETController, Node, 0);
    end;
  end;
end;

function TSyncListView.CanEdit(Item: TListItem): Boolean;
var
  Node: PVirtualNode;
begin
  if FPrevEditing or (toVETReadOnly in VETController.TreeOptions.VETMiscOptions) then begin
    FPrevEditing := false;  //see CMEnter
    Result := false;
  end
  else begin
    Result := inherited CanEdit(Item);
    if Assigned(Item) then begin
      if Assigned(VETController) and Assigned(VetController.OnEditing) then begin
        Node := PVirtualNode(Item.Data);
        VetController.OnEditing(VetController, Node, 0, Result);
      end;
    end;
  end;
end;

function GetSelectionBox(LV: TListview): TRect;
var
  Item: TListItem;
begin
  Result := Rect(0, 0, 0, 0);
  //Find the Top-Left and Bottom-Right of the selection box
  Item := LV.Selected;
  if Assigned(Item) then
  begin
    Result.TopLeft := Item.Position;
    Result.BottomRight := Item.Position;
    while Assigned(Item) do
    begin
      Result.Left := Min(Result.Left, Item.Position.X);
      Result.Top := Min(Result.Top, Item.Position.Y);
      Result.Right := Max(Result.Right, Item.Position.X);
      Result.Bottom := Max(Result.Bottom, Item.Position.Y);
      Item := LV.GetNextItem(Item, sdAll, [isSelected]);
    end;
  end;
end;

function IsPointInRect(R: TRect; P: TPoint): boolean;
begin
  // Alternative to PtInRect, since in PtInRect a point on the right or
  // bottom side is considered outside the rectangle.
  Result := (P.X >= R.Left) and (P.X <= R.Right) and (P.Y >= R.Top) and (P.Y <= R.Bottom);
end;

function KeyMultiSelectEx(LV: TListview; SD: TSearchDirection): boolean;
// The MS Listview control in virtual mode (OwnerData) has a bug when
// Shift-Selecting an item, it just selects all the items from the last
// selected to the current selected, index wise.
// This little function will solve this issue.
// From MSDN:
// For normal Listviews LVN_ITEMCHANGING is triggered for every item that is
// being selected:
// http://msdn.microsoft.com/library/default.asp?url=/library/en-us/shellcc/platform/commctls/listview/notifications/lvn_itemchanging.asp
//
// But for ownerdata listviews LVN_ODSTATECHANGED is triggered ONCE with all
// the selected items:
// http://msdn.microsoft.com/library/default.asp?url=/library/en-us/shellcc/platform/commctls/listview/notifications/lvn_odstatechanged.asp
//
// The NMLVODSTATECHANGE member has only 2 integers specifying the range of
// selected items, with no differentiation.
// http://msdn.microsoft.com/library/en-us/shellcc/platform/commctls/listview/structures/nmlvodstatechange.asp
var
  Item: TListItem;
  R, RFocused: TRect;
  P: TPoint;
  I, iClicked, iFocused: integer;
begin
  // TSearchDirection = (sdLeft, sdRight, sdAbove, sdBelow, or sdAll);
  // Note: when ItemFocused or Selected is called the TListItem pointer is changed, weird.
  Result := false;
  if Assigned(LV.ItemFocused) and (LV.Items.Count > 1) and (SD <> sdAll) then
  begin
    R := GetSelectionBox(LV);
    iFocused := LV.ItemFocused.Index;

    Item := LV.GetNextItem(LV.ItemFocused, SD, []);
    iClicked := Item.Index;
    if Assigned(LV.Items[iClicked]) and (iClicked <> iFocused) then
    begin
      P := LV.Items[iClicked].Position;
      if IsPointInRect(R, P) then
      begin
        // Contract the selection box
        Case SD of
          ComCtrls.sdLeft: R.Right := P.X;
          ComCtrls.sdAbove: R.Bottom := P.Y;
          ComCtrls.sdRight: R.Left := P.X;
          ComCtrls.sdBelow: R.Top := P.Y;
        end;
      end
      else begin
        // Expand the selection box
        if P.X < R.Left then R.Left := P.X
        else if P.X > R.Right then R.Right := P.X;
        if P.Y < R.Top then R.Top := P.Y
        else if P.Y > R.Bottom then R.Bottom := P.Y;
      end;

      // Update, select and focus the item
      ListView_GetItemRect(LV.Handle, iFocused, RFocused, LVIR_BOUNDS);
      LV.Items[iClicked].Selected := true;
      LV.Items[iClicked].Focused := true;
      InvalidateRect(LV.Handle, @RFocused, true);

      // Select all items in the selection box and unselect the rest
      for I := 0 to LV.Items.Count-1 do
      begin
        Item := LV.Items[I];
        Item.Selected := IsPointInRect(R, Item.Position);
      end;
      LV.Items[iClicked].MakeVisible(false);
      Result := true;
    end;
  end;
end;

function MouseMultiSelectEx(LV: TListview; iClickedItem, iFirstClicked: integer): boolean;
// The MS Listview control in virtual mode (OwnerData) has a bug when
// Shift-Selecting an item, it just selects all the items from the last
// selected to the current selected, index wise.
// This little function will solve this issue.
// From MSDN:
// For normal Listviews LVN_ITEMCHANGING is triggered for every item that is
// being selected:
// http://msdn.microsoft.com/library/default.asp?url=/library/en-us/shellcc/platform/commctls/listview/notifications/lvn_itemchanging.asp
//
// But for ownerdata listviews LVN_ODSTATECHANGED is triggered ONCE with all
// the selected items:
// http://msdn.microsoft.com/library/default.asp?url=/library/en-us/shellcc/platform/commctls/listview/notifications/lvn_odstatechanged.asp
//
// The NMLVODSTATECHANGE member has only 2 integers specifying the range of
// selected items, with no differentiation.
// http://msdn.microsoft.com/library/en-us/shellcc/platform/commctls/listview/structures/nmlvodstatechange.asp
var
  Item: TListItem;
  R, RFocused: TRect;
  I, iFocused: integer;
begin
  Result := false;
  if (iClickedItem > -1) and (iClickedItem < LV.Items.Count) then
  begin
    // When ItemFocused is called the Item parameter is changed, weird
    iFocused := LV.ItemFocused.Index;

    if Assigned(LV.ItemFocused) and (LV.Items.Count > 1) then
    begin
      // Update, select and focus the item
      ListView_GetItemRect(LV.Handle, iFocused, RFocused, LVIR_BOUNDS);
      LV.Items[iClickedItem].Selected := true;
      LV.Items[iClickedItem].Focused := true;
      InvalidateRect(LV.Handle, @RFocused, true);

      if (iFirstClicked < 0) or (iFirstClicked > LV.Items.Count-1) then
        iFirstClicked := iFocused;  // From the Focused if there's no FirstClicked

      // Set the selection box from the FirstClicked to the ClickedItem
      R.Left := Min(LV.Items[iClickedItem].Position.X, LV.Items[iFirstClicked].Position.X);
      R.Top := Min(LV.Items[iClickedItem].Position.Y, LV.Items[iFirstClicked].Position.Y);
      R.Right := Max(LV.Items[iClickedItem].Position.X, LV.Items[iFirstClicked].Position.X);
      R.Bottom := Max(LV.Items[iClickedItem].Position.Y, LV.Items[iFirstClicked].Position.Y);

      // Select all items in the selection box and unselect the rest
      for I := 0 to LV.Items.Count-1 do
      begin
        Item := LV.Items[I];
        Item.Selected := IsPointInRect(R, Item.Position);
      end;
      Result := true;
    end;
  end;
end;

procedure TSyncListView.WMLButtonDown(var Message: TWMLButtonDown);
var
  Shift: TShiftState;
  P: TPoint;
  R: TRect;
  Item: TListItem;
  Node: PVirtualNode;
  iClicked: integer;
begin
  //This will solve the MS Listview Shift-Select bug
  Shift := KeysToShiftState(Message.Keys);
  P := Point(Message.XPos, Message.YPos);
  if (MultiSelect) and (ViewStyle in [vsIcon, vsSmallIcon]) and (ssShift in Shift) then
  begin
    Item := GetItemAt(P.X, P.Y);
    if Assigned(Item) then
    begin
      iClicked := Item.Index;
      if SelCount = 0 then
        FFirstShiftClicked := -1
      else
        if SelCount = 1 then
          FFirstShiftClicked := Selected.Index;
      MouseMultiSelectEx(Self, iClicked, FFirstShiftClicked);
    end
    else
      inherited;
  end
  else begin
    // Enable Thumbnail checkbox clicks
    if (Shift = [ssLeft]) and (VETController.ViewStyle = vsxThumbs) and
      (toCheckSupport in VETController.TreeOptions.MiscOptions) then
    begin
      Item := GetItemAt(P.X, P.Y);
      if Assigned(Item) then begin
        Node := PVirtualNode(Item.Data);
        if Assigned(Node) and (Node.CheckType <> IDEVirtualTrees.ctNone) then begin
          R := Item.DisplayRect(drIcon);
          if FIsComCtl6 then
            R.Left := R.Left + VETController.ThumbsOptions.SpaceWidth div 2;
          R.Right := R.Left + 15;
          R.Bottom := R.Top + 15;
          if PtInRect(R, P) then begin
            VETController.CheckState[Node] := VETController.DetermineNextCheckState(Node.CheckType, Node.CheckState);
            InvalidateRect(Handle, @R, True);
            Exit;
          end;
        end;
      end;
    end;

    inherited;
  end;
end;

procedure TSyncListView.KeyDown(var Key: Word; Shift: TShiftState);
var
  I, H: integer;
  Node: PVirtualNode;
  DoDefault: boolean;
begin
  inherited;
  if IsEditing then Exit;

  //If the ClientHeight is to small to fit 2 thumbnails the PageUp/PageDown
  //key buttons won't work.
  //This is a VCL TListview bug, to workaround this I had to fake these keys
  //to Up/Down.
  if (VETController.ViewStyle = vsxThumbs) and (Items.Count > 0)
  and Assigned(ItemFocused) and (Key in [VK_NEXT, VK_PRIOR]) then begin
    H := (VETController.ThumbsOptions.Height + VETController.ThumbsOptions.SpaceHeight) * 2 + 5;
    if ClientHeight < H then begin // if there's not at least 2 full visible items
      Case Key of
        VK_NEXT: Key := VK_DOWN;
        VK_PRIOR: Key := VK_UP;
      end;
    end;
  end;

  //Corrected TListview bug in virtual mode, when the icon arrangement is iaLeft the arrow keys are scrambled
  if IconOptions.Arrangement = iaLeft then
    Case Key of
      VK_UP:    Key := VK_LEFT;
      VK_DOWN:  Key := VK_RIGHT;
      VK_LEFT:  Key := VK_UP;
      VK_RIGHT: Key := VK_DOWN;
    end;

  //This will solve the MS Listview Shift-Select bug
  if (MultiSelect) and (ViewStyle in [vsIcon, vsSmallIcon])
  and (ssShift in Shift) and (Key in [VK_UP, VK_DOWN, VK_LEFT, VK_RIGHT]) then
  begin
    Case Key of
      VK_UP:    KeyMultiSelectEx(Self, ComCtrls.sdAbove);
      VK_DOWN:  KeyMultiSelectEx(Self, ComCtrls.sdBelow);
      VK_LEFT:  KeyMultiSelectEx(Self, ComCtrls.sdLeft);
      VK_RIGHT: KeyMultiSelectEx(Self, ComCtrls.sdRight);
    end;
    Key := 0;
  end
  else
    if Assigned(VETController) then begin //call VET events
      DoDefault := true;
      if Assigned(VETController.OnKeyDown) then
        VETController.OnKeyDown(VETController, Key, Shift);
      if Assigned(VETController.OnKeyAction) then
        VETController.OnKeyAction(VETController, Key, Shift, DoDefault);
      if DoDefault then begin
        Case Key of
          VK_RETURN:
            if (ItemFocused <> nil) and (ItemFocused.Selected) then begin
              VETController.ClearSelection;
              Node := PVirtualNode(ItemFocused.Data);
              if Assigned(Node) then begin
                VETController.Selected[Node] := true;
                VETController.DoShellExecute(Node);
              end;
            end;
          VK_BACK:
            VETController.BrowseToPrevLevel;
          VK_F2:
            if (not ReadOnly) and (ItemFocused <> nil) then
              ItemFocused.EditCaption;
          VK_F5:
            VETController.RefreshTree(toRestoreTopNodeOnRefresh in VETController.TreeOptions.VETMiscOptions);
          VK_DELETE:
            VETController.SelectedFilesDelete;
          Ord('A'), Ord('a'):
            if ssCtrl in Shift then begin
              VETController.SelectAll(true);
              for I := 0 to Items.Count - 1 do
                Items[I].Selected := True;
            end;
          Ord('C'), Ord('c'):
            if ssCtrl in Shift then
              VETController.CopyToClipboard;
          Ord('X'), Ord('x'):
            if ssCtrl in Shift then
              VETController.CutToClipboard;
          Ord('V'), Ord('v'):
            if ssCtrl in Shift then
              VETController.PasteFromClipboard;
          Ord('I'), Ord('i'):
            if (ssCtrl in Shift) then begin
              VETController.InvertSelection(False);
              VETController.SyncSelectedItems;
            end;
          VK_INSERT: // Lefties favorite keys!
            if ssShift in Shift then
              VETController.PasteFromClipboard
            else
              if ssCtrl in Shift then
                VETController.CopyToClipboard;
        end;
      end;
    end;
end;

procedure TSyncListView.KeyPress(var Key: Char);
begin
  inherited;
  if Assigned(VETController) and Assigned(VETController.OnKeyPress) then
    VETController.OnKeyPress(VETController, Key);
end;

procedure TSyncListView.KeyUp(var Key: Word; Shift: TShiftState);
begin
  inherited;
  if Assigned(VETController) and Assigned(VETController.OnKeyUp) then
    VETController.OnKeyUp(VETController, Key, Shift);
end;

procedure TSyncListView.DblClick;
var
  Node: PVirtualNode;
begin
  inherited;
  if Assigned(VETController) then begin
    // Set the selection in the VETController
    Node := nil;
    if Assigned(ItemFocused) and (ItemFocused.Selected) then begin
      Node := PVirtualNode(ItemFocused.Data);
      if Assigned(Node) then begin
        VETController.ClearSelection;
        VETController.FocusedNode := Node;
        VETController.Selected[Node] := True;
      end;
    end;
    // Fire VETController.OnDblClick event
    if Assigned(VETController.OnDblClick) then
      VETController.OnDblClick(VETController);
    // Browse the Node
    if Assigned(Node) then
      VETController.DoShellExecute(Node);
  end;
end;

procedure TSyncListView.MouseDown(Button: TMouseButton; Shift: TShiftState;
  X, Y: Integer);
begin
  inherited;
  if Assigned(VETController) and Assigned(VETController.OnMouseDown) then
    VETController.OnMouseDown(VETController, Button, Shift, X, Y);
end;

procedure TSyncListView.MouseMove(Shift: TShiftState; X, Y: Integer);
begin
  inherited;
  if Assigned(VETController) and Assigned(VETController.OnMouseMove) then
    VETController.OnMouseMove(VETController, Shift, X, Y);
end;

procedure TSyncListView.MouseUp(Button: TMouseButton; Shift: TShiftState;
  X, Y: Integer);
begin
  inherited;
  if Assigned(VETController) then begin
    if Assigned(VETController.OnMouseUp) then
      VETController.OnMouseUp(VETController, Button, Shift, X, Y);
    if Assigned(VETController.OnClick) then
      VETController.OnClick(VETController);
  end;
end;

procedure TSyncListView.ContextMenuCmdCallback(Namespace: TNamespace;
  Verb: WideString; MenuItemID: Integer; var Handled: Boolean);
begin
  Handled := False;
  if Assigned(VETController) and Assigned(VETController.OnContextMenuCmd) then
    VETController.OnContextMenuCmd(VETController, Namespace, Verb, MenuItemID, Handled);
  if (Verb = 'rename') and not Handled then
  begin
    Handled := True;
    if (not ReadOnly) and (ItemFocused <> nil) then
      ItemFocused.EditCaption;
  end;
end;

procedure TSyncListView.ContextMenuShowCallback(Namespace: TNamespace;
  Menu: hMenu; var Allow: Boolean);
begin
  Allow := True;
  if Assigned(VETController) and Assigned(VETController.OnContextMenuShow) then
    VETController.OnContextMenuShow(VETController, Namespace, Menu, Allow);
end;

procedure TSyncListView.ContextMenuAfterCmdCallback(Namespace: TNamespace;
  Verb: WideString; MenuItemID: Integer; Successful: Boolean);
begin
  if Successful then
  begin
    if Verb = 'cut' then
      VETController.MarkNodesCut;
    if Verb = 'copy' then
      VETController.MarkNodesCopied;
    Invalidate;
  end
end;

procedure TSyncListView.DoContextPopup(MousePos: TPoint; var Handled: Boolean);
var
  Node: PVirtualNode;
  Pt: TPoint;
begin
  if Assigned(VETController) and not(toVETReadOnly in VETController.TreeOptions.VETMiscOptions) then begin
    Pt := ClientToScreen(MousePos);
    if (toContextMenus in VETController.TreeOptions.VETShellOptions) and Assigned(Selected) then begin
      Handled := true;
      //We are going to work on ExplorerLV, first sync the selected items
      VETController.SyncSelectedItems(false);
      Node := VETController.GetFirstSelected;
      //Save the namespace for WM_INITMENUPOPUP, WM_DRAWITEM, WM_MEASUREITEM messages
      if VETController.ValidateNamespace(Node, FSavedPopupNamespace) then
          FSavedPopupNamespace.ShowContextMenuMulti(Self, ContextMenuCmdCallback,
            ContextMenuShowCallback, ContextMenuAfterCmdCallback, VETController.SelectedToNamespaceArray, @Pt,
            VETController.ShellContextSubMenu, VETController.ShellContextSubMenuCaption);
    end
    else
      if Assigned(VETController.PopupMenu) then begin
        Handled := True;
        VETController.PopupMenu.Popup(Pt.x, Pt.y);
      end;
  end;

  if not Handled then
    inherited;
end;

procedure TSyncListView.SetDetailedHints(const Value: Boolean);
begin
  //Disable the tooltip that is shown when an item caption is truncated
  if FDetailedHints <> Value then begin
    if Value then
      VETController.ShowHint := True;
    FDetailedHints := Value;
    UpdateHintHandle;
  end;
end;

procedure TSyncListView.WMPaint(var Message: TWMPaint);
begin
  FInPaintCycle := True;
  inherited;
  FInPaintCycle := False;
end;

procedure TSyncListView.WMVScroll(var Message: TWMVScroll);
  // Local function by Peter Bellow
  // http://groups.google.com/groups?hl=en&lr=&ie=UTF-8&selm=VA.00007ac9.007e7c13%40antispam.compuserve.com
  function FindDynamicMethod( aClass: TClass; anIndex: SmallInt ): Pointer;
  type
    TIndices= Array [1..1024] of SmallInt;
    TProcs  = Array [1..1024] of Pointer;
  var
    pDMT : PWord;
    i, count: Word;
    pIndices : ^TIndices;
    pProcs : ^TProcs;
  begin
    Result := nil;
    If aClass = nil Then
      Exit;
    pDMT := Pointer(aClass);
    // Find pointer to DMT in VMT, first Word is the count of dynamic
    // methods
    pDMT := Pointer(PDword( Integer(pDMT) + vmtDynamicTable )^);
    count := pDMT^;
    pIndices := Pointer( Integer( pDMT ) + 2 );
    pProcs   := Pointer( Integer( pDMT ) + 2 + count * sizeof( smallint ));
    // find handler for anIndex
    for i:= 1 to count do
      if pIndices^[i] = anIndex then begin
        Result := pProcs^[i];
        Break;
      end;
    If Result = nil Then
      Result := FindDynamicMethod( aClass.Classparent, anIndex );
  end;
{$IFDEF DELPHI_7_UP}
var
  oldWMVScroll: procedure(var Message: TWMVScroll) of object;
{$ENDIF}
begin
  // Delphi 7 bug: the Listview invalidates the canvas when Scrolling :(
  // The problem is in ComCtrls.pas: TCustomListView.WMVScroll
  {$IFDEF DELPHI_7_UP}
    // Call the grandfather WM_VSCROLL handler, sort of inherited-inherited
    TMethod(oldWMVScroll).code := FindDynamicMethod(TWincontrol, WM_VSCROLL);
    TMethod(oldWMVScroll).data := Self;
    oldWMVScroll(Message);
  {$ELSE}
    inherited;
  {$ENDIF}

  // Set the focus when the scrollbar is clicked
  if not Focused then
    SetFocus;
end;

procedure TSyncListView.WMHScroll(var Message: TWMHScroll);
begin
  inherited;
  // Set the focus when the scrollbar is clicked
  if not Focused then
    SetFocus;
end;

procedure TSyncListView.UpdateHintHandle;
begin
  if HandleAllocated then begin
    if not ShowHint then
      // VCL TListview bug, setting ShowHint to False doesn't disable the hints
      // We must do this explicitly
      ListView_SetToolTips(Handle, 0)
    else
      if FDetailedHints then
        ListView_SetToolTips(Handle, 0)
      else
        ListView_SetToolTips(Handle, FDefaultTooltipsHandle);
  end;
end;

procedure TSyncListView.CMShowHintChanged(var Message: TMessage);
begin
  inherited;
  // VCL TListview bug, setting ShowHint to False doesn't disable the hints
  // We must do this explicitly
  UpdateHintHandle;
end;

procedure TSyncListView.CMHintShow(var Message: TCMHintShow);
var
  HintInfo: PHintInfo;
  Item: TListItem;
  Node: PVirtualNode;
  NS: TNamespace;
  S: WideString;
  R: TRect;
  P: TPoint;
  OverlayI: integer;
  Style: Cardinal;
begin
  if FDetailedHints and Assigned(VETController) and (VETController.ViewStyle <> vsxReport) then begin
    HintInfo := TCMHintShow(Message).HintInfo;
    Item := GetItemAt(HintInfo.CursorPos.X, HintInfo.CursorPos.Y);
    if Assigned(Item) and VETController.ValidateNamespace(PVirtualNode(Item.Data), NS) then begin
      Node := PVirtualNode(Item.Data);
      //Set the Hint
      HintInfo.HintWindowClass := TBitmapHint; //custom HintWindow class
      HintInfo.HintData := FThumbnailHintBitmap; //TApplication.ActivateHint will pass the data to the HintWindow
      HintInfo.HintStr := Item.Caption;
      HintInfo.CursorRect := GetLVItemRect(Item.Index, drBounds);
      HintInfo.CursorRect.TopLeft := ClientToScreen(HintInfo.CursorRect.TopLeft);
      HintInfo.CursorRect.BottomRight := ClientToScreen(HintInfo.CursorRect.BottomRight);
//        HintInfo.HintPos.X := HintInfo.CursorRect.Left + GetSystemMetrics(SM_CXCURSOR) - 5;
//        HintInfo.HintPos.Y := HintInfo.CursorRect.Top + GetSystemMetrics(SM_CYCURSOR) ;
      HintInfo.HintMaxWidth := ClientWidth;
      HintInfo.HideTimeout := 60000; //1 minute

      //Draw in the hint
      S := VETController.DoThumbsGetDetails(Node, true);
      if S <> '' then begin
        R := Rect(0, 0, 0, 0);

        if Win32Platform = VER_PLATFORM_WIN32_WINDOWS then
          Windows.DrawTextA(FThumbnailHintBitmap.Canvas.Handle, PAnsiChar(AnsiString(S)), -1, R, DT_CALCRECT)
        else
          Windows.DrawTextW(FThumbnailHintBitmap.Canvas.Handle, PWideChar(S), -1, R, DT_CALCRECT);

        FThumbnailHintBitmap.Width := R.Right + LargeSysImages.Width + 16;
        if R.Bottom >= LargeSysImages.Height then
          FThumbnailHintBitmap.Height := R.Bottom + 8
        else
          FThumbnailHintBitmap.Height := LargeSysImages.Height + 8;
        FThumbnailHintBitmap.Canvas.Font.Color := clInfoText;
        FThumbnailHintBitmap.Canvas.Pen.Color := clBlack;
        FThumbnailHintBitmap.Canvas.Brush.Color := clInfoBk;
        FThumbnailHintBitmap.Canvas.FillRect(Rect(0, 0, FThumbnailHintBitmap.Width, FThumbnailHintBitmap.Height));

        //Custom drawing
        if VETController.DoThumbsDrawHint(FThumbnailHintBitmap, Node) then begin
          P.x := 4;
          P.y := (FThumbnailHintBitmap.Height - LargeSysImages.Height) div 2;
          Style := ILD_TRANSPARENT;

          OverlayI := -1;
          if not (toHideOverlay in VETController.TreeOptions.VETImageOptions) and Assigned(NS.ShellIconOverlayInterface) then
            OverlayI := NS.OverlayIndex - 1
          else
            if NS.Link then
              OverlayI := 1
            else
              if NS.Share then OverlayI := 0;

          if OverlayI > -1 then
            Style := Style or ILD_OVERLAYMASK and Cardinal(IndexToOverlayMask(OverlayI + 1));
          ImageList_DrawEx(LargeSysImages.Handle, NS.GetIconIndex(false, icLarge), FThumbnailHintBitmap.Canvas.Handle, P.x, P.y, 0, 0, CLR_NONE, CLR_NONE, Style);

          OffsetRect(R, LargeSysImages.Width + 8, (FThumbnailHintBitmap.Height - R.Bottom) div 2);
          if Win32Platform = VER_PLATFORM_WIN32_WINDOWS then
            Windows.DrawTextA(FThumbnailHintBitmap.Canvas.Handle, PAnsiChar(AnsiString(S)), -1, R, 0)
          else
            Windows.DrawTextW(FThumbnailHintBitmap.Canvas.Handle, PWideChar(S), -1, R, 0);
        end;

        Message.Result := 0;
      end;
    end;
  end
  else
    inherited;
end;

procedure TSyncListView.WMEraseBkgnd(var Message: TWMEraseBkgnd);
begin
  if IsBackgroundValid then begin
    DefaultHandler(Message);
    Message.Result := 1;
  end
  else
    inherited;
end;

//WMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWM
{ TOLEListview }

procedure TOLEListview.AutoScrollTimerCallback(Window: hWnd; Msg,
  idEvent: integer; dwTime: Longword);
var
  Pt: TPoint;
begin
  FAutoScrolling := True;
  try
    Pt := Mouse.CursorPos;
    Pt := ScreenToClient(Pt);
    if Pt.y < 20 then
      Scroll(0, -(20 - Pt.y)*2);
    if Pt.y > ClientHeight - 20 then
      Scroll(0, (20 - (ClientHeight - Pt.y)) * 2);
    if Pt.x < 20 then
      Scroll(-(20 - Pt.x)*2, 0);
    if Pt.x > ClientWidth - 20 then
      Scroll((20 - (ClientWidth - Pt.x)) * 2, 0);
  finally
    FAutoScrolling := False;
  end
end;

procedure TOLEListview.ClearTimers;
begin
  if FScrollTimer <> 0 then
  begin
    KillTimer(Handle, FScrollTimer);
    FScrollTimer := 0;
  end;
  if FScrollDelayTimer <> 0 then
  begin
    KillTimer(Handle, FScrollDelayTimer);
    FScrollDelayTimer := 0;
  end;
end;

constructor TOLEListview.Create(AOwner: TComponent);
begin
  inherited;
  CurrentDropIndex := -2;  // -1 = backgound -2 = nothing
  FDragItemIndex := -1;
  FAutoScrollTimerStub := CreateStub(Self, @TOLEListview.AutoScrollTimerCallback);
end;

procedure TOLEListview.CreateDragImage(TotalDragRect: TRect; RectArray: TRectArray; var Bitmap: TBitmap);
var
  i: integer;
  LVBitmap: TBitmap;
  Offset: TPoint;
  R: TRect;
begin
  if Assigned(Bitmap) and (SelCount > 0) then
  begin
    //Get a Bitmap with all the selected items
    LVBitmap := TBitmap.Create;
    Bitmap.Canvas.Lock;
    try
      //update Bitmap size
      Bitmap.Width := TotalDragRect.Right - TotalDragRect.Left;
      Bitmap.Height := TotalDragRect.Bottom - TotalDragRect.Top;
      Bitmap.Canvas.Brush.Color := Self.Color;
      Bitmap.Canvas.FillRect(Rect(0, 0, Bitmap.Width, Bitmap.Height));

      LVBitmap.Width := ClientWidth;
      LVBitmap.Height := ClientHeight;
      LVBitmap.Canvas.Lock; //we need to Lock the canvas in order to use the PaintTo method
      try
        //Don't use PaintTo, WinXP doesn't support it.
        BitBlt(LVBitmap.Canvas.Handle, 0, 0, ClientWidth, ClientHeight, Canvas.Handle, 0, 0, srcCopy);
      finally
        LVBitmap.Canvas.UnLock;
      end;

      Offset.X := TotalDragRect.Left;
      Offset.Y := TotalDragRect.Top;
      //Iterate and CopyRect the selected items
      for I := 0 to Length(RectArray) - 1 do begin
        R := RectArray[I];
        if R.Top <= 0 then R.Top := 0;  //don't draw borders
        if R.Left <= 0 then R.Left := 0;  //don't draw borders
        StretchBlt(Bitmap.Canvas.Handle, R.Left - Offset.x, R.top - Offset.y, R.Right - R.Left, R.Bottom - R.Top,
          LVBitmap.Canvas.Handle, R.Left+2, R.Top+2, R.Right - R.Left+2, R.Bottom - R.Top+2, cmSrcCopy);
      end;
    finally
      Bitmap.Canvas.Unlock;
      LVBitmap.Free;
    end;
  end;
end;

procedure TOLEListview.CreateWnd;
begin
  inherited;
  if not (csDesigning in ComponentState) then
  begin
    CoCreateInstance(CLSID_DragDropHelper, nil, CLSCTX_INPROC_SERVER, IID_IDropTargetHelper, FDropTargetHelper);
    RegisterDragDrop(Handle, Self)
  end
end;

destructor TOLEListview.Destroy;
begin
  ClearTimers;
  if Assigned(FAutoScrollTimerStub) then
    DisposeStub(FAutoScrollTimerStub);
  inherited;
end;

procedure TOLEListview.DoContextPopup(MousePos: TPoint; var Handled: Boolean);
begin
  // Clear the mouse states as the context menu grabs the mouse and never sends us a WM_xMOUSEUP message
  MouseButtonState := [];
  inherited;
end;

procedure TOLEListview.DestroyWnd;
begin
  if not (csDesigning in ComponentState) then
    RevokeDragDrop(Handle);
  inherited;
end;

function TOLEListview.DragEnter(const dataObj: IDataObject;
  grfKeyState: Integer; pt: TPoint; var dwEffect: Integer): HResult;
begin
  if Assigned(DropTargetHelper) then
    DropTargetHelper.DragEnter(Handle, dataObj, Pt, dwEffect);

  DragDataObject := DataObj;

  FScrollDelayTimer := SetTimer(Handle, SCROLL_DELAY_TIMER, VETController.AutoScrollDelay, nil);

  Result := S_OK;
end;

function TOLEListview.Dragging: Boolean;
begin
  Result := FDragging
end;

function TOLEListview.DragLeave: HResult;
var
  TempNS: TNamespace;
  TempItem: TListItem;
begin
  if Assigned(DropTargetHelper) then
    DropTargetHelper.DragLeave;

  ClearTimers;

  TempNS := nil;
  if CurrentDropIndex > -2 then
  begin
    if (CurrentDropIndex > -1) and (CurrentDropIndex < Items.Count) then
    begin
      TempItem := Items[CurrentDropIndex];
      TempItem.DropTarget := False;   // DropTarget only hilight caption
      TempNS := ListItemToNamespace(TempItem, True);
    end else
       VETController.ValidateNamespace(VETController.RootNode, TempNS);
    if Assigned(TempNS) then
      TempNS.DragLeave;
  end;
  CurrentDropIndex := -2;
  DragDataObject := nil;
  Result := S_OK;
end;


function TOLEListview.DragOverOLE(grfKeyState: Integer; pt: TPoint; var dwEffect: Integer): HResult;
var
  HitNS, TempNS: TNamespace;
  HitItem, TempItem: TListItem;
  HitIndex: integer;
  ShiftState: TShiftState;
begin
  // Update any drag image
  if Assigned(DropTargetHelper) then
    DropTargetHelper.DragOver(pt, dwEffect);

  Result := S_OK;
  if AutoScrolling or not (toAcceptOLEDrop in VETController.TreeOptions.MiscOptions) or
  (toVETReadOnly in VETController.TreeOptions.VETMiscOptions) then begin
    dwEffect := DROPEFFECT_NONE;
    Exit;
  end;

  ShiftState := KeysToShiftState(grfKeyState);

  // Fire VETController.OnDragOver event
  VETController.DoDragOver(Self, ShiftState, dsDragMove, Pt, dmOnNode, dwEffect);

  Pt := ScreenToClient(Pt);

  HitItem := GetItemAt(Pt.X, Pt.Y);
  HitNS := ListItemToNamespace(HitItem, True);
  if Assigned(HitItem) then
    HitIndex := HitItem.Index
  else
    HitIndex := -1;

  // Don't allow to drop in the dragging item unless Shift, Alt or Ctrl is
  // pressed and the item is inside the listview
  if (HitIndex = FDragItemIndex) and (FDragItemIndex > -1) and (ShiftState * [ssRight, ssShift, ssAlt, ssCtrl] = []) then begin
    if (CurrentDropIndex > -1) and (CurrentDropIndex < Items.Count) then
      Items[CurrentDropIndex].DropTarget := False;  // reset highlight caption
    dwEffect := DROPEFFECT_NONE;
    CurrentDropIndex := -2;
    exit;
  end;

  // If the HitIndex  is different that the current drop target then
  // update everything to select the new item (or parent if the drop is "into" the list view
  if (HitIndex <> CurrentDropIndex) then begin
    //<<<<<if GetHitTestInfoAt(Pt.X, Pt.Y) * [htOnIcon, htOnLabel] <> [] then begin
    if HitIndex > -1 then begin
      // Try to enter the new namespace
      Result := HitNS.DragEnter(DragDataObject, grfKeyState, pt, dwEffect);

      // If we can't drop on that namespace then we need to just default to dropping
      // "into" the current list view, i.e. the RootNode of the VT. Otherwise every
      // thing is hunky dory and the HitItem will be selected
      if dwEffect <> DROPEFFECT_NONE then
        HitItem.DropTarget := True    // DropTarget only hilight caption
      else begin
        HitNS.DragLeave;  // Leave the HitItem namespace, not going to use it
        exit; // cancel the drag
      end;
    end;

    // If we were on the background and the hit node does not take drops leave it
    // in the backgound with making any changes
    if CurrentDropIndex <> HitIndex then
    begin
      TempNS := nil;
      if (CurrentDropIndex > -1) and (CurrentDropIndex < Items.Count) then begin
        TempItem := Items[CurrentDropIndex];
        TempItem.DropTarget := False;   // reset highlight caption
        TempNS := ListItemToNamespace(TempItem, False);
      end
      else
        VETController.ValidateNamespace(VETController.RootNode, TempNS);

      if Assigned(TempNS) then begin
        // Only drag leave if the current actually was somewhere (-2 means current was over nothing)
        if (CurrentDropIndex > -2) and (CurrentDropIndex < Items.Count) then
          TempNS.DragLeave;
        TempNS := ListIndexToNamespace(HitIndex);
        TempNS.DragEnter(DragDataObject, grfKeyState, pt, dwEffect);
      end;
      CurrentDropIndex := HitIndex
    end
    else begin
      dwEffect := DROPEFFECT_NONE;
    end;
  end
  else begin
    // Don't allow to drop in the background unless Shift, Alt or Ctrl is
    // pressed and the drag item is INSIDE the Listview
    if (HitIndex = -1) and (FDragItemIndex > -1) and (ShiftState * [ssRight, ssShift, ssAlt, ssCtrl] = []) then
      dwEffect := DROPEFFECT_NONE
    else begin
      TempNS := ListIndexToNamespace(CurrentDropIndex);
      if Assigned(TempNS) then
        TempNS.DragOver(grfKeyState, pt, dwEffect);
    end;
  end;
end;

function TOLEListview.Drop(const dataObj: IDataObject;
  grfKeyState: Integer; Pt: TPoint; var dwEffect: Integer): HResult;
var
  TempNS: TNamespace;
  TempItem: TListItem;
  ClientPt: TPoint;
  I: Integer;
begin
  FDropped := True;
  try
    if Assigned(DropTargetHelper) then
      DropTargetHelper.Drop(dataObj, Pt, dwEffect);
    ClearTimers;
    TempNS := nil;

    // Fire VETController.OnDragDrop event
    I := dwEffect;
    ClientPt := ScreenToClient(Pt);
    VETController.DoDragDrop(Self, dataObj, nil, KeysToShiftState(grfKeyState), ClientPt, I, dmOnNode);

    if (CurrentDropIndex > -2) and (I <> DROPEFFECT_NONE) then
    begin
      if (CurrentDropIndex > -1) and (CurrentDropIndex < Items.Count) then
      begin
        TempItem := Items[CurrentDropIndex];
        TempItem.DropTarget := False;   // DropTarget only highlight caption
        TempNS := ListItemToNamespace(TempItem, True);
      end else
         VETController.ValidateNamespace(VETController.RootNode, TempNS);
      if Assigned(TempNS) then
        TempNS.Drop(dataObj, grfKeyState, pt, dwEffect);
    end;
    CurrentDropIndex := -2;
    DragDataObject := nil;
    Result := S_OK;

    // Fire OnDragDrop for the TOLEListview control
    if I <> DROPEFFECT_NONE then
      DragDrop(Self, ClientPt.X, ClientPt.Y);
  finally
    FDropped := False;
  end;
end;

function TOLEListview.GiveFeedback(dwEffect: Integer): HResult;
begin
  Result := DRAGDROP_S_USEDEFAULTCURSORS
end;

function TOLEListview.ListIndexToNamespace(ItemIndex: integer): TNamespace;
// use -1 to get the Listview background namespace
var
  Node: PVirtualNode;
begin
  Result := nil;
  if (ItemIndex > -1) and (ItemIndex < Items.Count) then begin
    Node := VETController.GetNodeByIndex(ItemIndex);  //this is fast enough
    VETController.ValidateNamespace(Node, Result);
  end
  else
    if ItemIndex = -1 then
      VETController.ValidateNamespace(VETController.RootNode, Result);
end;

function TOLEListview.ListItemToNamespace(Item: TListItem; BackGndIfNIL: Boolean): TNamespace;
var
  Node: PVirtualNode;
begin
  Result := nil;
  if Assigned(Item) then begin
    Node := VETController.GetNodeByIndex(Item.Index);  //this is fast enough
    VETController.ValidateNamespace(Node, Result);
  end
  else
    if BackGndIfNIL then
      VETController.ValidateNamespace(VETController.RootNode, Result);
end;

function TOLEListview.QueryContinueDrag(fEscapePressed: BOOL;
  grfKeyState: Integer): HResult;
begin
  Result := S_OK;

  if fEscapePressed then
    Result := DRAGDROP_S_CANCEL
  else
    if LButtonDown in MouseButtonState then begin
      if grfKeyState and MK_LBUTTON > 0 then  // is the LButton flag set?
        Result := S_OK                        // Button is still down
      else
        Result := DRAGDROP_S_DROP;            // Button has been released
    end
    else
      if RButtonDown in MouseButtonState then begin
        if grfKeyState and MK_RBUTTON > 0 then  // is the RButton flag set?
          Result := S_OK                        // Button is still down
        else
          Result := DRAGDROP_S_DROP;            // Button has been released
      end;
end;

procedure TOLEListview.WMLButtonDown(var Message: TWMLButtonDown);
begin
  Include(FMouseButtonState, LButtonDown);
  inherited;
end;

procedure TOLEListview.WMLButtonUp(var Message: TWMLButtonUp);
begin
  Exclude(FMouseButtonState, LButtonDown);
  inherited
end;

procedure TOLEListview.WMMouseMove(var Message: TWMMouseMove);
var
  dwOkEffects, dwEffectResult: LongInt;
  DataObject: IDataObject;
  NSArray: TNamespaceArray;
  i: integer;
  Item: TListItem;
  Pt: TPoint;
  DoDrag: Boolean;
  Bitmap: TBitmap;
  DragSourceHelper: IDragSourceHelper;
  SHDragImage: TSHDragImage;
  TotalDragRect, R: TRect;
  RectArray: TRectArray;
  DummyDragObject: TDragObject;
begin
  if MouseButtonState * [LButtonDown, RButtonDown] <> [] then
  begin
    DoDrag := False;
    Pt := SmallPointToPoint(Message.Pos);
    Item := GetItemAt(Pt.X, Pt.Y);

    if Assigned(Item) then
      DoDrag := (GetHitTestInfoAt(Pt.X, Pt.Y) * [htOnLabel, htOnIcon] <> []) and VETController.DoBeforeDrag(Item.Data, -1);

    if DoDrag and (SelCount > 0) then
    begin
      FDragging := DragDetectPlus(Parent.Handle, Pt);
      if Dragging then
      begin
        DummyDragObject := nil;
        // Fire OnStartDrag for the TOLEListview control
        DoStartDrag(DummyDragObject);
        // Fire VETController.OnStartDrag
        VETController.DoStartDrag(DummyDragObject);

        FDragItemIndex := Item.Index;
        Bitmap := TBitmap.Create;
        try
          SetLength(NSArray, SelCount);
          SetLength(RectArray, 1);
          Item := Selected;
          NSArray[0] := ListItemToNamespace(Item, False);
          RectArray[0] := GetLVItemRect(Item.Index, drSelectBounds);
          TotalDragRect := RectArray[0];
          if Assigned(NSArray[0]) then
          begin
            i := 1;
            while (i < SelCount) do
            begin
              Item := GetNextItem(Item, sdAll, [isSelected]);
              NSArray[i] := ListItemToNamespace(Item, False);

              //Add visible items bounds to the RectArray
              R := GetLVItemRect(Item.Index, drSelectBounds);
              if PtInRect(ClientRect, R.TopLeft) then begin
                SetLength(RectArray, i + 1);
                RectArray[i] := R;
                //update TotalDragRect size
                if R.Left < TotalDragRect.Left then TotalDragRect.Left := R.Left;
                if R.Top < TotalDragRect.Top then TotalDragRect.Top := R.Top;
                if R.Right > TotalDragRect.Right then TotalDragRect.Right := R.Right;
                if R.Bottom > TotalDragRect.Bottom then TotalDragRect.Bottom := R.Bottom;
              end;

              Inc(i)
            end;

            DataObject := NSArray[0].DataObjectMulti(NSArray);

            if VETController.CanShowDragImage and Succeeded(CoCreateInstance(CLSID_DragDropHelper, nil, CLSCTX_INPROC_SERVER, IID_IDragSourceHelper, DragSourceHelper)) then
            begin
              FillChar(SHDragImage, SizeOf(SHDragImage), #0);
              Bitmap.Width := VETController.DragWidth;
              Bitmap.Height := VETController.DragHeight;
              CreateDragImage(TotalDragRect, RectArray, Bitmap);
              SHDragImage.sizeDragImage.cx := Bitmap.Width;
              SHDragImage.sizeDragImage.cy := Bitmap.Height;
              SHDragImage.ptOffset.X := SmallPointToPoint(Message.Pos).X - TotalDragRect.Left;
              SHDragImage.ptOffset.Y := SmallPointToPoint(Message.Pos).Y - TotalDragRect.Top;
              SHDragImage.ColorRef := ColorToRGB(Color);
              SHDragImage.hbmpDragImage := CopyImage(Bitmap.Handle, IMAGE_BITMAP, 0, 0, LR_COPYRETURNORG);
              if SHDragImage.hbmpDragImage <> 0 then
                if not Succeeded(DragSourceHelper.InitializeFromBitmap(SHDragImage, DataObject)) then
                  DeleteObject(SHDragImage.hbmpDragImage);
            end;

            dwOkEffects := DROPEFFECT_COPY or DROPEFFECT_MOVE or DROPEFFECT_LINK;

            if not FDropped then
              DoDragDrop(DataObject, Self, dwOkEffects, dwEffectResult);
            MouseButtonState := [];
          end
        finally
          FDragging := False;
          Bitmap.Free;
          FDragItemIndex := -1;
          // Fire OnEndDrag for the TOLEListview control
          DoEndDrag(Self, Pt.X, Pt.Y);
          // Fire VETController.OnEndDrag
          VETController.DoEndDrag(Self, Pt.X, Pt.Y);
        end
      end
    end;
  end;
  inherited;
end;

procedure TOLEListview.WMRButtonDown(var Message: TWMRButtonDown);
begin
  Include(FMouseButtonState, RButtonDown);
  inherited;
end;

procedure TOLEListview.WMRButtonUp(var Message: TWMRButtonUp);
begin
  Exclude(FMouseButtonState, RButtonDown);
  inherited;
end;

procedure TOLEListview.WMTimer(var Message: TWMTimer);
begin
  inherited;
  case Message.TimerID of
    SCROLL_DELAY_TIMER:
      begin
        KillTimer(Handle, FScrollDelayTimer);
        FScrollDelayTimer := 0;
        FScrollTimer := SetTimer(Handle, SCROLL_TIMER, VETController.AutoScrollInterval, FAutoScrollTimerStub);
      end;
  end
end;

//WMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWM
{ TCustomVirtualExplorerListviewEx }

constructor TCustomVirtualExplorerListviewEx.Create(AOwner: TComponent);
begin
  inherited Create(AOwner);
  FImageLibrary := timNone;
  {$IFDEF USEGRAPHICEX} FImageLibrary := timGraphicEx; {$ELSE}
    {$IFDEF USEIMAGEEN} FImageLibrary := timImageEn; {$ELSE}
      {$IFDEF USEENVISION} FImageLibrary := timImageMagick; {$ELSE}
        {$IFDEF USEIMAGEMAGICK} FImageLibrary := timImageMagick; {$ENDIF}
      {$ENDIF}
    {$ENDIF}
  {$ENDIF}

  FInternalDataOffset := AllocateInternalDataArea(SizeOf(TThumbnailData));

  FThumbsOptions := TThumbsOptions.Create(Self);

  FListview := TOLEListview.Create(Self);
  FListview.VETController := Self;
  FListview.OnAdvancedCustomDrawItem := LVOnAdvancedCustomDrawItem;
  FListview.SmallImages := IDEVirtualSystemImageLists.SmallSysImages;

  FDummyIL := TImageList.Create(Self);

  FExtensionsList := TExtensionsList.Create;
  FShellExtractExtensionsList := TExtensionsList.Create;
  FExtensionsExclusionList := TExtensionsList.Create;
  FillExtensionsList;

  FVisible := true;
  FViewStyle := vsxReport;
  FAccumulatedChanging := false;
end;

destructor TCustomVirtualExplorerListviewEx.Destroy;
begin
  {$IFDEF THREADEDICONS}
  if ThreadedImagesEnabled then
    ImageThreadManager.ClearPendingItems(Self, WM_VTSETICONINDEX, Malloc);
  {$ENDIF}
  //The Listview is automatically freed.
  //FreeAndNil(FListview);
  FDummyIL.Free;
  FExtensionsList.Free;
  FShellExtractExtensionsList.Free;
  FExtensionsExclusionList.Free;
  FThumbsOptions.Free;
  FThumbsOptions := nil;


  if Assigned(FThumbThread) then
  begin
    FThumbThread.Priority := tpNormal; //D6 has a Thread bug, we must set the priority to tpNormal before destroying
    FThumbThread.ClearPendingItems(Self, WM_VLVEXTHUMBTHREAD, Malloc);
    FThumbThread.Terminate;
    FThumbThread.SetEvent;
    FThumbThread.WaitFor;
    FreeAndNil(FThumbThread);
  end;

  inherited;
end;

function TCustomVirtualExplorerListviewEx.GetAnimateWndParent: TWinControl;
begin
  if ViewStyle <> vsxReport then
    Result := ChildListview
  else
    Result := Self
end;

procedure TCustomVirtualExplorerListviewEx.CreateWnd;
begin
  inherited;
  SyncOptions;
end;

procedure TCustomVirtualExplorerListviewEx.Loaded;
begin
  inherited;
  SyncOptions;
end;

procedure TCustomVirtualExplorerListviewEx.Notification(
  AComponent: TComponent; Operation: TOperation);
begin
  inherited;
  if (Operation = opRemove) and (AComponent = FListView) then
    FListView := nil;
end;

procedure TCustomVirtualExplorerListviewEx.RequestAlign;
begin
  inherited;
  if IsValidChildListview then
    if (FListview.Align <> Align) or (FListview.Anchors <> Anchors) or
       (FListview.Constraints.MaxWidth <> Constraints.MaxWidth) or
       (FListview.Constraints.MaxHeight <> Constraints.MaxHeight) or
       (FListview.Constraints.MinWidth <> Constraints.MinWidth) or
       (FListview.Constraints.MinHeight <> Constraints.MinHeight) then
      SyncOptions;
end;

procedure TCustomVirtualExplorerListviewEx.SetParent(AParent: TWinControl);
begin
  inherited;
  //This is not a compound component, a compound component is a container
  //with 1 or more controls in it.
  //The parent of the child VCL Listview is the Self.Parent, this is so
  //to retain all the properties of TExplorerListview, that means I don't have
  //to copy all these properties and you don't loose usability.
  if Assigned(FListview) and (AParent <> nil) then
    FListview.Parent := AParent;
end;

procedure TCustomVirtualExplorerListviewEx.SetZOrder(TopMost: Boolean);
begin
  inherited;
  if (ViewStyle <> vsxReport) and Assigned(FListview) then
    FListview.SetZOrder(TopMost);
end;

function TCustomVirtualExplorerListviewEx.GetClientRect: TRect;
begin
  if ViewStyle = vsxReport then
    Result := inherited GetClientRect
  else
    Result := FListview.GetClientRect;  
end;

procedure TCustomVirtualExplorerListviewEx.CMShowHintChanged(var Message: TMessage);
begin
  inherited;
  if Assigned(FListview) then FListview.ShowHint := ShowHint;
end;

procedure TCustomVirtualExplorerListviewEx.CMBorderChanged(var Message: TMessage);
begin
  inherited;
  if Assigned(FListview) then begin
    FListview.BevelEdges := BevelEdges;
    FListview.BevelInner := BevelInner;
    FListview.BevelKind := BevelKind;
    FListview.BevelOuter := BevelOuter;
    FListview.BevelWidth := BevelWidth;
    FListview.BorderWidth := BorderWidth;
  end;
end;

procedure TCustomVirtualExplorerListviewEx.CMBidimodechanged(var Message: TMessage);
begin
  inherited;
  if Assigned(FListview) then FListview.BiDiMode := BiDiMode;
end;

procedure TCustomVirtualExplorerListviewEx.CMCtl3DChanged(var Message: TMessage);
begin
  inherited;
  if Assigned(FListview) then FListview.Ctl3D := Ctl3D;
end;

procedure TCustomVirtualExplorerListviewEx.CMColorChanged(var Message: TMessage);
begin
  inherited;
  if Assigned(FListview) then FListview.Color := Color;
end;

procedure TCustomVirtualExplorerListviewEx.CMCursorChanged(var Message: TMessage);
begin
  inherited;
  if Assigned(FListview) then FListview.Cursor := Cursor;
end;

procedure TCustomVirtualExplorerListviewEx.CMEnabledchanged(var Message: TMessage);
begin
  inherited;
  if Assigned(FListview) then FListview.Enabled := Enabled;
end;

procedure TCustomVirtualExplorerListviewEx.CMFontChanged(var Message: TMessage);
begin
  inherited;
  if Assigned(FListview) then FListview.Font.Assign(Font);
end;

procedure TCustomVirtualExplorerListviewEx.WMEnumThreadFinished(
  var Msg: TMessage);
begin
  inherited;
  if ViewStyle <> vsxReport then
  begin
    SyncItemsCount;
    SyncInvalidate;
    ChildListview.Cursor := crArrow
  end
end;

procedure TCustomVirtualExplorerListviewEx.WMEnumThreadStart(var Msg: TMessage);
begin
  inherited;
  if ViewStyle <> vsxReport then
  begin
    ChildListview.HandleNeeded;
    ChildListview.Cursor := crHourGlass
  end
end;

procedure TCustomVirtualExplorerListviewEx.WMNCDestroy(var Message: TWMNCDestroy);
begin
  if Assigned(FThumbThread) then
    FThumbThread.ClearPendingItems(Self, WM_VLVEXTHUMBTHREAD, Malloc);
  inherited;
end;

{$IFDEF THREADEDICONS}
procedure TCustomVirtualExplorerListviewEx.WMVTSetIconIndex(var Msg: TWMVTSetIconIndex);
var
  NS: TNamespace;
  R: TRect;
begin
  if FViewStyle = vsxReport then
    inherited
  else begin
    try
      if ValidateNamespace(Msg.IconInfo.UserData, NS) then
      begin
        NS.SetIconIndexByThread(Msg.IconInfo.IconIndex, True);
        R := FListview.GetLVItemRect(Msg.IconInfo.Tag, drIcon);
        InvalidateRect(FListview.Handle, @R, False);
      end
    finally
      ImageThreadManager.ReleaseItem(Msg.IconInfo, Malloc)
    end
  end;
end;
{$ENDIF}

procedure TCustomVirtualExplorerListviewEx.WMVLVExThumbThread(var Message: TMessage);
var
  TD: PThumbnailData;
  ThumbThreadData: PThumbnailThreadData;
  NS: TNamespace;
  Node: PVirtualNode;
  Info: PVirtualThreadIconInfo;
  B: TBitmap;
  CI: TThumbsCacheItem;
  I: integer;
begin
  Info := PVirtualThreadIconInfo(Message.wParam);

  if Assigned(Info) then
  begin
    try
      ThumbThreadData := PThumbnailThreadData(Info.UserData2);
      if Assigned(ThumbThreadData) then
      begin
        Node := Info.UserData;
        if (ThumbThreadData.State = tsValid) and ValidateNamespace(Node, NS) then
        begin
          if ValidateThumbnail(Node, TD) then
          begin
            TD.State := ThumbThreadData.State;

            if TD.Reloading then begin
              TD.Reloading := False;
              if ThumbsOptions.CacheOptions.FThumbsCache.Read(TD.CachePos, CI) then begin
                I := ThumbsOptions.CacheOptions.FThumbsCache.FScreenBuffer.IndexOf(NS.NameForParsing);
                if I > -1 then
                  ThumbsOptions.CacheOptions.FThumbsCache.FScreenBuffer.Delete(I);
                CI.CompressedThumbImageStream := ThumbThreadData.CompressedStream;
                CI.FThumbImageStream.LoadFromStream(ThumbThreadData.MemStream);
              end;
            end
            else
              if Assigned(OnThumbsCacheItemAdd) then begin
                B := TBitmap.Create;
                try
                  if ThumbThreadData.CompressedStream then
                    ConvertJPGStreamToBitmap(ThumbThreadData.MemStream, B)
                  else
                    B.LoadFromStream(ThumbThreadData.MemStream);
                  if DoThumbsCacheItemAdd(NS, B, ThumbThreadData.ImageWidth, ThumbThreadData.ImageHeight) then
                    TD.CachePos := ThumbsOptions.CacheOptions.FThumbsCache.Add(NS.NameForParsing,
                      '', '', NS.LastWriteDateTime, ThumbThreadData.ImageWidth, ThumbThreadData.ImageHeight,
                      ThumbThreadData.CompressedStream, ThumbThreadData.MemStream);
                finally
                  B.Free;
                end;
              end
              else
                TD.CachePos := ThumbsOptions.CacheOptions.FThumbsCache.Add(NS.NameForParsing,
                  '', '', NS.LastWriteDateTime, ThumbThreadData.ImageWidth, ThumbThreadData.ImageHeight,
                  ThumbThreadData.CompressedStream, ThumbThreadData.MemStream);
            //redraw the item
            FListview.UpdateItems(Info.Tag, Info.Tag);
            //Update inmediatly, from ListView_RedrawItems windows help
            FListview.Update;
          end;
        end;
      end;
    finally
      ThumbThread.ReleaseItem(Info, Malloc);
    end;
  end;
end;

procedure TCustomVirtualExplorerListviewEx.DoInitNode(Parent, Node: PVirtualNode;
  var InitStates: TVirtualNodeInitStates);
var
  Data: PThumbnailData;
begin
  if ValidateThumbnail(Node, Data) then begin
    Data.CachePos := -1;
    Data.Reloading := False;
    Data.State := tsEmpty;
  end;
  inherited;
end;

procedure TCustomVirtualExplorerListviewEx.DoFreeNode(Node: PVirtualNode);
begin
  FLastDeletedNode := Node;
  if Assigned(Node) and Assigned(FThumbThread) then
    FThumbThread.ClearPendingItem(Self, Node, WM_VLVEXTHUMBTHREAD, Malloc);
  inherited;
end;

function TCustomVirtualExplorerListviewEx.DoKeyAction(var CharCode: Word;
  var Shift: TShiftState): Boolean;
begin
  Result := inherited DoKeyAction(CharCode, Shift);
  if Result then
    Case CharCode of
      Ord('I'), Ord('i'):
        if (ssCtrl in Shift) then
          InvertSelection(False);
    end;
end;

procedure TCustomVirtualExplorerListviewEx.Clear;
begin
  // Clear the cache before we rebuild the tree, called by RebuildRootNamespace
  if Assigned(ThumbsOptions) then begin
    ThumbsOptions.CacheOptions.FThumbsCache.Clear;
    ThumbsOptions.CacheOptions.FThumbsCache.ThumbWidth := ThumbsOptions.Width;
    ThumbsOptions.CacheOptions.FThumbsCache.ThumbHeight := ThumbsOptions.Height;
  end;
  inherited
end;

procedure TCustomVirtualExplorerListviewEx.RebuildRootNamespace;
var
  NS: TNamespace;
  C: TThumbsCache;
begin
  //I was overriding DoRootRebuild to do this, but I need to do it here
  //because in TCustomVirtualExplorerTree.RebuildRootNamespace there's a call
  //to EndUpdate and this fires FListview.OwnerDataHint (via DoStructureChange, Accumulated event).
  if (RebuildRootNamespaceCount = 0) and not (csLoading in ComponentState)
    and Assigned(RootFolderNamespace) and Active then
  begin
    FListview.OwnerDataPause := True; //avoid generating data
    try
      inherited; // We're clear to rebuild the tree, it will call Clear
      FListview.Items.BeginUpdate;
      try
        ThumbsOptions.CacheOptions.BrowsingFolder := RootFolderNamespace.NameForParsing;
        NS := RootFolderNamespace;
        if Assigned(ThumbsOptions.CacheOptions) and ThumbsOptions.CacheOptions.AutoLoad and Assigned(NS) and NS.Folder and NS.FileSystem then
        begin
          ThumbsOptions.CacheOptions.Load;
          C := ThumbsOptions.CacheOptions.FThumbsCache;
          if ThumbsOptions.CacheOptions.AdjustIconSize and C.LoadedFromFile and ((ThumbsOptions.Width <> C.ThumbWidth) or (ThumbsOptions.Height <> C.ThumbHeight)) then begin
            ThumbsThreadPause := True;
            try
              ThumbsOptions.Width := C.ThumbWidth;
              ThumbsOptions.Height := C.ThumbHeight;
            finally
              ThumbsThreadPause := False;
            end;
          end;
        end;

        SyncItemsCount;
        FListview.Selected := nil;
      finally
        FListview.Items.EndUpdate;
      end;
    finally
      FListview.OwnerDataPause := False;
      FListview.UpdateArrangement;
      if ThumbsOptions.LoadAllAtOnce then
        FListview.FetchThumbs(0, FListview.Items.Count - 1);
      if FListview.items.count > 0 then
        FListview.ItemFocused := FListview.Items[0];
    end;
  end;
  FlushSearchCache;
end;

procedure TCustomVirtualExplorerListviewEx.DoRootChanging(
  const NewRoot: TRootFolder; Namespace: TNamespace; var Allow: Boolean);
var
  NS: TNamespace;
begin
  if (RebuildRootNamespaceCount = 0) and not (csLoading in ComponentState) and Active then
    FListview.OwnerDataPause := True; //avoid generating data
  FlushSearchCache;

  inherited DoRootChanging(NewRoot, Namespace, Allow);

  if not Allow and FListview.OwnerDataPause then
    FListview.OwnerDataPause := False;

  if Allow and not (csLoading in Componentstate) and Assigned(ThumbsOptions) and Assigned(ThumbsOptions.CacheOptions) then begin
    NS := RootFolderNamespace;
    if ThumbsOptions.CacheOptions.AutoSave and Assigned(NS) and NS.Folder and NS.FileSystem then
      if (ThumbsOptions.CacheOptions.StorageType = tcsCentral) or not NS.ReadOnly then
        ThumbsOptions.CacheOptions.Save;
  end;
end;

procedure TCustomVirtualExplorerListviewEx.DoStructureChange(Node: PVirtualNode; Reason: TChangeReason);
begin
  FlushSearchCache;
  inherited;
  if FViewStyle <> vsxReport then begin
    Case Reason of
      crChildAdded:
        begin
          SyncSelectedItems;
          SyncItemsCount;
        end;
      crChildDeleted:
        begin
          SyncSelectedItems;
          SyncItemsCount;
          // Don't catch this when the RootNode is deleted (happens when changing dir)
          if Node <> RootNode then begin
            // Focus the last item if required
            if (FListview.ItemFocused = nil) and (Flistview.Items.Count > 0) then
              FListview.ItemFocused := FListview.Items[Flistview.Items.Count - 1];
            if ViewStyle <> vsxReport then SyncInvalidate;
          end;
        end;
      crAccumulated:
        begin
          SyncSelectedItems;
          if FAccumulatedChanging then begin // Take a look at ReReadAndRefreshNode
            SyncItemsCount;
            // Focus the last item if required
            if (FListview.ItemFocused = nil) and (Flistview.Items.Count > 0) then
              FListview.ItemFocused := FListview.Items[Flistview.Items.Count - 1];
            if ViewStyle <> vsxReport then SyncInvalidate;
          end;
        end;
    end;
  end;
end;

procedure TCustomVirtualExplorerListviewEx.ReReadAndRefreshNode(Node: PVirtualNode; SortNode: Boolean);
begin
  // This method is called by WM_SHELLNOTIFY and it's responsible for updating
  // the nodes, looking if they were added or deleted.
  // The nodes are updated inside a BeginUpdate/EndUpdate block, when EndUpdate
  // is reached DoStructureChange (with crAccumulated) is called once.
  // We need a flag so DoStructureChange knows who's calling it.
  FAccumulatedChanging := true;
  inherited;
  FAccumulatedChanging := false;
end;

procedure TCustomVirtualExplorerListviewEx.DoBeforeCellPaint(Canvas: TCanvas;
  Node: PVirtualNode; Column: TColumnIndex; CellRect: TRect);
var
  NS: TNamespace;
  I: integer;
begin
  if ValidateNamespace(Node, NS) then
    if ThumbsOptions.Highlight = thMultipleColors then begin
      I := IsImageFileIndex(NS.NameForParsing);
      if I > -1 then
        if FExtensionsList.Colors[I] <> clNone then begin
          Canvas.Brush.Color := FExtensionsList.Colors[I];
          Canvas.FillRect(CellRect);
        end;
    end
    else
      if ThumbsOptions.Highlight = thSingleColor then
        if IsImageFile(NS.NameForParsing) then begin
          Canvas.Brush.Color := ThumbsOptions.HighlightColor;
          Canvas.FillRect(CellRect);
        end;
  inherited;
end;

function TCustomVirtualExplorerListviewEx.InternalData(Node: PVirtualNode): Pointer;
begin
  if Node = nil then
    Result := nil
  else
    Result := PByte(Node) + FInternalDataOffset;
end;

function TCustomVirtualExplorerListviewEx.IsAnyEditing: Boolean;
begin
  Result := (inherited IsAnyEditing) or (IsValidChildListview and FListview.IsEditing);
end;

function TCustomVirtualExplorerListviewEx.GetNodeByIndex(Index: Cardinal): PVirtualNode;
begin
  Result := GetChildByIndex(RootNode, Index, FSearchCache);
  if Result = FLastDeletedNode then
    Result := nil;
end;

procedure TCustomVirtualExplorerListviewEx.FlushSearchCache;
begin
  FSearchCache := nil;
  FLastDeletedNode := nil;
end;

function TCustomVirtualExplorerListviewEx.ValidateThumbnail(Node: PVirtualNode; var ThumbData: PThumbnailData): Boolean;
begin
  Result := False;
  ThumbData := nil;
  if Assigned(Node) then
  begin
    ThumbData := InternalData(Node);
    Result := Assigned(ThumbData);
  end
end;

function TCustomVirtualExplorerListviewEx.ValidateListItem(Node: PVirtualNode; var ListItem: TListItem): Boolean;
var
  C: Cardinal;
begin
  Result := False;
  ListItem := nil;
  if Assigned(Node) and Assigned(FListview) then
  begin
    if not (vsInitialized in Node.States) then
      InitNode(Node);
    C := FListview.Items.Count;
    if (C > 0) and (Node.Index < C) then
      ListItem := FListview.Items[Node.Index];
    Result := Assigned(ListItem);
  end
end;

function TCustomVirtualExplorerListviewEx.IsValidChildListview: boolean;
begin
  Result := Assigned(FListview) and FListview.HandleAllocated;
end;

procedure TCustomVirtualExplorerListviewEx.SetBounds(ALeft, ATop, AWidth,
  AHeight: Integer);
begin
  inherited;
  if IsValidChildListview then
    if not EqualRect(FListview.BoundsRect, Rect(ALeft, ATop, AWidth, AHeight)) then
      FListview.SetBounds(ALeft, ATop, AWidth, AHeight);
end;

function TCustomVirtualExplorerListviewEx.Focused: Boolean;
begin
  if (ViewStyle <> vsxReport) and IsValidChildListview then
    Result := FListview.Focused
  else
    Result := inherited Focused;
end;

procedure TCustomVirtualExplorerListviewEx.SetFocus;
begin
  if Parent.Visible then begin
    if ViewStyle = vsxReport then
      inherited
    else
      if IsValidChildListview then
        FListview.SetFocus;
  end;
end;

function TCustomVirtualExplorerListviewEx.BrowseToByPIDL(APIDL: PItemIDList;
  ExpandTarget, SelectTarget, SetFocusToVET, CollapseAllFirst: Boolean;
  ShowAllSiblings: Boolean = True): Boolean;
begin
  Result := inherited BrowseToByPIDL(APIDL, ExpandTarget, SelectTarget,
    SetFocusToVET, CollapseAllFirst, ShowAllSiblings);
  if FViewStyle <> vsxReport then SyncSelectedItems;
end;

procedure TCustomVirtualExplorerListviewEx.CopyToClipBoard;
begin
  if FViewStyle <> vsxReport then SyncSelectedItems(false);
  inherited;
end;

procedure TCustomVirtualExplorerListviewEx.CutToClipBoard;
begin
  if FViewStyle <> vsxReport then SyncSelectedItems(false);
  inherited;
end;

function TCustomVirtualExplorerListviewEx.PasteFromClipboard: Boolean;
begin
  if FViewStyle <> vsxReport then SyncSelectedItems(false);
  Result := inherited PasteFromClipboard;
end;

procedure TCustomVirtualExplorerListviewEx.SelectedFilesDelete;
begin
  if FViewStyle <> vsxReport then SyncSelectedItems(false);
  inherited;
end;

procedure TCustomVirtualExplorerListviewEx.SelectedFilesPaste(AllowMultipleTargets: Boolean);
begin
  if not (toVETReadOnly in TreeOptions.VETMiscOptions) then begin
    if FViewStyle <> vsxReport then SyncSelectedItems(false);
    inherited;
  end;
end;

procedure TCustomVirtualExplorerListviewEx.SelectedFilesShowProperties;
begin
  if FViewStyle <> vsxReport then SyncSelectedItems(false);
  inherited;
end;

function TCustomVirtualExplorerListviewEx.EditNode(Node: PVirtualNode;
  Column: TColumnIndex): Boolean;
begin
  Result := false;
  if FViewStyle = vsxReport then
    Result := inherited EditNode(Node, Column)
  else begin
    if Assigned(Node) and not (vsDisabled in Node.States) and
      not (toReadOnly in TreeOptions.MiscOptions) then
    begin
      if not Focused then
        SetFocus;
      FocusedNode := Node;
      if not (vsInitialized in Node.States) then
        InitNode(Node);
      SyncSelectedItems;
      Result := FListview.Items[Node.index].EditCaption;
    end;
  end;
end;

function TCustomVirtualExplorerListviewEx.EditFile(APath: WideString): Boolean;
// Selects a file to edit it.
// APath parameter can be a filename or a full pathname to a file.
// If APath is a filename it searches the file in the current directory.
// If APath is a full pathname it changes the current directory to the APath
// dir and searches the file.
var
  Node: PVirtualNode;
  D: WideString;
begin
  Result := False;
  if Pos(':', APath) = 0 then
    // It's a file, include the current directory
    APath := IncludeTrailingBackslashW(RootFolderNamespace.NameForParsing) + APath
  else begin
    // Browse to the directory if the root is incorrect
    D := ExtractFileDirW(APath);
    if not SpCompareText(RootFolderNamespace.NameForParsing, D) then
      BrowseTo(D);
  end;

  Node := FindNode(APath);
  if Assigned(Node) then begin
    ClearSelection;
    Selected[Node] := True;
    EditNode(Node, 0); // EditNode calls SyncSelected
    Result := True;
  end;
end;

function TCustomVirtualExplorerListviewEx.InvalidateNode(Node: PVirtualNode): TRect;
var
  L: TListItem;
  R: TRect;
begin
  Result := inherited InvalidateNode(Node);
  if Assigned(Node) and (Node.CheckType <> IDEVirtualTrees.ctNone) and not (csDesigning in ComponentState) and
    HandleAllocated and IsValidChildListview and (ViewStyle <> vsxReport) and ValidateListItem(Node, L) then
  begin
    R := L.DisplayRect(drIcon);
    InvalidateRect(FListview.Handle, @R, True);
  end;
end;

procedure TCustomVirtualExplorerListviewEx.SyncInvalidate;
begin
  if ViewStyle = vsxReport then
    Invalidate
  else
    if HandleAllocated and (ViewStyle <> vsxReport) and IsValidChildListview and not (csDesigning in ComponentState) then
      FListview.Invalidate;
end;

procedure TCustomVirtualExplorerListviewEx.SyncItemsCount;
begin
  if HandleAllocated then
    FListview.Items.Count := RootNode.ChildCount;
end;

procedure TCustomVirtualExplorerListviewEx.SyncOptions;
begin
  if Assigned(FListview) then begin
    FListview.SetBounds(left, top, width, height);
    FListview.Align := Align;
    FListview.Anchors := Anchors;
    FListview.Constraints.Assign(Constraints);
    FListview.ReadOnly := not (toEditable in TreeOptions.MiscOptions);
    FListview.MultiSelect := toMultiSelect in TreeOptions.SelectionOptions;
    FListview.PopupMenu := PopupMenu;
    FListview.BorderStyle := BorderStyle;
  end;
end;

procedure TCustomVirtualExplorerListviewEx.SyncSelectedItems(UpdateChildListview: Boolean = True);
var
  Node: PVirtualNode;
  LItem: TListItem;
begin
  if not FListview.HandleAllocated then
    Exit;
  //Sync focused and selected items
  if UpdateChildListview then begin
    //Clear the selection, but pause automatic selection sync of TSyncListView.CNNotify
    FListview.FSelectionPause := True; //pause the selection
    try
      //Long captions items doen't get refreshed, we can't use FListview.Selected := nil;
      LItem := FListview.Selected;
      while Assigned(LItem) do begin
        LItem.Selected := False;
        FListview.UpdateItems(LItem.index, LItem.Index);
        LItem := FListview.GetNextItem(LItem, sdAll, [isSelected]);
      end;
    finally
      FListview.FSelectionPause := False; //restore
    end;

    Node := GetFirstSelected;
    while Assigned(Node) do begin
      FListview.Items[Node.Index].Selected := True;
      Node := GetNextSelected(Node);
    end;
    if Assigned(FocusedNode) then begin
      LItem := FListview.ItemFocused;
      if Assigned(LItem) then
        FListview.UpdateItems(LItem.index, LItem.Index);
      FListview.Items[FocusedNode.Index].Focused := True;
      FListview.Items[FocusedNode.Index].MakeVisible(False);
    end;
  end
  else begin
    ClearSelection;
    LItem := FListview.Selected;
    while Assigned(LItem) do begin
      Node := PVirtualNode(LItem.Data);
      if Assigned(Node) then Selected[Node] := true;
      LItem := FListview.GetNextItem(LItem, sdAll, [isSelected]);
    end;
    if Assigned(FListview.ItemFocused) then
      FocusedNode := PVirtualNode(FListview.ItemFocused.Data);
  end;
end;

function TCustomVirtualExplorerListviewEx.DoThumbsCacheItemAdd(NS: TNamespace;
  Thumbnail: TBitmap; ImageWidth, ImageHeight: Integer): Boolean;
begin
  Result := True;
  if Assigned(OnThumbsCacheItemAdd) then
    FOnThumbsCacheItemAdd(Self, NS, Thumbnail, ImageWidth, ImageHeight, Result);
end;

function TCustomVirtualExplorerListviewEx.DoThumbsCacheItemLoad(NS: TNamespace;
  var CacheItem: TThumbsCacheItem): Boolean;
begin
  // Update the cache entry if the file was changed
  Result := NS.LastWriteDateTime = CacheItem.FileDateTime;
  if Assigned(OnThumbsCacheItemLoad) then
    FOnThumbsCacheItemLoad(Self, NS, CacheItem, Result);
  CacheItem.FileDateTime := NS.LastWriteDateTime;
end;

function TCustomVirtualExplorerListviewEx.DoThumbsCacheItemRead(Filename: WideString;
  Thumbnail: TBitmap): Boolean;
begin
  Result := True;
  if Assigned(OnThumbsCacheItemRead) then
    FOnThumbsCacheItemRead(Self, Filename, Thumbnail, Result);
end;

function TCustomVirtualExplorerListviewEx.DoThumbsCacheItemProcessing(NS: TNamespace;
  Thumbnail: TBitmap; var ImageWidth, ImageHeight: integer): Boolean;
begin
  Result := True;
  if Assigned(OnThumbsCacheItemProcessing) then
    FOnThumbsCacheItemProcessing(Self, NS, Thumbnail, ImageWidth, ImageHeight, Result);
end;

function TCustomVirtualExplorerListviewEx.DoThumbsCacheLoad(Sender: TThumbsCache;
  CacheFilePath: WideString; Comments: TWideStringList): Boolean;
begin
  Result := True;
  if Assigned(FOnThumbsCacheLoad) then FOnThumbsCacheLoad(Sender, CacheFilePath, Comments, Result);
end;

function TCustomVirtualExplorerListviewEx.DoThumbsCacheSave(Sender: TThumbsCache;
  CacheFilePath: WideString; Comments: TWideStringList): Boolean;
begin
  Result := True;
  if Assigned(FOnThumbsCacheSave) then FOnThumbsCacheSave(Sender, CacheFilePath, Comments, Result);
end;

function TCustomVirtualExplorerListviewEx.GetDetailsString(Node: PVirtualNode; ThumbFormatting: Boolean = True): WideString;
var
  NS: TNamespace;
  TD: PThumbnailData;
  CI: TThumbsCacheItem;
  S, K, D: WideString;
  I: integer;
begin
  //When ThumbFormatting is true it returns the minimum info available,
  //mantaining the line order.
  Result := '';
  if ValidateNamespace(Node, NS) then begin
    //Image size
    S := '';
    if ValidateThumbnail(Node, TD) and (TD.State = tsValid) and (TD.CachePos > -1) then begin
      if ThumbsOptions.CacheOptions.FThumbsCache.Read(TD.CachePos, CI) and (CI.ImageWidth > 0) and (CI.ImageHeight > 0) then
        S := Format('%dx%d', [CI.ImageWidth, CI.ImageHeight])
    end;
    //File size
    if not NS.Folder then
      K := NS.SizeOfFileKB
    else
      K := '';
    //Modified date
    D := NS.LastWriteTime;
    if D <> '' then begin  //Delete the seconds from the LastWriteTime
      for I := Length(D) downto 0 do
        if D[I] = ':' then break;
      if I > 0 then
        Delete(D, I, Length(D));
    end;

    if ThumbFormatting then begin
      Result := Format('%s' + #13 + '%s' + #13 + '%s', [S, K, D]);
      if Result = #13 + #13 then Result := '';
    end
    else begin
      if NS.NameInFolder <> '' then
        Result := Result + NS.NameInFolder;
      if S <> '' then
        Result := Result + #13 + S;
      if K <> '' then
        Result := Result + #13 + K;
      if D <> '' then
        Result := Result + #13 + D;
    end;
  end;
end;

function TCustomVirtualExplorerListviewEx.DoThumbsGetDetails(Node: PVirtualNode; HintDetails: Boolean): WideString;
begin
  Result := GetDetailsString(Node, not HintDetails);
  if Assigned(FOnThumbsGetDetails) then FOnThumbsGetDetails(Self, Node, HintDetails, Result);
end;

procedure TCustomVirtualExplorerListviewEx.DoThumbsDrawBefore(ACanvas: TCanvas;
  ListItem: TListItem; ThumbData: PThumbnailData;
  AImageRect, ADetailsRect: TRect; var DefaultDraw: Boolean);
begin
  if Assigned(OnThumbsDrawBefore) then
    FOnThumbsDrawBefore(Self, ACanvas, ListItem, ThumbData, AImageRect, ADetailsRect, DefaultDraw);
end;

procedure TCustomVirtualExplorerListviewEx.DoThumbsDrawAfter(ACanvas: TCanvas;
  ListItem: TListItem; ThumbData: PThumbnailData;
  AImageRect, ADetailsRect: TRect; var DefaultDraw: Boolean);
begin
  if Assigned(OnThumbsDrawAfter) then
    FOnThumbsDrawAfter(Self, ACanvas, ListItem, ThumbData, AImageRect, ADetailsRect, DefaultDraw);
end;

function TCustomVirtualExplorerListviewEx.DoThumbsDrawHint(HintBitmap: TBitmap; Node: PVirtualNode): Boolean;
begin
  Result := true;
  if Assigned(FOnThumbsDrawHint) then FOnThumbsDrawHint(Self, HintBitmap, Node, Result);
end;

procedure TCustomVirtualExplorerListviewEx.DrawThumbBG(ACanvas: TCanvas;
  Item: TListItem; ThumbData: PThumbnailData; R: TRect);
begin
  if (ThumbData.State = tsValid) or (ThumbsOptions.BorderOnFiles) then begin
    if not (ThumbsOptions.Border in [tbWindowed, tbFit]) and Item.Selected and (not FListview.IsEditing) then begin
      if FListview.Focused then ACanvas.Brush.Color := Colors.FocusedSelectionColor
      else ACanvas.Brush.Color := Colors.UnFocusedSelectionColor;
    end;
    ACanvas.Fillrect(R);
  end;
end;

procedure DrawFocusRect2(ACanvas: TCanvas; const R: TRect);
var
  DC: HDC;
  C1, C2: TColor;
begin
  DC := ACanvas.Handle;
  C1 := SetTextColor(DC, clBlack);
  C2 := SetBkColor(DC, clWhite);
  Windows.DrawFocusRect(DC, R);
  SetTextColor(DC, C1);
  SetBkColor(DC, C2);
end;

procedure TCustomVirtualExplorerListviewEx.DrawThumbFocus(ACanvas: TCanvas;
  Item: TListItem; ThumbData: PThumbnailData; R: TRect);
begin
  if (ThumbData.State = tsValid) or (ThumbsOptions.BorderOnFiles) then begin
    ACanvas.Brush.Style := bsSolid;
    case ThumbsOptions.Border of
      tbWindowed:
        DrawThumbBorder(ACanvas, ThumbsOptions.Border, R, Item.Selected);
      tbFit:
        if ThumbData.State = tsValid then
          DrawThumbBorder(ACanvas, ThumbsOptions.Border, R, Item.Selected);
    else
      if FListview.Focused and Item.Focused then
        DrawFocusRect2(ACanvas, R)
      else
        DrawThumbBorder(ACanvas, ThumbsOptions.Border, R, Item.Selected);
    end;
  end;
end;

procedure DrawImageListIcon(ACanvas: TCanvas; ImgL: TImageList; X, Y: integer; LV: TCustomVirtualExplorerListviewEx; Item: TListItem);
var
  ForeColor, Style: Cardinal;
begin
  if LV.ChildListview.IsItemGhosted(Item) then
    ForeColor := ColorToRGB(LV.ChildListview.Color)
  else
    if ((LV.ViewStyle <> vsxThumbs) or not LV.ThumbsOptions.BorderOnFiles) and LV.ChildListview.Focused and Item.Selected then
      ForeColor := ColorToRGB(clHighlight)
    else
      ForeColor := CLR_NONE;

  if ForeColor = CLR_NONE then Style := ILD_TRANSPARENT
  else Style := ILD_TRANSPARENT or ILD_BLEND;

  if Item.OverlayIndex > -1 then
    Style := Style or ILD_OVERLAYMASK and Cardinal(IndexToOverlayMask(Item.OverlayIndex + 1));

  ImageList_DrawEx(ImgL.Handle, Item.ImageIndex, ACanvas.Handle, X, Y, 0, 0, CLR_NONE, ForeColor, Style);
end;

procedure TCustomVirtualExplorerListviewEx.DrawIcon(ACanvas: TCanvas; Item: TListItem;
  ThumbData: PThumbnailData; RThumb, RDetails: TRect; var RDestination: TRect);
var
  X, Y: integer;
  IL: TImageList;
  CacheThumb: TBitmap;
  S: WideString;
  NS: TNamespace;
begin
  RDestination := Rect(0, 0, 0, 0);

  //Paint the Thumbs details
  if ThumbsOptions.Details and (ViewStyle = vsxThumbs) then begin
    S := DoThumbsGetDetails(PVirtualNode(Item.Data), false);
    if S <> '' then begin
      ACanvas.Brush.Style := bsClear;
      ACanvas.Font.Size := 8;
      if FListview.Focused and Item.selected then begin
        ACanvas.Font.Color := clHighlightText;
        ACanvas.Brush.Color := clHighlight;
      end;
      if Win32Platform = VER_PLATFORM_WIN32_WINDOWS then
        Windows.DrawTextA(ACanvas.Handle, PAnsiChar(AnsiString(S)), -1, RDetails, DT_CENTER)
      else
        Windows.DrawTextW(ACanvas.Handle, PWideChar(S), -1, RDetails, DT_CENTER);
    end;
  end;

  //Paint the Thumbs or Icons
  if (ViewStyle = vsxThumbs) and (ThumbData.State = tsValid) then begin
    CacheThumb := TBitmap.Create;
    try
      if (Assigned(OnThumbsCacheItemRead) and ValidateNamespace(PVirtualNode(Item.Data), NS) and not DoThumbsCacheItemRead(NS.NameForParsing, CacheThumb))
        or ((ThumbData.CachePos > -1) and ThumbsOptions.CacheOptions.FThumbsCache.Read(ThumbData.CachePos, CacheThumb)) then
      begin
        X := RThumb.left + (RThumb.right - RThumb.left - CacheThumb.Width) div 2;
        Y := RThumb.top + (RThumb.bottom - RThumb.top -  CacheThumb.Height) div 2;
        ACanvas.Draw(X, Y, CacheThumb);
        RDestination := Bounds(X, Y, CacheThumb.Width, CacheThumb.Height);
      end;
    finally
      CacheThumb.Free;
    end;
  end
  else begin
    Case ViewStyle of
      vsxThumbs:
        if ThumbsOptions.ShowXLIcons and (ThumbsOptions.Width >= 48) and (ThumbsOptions.Height >= 48) then
          IL := IDEVirtualSystemImageLists.ExtraLargeSysImages
        else
          IL := IDEVirtualSystemImageLists.LargeSysImages;
      vsxIcon:
        IL := IDEVirtualSystemImageLists.LargeSysImages
    else
      IL := IDEVirtualSystemImageLists.SmallSysImages;
    end;

    X := RThumb.left + (RThumb.right - RThumb.left - IL.Width) div 2;
    Y := RThumb.top + (RThumb.bottom - RThumb.top - IL.Height) div 2;

    DrawImageListIcon(ACanvas, IL, X, Y, Self, Item);
    RDestination := Bounds(X, Y, IL.Width, IL.Height);
  end;
end;

procedure TCustomVirtualExplorerListviewEx.LVOnAdvancedCustomDrawItem(Sender: TCustomListView;
  Item: TListItem; State: TCustomDrawState; Stage: TCustomDrawStage; var DefaultDraw: Boolean);
var
  ThumbDefaultDraw: boolean;
  RIcon, R, RThumb, RDetails, RThumbDestination: TRect;
  Node: PVirtualNode;
  ThumbData: PThumbnailData;
  PaintInfo: TVTPaintInfo;
  B: TBitmap;
  NS: TNamespace;
  I: Integer;
  WS: WideString;
begin
  if (ViewStyle = vsxReport) or (not Assigned(FListview)) or (FListview.OwnerDataPause) or
    (FListview.Items.Count = 0) or (Item = nil) or (Item.Index < 0) or (Item.Index >= FListview.Items.Count) then
      Exit;

  case Stage of
    cdPrePaint:
      if not Item.Selected then begin
        Node := PVirtualNode(Item.Data);
        if ValidateNamespace(Node, NS) then begin
          Case ThumbsOptions.Highlight of
            thSingleColor:
              if IsImageFile(NS.NameForParsing) then
                Sender.Canvas.Brush.Color := ThumbsOptions.HighlightColor;
            thMultipleColors:
              begin
                I := IsImageFileIndex(NS.NameForParsing);
                if I > -1 then
                  if FExtensionsList.Colors[I] <> clNone then
                    Sender.Canvas.Brush.Color := FExtensionsList.Colors[I];
              end;
          end;
          if not (toNoUseVETColorsProp in TreeOptions.VETFolderOptions) then begin
            if NS.Compressed then Sender.Canvas.Font.Color := VETColors.CompressedTextColor
            else
              if NS.Folder then Sender.Canvas.Font.Color := VETColors.FolderTextColor
              else Sender.Canvas.Font.Color := VETColors.FileTextColor;
          end;
          if Assigned(OnPaintText) then OnPaintText(Self, Sender.Canvas, Node, 0, ttNormal);
        end;
      end;
    cdPostPaint:
      begin
        //In a normal ViewStyle only draw when there's an overlay icon
        if ViewStyle in [vsxIcon, vsxSmallIcon, vsxList] then
          if not (FListview.IsItemGhosted(Item) or (Item.OverlayIndex > -1)) then
            Exit;
        //Do the drawing in a bitmap buffer
        B := TBitmap.Create;
        B.Canvas.Lock;
        try
          RThumbDestination := Rect(0, 0, 0, 0);
          RIcon := FListview.GetLVItemRect(Item.index, drIcon);
          R := Rect(0, 0, RIcon.Right - RIcon.Left, RIcon.Bottom - RIcon.Top);
          InitBitmap(B, R.Right, R.Bottom, Self.Color);

          if FListview.IsBackgroundValid then
            B.Canvas.CopyRect(R, Sender.Canvas, RIcon);

          B.Canvas.Brush.Color := Sender.Canvas.Brush.Color;
          if ViewStyle = vsxThumbs then begin
            Node := PVirtualNode(Item.Data);
            if ValidateThumbnail(Node, ThumbData) then begin
              DrawThumbBG(B.Canvas, Item, ThumbData, R);

              if ThumbsOptions.Details then begin
                RThumb :=   Rect(R.Left, R.Top, R.Right, R.Bottom - ThumbsOptions.DetailsHeight);
                RDetails := Rect(R.Left, R.Bottom - ThumbsOptions.DetailsHeight, R.Right, R.Bottom);
              end
              else begin
                RThumb := R;
                RDetails := Rect(0, 0, 0, 0);
              end;

              ThumbDefaultDraw := true;
              DoThumbsDrawBefore(B.Canvas, Item, ThumbData, RThumb, RDetails, ThumbDefaultDraw);
              //if ThumbDefaultDraw and (ThumbData.State <> tsProcessing) then //this will increase the rendering speed, but it looks odd
              if ThumbDefaultDraw then begin
                DrawIcon(B.Canvas, Item, ThumbData, RThumb, RDetails, RThumbDestination);
                // Draw the checkbox
                if (toCheckSupport in TreeOptions.MiscOptions) and (Node.CheckType <> IDEVirtualTrees.ctNone) then begin
                  PaintInfo.Node := Node;
                  PaintInfo.Canvas := B.Canvas;
                  PaintInfo.ImageInfo[iiCheck].Index := GetCheckImage(Node);
                  PaintInfo.ImageInfo[iiCheck].XPos := 0;
                  PaintInfo.ImageInfo[iiCheck].YPos := 0;
                  PaintCheckImage(PaintInfo);
                end;
              end;

              ThumbDefaultDraw := true;
              DoThumbsDrawAfter(B.Canvas, Item, ThumbData, RThumb, RDetails, ThumbDefaultDraw);
              if ThumbDefaultDraw then begin
                if ThumbsOptions.ShowSmallIcon and (ThumbData.State = tsValid) then
                  DrawImageListIcon(B.Canvas, IDEVirtualSystemImageLists.SmallSysImages, (RThumb.right - RThumb.left - IDEVirtualSystemImageLists.SmallSysImages.Width) - 2, 2, Self, Item);
                if ThumbsOptions.Border = tbFit then
                  DrawThumbFocus(B.Canvas, Item, ThumbData, RThumbDestination)
                else
                  DrawThumbFocus(B.Canvas, Item, ThumbData, R);
              end;
            end;
          end
          else begin
            //Listviews in virtual mode doesn't draw overlay images, from:
            //http://groups.google.com/groups?hl=en&selm=7gftob%24aq4%40forums.borland.com
            DrawIcon(B.Canvas, Item, nil, R, R, RThumbDestination);
          end;

          Sender.Canvas.Lock;
          try
            // The TListview Canvas is very delicate
            // We must set the font color for the default focus painting
            if ColorToRGB(Sender.Canvas.Brush.Color) = 0 then
              Sender.Canvas.Font.Color := clWhite
            else
              Sender.Canvas.Font.Color := clWindowText;
            Sender.Canvas.Draw(RIcon.Left, RIcon.Top, B);

            if not FListview.IsEditing and Item.Selected then begin
              R := FListview.GetLVItemRect(Item.index, drLabel);
              if FListview.Focused then begin
                Sender.Canvas.Brush.Color := Colors.FocusedSelectionColor;
                Sender.Canvas.Font.Color := clHighlightText;
              end
              else
                Sender.Canvas.Brush.Color := Colors.UnFocusedSelectionColor;
              Sender.Canvas.FillRect(R);

              WS := FListview.GetItemCaption(Item.Index);
              if Win32Platform = VER_PLATFORM_WIN32_WINDOWS then
                Windows.DrawTextA(Sender.Canvas.Handle, PAnsiChar(AnsiString(WS)), -1, R, DT_CENTER)
              else
                Windows.DrawTextW(Sender.Canvas.Handle, PWideChar(WS), -1, R, DT_CENTER or DT_WORDBREAK or DT_EDITCONTROL);

              if FListview.Focused and Item.Focused then
                DrawFocusRect2(Sender.Canvas, R);
            end;
          finally
            Sender.Canvas.UnLock;
          end;
        finally
          B.Canvas.UnLock;
          B.Free;
        end;
      end;
  end;
end;

function TCustomVirtualExplorerListviewEx.GetThumbDrawingBounds(IncludeThumbDetails, IncludeBorderSize: Boolean): TRect;
begin
  // Obtains the REAL Thumbnail drawing Bounds Rect
  // It's like Item.displayrect(dricon)
  if IncludeBorderSize then
    Result := Rect(0, 0, ThumbsOptions.Width + ThumbsOptions.BorderSize * 2, ThumbsOptions.Height + ThumbsOptions.BorderSize * 2)
  else
    Result := Rect(0, 0, ThumbsOptions.Width, ThumbsOptions.Height);
  if IncludeThumbDetails and ThumbsOptions.Details then
    Result.Bottom := Result.Bottom + ThumbsOptions.DetailsHeight;
end;

function TCustomVirtualExplorerListviewEx.IsImageFileIndex(FileName: WideString): integer;
var
  Ext: WideString;
begin
  Result := -1;
  Ext := ExtractFileExtW(FileName);
  if ExtensionsExclusionList.IndexOf(Ext) < 0 then
  begin
    Result := ExtensionsList.IndexOf(Ext);
    if Result < 0 then
      Result := ShellExtractExtensionsList.IndexOf(Ext);
  end;
end;

function TCustomVirtualExplorerListviewEx.IsImageFile(FileName: WideString): Boolean;
begin
  Result := IsImageFileIndex(FileName) > -1;
end;

function TCustomVirtualExplorerListviewEx.IsImageFile(Node: PVirtualNode): TNamespace;
begin
  // Returns the namespace if it's an Image file
  Result := nil;
  if ValidateNamespace(Node, Result) and Result.FileSystem and not Result.Folder then begin
    if not IsImageFile(Result.NameForParsing) then Result := nil;
  end
  else
    Result := nil;
end;

procedure TCustomVirtualExplorerListviewEx.ResetThumbSpacing;
var
  R: TRect;
  W, H: integer;
begin
  //The cx and cy parameters of ListView_SetIconSpacing are relative to the
  //upper-left corner of an icon.
  //Therefore, to set spacing between icons that do not overlap, the cx or cy
  //values must include the size of the icon + the amount of empty space
  //desired between icons. Values that do not include the width of the icon
  //will result in overlaps.
  //When defining the icon spacing, cx and cy must set to 4 or larger.
  //Smaller values will not yield the desired layout.
  //To reset cx and cy to the default spacing, set the lParam value to -1.
  //i.e  SendMessage(FListview.Handle, LVM_SETICONSPACING, 0, -1);
  if ViewStyle = vsxThumbs then begin
    R := GetThumbDrawingBounds(true, true);
    W := R.Right + ThumbsOptions.SpaceWidth;
    H := R.Bottom + ThumbsOptions.SpaceHeight;
    ListView_SetIconSpacing(FListview.Handle, W, H);
    FListview.UpdateArrangement;
    if Assigned(FListview.ItemFocused) then
      FListview.ItemFocused.MakeVisible(false);
  end;
end;

procedure TCustomVirtualExplorerListviewEx.ResetThumbThread;
var
  N: PVirtualNode;
  TD: PThumbnailData;
begin
  if FThumbsThreadPause then Exit;
  //Reset the thread properties and reload
  ThumbThread.QueryList.LockList;
  try
    ThumbThread.ClearPendingItems(Self, WM_VLVEXTHUMBTHREAD, Malloc);
    ThumbThread.ResetThumbOptions;
    ThumbsOptions.CacheOptions.FThumbsCache.Clear;
    ThumbsOptions.CacheOptions.FThumbsCache.ThumbWidth := ThumbsOptions.Width;
    ThumbsOptions.CacheOptions.FThumbsCache.ThumbHeight := ThumbsOptions.Height;

    //Iterate through the nodes and reset the thumb state of valid items
    N := RootNode.FirstChild;
    while Assigned(N) do begin
      if (vsInitialized in N.States) and ValidateThumbnail(N, TD) then
        if IsThumbnailActive(TD.State) then begin
          TD.CachePos := -1;
          TD.Reloading := False;
          TD.State := tsEmpty;
        end;
      N := N.NextSibling;
    end;

  finally
    ThumbThread.QueryList.UnlockList;
  end;

  if ViewStyle = vsxThumbs then
    SyncInvalidate;
end;

procedure TCustomVirtualExplorerListviewEx.SetViewStyle(const Value: TViewStyleEx);
var
  PrevFocused: boolean;
begin
  if FViewStyle <> Value then begin
    SyncOptions;
    FListview.Items.BeginUpdate;
    try
      if not (csDesigning in ComponentState) and FVisible and
      ((FViewStyle = vsxReport) and (Value <> vsxReport)) or
      ((FViewStyle <> vsxReport) and (Value = vsxReport)) then begin
        PrevFocused := Focused;
        Parent.DisableAlign;
        try
          FListview.Visible := Value <> vsxReport;
          inherited Visible := not FListview.Visible;
          if FListview.Visible then //force hiding
            SetWindowPos(Self.Handle, 0, 0, 0, 0, 0, SWP_NOSIZE + SWP_NOMOVE + SWP_NOZORDER + SWP_NOACTIVATE + SWP_HIDEWINDOW);
        finally
          Parent.EnableAlign;
        end;
        //Sync focused and selected items
        if (FViewStyle = vsxReport) and (Value <> vsxReport) then begin
          SyncItemsCount;
          SyncSelectedItems(true);
        end
        else
          if (FViewStyle <> vsxReport) and (Value = vsxReport) then
            SyncSelectedItems(false);
        //Sync the focus
        if PrevFocused then begin
          FViewStyle := Value;
          SetFocus;
        end;
      end;

      FViewStyle := Value;

      //Set the child Listview.ViewStyle, this might look simple but the
      //icon spacing is incorrect if you don't force it.
      //To force correct spacing you should:
      // - Make sure the Listview is visible
      // - Set the correct icon spacing
      //I haven't found a better way, if you do just let me know.

      FListview.LargeImages := nil;

      Case Value of
        vsxThumbs: begin
          ResetThumbImageList(false);
          FListview.LargeImages := FDummyIL;
          FListview.ViewStyle := vsIcon;
          ResetThumbSpacing; //reset the spacing after vsIcon is setted
        end;
        vsxIcon: begin
          FListview.LargeImages := IDEVirtualSystemImageLists.LargeSysImages;
          FListview.ViewStyle := vsIcon;
          ListView_SetIconSpacing(FListview.Handle, GetSystemMetrics(SM_CXICONSPACING),
            GetSystemMetrics(SM_CYICONSPACING));
        end;
        vsxList:
          FListview.ViewStyle := vsList;
        vsxSmallIcon:
          FListview.ViewStyle := vsSmallIcon;
      end;

      FListview.UpdateSingleLineMaxChars;
    finally
      FListview.Items.EndUpdate;
      if Value <> vsxThumbs then
        FListview.UpdateArrangement;
    end;
  end;
end;

procedure TCustomVirtualExplorerListviewEx.SetVisible(const Value: boolean);
begin
  if FVisible <> Value then begin
    FListview.Visible := Value and (FViewStyle <> vsxReport);
    inherited Visible := Value and (not FListview.Visible);
    FVisible := Value;
  end;
end;

function TCustomVirtualExplorerListviewEx.GetThumbThread: TThumbThread;
begin
  if not Assigned(FThumbThread) then
    FThumbThread := GetThumbThreadClass.Create(Self);
  Result := FThumbThread;
end;

function TCustomVirtualExplorerListviewEx.GetThumbThreadClass: TThumbThreadClass;
begin
  Result := nil;
  if Assigned(OnThumbThreadClass) then FOnThumbThreadClass(Self, Result);
  if not Assigned(Result) then
    Result := TThumbThread;
end;

procedure TCustomVirtualExplorerListviewEx.SetThumbThreadClassEvent(const Value: TThumbThreadClassEvent);
begin
  if Assigned(FThumbThread) then
  begin
    FThumbThread.Priority := tpNormal; //D6 has a Thread bug, we must set the priority to tpNormal before destroying
    ResetThumbThread;
    FThumbThread.Terminate;
    FThumbThread.SetEvent;
    FThumbThread.WaitFor;
    FreeAndNil(FThumbThread);
  end;
  FOnThumbThreadClass := Value;
end;

procedure TCustomVirtualExplorerListviewEx.ResetThumbImageList(ResetSpacing: boolean = True);
begin
  if ViewStyle = vsxThumbs then begin
    FDummyIL.Width := ThumbsOptions.Width - 16 + (ThumbsOptions.BorderSize * 2);
    if ThumbsOptions.Details then
      FDummyIL.Height := ThumbsOptions.Height - 4 + (ThumbsOptions.BorderSize * 2) + ThumbsOptions.DetailsHeight
    else
      FDummyIL.Height := ThumbsOptions.Height - 4 + (ThumbsOptions.BorderSize * 2);
    if ResetSpacing then
      ResetThumbSpacing;
  end;
end;

procedure TCustomVirtualExplorerListviewEx.FillExtensionsList(FillColors: Boolean = true);
var
  I: integer;
  Ext: WideString;
{$IFDEF USEGRAPHICEX}
  L: TStringList;
{$ELSE}
  {$IFDEF USEIMAGEMAGICK}
     L: TStringList;
  {$ENDIF}
{$ENDIF}
begin
  FExtensionsList.Clear;
  {$IFDEF USEGRAPHICEX}
  L := TStringList.Create;
  try
    FileFormatList.GetExtensionList(L);
    FExtensionsList.AddStrings(L);
    FExtensionsList.DeleteString('ico'); // Don't add ico
  finally
    L.Free;
  end;
  {$ELSE}
  {$IFDEF USEIMAGEMAGICK}
  L := TStringList.Create;
  try
    if Assigned(MagickFileFormatList) then begin
      MagickFileFormatList.GetExtensionList(L);
      FExtensionsList.AddStrings(L);
      FExtensionsList.DeleteString('ico'); // Don't add ico

      FExtensionsList.DeleteString('pdf'); // TODO -cImageMagick : Stack overflow exception in TMagickImage.LoadFromStream, MagickImage.pas line 744
      FExtensionsList.DeleteString('txt'); // TODO -cImageMagick : 'Stream size must be defined' exception in TMagickImage.LoadFromStream (ASize <= 0), line 728
      FExtensionsList.DeleteString('avi'); // TODO -cImageMagick : infinite loop in BlobToImage: Trace: TMagickImage.LoadFromStream -> StreamToImage -> BlobToImage.
      FExtensionsList.DeleteString('mpg'); // TODO -cImageMagick : 'Stream size must be defined' exception in TMagickImage.LoadFromStream (ASize <= 0), line 728
      FExtensionsList.DeleteString('mpeg'); // TODO -cImageMagick : 'Stream size must be defined' exception in TMagickImage.LoadFromStream (ASize <= 0), line 728
      FExtensionsList.DeleteString('htm'); // TODO -cImageMagick : delegate not supported
      FExtensionsList.DeleteString('html'); // TODO -cImageMagick : delegate not supported
    end;
  finally
    L.Free;
  end;
  {$ELSE}
  with FExtensionsList do begin
    CommaText := '.jpg, .jpeg, .jif, .bmp, .emf, .wmf';
    {$IFDEF USEIMAGEEN}
    CommaText := CommaText + ', .png, .pcx, .tif, .tiff, .gif';
    {$ELSE}
    {$IFDEF USEENVISION}
      //version 1.1
      CommaText := CommaText + ', .png, .pcx, .pcc, .tif, .tiff, .dcx, .tga, .vst, .afi';
      //version 2.0, eps (Encapsulated Postscript) and jp2 (JPEG2000 version)
      //CommaText := CommaText + ', .eps, .jp2'; <<<<<<< still in beta
    {$ENDIF}
    {$ENDIF}    
  end;
  {$ENDIF}
  {$ENDIF}

  if FillColors then begin
    for I := 0 to FExtensionsList.Count - 1 do begin
      Ext := FExtensionsList[I];
      if (Ext = '.jpg') or (Ext = '.jpeg') or (Ext = '.jif') or (Ext = '.jfif') or (Ext = '.jpe') then FExtensionsList.Colors[I] := $BADDDD
      else if (Ext = '.bmp') or (Ext = '.rle') or (Ext = '.dib') then FExtensionsList.Colors[I] := $EFD3D3
      else if (Ext = '.emf') or (Ext = '.wmf') then FExtensionsList.Colors[I] := $7DC7B0
      else if (Ext = '.gif') then FExtensionsList.Colors[I] := $CCDBCC
      else if (Ext = '.png') then FExtensionsList.Colors[I] := $DAB6DA
      else if (Ext = '.tif') or (Ext = '.tiff') or (Ext = '.fax') or (Ext = '.eps') then FExtensionsList.Colors[I] := $DBB5B0
      else if (Ext = '.pcx') or (Ext = '.dcx') or (Ext = '.pcc') or (Ext = '.scr') then FExtensionsList.Colors[I] := $D6D6DB
      else if (Ext = '.tga') or (Ext = '.vst') or (Ext = '.vda') or (Ext = '.win') or (Ext = '.icb') or (Ext = '.afi') then FExtensionsList.Colors[I] := $EFEFD6
      else if (Ext = '.psd') or (Ext = '.pdd') then FExtensionsList.Colors[I] := $D3EFEF
      else if (Ext = '.psp') then FExtensionsList.Colors[I] := $93C0DD
      else if (Ext = '.sgi') or (Ext = '.rgba') or (Ext = '.rgb') or (Ext = '.bw') then FExtensionsList.Colors[I] := $C2BBE3
      else if (Ext = '.rla') or (Ext = '.rpf') then FExtensionsList.Colors[I] := $D3EFEF
      else if (Ext = '.ppm') or (Ext = '.pgm') or (Ext = '.pbm') then FExtensionsList.Colors[I] := $95D4DD
      else if (Ext = '.cel') or (Ext = '.pic') then FExtensionsList.Colors[I] := $AFEFEE
      else if (Ext = '.cut') then FExtensionsList.Colors[I] := $AFEFEE
      else if (Ext = '.pcd') then FExtensionsList.Colors[I] := $AFEFEE;
      // $7DC7B0 = green, $FFBD0B = orange, CFCFCF = grey
    end;
  end;

  ExtensionsExclusionList.CommaText := '.url, .lnk, .ico';
end;

//WMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWM
{ TThumbsCacheItem }

constructor TThumbsCacheItem.Create(AFilename: WideString);
begin
  FThumbImageStream := TMemoryStream.Create;
  FStreamSignature := DefaultStreamSignature;
  FFilename := AFilename;
end;

constructor TThumbsCacheItem.CreateFromStream(ST: TStream);
begin
  FThumbImageStream := TMemoryStream.Create;
  FStreamSignature := DefaultStreamSignature;
  LoadFromStream(ST);
end;

destructor TThumbsCacheItem.Destroy;
begin
  FThumbImageStream.Free;
  inherited;
end;

procedure TThumbsCacheItem.Fill(AFileDateTime: TDateTime; AExif, AComment: WideString;
  AImageWidth, AImageHeight: Integer; ACompressed: Boolean;
  AThumbImageStream: TMemoryStream);
begin
  FFileDateTime := AFileDateTime;
  FExif := AExif;
  FComment := AComment;
  FImageWidth := AImageWidth;
  FImageHeight := AImageHeight;
  FCompressed := ACompressed;
  if Assigned(AThumbImageStream) then
    FThumbImageStream.LoadFromStream(AThumbImageStream);

  Changed;
end;

function TThumbsCacheItem.DefaultStreamSignature: WideString;
begin
  // Override this method to change the default stream signature
  // Use the StreamSignature to load or not the custom properties
  // in LoadFromStream.
  Result := '1.4';
end;

procedure TThumbsCacheItem.Assign(CI: TThumbsCacheItem);
begin
  // Override this method to Assign the the custom properties.
  Fill(CI.FileDateTime, CI.Exif, CI.Comment, CI.ImageWidth,
    CI.ImageHeight, CI.CompressedThumbImageStream, CI.ThumbImageStream);
end;

procedure TThumbsCacheItem.Changed;
begin
  // Override this method to set the custom properties.
  // At this point all the properties are filled and valid.
  // The protected FFilename variable is also valid.
end;

function TThumbsCacheItem.LoadFromStream(ST: TStream): Boolean;
begin
  // Override this method to read the properties from the stream
  // Use the StreamSignature to load or not the custom properties
  Result := True;
  FStreamSignature := ReadWideStringFromStream(ST);
  FFilename := ReadWideStringFromStream(ST);
  FFileDateTime := ReadDateTimeFromStream(ST);
  FImageWidth := ReadIntegerFromStream(ST);
  FImageHeight := ReadIntegerFromStream(ST);
  FExif := ReadWideStringFromStream(ST);
  FComment := ReadWideStringFromStream(ST);
  FCompressed := Boolean(ReadIntegerFromStream(ST));
  ReadMemoryStreamFromStream(ST, FThumbImageStream);
end;

procedure TThumbsCacheItem.SaveToStream(ST: TStream);
begin
  // Override this method to write the properties to the stream
  WriteWideStringToStream(ST, FStreamSignature);
  WriteWideStringToStream(ST, FFilename);
  WriteDateTimeToStream(ST, FFileDateTime);
  WriteIntegerToStream(ST, FImageWidth);
  WriteIntegerToStream(ST, FImageHeight);
  WriteWideStringToStream(ST, FExif);
  WriteWideStringToStream(ST, FComment);
  WriteIntegerToStream(ST, Integer(FCompressed));
  WriteMemoryStreamToStream(ST, FThumbImageStream);
end;

function TThumbsCacheItem.ReadBitmap(OutBitmap: TBitmap): Boolean;
begin
  Result := False;
  if Assigned(FThumbImageStream) then
    if FCompressed then
      ConvertJPGStreamToBitmap(FThumbImageStream, OutBitmap) // JPEG compressed, convert to Bitmap
    else begin
      OutBitmap.LoadFromStream(FThumbImageStream);
      FThumbImageStream.Position := 0;
    end;
end;

procedure TThumbsCacheItem.WriteBitmap(ABitmap: TBitmap; CompressIt: Boolean);
begin
  FCompressed := CompressIt;
  ABitmap.SaveToStream(FThumbImageStream);
  if FCompressed then
    ConvertBitmapStreamToJPGStream(FThumbImageStream, 60); //JPEG compressed
end;

end.
