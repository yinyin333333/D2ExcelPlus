unit uTableFrame;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, System.Math,
  Vcl.Graphics, Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.Grids, Vcl.ComCtrls,
  Vcl.ExtCtrls, Vcl.StdCtrls, Vcl.Samples.Spin, Vcl.Menus, Clipbrd, Generics.Collections,
  Vcl.Buttons;

type
  TEditorState = class
  strict private
    FArea: TGridRect;
    FValues: TStrings;
    FRowCount: Integer;
    FColCount: Integer;
  public
    constructor Create;
    destructor Destroy; override;
    function Clone: TEditorState;
    property Area: TGridRect read FArea write FArea;
    property Values: TStrings read FValues;
    property RowCount: Integer read FRowCount write FRowCount;
    property ColCount: Integer read FColCount write FColCount;
  end;

  TGridSnapshot = class
  strict private
    FCells: TStringList;
    FRowCount: Integer;
    FColCount: Integer;
    FFixedCols: Integer;
    FFixedRows: Integer;
    FSelection: TGridRect;
    FLeftCol: Integer;
    FTopRow: Integer;
    FBaseColWidths: TArray<Integer>;
    FHiddenRows: TArray<Boolean>;
    FHiddenCols: TArray<Boolean>;
  public
    constructor Create;
    destructor Destroy; override;
    function GetCell(ACol, ARow: Integer): string;
    procedure SetCell(ACol, ARow: Integer; const AValue: string);
    property Cells[ACol, ARow: Integer]: string read GetCell write SetCell; default;
    property CellData: TStringList read FCells;
    property RowCount: Integer read FRowCount write FRowCount;
    property ColCount: Integer read FColCount write FColCount;
    property FixedCols: Integer read FFixedCols write FFixedCols;
    property FixedRows: Integer read FFixedRows write FFixedRows;
    property Selection: TGridRect read FSelection write FSelection;
    property LeftCol: Integer read FLeftCol write FLeftCol;
    property TopRow: Integer read FTopRow write FTopRow;
    property BaseColWidths: TArray<Integer> read FBaseColWidths write FBaseColWidths;
    property HiddenRows: TArray<Boolean> read FHiddenRows write FHiddenRows;
    property HiddenCols: TArray<Boolean> read FHiddenCols write FHiddenCols;
  end;

  THistoryCellChange = class
  strict private
    FRow: Integer;
    FCol: Integer;
    FRowCaption: string;
    FColumnName: string;
    FOldValue: string;
    FNewValue: string;
    FDisplayText: string;
  public
    property Row: Integer read FRow write FRow;
    property Col: Integer read FCol write FCol;
    property RowCaption: string read FRowCaption write FRowCaption;
    property ColumnName: string read FColumnName write FColumnName;
    property OldValue: string read FOldValue write FOldValue;
    property NewValue: string read FNewValue write FNewValue;
    property DisplayText: string read FDisplayText write FDisplayText;
  end;

  THistoryEntry = class
  strict private
    FEntryID: Int64;
    FFileName: string;
    FActionText: string;
    FTimeStamp: TDateTime;
    FBeforeState: TEditorState;
    FAfterState: TEditorState;
    FAfterSnapshot: TGridSnapshot;
    FChanges: TObjectList<THistoryCellChange>;
    FRowCountBefore: Integer;
    FRowCountAfter: Integer;
    FColCountBefore: Integer;
    FColCountAfter: Integer;
  public
    constructor Create;
    destructor Destroy; override;
    property EntryID: Int64 read FEntryID write FEntryID;
    property FileName: string read FFileName write FFileName;
    property ActionText: string read FActionText write FActionText;
    property TimeStamp: TDateTime read FTimeStamp write FTimeStamp;
    property BeforeState: TEditorState read FBeforeState write FBeforeState;
    property AfterState: TEditorState read FAfterState write FAfterState;
    property AfterSnapshot: TGridSnapshot read FAfterSnapshot write FAfterSnapshot;
    property Changes: TObjectList<THistoryCellChange> read FChanges;
    property RowCountBefore: Integer read FRowCountBefore write FRowCountBefore;
    property RowCountAfter: Integer read FRowCountAfter write FRowCountAfter;
    property ColCountBefore: Integer read FColCountBefore write FColCountBefore;
    property ColCountAfter: Integer read FColCountAfter write FColCountAfter;
  end;

  THistoryNodeKind = (hnkEntry, hnkRow, hnkCell);

  THistoryNodeRef = class
  strict private
    FKind: THistoryNodeKind;
    FEntryIndex: Integer;
    FRow: Integer;
    FCol: Integer;
  public
    property Kind: THistoryNodeKind read FKind write FKind;
    property EntryIndex: Integer read FEntryIndex write FEntryIndex;
    property Row: Integer read FRow write FRow;
    property Col: Integer read FCol write FCol;
  end;


  TDisplayedHistoryBlock = class
  strict private
    FKey: string;
    FHeader: string;
    FLines: TStringList;
    FLastEntryIndex: Integer;
    FTargetRow: Integer;
    FTargetCol: Integer;
  public
    constructor Create;
    destructor Destroy; override;
    property Key: string read FKey write FKey;
    property Header: string read FHeader write FHeader;
    property Lines: TStringList read FLines;
    property LastEntryIndex: Integer read FLastEntryIndex write FLastEntryIndex;
    property TargetRow: Integer read FTargetRow write FTargetRow;
    property TargetCol: Integer read FTargetCol write FTargetCol;
  end;

  TTableFrame = class(TFrame)
    {$region 'Components'}
    sgTable: TStringGrid;
    Panel1: TPanel;
    cbFixColumns: TCheckBox;
    seFixedColumns: TSpinEdit;
    cbFixRows: TCheckBox;
    sbToggleHistory: TSpeedButton;
    spHistory: TSplitter;
    plHistory: TPanel;
    pnlHistoryHeader: TPanel;
    lblHistoryTitle: TLabel;
    tvHistory: TListBox;
    pmGrid: TPopupMenu;
    miColumnOps: TMenuItem;
    miRowOps: TMenuItem;
    miResizeToFit: TMenuItem;
    miResizeToFitThisColumn: TMenuItem;
    miUnhideAll: TMenuItem;
    miFill: TMenuItem;
    miMath: TMenuItem;
    miColumnAdd: TMenuItem;
    miColumnInsert: TMenuItem;
    miColumnHide: TMenuItem;
    miColumnDelete: TMenuItem;
    miRowAdd: TMenuItem;
    miRowInsert: TMenuItem;
    miRowHide: TMenuItem;
    miRowDelete: TMenuItem;
    miRowClone: TMenuItem;
    miFillCells: TMenuItem;
    miFillIncrement: TMenuItem;
    miMathMultiply: TMenuItem;
    miMathDivide: TMenuItem;
    miMathAdd: TMenuItem;
    miMathSubtract: TMenuItem;
    {$endregion}
    {$region 'Eventhandler'}
    procedure sgTableMouseWheelDown(Sender: TObject; Shift: TShiftState;
      MousePos: TPoint; var Handled: Boolean);
    procedure sgTableMouseWheelUp(Sender: TObject; Shift: TShiftState;
      MousePos: TPoint; var Handled: Boolean);
    procedure sgTableKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure sgTableGetEditText(Sender: TObject; ACol, ARow: Integer;
      var Value: string);
    procedure cbFixColumnsClick(Sender: TObject);
    procedure cbFixRowsClick(Sender: TObject);
    procedure miGridAddNewClick(Sender: TObject);
    procedure miGridAddCopyClick(Sender: TObject);
    procedure miGridDeleteRowClick(Sender: TObject);
    procedure sgTableContextPopup(Sender: TObject; MousePos: TPoint;
      var Handled: Boolean);
    procedure miGridInsertNewClick(Sender: TObject);
    procedure miGridInsertCopyClick(Sender: TObject);
    procedure sgTableSelectCell(Sender: TObject; ACol, ARow: Integer;
      var CanSelect: Boolean);
    procedure sgTableDrawCell(Sender: TObject; ACol, ARow: Integer; Rect: TRect;
      State: TGridDrawState);
    procedure sgTableMouseDown(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure sgTableMouseMove(Sender: TObject; Shift: TShiftState; X,
      Y: Integer);
    procedure sgTableFixedCellClick(Sender: TObject; ACol, ARow: Integer);
    procedure miRowHideClick(Sender: TObject);
    procedure miColumnAddClick(Sender: TObject);
    procedure miColumnInsertClick(Sender: TObject);
    procedure miColumnHideClick(Sender: TObject);
    procedure miColumnDeleteClick(Sender: TObject);
    procedure miResizeToFitClick(Sender: TObject);
    procedure miResizeToFitThisColumnClick(Sender: TObject);
    procedure miUnhideAllClick(Sender: TObject);
    procedure miFillCellsClick(Sender: TObject);
    procedure miFillIncrementClick(Sender: TObject);
    procedure miMathMultiplyClick(Sender: TObject);
    procedure miMathDivideClick(Sender: TObject);
    procedure miMathAddClick(Sender: TObject);
    procedure miMathSubtractClick(Sender: TObject);
    procedure sgTableMouseUp(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure sbToggleHistoryClick(Sender: TObject);
    procedure tvHistoryDblClick(Sender: TObject);
    procedure tvHistoryDrawItem(Control: TWinControl; Index: Integer;
      Rect: TRect; State: TOwnerDrawState);
    procedure tvHistoryMeasureItem(Control: TWinControl; Index: Integer;
      var Height: Integer);
    {$endregion}

  private
    FFilename: String;
    FOldValue: String;
    FModified: Boolean;
    FSelStart: TPoint;
    FOnModifiedChanged: TNotifyEvent;
    FUndoStack: TStack<TEditorState>;
    FRedoStack: TStack<TEditorState>;
    FZoomPercent: Integer;
    FBaseFontSize: Integer;
    FBaseRowHeight: Integer;
    FBaseColWidths: TArray<Integer>;
    FHiddenRows: TArray<Boolean>;
    FHiddenCols: TArray<Boolean>;
    FPopupCol: Integer;

    FResizingCol: Integer;
    FResizeStartX: Integer;
    FResizeStartWidth: Integer;

    FHistoryEntries: TObjectList<THistoryEntry>;
    FHistoryNodeRefs: TObjectList<THistoryNodeRef>;
    FHistoryIndex: Integer;
    FHistoryWidth: Integer;
    FHistoryVisible: Boolean;
    FNextHistoryID: Int64;

    FPendingHistoryActive: Boolean;
    FPendingHistoryArea: TGridRect;
    FPendingHistoryAction: string;
    FPendingBeforeSnapshot: TGridSnapshot;
    FPendingBeforeState: TEditorState;
    FInitialSnapshot: TGridSnapshot;
    FApplyingHistorySnapshot: Boolean;
    FDarkMode: Boolean;
    FLockColsPrefixLabel: TLabel;
    FLockColsSuffixLabel: TLabel;
    FLockRowsLabel: TLabel;

    function GetResizeHitCol(X, Y: Integer): Integer;
    procedure ApplyTheme;
    procedure EnsureToolbarLabels;
    procedure UpdateToolbarLayout;

    procedure DeleteCurrentRow;
    procedure SetModified(AModified: Boolean);
    function GetTabsheet: TTabsheet;
    function GetCanUndo: Boolean;
    function GetCanRedo: Boolean;
    function ScaleValue(AValue: Integer): Integer;
    procedure CaptureBaseMetrics;
    procedure ApplyZoom;
    procedure ApplyHistoryVisibility;
    function GetClampedHistoryWidth(ARequestedWidth: Integer): Integer;
    function GetSettingsFileName: string;
    procedure LoadHistorySettings;
    procedure SaveHistorySettings;
    procedure RefreshHistoryView;
    procedure ClearHistoryBranchAfterCurrent;
    procedure RebuildUndoRedoStacksFromHistory;
    procedure PrepareHistoryAction(const AActionText: string);
    procedure CommitPendingHistory;
    procedure CancelPendingHistory(ARemoveUndoState: Boolean);
    procedure JumpToHistory(AHistoryIndex: Integer);
    function CaptureGridSnapshot: TGridSnapshot;
    procedure ApplyGridSnapshot(ASnapshot: TGridSnapshot);
    procedure BuildHistoryChanges(AEntry: THistoryEntry; ABefore, AAfter: TGridSnapshot; const AArea: TGridRect);
    procedure BuildDeletedRowChanges(AEntry: THistoryEntry; ABefore, AAfter: TGridSnapshot);
    procedure BuildDeletedColumnChanges(AEntry: THistoryEntry; ABefore, AAfter: TGridSnapshot);
    function BuildHistoryRowCaption(ABefore, AAfter: TGridSnapshot; ARow: Integer): string;
    function BuildHistoryColumnCaption(ABefore, AAfter: TGridSnapshot; ACol: Integer): string;
    function BuildGroupedChangeText(AEntry: THistoryEntry; ARow: Integer; out AFirstCol: Integer): string;
    procedure UpdateGridScrollbars;
    function GetWholeGridRect: TGridRect;
    procedure AutoFitColumn(ACol: Integer; const AUseAllRows: Boolean);
    procedure AutoFitColumns(const AUseAllRows: Boolean);
    procedure ZoomIn;
    procedure ZoomOut;
    procedure ZoomReset;

    procedure ClearStack(AStack: TStack<TEditorState>);
    function CreateEditorState(AArea: TGridRect): TEditorState;
    procedure ProcessEditorState(AFromStack: TStack<TEditorState>; AToStack: TStack<TEditorState>);

    procedure SyncHiddenArrays;
    procedure InsertRowState(AIndex, ACount: Integer);
    procedure DeleteRowState(AIndex, ACount: Integer);
    procedure InsertColState(AIndex, ACount: Integer; ADefaultWidth: Integer);
    procedure DeleteColState(AIndex, ACount: Integer);

    function PromptForInteger(const ACaption, APrompt: string; ADefault: Integer; out AValue: Integer): Boolean;
    function PromptForFloat(const ACaption, APrompt: string; ADefault: Double; out AValue: Double): Boolean;
    function TryParseCellFloat(const S: string; out AValue: Double): Boolean;
    function FormatCellNumber(const AValue: Double): string;

    procedure ApplyMathToSelection(const AOperand: Double; const AOp: Char);

  protected
    procedure Resize; override;

  public
    constructor Create(AOwner: TComponent; AFilename: String); reintroduce;
    destructor Destroy; override;

    procedure LoadFile(const AFilename: String);
    procedure SaveFile();

    procedure RefreshVisualMetrics;
    function Find(AText: String; AStartPos: TPoint; AOptions: TFindOptions; out ResultPosition: TPoint): Boolean;
    procedure Select(ACol, ARow: Integer; AScrollTo: Boolean = false);

    procedure CreateUndo(); overload;
    procedure CreateUndo(AArea: TGridRect); overload;
    procedure Undo();
    procedure Redo();
    procedure ToggleHistoryPanel;
    procedure SetDarkModeEnabled(AEnabled: Boolean);
    function HandleZoomShortcut(AKey: Word; Shift: TShiftState): Boolean;

    property Filename: String read FFilename;
    property Modified: Boolean read FModified;
    property CanUndo: Boolean read GetCanUndo;
    property CanRedo: Boolean read GetCanRedo;
    property Tabsheet: TTabsheet read GetTabsheet;

    property OnModifiedChanged: TNotifyEvent read FOnModifiedChanged write FOnModifiedChanged;
  end;

implementation

{$R *.dfm}

uses
  StrUtils, System.Types, System.IOUtils, System.Hash, IniFiles;

const
  clAlternatingRow = $00FAFAFA;

  WSB_PROP_CYVSCROLL = $00000001;
  WSB_PROP_CXHSCROLL = $00000002;
  WSB_PROP_CYHSCROLL = $00000004;
  WSB_PROP_CXVSCROLL = $00000008;

function InitializeFlatSB(hWnd: HWND): BOOL; stdcall; external 'comctl32.dll' name 'InitializeFlatSB';
function FlatSB_SetScrollProp(hWnd: HWND; Index, NewValue: Integer; Redraw: BOOL): BOOL; stdcall; external 'comctl32.dll' name 'FlatSB_SetScrollProp';
function FlatSB_ShowScrollBar(hWnd: HWND; code: Integer; fShow: BOOL): BOOL; stdcall; external 'comctl32.dll' name 'FlatSB_ShowScrollBar';

function SnapshotsRowsEqual(ABefore: TGridSnapshot; ABeforeRow: Integer;
  AAfter: TGridSnapshot; AAfterRow: Integer): Boolean;
var
  X, MaxCol: Integer;
begin
  if not Assigned(ABefore) or not Assigned(AAfter) then
    Exit(False);

  MaxCol := Min(ABefore.ColCount, AAfter.ColCount) - 1;
  for X := 0 to MaxCol do
    if ABefore[X, ABeforeRow] <> AAfter[X, AAfterRow] then
      Exit(False);

  Result := True;
end;

function SnapshotsColsEqual(ABefore: TGridSnapshot; ABeforeCol: Integer;
  AAfter: TGridSnapshot; AAfterCol: Integer): Boolean;
var
  Y, MaxRow: Integer;
begin
  if not Assigned(ABefore) or not Assigned(AAfter) then
    Exit(False);

  MaxRow := Min(ABefore.RowCount, AAfter.RowCount) - 1;
  for Y := 0 to MaxRow do
    if ABefore[ABeforeCol, Y] <> AAfter[AAfterCol, Y] then
      Exit(False);

  Result := True;
end;

{ TGridSnapshot }

constructor TGridSnapshot.Create;
begin
  inherited Create;
  FCells := TStringList.Create;
end;

destructor TGridSnapshot.Destroy;
begin
  FreeAndNil(FCells);
  inherited;
end;

function TGridSnapshot.GetCell(ACol, ARow: Integer): string;
var
  Index: Integer;
begin
  Result := '';
  if (ACol < 0) or (ARow < 0) or (ACol >= FColCount) or (ARow >= FRowCount) then
    Exit;

  Index := (ARow * FColCount) + ACol;
  if (Index >= 0) and (Index < FCells.Count) then
    Result := FCells[Index];
end;

procedure TGridSnapshot.SetCell(ACol, ARow: Integer; const AValue: string);
var
  Index: Integer;
begin
  if (ACol < 0) or (ARow < 0) or (ACol >= FColCount) or (ARow >= FRowCount) then
    Exit;

  while FCells.Count < (FRowCount * FColCount) do
    FCells.Add('');

  Index := (ARow * FColCount) + ACol;
  FCells[Index] := AValue;
end;

{ THistoryEntry }

constructor THistoryEntry.Create;
begin
  inherited Create;
  FChanges := TObjectList<THistoryCellChange>.Create(True);
end;

destructor THistoryEntry.Destroy;
begin
  FreeAndNil(FBeforeState);
  FreeAndNil(FAfterState);
  FreeAndNil(FAfterSnapshot);
  FreeAndNil(FChanges);
  inherited;
end;

{ TDisplayedHistoryBlock }

constructor TDisplayedHistoryBlock.Create;
begin
  inherited Create;
  FLines := TStringList.Create;
  FLastEntryIndex := -1;
  FTargetRow := -1;
  FTargetCol := -1;
end;

destructor TDisplayedHistoryBlock.Destroy;
begin
  FreeAndNil(FLines);
  inherited;
end;

{ TTableFrame }


procedure TTableFrame.ApplyHistoryVisibility;
var
  NewWidth: Integer;
begin
  if FHistoryVisible then
  begin
    plHistory.Visible := True;
    spHistory.Visible := True;

    NewWidth := GetClampedHistoryWidth(FHistoryWidth);
    if NewWidth <> FHistoryWidth then
      FHistoryWidth := NewWidth;

    if plHistory.Width <> FHistoryWidth then
      plHistory.Width := FHistoryWidth;

    sbToggleHistory.Caption := 'Hide History';
  end
  else
  begin
    if plHistory.Visible and (plHistory.Width > 0) then
      FHistoryWidth := GetClampedHistoryWidth(plHistory.Width);

    plHistory.Visible := False;
    spHistory.Visible := False;
    sbToggleHistory.Caption := 'Show History';
  end;
end;


function TTableFrame.GetSettingsFileName: string;
var
  Path: string;
begin
  Path := IncludeTrailingPathDelimiter(GetEnvironmentVariable('APPDATA')) + 'D2ExcelPlus\';
  ForceDirectories(Path);
  Result := Path + 'settings.ini';
end;

procedure TTableFrame.LoadHistorySettings;
var
  Ini: TIniFile;
begin
  Ini := TIniFile.Create(GetSettingsFileName);
  try
    FHistoryVisible := Ini.ReadBool('SETTINGS', 'history_visible', True);
    FHistoryWidth := Ini.ReadInteger('SETTINGS', 'history_width', 240);
    FHistoryWidth := EnsureRange(FHistoryWidth, 180, 280);
  finally
    Ini.Free;
  end;
end;


procedure TTableFrame.SaveHistorySettings;
var
  Ini: TIniFile;
begin
  Ini := TIniFile.Create(GetSettingsFileName);
  try
    if plHistory.Visible and (plHistory.Width > 0) then
      FHistoryWidth := GetClampedHistoryWidth(plHistory.Width)
    else
      FHistoryWidth := GetClampedHistoryWidth(FHistoryWidth);

    Ini.WriteBool('SETTINGS', 'history_visible', FHistoryVisible);
    Ini.WriteInteger('SETTINGS', 'history_width', FHistoryWidth);
    Ini.UpdateFile;
  finally
    Ini.Free;
  end;
end;


procedure TTableFrame.PrepareHistoryAction(const AActionText: string);
begin
  FPendingHistoryAction := AActionText;
end;


procedure TTableFrame.UpdateGridScrollbars;
var
  ScrollSize: Integer;
begin
  sgTable.HandleNeeded;
  InitializeFlatSB(sgTable.Handle);

  ScrollSize := Max(18, ScaleValue(18));
  FlatSB_SetScrollProp(sgTable.Handle, WSB_PROP_CYHSCROLL, ScrollSize, True);
  FlatSB_SetScrollProp(sgTable.Handle, WSB_PROP_CXVSCROLL, ScrollSize, True);
  FlatSB_SetScrollProp(sgTable.Handle, WSB_PROP_CYVSCROLL, ScrollSize, True);
  FlatSB_SetScrollProp(sgTable.Handle, WSB_PROP_CXHSCROLL, ScrollSize, True);
  FlatSB_ShowScrollBar(sgTable.Handle, SB_BOTH, True);
end;

function TTableFrame.GetWholeGridRect: TGridRect;
begin
  Result.Left := 0;
  Result.Top := 0;
  Result.Right := Max(0, sgTable.ColCount - 1);
  Result.Bottom := Max(0, sgTable.RowCount - 1);
end;

function TTableFrame.CaptureGridSnapshot: TGridSnapshot;
var
  X, Y: Integer;
begin
  Result := TGridSnapshot.Create;
  Result.ColCount := sgTable.ColCount;
  Result.RowCount := sgTable.RowCount;
  Result.FixedCols := sgTable.FixedCols;
  Result.FixedRows := sgTable.FixedRows;
  Result.Selection := sgTable.Selection;
  Result.LeftCol := sgTable.LeftCol;
  Result.TopRow := sgTable.TopRow;
  Result.BaseColWidths := Copy(FBaseColWidths);
  Result.HiddenRows := Copy(FHiddenRows);
  Result.HiddenCols := Copy(FHiddenCols);

  for Y := 0 to sgTable.RowCount - 1 do
    for X := 0 to sgTable.ColCount - 1 do
      Result[X, Y] := sgTable.Cells[X, Y];
end;

procedure TTableFrame.ApplyGridSnapshot(ASnapshot: TGridSnapshot);
var
  X, Y: Integer;
begin
  if not Assigned(ASnapshot) then
    Exit;

  sgTable.HandleNeeded;
  SendMessage(sgTable.Handle, WM_SETREDRAW, 0, 0);
  try
    sgTable.ColCount := Max(1, ASnapshot.ColCount);
    sgTable.RowCount := Max(1, ASnapshot.RowCount);
    sgTable.FixedCols := Min(ASnapshot.FixedCols, sgTable.ColCount - 1);
    sgTable.FixedRows := Min(ASnapshot.FixedRows, sgTable.RowCount - 1);

    for Y := 0 to sgTable.RowCount - 1 do
      for X := 0 to sgTable.ColCount - 1 do
        sgTable.Cells[X, Y] := ASnapshot[X, Y];

    FBaseColWidths := Copy(ASnapshot.BaseColWidths);
    FHiddenRows := Copy(ASnapshot.HiddenRows);
    FHiddenCols := Copy(ASnapshot.HiddenCols);

    if Length(FBaseColWidths) <> sgTable.ColCount then
    begin
      SetLength(FBaseColWidths, sgTable.ColCount);
      for X := 0 to sgTable.ColCount - 1 do
        if FBaseColWidths[X] <= 0 then
          FBaseColWidths[X] := Max(12, sgTable.ColWidths[X]);
    end;

    if Length(FHiddenRows) <> sgTable.RowCount then
      SetLength(FHiddenRows, sgTable.RowCount);

    if Length(FHiddenCols) <> sgTable.ColCount then
      SetLength(FHiddenCols, sgTable.ColCount);

    ApplyZoom;

    sgTable.Selection := ASnapshot.Selection;
    sgTable.LeftCol := Max(ASnapshot.LeftCol, sgTable.FixedCols);
    sgTable.TopRow := Max(ASnapshot.TopRow, sgTable.FixedRows);
  finally
    SendMessage(sgTable.Handle, WM_SETREDRAW, 1, 0);
    sgTable.Invalidate;
  end;
end;

function TTableFrame.BuildHistoryColumnCaption(ABefore, AAfter: TGridSnapshot; ACol: Integer): string;
begin
  Result := '';
  if Assigned(AAfter) then
    Result := AAfter[ACol, 0];

  if (Result = '') and Assigned(ABefore) then
    Result := ABefore[ACol, 0];

  if Result = '' then
    Result := 'Column ' + IntToStr(ACol);
end;

function TTableFrame.BuildHistoryRowCaption(ABefore, AAfter: TGridSnapshot; ARow: Integer): string;
var
  RowName: string;
begin
  if ARow = 0 then
    Exit('[Header Row]');

  RowName := '';
  if Assigned(AAfter) then
    RowName := AAfter[1, ARow];

  if (RowName = '') and Assigned(ABefore) then
    RowName := ABefore[1, ARow];

  if RowName <> '' then
    Result := Format('%s (Row %d)', [RowName, ARow])
  else
    Result := Format('(Row %d)', [ARow]);
end;



procedure TTableFrame.BuildHistoryChanges(AEntry: THistoryEntry; ABefore, AAfter: TGridSnapshot; const AArea: TGridRect);
var
  X, Y: Integer;
  LeftCol, TopRow, RightCol, BottomRow: Integer;
  Change: THistoryCellChange;
  OldValue, NewValue: string;
begin
  if not Assigned(AEntry) or not Assigned(ABefore) or not Assigned(AAfter) then
    Exit;

  if SameText(AEntry.ActionText, 'Delete Row') then
  begin
    BuildDeletedRowChanges(AEntry, ABefore, AAfter);
    Exit;
  end;

  if SameText(AEntry.ActionText, 'Delete Column(s)') then
  begin
    BuildDeletedColumnChanges(AEntry, ABefore, AAfter);
    Exit;
  end;

  LeftCol := Max(AArea.Left, 1);
  TopRow := Max(AArea.Top, 0);
  RightCol := Min(AArea.Right, Max(ABefore.ColCount, AAfter.ColCount) - 1);
  BottomRow := Min(AArea.Bottom, Max(ABefore.RowCount, AAfter.RowCount) - 1);

  if (ABefore.RowCount <> AAfter.RowCount) or (ABefore.ColCount <> AAfter.ColCount) then
  begin
    LeftCol := 1;
    TopRow := 0;
    RightCol := Max(ABefore.ColCount, AAfter.ColCount) - 1;
    BottomRow := Max(ABefore.RowCount, AAfter.RowCount) - 1;
  end;

  for Y := TopRow to BottomRow do
  begin
    for X := LeftCol to RightCol do
    begin
      OldValue := ABefore[X, Y];
      NewValue := AAfter[X, Y];

      if OldValue = NewValue then
        Continue;

      Change := THistoryCellChange.Create;
      Change.Row := Y;
      Change.Col := X;
      Change.RowCaption := BuildHistoryRowCaption(ABefore, AAfter, Y);
      Change.ColumnName := BuildHistoryColumnCaption(ABefore, AAfter, X);
      Change.OldValue := OldValue;
      Change.NewValue := NewValue;
      AEntry.Changes.Add(Change);
    end;
  end;
end;

procedure TTableFrame.BuildDeletedRowChanges(AEntry: THistoryEntry; ABefore, AAfter: TGridSnapshot);
var
  DeletedCount, StartRow, Y: Integer;
  Change: THistoryCellChange;
begin
  if not Assigned(AEntry) or not Assigned(ABefore) or not Assigned(AAfter) then
    Exit;

  DeletedCount := Max(0, ABefore.RowCount - AAfter.RowCount);
  if DeletedCount <= 0 then
    Exit;

  StartRow := AAfter.RowCount;
  for Y := Max(AAfter.FixedRows, 0) to AAfter.RowCount - 1 do
    if not SnapshotsRowsEqual(ABefore, Y, AAfter, Y) then
    begin
      StartRow := Y;
      Break;
    end;

  for Y := StartRow to Min(ABefore.RowCount - 1, StartRow + DeletedCount - 1) do
  begin
    Change := THistoryCellChange.Create;
    Change.Row := Min(Y, Max(AAfter.RowCount - 1, AAfter.FixedRows));
    Change.Col := Max(AAfter.FixedCols, 1);
    Change.RowCaption := BuildHistoryRowCaption(ABefore, nil, Y);
    Change.DisplayText := '[Deleted Row]';
    AEntry.Changes.Add(Change);
  end;
end;

procedure TTableFrame.BuildDeletedColumnChanges(AEntry: THistoryEntry; ABefore, AAfter: TGridSnapshot);
var
  DeletedCount, StartCol, X: Integer;
  Change: THistoryCellChange;
begin
  if not Assigned(AEntry) or not Assigned(ABefore) or not Assigned(AAfter) then
    Exit;

  DeletedCount := Max(0, ABefore.ColCount - AAfter.ColCount);
  if DeletedCount <= 0 then
    Exit;

  StartCol := AAfter.ColCount;
  for X := Max(AAfter.FixedCols, 1) to AAfter.ColCount - 1 do
    if not SnapshotsColsEqual(ABefore, X, AAfter, X) then
    begin
      StartCol := X;
      Break;
    end;

  for X := StartCol to Min(ABefore.ColCount - 1, StartCol + DeletedCount - 1) do
  begin
    Change := THistoryCellChange.Create;
    Change.Row := -1;
    Change.Col := Min(X, Max(AAfter.ColCount - 1, AAfter.FixedCols));
    Change.RowCaption := '[Columns]';
    Change.ColumnName := BuildHistoryColumnCaption(ABefore, nil, X);
    Change.DisplayText := Format('%s [Deleted]', [Change.ColumnName]);
    AEntry.Changes.Add(Change);
  end;
end;

function TTableFrame.BuildGroupedChangeText(AEntry: THistoryEntry; ARow: Integer; out AFirstCol: Integer): string;
var
  I: Integer;
  Change: THistoryCellChange;
  Part: string;
begin
  Result := '';
  AFirstCol := -1;

  if not Assigned(AEntry) then
    Exit;

  for I := 0 to AEntry.Changes.Count - 1 do
  begin
    Change := AEntry.Changes[I];
    if Change.Row <> ARow then
      Continue;

    if AFirstCol < 0 then
      AFirstCol := Change.Col;

    if Change.DisplayText <> '' then
      Part := Change.DisplayText
    else
      Part := Format('%s: %s -> %s',
        [Change.ColumnName, QuotedStr(Change.OldValue), QuotedStr(Change.NewValue)]);

    if Result <> '' then
      Result := Result + sLineBreak;
    Result := Result + Part;
  end;
end;

procedure TTableFrame.ClearHistoryBranchAfterCurrent;
begin
  while FHistoryEntries.Count - 1 > FHistoryIndex do
    FHistoryEntries.Delete(FHistoryEntries.Count - 1);
end;

procedure TTableFrame.CommitPendingHistory;
var
  Entry: THistoryEntry;
  AfterState: TEditorState;
  AfterSnapshot: TGridSnapshot;
begin
  if not FPendingHistoryActive then
    Exit;

  AfterState := CreateEditorState(FPendingHistoryArea);
  AfterSnapshot := CaptureGridSnapshot;
  Entry := THistoryEntry.Create;
  try
    Entry.EntryID := FNextHistoryID;
    Inc(FNextHistoryID);

    Entry.FileName := ExtractFileName(FFilename);
    Entry.ActionText := FPendingHistoryAction;
    if Entry.ActionText = '' then
      Entry.ActionText := 'Edit';
    Entry.TimeStamp := Now;

    Entry.RowCountBefore := FPendingBeforeSnapshot.RowCount;
    Entry.ColCountBefore := FPendingBeforeSnapshot.ColCount;
    Entry.RowCountAfter := AfterSnapshot.RowCount;
    Entry.ColCountAfter := AfterSnapshot.ColCount;

    Entry.BeforeState := FPendingBeforeState;
    FPendingBeforeState := nil;

    Entry.AfterState := AfterState;
    AfterState := nil;

    Entry.AfterSnapshot := AfterSnapshot;
    AfterSnapshot := nil;

    BuildHistoryChanges(Entry, FPendingBeforeSnapshot, Entry.AfterSnapshot, FPendingHistoryArea);

    if (Entry.Changes.Count = 0) and
       (Entry.RowCountBefore = Entry.RowCountAfter) and
       (Entry.ColCountBefore = Entry.ColCountAfter) then
    begin
      if FUndoStack.Count > 0 then
        FUndoStack.Pop.Free;
      Exit;
    end;

    ClearStack(FRedoStack);
    ClearHistoryBranchAfterCurrent;

    FHistoryEntries.Add(Entry);
    Entry := nil;
    FHistoryIndex := FHistoryEntries.Count - 1;

    RefreshHistoryView;
  finally
    Entry.Free;
    AfterState.Free;
    AfterSnapshot.Free;
    FreeAndNil(FPendingBeforeSnapshot);

    FPendingHistoryActive := False;
    FPendingHistoryAction := '';

    if Assigned(FOnModifiedChanged) then
      FOnModifiedChanged(Tabsheet);
  end;
end;

procedure TTableFrame.CancelPendingHistory(ARemoveUndoState: Boolean);
begin
  if not FPendingHistoryActive then
    Exit;

  if ARemoveUndoState and (FUndoStack.Count > 0) then
    FUndoStack.Pop.Free;

  FreeAndNil(FPendingBeforeSnapshot);
  FreeAndNil(FPendingBeforeState);

  FPendingHistoryActive := False;
  FPendingHistoryAction := '';

  if Assigned(FOnModifiedChanged) then
    FOnModifiedChanged(Tabsheet);
end;

procedure TTableFrame.RebuildUndoRedoStacksFromHistory;
var
  I: Integer;
begin
  ClearStack(FUndoStack);
  ClearStack(FRedoStack);

  for I := 0 to FHistoryIndex do
    FUndoStack.Push(FHistoryEntries[I].BeforeState.Clone);

  for I := FHistoryEntries.Count - 1 downto FHistoryIndex + 1 do
    FRedoStack.Push(FHistoryEntries[I].AfterState.Clone);

  if Assigned(FOnModifiedChanged) then
    FOnModifiedChanged(Tabsheet);
end;




function TTableFrame.GetClampedHistoryWidth(ARequestedWidth: Integer): Integer;
var
  MaxWidth: Integer;
begin
  MaxWidth := ClientWidth - 420;
  if MaxWidth < 160 then
    MaxWidth := ClientWidth div 2;

  MaxWidth := Min(MaxWidth, 280);
  if MaxWidth < 140 then
    MaxWidth := 140;

  Result := EnsureRange(ARequestedWidth, 140, MaxWidth);
end;

procedure TTableFrame.Resize;
var
  NewWidth: Integer;
begin
  inherited;

  UpdateToolbarLayout;

  if csDestroying in ComponentState then
    Exit;

  if FHistoryVisible then
  begin
    plHistory.Visible := True;
    spHistory.Visible := True;

    if plHistory.Width > 0 then
      NewWidth := GetClampedHistoryWidth(plHistory.Width)
    else
      NewWidth := GetClampedHistoryWidth(FHistoryWidth);

    if FHistoryWidth <> NewWidth then
      FHistoryWidth := NewWidth;

    if plHistory.Width <> NewWidth then
      plHistory.Width := NewWidth;
  end
  else
  begin
    plHistory.Visible := False;
    spHistory.Visible := False;
  end;

  if tvHistory.HandleAllocated then
    tvHistory.Invalidate;
end;

procedure TTableFrame.RefreshHistoryView;
var
  I, J: Integer;
  Entry: THistoryEntry;
  Change: THistoryCellChange;
  NodeRef: THistoryNodeRef;
  Map: TStringList;
  Order: TObjectList<TDisplayedHistoryBlock>;
  Block: TDisplayedHistoryBlock;
  CurrentItemIndex: Integer;
  ItemText, LineText, BlockKey, BlockHeader: string;

  function GetOrCreateBlock(const AKey, AHeader: string; AEntryIndex, ARow, ACol: Integer): TDisplayedHistoryBlock;
  var
    Idx, OrderIdx: Integer;
  begin
    Idx := Map.IndexOf(AKey);
    if Idx >= 0 then
      Result := TDisplayedHistoryBlock(Map.Objects[Idx])
    else
    begin
      Result := TDisplayedHistoryBlock.Create;
      Result.Key := AKey;
      Result.Header := AHeader;
      Order.Add(Result);
      Map.AddObject(AKey, Result);
    end;

    if AHeader <> '' then
      Result.Header := AHeader;

    Result.LastEntryIndex := AEntryIndex;
    Result.TargetRow := ARow;
    Result.TargetCol := ACol;

    OrderIdx := Order.IndexOf(Result);
    if (OrderIdx >= 0) and (OrderIdx <> Order.Count - 1) then
    begin
      Order.Extract(Result);
      Order.Add(Result);
    end;
  end;

begin
  tvHistory.Items.BeginUpdate;
  Map := TStringList.Create;
  Order := TObjectList<TDisplayedHistoryBlock>.Create(True);
  try
    FHistoryNodeRefs.Clear;
    tvHistory.Items.Clear;
    CurrentItemIndex := -1;

    lblHistoryTitle.Caption := 'History - ' + ExtractFileName(FFilename);

    for I := 0 to FHistoryIndex do
    begin
      if (I < 0) or (I >= FHistoryEntries.Count) then
        Continue;

      Entry := FHistoryEntries[I];

      if Entry.Changes.Count = 0 then
      begin
        BlockKey := Format('ENTRY|%d', [I]);
        BlockHeader := Entry.ActionText;
        if (Entry.RowCountBefore <> Entry.RowCountAfter) or
           (Entry.ColCountBefore <> Entry.ColCountAfter) then
          LineText := Format('Rows %d / Cols %d -> Rows %d / Cols %d',
            [Entry.RowCountBefore, Entry.ColCountBefore, Entry.RowCountAfter, Entry.ColCountAfter])
        else
          LineText := '';

        Block := GetOrCreateBlock(BlockKey, BlockHeader, I, sgTable.FixedRows, sgTable.FixedCols);
        if LineText <> '' then
          Block.Lines.Add(LineText);
        Continue;
      end;

      for J := 0 to Entry.Changes.Count - 1 do
      begin
        Change := Entry.Changes[J];

        if Change.DisplayText <> '' then
          LineText := Change.DisplayText
        else
          LineText := Format('%s: %s -> %s',
            [Change.ColumnName, QuotedStr(Change.OldValue), QuotedStr(Change.NewValue)]);

        if (Change.Row >= sgTable.FixedRows) and (Trim(Change.RowCaption) <> '') and
           not SameText(Change.RowCaption, '[Columns]') then
        begin
          BlockKey := 'ROW|' + Change.RowCaption;
          BlockHeader := Change.RowCaption;
        end
        else
        begin
          BlockKey := Format('ENTRY|%d|%d', [I, J]);
          if Trim(Change.RowCaption) <> '' then
            BlockHeader := Change.RowCaption
          else if Trim(Entry.ActionText) <> '' then
            BlockHeader := Entry.ActionText
          else
            BlockHeader := 'History';
        end;

        Block := GetOrCreateBlock(BlockKey, BlockHeader, I, Change.Row, Change.Col);
        if LineText <> '' then
          Block.Lines.Add(LineText);
      end;
    end;

    for I := 0 to Order.Count - 1 do
    begin
      Block := Order[I];
      ItemText := Block.Header;

      for J := 0 to Block.Lines.Count - 1 do
        ItemText := ItemText + sLineBreak + '  ' + Block.Lines[J];

      NodeRef := THistoryNodeRef.Create;
      if Pos('ROW|', Block.Key) = 1 then
        NodeRef.Kind := hnkRow
      else
        NodeRef.Kind := hnkEntry;
      NodeRef.EntryIndex := Block.LastEntryIndex;
      NodeRef.Row := Block.TargetRow;
      NodeRef.Col := Block.TargetCol;
      FHistoryNodeRefs.Add(NodeRef);

      tvHistory.Items.AddObject(ItemText, NodeRef);

      if (Block.LastEntryIndex = FHistoryIndex) and (CurrentItemIndex < 0) then
        CurrentItemIndex := tvHistory.Items.Count - 1;
    end;

    if CurrentItemIndex >= 0 then
    begin
      tvHistory.ItemIndex := CurrentItemIndex;
      tvHistory.TopIndex := Max(0, CurrentItemIndex - 2);
    end
    else
      tvHistory.ItemIndex := -1;
  finally
    Order.Free;
    Map.Free;
    tvHistory.Items.EndUpdate;
    tvHistory.Invalidate;
  end;
end;

procedure TTableFrame.JumpToHistory(AHistoryIndex: Integer);
begin
  if (AHistoryIndex < -1) or (AHistoryIndex >= FHistoryEntries.Count) then
    Exit;

  if AHistoryIndex = FHistoryIndex then
    Exit;

  FApplyingHistorySnapshot := True;
  try
    if AHistoryIndex = -1 then
      ApplyGridSnapshot(FInitialSnapshot)
    else
      ApplyGridSnapshot(FHistoryEntries[AHistoryIndex].AfterSnapshot);

    FHistoryIndex := AHistoryIndex;
    RebuildUndoRedoStacksFromHistory;
    RefreshHistoryView;

    if AHistoryIndex = -1 then
      SetModified(False)
    else
      SetModified(True);
  finally
    FApplyingHistorySnapshot := False;
  end;
end;

procedure TTableFrame.ToggleHistoryPanel;
begin
  FHistoryVisible := not FHistoryVisible;
  ApplyHistoryVisibility;
  SaveHistorySettings;
end;

procedure TTableFrame.sbToggleHistoryClick(Sender: TObject);
begin
  ToggleHistoryPanel;
end;


procedure TTableFrame.tvHistoryDblClick(Sender: TObject);
var
  NodeRef: THistoryNodeRef;
  TargetRow, TargetCol: Integer;
begin
  if (tvHistory.ItemIndex < 0) or
     (tvHistory.ItemIndex >= tvHistory.Items.Count) or
     not Assigned(tvHistory.Items.Objects[tvHistory.ItemIndex]) then
    Exit;

  NodeRef := THistoryNodeRef(tvHistory.Items.Objects[tvHistory.ItemIndex]);
  JumpToHistory(NodeRef.EntryIndex);

  TargetCol := Max(NodeRef.Col, sgTable.FixedCols);
  TargetRow := NodeRef.Row;

  if TargetRow < sgTable.FixedRows then
    TargetRow := sgTable.FixedRows;
  if TargetRow >= sgTable.RowCount then
    TargetRow := sgTable.RowCount - 1;

  if TargetCol < sgTable.FixedCols then
    TargetCol := sgTable.FixedCols;
  if TargetCol >= sgTable.ColCount then
    TargetCol := sgTable.ColCount - 1;

  if (TargetCol >= 0) and (TargetRow >= 0) then
    Select(TargetCol, TargetRow, True);

  if sgTable.CanFocus then
    sgTable.SetFocus;
end;


procedure TTableFrame.EnsureToolbarLabels;
begin
  if not Assigned(FLockColsPrefixLabel) then
  begin
    FLockColsPrefixLabel := TLabel.Create(Self);
    FLockColsPrefixLabel.Parent := Panel1;
    FLockColsPrefixLabel.Caption := 'Lock first';
    FLockColsPrefixLabel.Transparent := True;
    FLockColsPrefixLabel.FocusControl := cbFixColumns;
  end;

  if not Assigned(FLockColsSuffixLabel) then
  begin
    FLockColsSuffixLabel := TLabel.Create(Self);
    FLockColsSuffixLabel.Parent := Panel1;
    FLockColsSuffixLabel.Caption := 'Columns';
    FLockColsSuffixLabel.Transparent := True;
    FLockColsSuffixLabel.FocusControl := cbFixColumns;
  end;

  if not Assigned(FLockRowsLabel) then
  begin
    FLockRowsLabel := TLabel.Create(Self);
    FLockRowsLabel.Parent := Panel1;
    FLockRowsLabel.Caption := 'Lock first row';
    FLockRowsLabel.Transparent := True;
    FLockRowsLabel.FocusControl := cbFixRows;
  end;

  cbFixColumns.Caption := '';
  cbFixColumns.Width := 17;

  cbFixRows.Caption := '';
  cbFixRows.Width := 17;
end;

procedure TTableFrame.UpdateToolbarLayout;
begin
  EnsureToolbarLabels;

  cbFixColumns.SetBounds(8, 6, 17, 17);

  FLockColsPrefixLabel.Left := cbFixColumns.Left + cbFixColumns.Width + 4;
  FLockColsPrefixLabel.Top := 7;

  seFixedColumns.Left := FLockColsPrefixLabel.Left + FLockColsPrefixLabel.Width + 6;
  seFixedColumns.Top := 4;

  FLockColsSuffixLabel.Left := seFixedColumns.Left + seFixedColumns.Width + 8;
  FLockColsSuffixLabel.Top := 7;

  cbFixRows.Left := FLockColsSuffixLabel.Left + FLockColsSuffixLabel.Width + 18;
  cbFixRows.Top := 8;

  FLockRowsLabel.Left := cbFixRows.Left + cbFixRows.Width + 4;
  FLockRowsLabel.Top := 7;
end;

procedure TTableFrame.SetDarkModeEnabled(AEnabled: Boolean);

begin

  if FDarkMode = AEnabled then

    Exit;



  FDarkMode := AEnabled;

  ApplyTheme;

end;



procedure TTableFrame.ApplyTheme;

var

  PanelColor: TColor;

  SurfaceColor: TColor;

  AltSurfaceColor: TColor;

  FixedColor: TColor;

  TextColor: TColor;

begin

  ParentBackground := False;



  if FDarkMode then

  begin

    Color := RGB(45, 45, 48);

    PanelColor := RGB(45, 45, 48);

    SurfaceColor := RGB(30, 30, 30);

    AltSurfaceColor := RGB(38, 38, 40);

    FixedColor := RGB(56, 56, 58);

    TextColor := RGB(240, 240, 240);

  end

  else

  begin

    Color := clBtnFace;

    PanelColor := clBtnFace;

    SurfaceColor := clWindow;

    AltSurfaceColor := clAlternatingRow;

    FixedColor := clBtnFace;

    TextColor := clWindowText;

  end;



  Panel1.ParentBackground := False;

  Panel1.Color := PanelColor;

  Panel1.Font.Color := TextColor;

  UpdateToolbarLayout;

  cbFixColumns.ParentColor := False;
  cbFixColumns.ParentFont := False;
  cbFixColumns.Color := PanelColor;
  cbFixColumns.Font.Color := TextColor;

  cbFixRows.ParentColor := False;
  cbFixRows.ParentFont := False;
  cbFixRows.Color := PanelColor;
  cbFixRows.Font.Color := TextColor;

  if Assigned(FLockColsPrefixLabel) then
  begin
    FLockColsPrefixLabel.Font.Color := TextColor;
    FLockColsPrefixLabel.Color := PanelColor;
  end;

  if Assigned(FLockColsSuffixLabel) then
  begin
    FLockColsSuffixLabel.Font.Color := TextColor;
    FLockColsSuffixLabel.Color := PanelColor;
  end;

  if Assigned(FLockRowsLabel) then
  begin
    FLockRowsLabel.Font.Color := TextColor;
    FLockRowsLabel.Color := PanelColor;
  end;

  seFixedColumns.ParentColor := False;
  seFixedColumns.ParentFont := False;
  seFixedColumns.Color := SurfaceColor;
  seFixedColumns.Font.Color := TextColor;



  sbToggleHistory.Font.Color := TextColor;



  plHistory.ParentBackground := False;

  plHistory.Color := PanelColor;

  plHistory.Font.Color := TextColor;



  pnlHistoryHeader.ParentBackground := False;

  pnlHistoryHeader.Color := PanelColor;

  pnlHistoryHeader.Font.Color := TextColor;

  lblHistoryTitle.Font.Color := TextColor;



  tvHistory.ParentColor := False;

  tvHistory.Color := SurfaceColor;

  tvHistory.Font.Color := TextColor;



  sgTable.ParentColor := False;

  sgTable.Color := SurfaceColor;

  sgTable.FixedColor := FixedColor;

  sgTable.Font.Color := TextColor;



  Invalidate;

  Panel1.Invalidate;

  sgTable.Invalidate;

  tvHistory.Invalidate;

end;



procedure TTableFrame.tvHistoryDrawItem(Control: TWinControl; Index: Integer;
  Rect: TRect; State: TOwnerDrawState);
var
  LBox: TListBox;
  R: TRect;
  S: string;
begin
  LBox := Control as TListBox;
  S := LBox.Items[Index];

  LBox.Canvas.Font.Assign(LBox.Font);
  LBox.Canvas.Font.Color := LBox.Font.Color;

  if odSelected in State then
  begin
    if FDarkMode then
      LBox.Canvas.Brush.Color := RGB(70, 90, 120)
    else
      LBox.Canvas.Brush.Color := $00E6D8C5;

    LBox.Canvas.Font.Style := [fsBold];
  end
  else
  begin
    if FDarkMode then
    begin
      if Odd(Index) then
        LBox.Canvas.Brush.Color := RGB(38, 38, 40)
      else
        LBox.Canvas.Brush.Color := RGB(30, 30, 30);
    end
    else if Odd(Index) then
      LBox.Canvas.Brush.Color := clAlternatingRow
    else
      LBox.Canvas.Brush.Color := clWindow;

    LBox.Canvas.Font.Style := [];
  end;

  LBox.Canvas.FillRect(Rect);

  R := Rect;
  InflateRect(R, -6, -4);
  DrawText(LBox.Canvas.Handle, PChar(S), Length(S), R,
    DT_LEFT or DT_NOPREFIX or DT_WORDBREAK or DT_EDITCONTROL);

  if odFocused in State then
    LBox.Canvas.DrawFocusRect(Rect);
end;

procedure TTableFrame.tvHistoryMeasureItem(Control: TWinControl; Index: Integer;
  var Height: Integer);
var
  LBox: TListBox;
  R: TRect;
  S: string;
begin
  LBox := Control as TListBox;
  S := LBox.Items[Index];

  R := Rect(0, 0, Max(40, LBox.ClientWidth - GetSystemMetrics(SM_CXVSCROLL) - 12), 0);
  LBox.Canvas.Font.Assign(LBox.Font);
  DrawText(LBox.Canvas.Handle, PChar(S), Length(S), R,
    DT_LEFT or DT_NOPREFIX or DT_WORDBREAK or DT_EDITCONTROL or DT_CALCRECT);

  Height := Max(LBox.ItemHeight, (R.Bottom - R.Top) + 10);
end;

function TTableFrame.ScaleValue(AValue: Integer): Integer;
begin
  Result := Max(1, MulDiv(AValue, FZoomPercent, 100));
end;

procedure TTableFrame.CaptureBaseMetrics;
var
  i: Integer;
begin
  FBaseFontSize := sgTable.Font.Size;
  FBaseRowHeight := sgTable.DefaultRowHeight;

  SetLength(FBaseColWidths, sgTable.ColCount);
  for i := 0 to sgTable.ColCount - 1 do
    FBaseColWidths[i] := sgTable.ColWidths[i];
end;


function TTableFrame.GetResizeHitCol(X, Y: Integer): Integer;
var
  Coord: TGridCoord;
  R: TRect;
  Tol: Integer;
begin
  Result := -1;

  Coord := sgTable.MouseCoord(X, Y);
  if (Coord.X < sgTable.FixedCols) or (Coord.Y < 0) then
    Exit;

  if (Coord.X >= 0) and (Coord.X < Length(FHiddenCols)) and FHiddenCols[Coord.X] then
    Exit;

  R := sgTable.CellRect(Coord.X, Coord.Y);
  Tol := Max(3, ScaleValue(4));

  if Abs(X - R.Right) <= Tol then
    Result := Coord.X;
end;

procedure TTableFrame.RefreshVisualMetrics;
begin
  sgTable.HandleNeeded;
  CaptureBaseMetrics;
  SyncHiddenArrays;
  ApplyZoom;
end;

procedure TTableFrame.ApplyZoom;
var
  i: Integer;
  NewFontSize: Integer;
  NewDefaultRowHeight: Integer;
  NewColWidth: Integer;
  NewRowHeight: Integer;
  HeaderMinWidth: Integer;
begin
  sgTable.HandleNeeded;

  NewFontSize := Max(1, MulDiv(FBaseFontSize, FZoomPercent, 100));
  NewDefaultRowHeight := ScaleValue(FBaseRowHeight);

  SendMessage(sgTable.Handle, WM_SETREDRAW, 0, 0);
  try
    if sgTable.Font.Size <> NewFontSize then
      sgTable.Font.Size := NewFontSize;

    if sgTable.DefaultRowHeight <> NewDefaultRowHeight then
      sgTable.DefaultRowHeight := NewDefaultRowHeight;

    sgTable.Canvas.Font.Assign(sgTable.Font);
    sgTable.Canvas.Font.Quality := fqDefault;

    for i := 0 to High(FBaseColWidths) do
    begin
      if (i < Length(FHiddenCols)) and FHiddenCols[i] then
        NewColWidth := 0
      else
      begin
        NewColWidth := ScaleValue(FBaseColWidths[i]);

        if (sgTable.FixedRows > 0) and (i >= sgTable.FixedCols) and (sgTable.RowCount > 0) then
        begin
          sgTable.Canvas.Font.Style := [fsBold];
          HeaderMinWidth := sgTable.Canvas.TextWidth(sgTable.Cells[i, 0]) + ScaleValue(12);
          sgTable.Canvas.Font.Style := [];

          if NewColWidth < HeaderMinWidth then
            NewColWidth := HeaderMinWidth;
        end;
      end;

      if sgTable.ColWidths[i] <> NewColWidth then
        sgTable.ColWidths[i] := NewColWidth;
    end;

    for i := 0 to sgTable.RowCount - 1 do
    begin
      if (i < Length(FHiddenRows)) and FHiddenRows[i] then
        NewRowHeight := 0
      else
        NewRowHeight := NewDefaultRowHeight;

      if sgTable.RowHeights[i] <> NewRowHeight then
        sgTable.RowHeights[i] := NewRowHeight;
    end;
  finally
    SendMessage(sgTable.Handle, WM_SETREDRAW, 1, 0);
    UpdateGridScrollbars;
    sgTable.Invalidate;
    tvHistory.Invalidate;
    pnlHistoryHeader.Invalidate;
  end;
end;

procedure TTableFrame.AutoFitColumn(ACol: Integer; const AUseAllRows: Boolean);
var
  Y: Integer;
  W: Integer;
  LastRow: Integer;
  HeaderPad: Integer;
  CellPad: Integer;
  MinW: Integer;
begin
  if (ACol < sgTable.FixedCols) or (ACol >= sgTable.ColCount) then
    Exit;

  sgTable.HandleNeeded;
  sgTable.Canvas.Font.Assign(sgTable.Font);
  sgTable.Canvas.Font.Quality := fqDefault;

  HeaderPad := 12;
  CellPad := 6;
  MinW := 12;

  if AUseAllRows then
    LastRow := sgTable.RowCount - 1
  else
    LastRow := Min(sgTable.RowCount - 1, 3);

  sgTable.Canvas.Font.Style := [fsBold];
  W := Max(MinW, sgTable.Canvas.TextWidth(sgTable.Cells[ACol, 0]) + HeaderPad);

  sgTable.Canvas.Font.Style := [];
  if AUseAllRows then
  begin
    for Y := 1 to LastRow do
      W := Max(W, sgTable.Canvas.TextWidth(sgTable.Cells[ACol, Y]) + CellPad);
  end;

  sgTable.ColWidths[ACol] := W;
end;

procedure TTableFrame.AutoFitColumns(const AUseAllRows: Boolean);
var
  X: Integer;
begin
  for X := sgTable.FixedCols to sgTable.ColCount - 1 do
    AutoFitColumn(X, AUseAllRows);

  sgTable.HandleNeeded;
  sgTable.Canvas.Font.Assign(sgTable.Font);
  sgTable.Canvas.Font.Quality := fqDefault;
  sgTable.Canvas.Font.Style := [fsBold];

  sgTable.ColWidths[0] :=
    Max(20, sgTable.Canvas.TextWidth(IntToStr(Max(1, sgTable.RowCount - 1))) + 6);

  sgTable.Canvas.Font.Style := [];
end;

procedure TTableFrame.SyncHiddenArrays;
begin
  SetLength(FHiddenRows, sgTable.RowCount);
  SetLength(FHiddenCols, sgTable.ColCount);
end;

procedure TTableFrame.InsertRowState(AIndex, ACount: Integer);
var
  OldLen, i: Integer;
begin
  OldLen := Length(FHiddenRows);
  SetLength(FHiddenRows, sgTable.RowCount);

  for i := OldLen - 1 downto AIndex do
    FHiddenRows[i + ACount] := FHiddenRows[i];

  for i := AIndex to AIndex + ACount - 1 do
    FHiddenRows[i] := False;
end;

procedure TTableFrame.DeleteRowState(AIndex, ACount: Integer);
var
  OldLen, i: Integer;
begin
  OldLen := Length(FHiddenRows);
  for i := AIndex to OldLen - ACount - 1 do
    FHiddenRows[i] := FHiddenRows[i + ACount];

  SetLength(FHiddenRows, sgTable.RowCount);
end;

procedure TTableFrame.InsertColState(AIndex, ACount: Integer; ADefaultWidth: Integer);
var
  OldLen, i: Integer;
begin
  OldLen := Length(FBaseColWidths);
  SetLength(FBaseColWidths, sgTable.ColCount);
  SetLength(FHiddenCols, sgTable.ColCount);

  for i := OldLen - 1 downto AIndex do
  begin
    FBaseColWidths[i + ACount] := FBaseColWidths[i];
    FHiddenCols[i + ACount] := FHiddenCols[i];
  end;

  for i := AIndex to AIndex + ACount - 1 do
  begin
    FBaseColWidths[i] := ADefaultWidth;
    FHiddenCols[i] := False;
  end;
end;

procedure TTableFrame.DeleteColState(AIndex, ACount: Integer);
var
  OldLen, i: Integer;
begin
  OldLen := Length(FBaseColWidths);

  for i := AIndex to OldLen - ACount - 1 do
  begin
    FBaseColWidths[i] := FBaseColWidths[i + ACount];
    FHiddenCols[i] := FHiddenCols[i + ACount];
  end;

  SetLength(FBaseColWidths, sgTable.ColCount);
  SetLength(FHiddenCols, sgTable.ColCount);
end;

function TTableFrame.PromptForInteger(const ACaption, APrompt: string; ADefault: Integer; out AValue: Integer): Boolean;
var
  S: string;
begin
  S := IntToStr(ADefault);
  Result := InputQuery(ACaption, APrompt, S) and TryStrToInt(S, AValue) and (AValue > 0);
end;

function TTableFrame.PromptForFloat(const ACaption, APrompt: string; ADefault: Double; out AValue: Double): Boolean;
var
  S: string;
begin
  S := FloatToStr(ADefault);
  Result := InputQuery(ACaption, APrompt, S);
  if not Result then
    Exit(False);

  Result := TryStrToFloat(S, AValue);
  if not Result then
  begin
    S := StringReplace(S, '.', FormatSettings.DecimalSeparator, [rfReplaceAll]);
    Result := TryStrToFloat(S, AValue);
  end;
end;

function TTableFrame.TryParseCellFloat(const S: string; out AValue: Double): Boolean;
var
  T: string;
begin
  Result := TryStrToFloat(S, AValue);
  if Result then
    Exit;

  T := StringReplace(S, '.', FormatSettings.DecimalSeparator, [rfReplaceAll]);
  Result := TryStrToFloat(T, AValue);
end;

function TTableFrame.FormatCellNumber(const AValue: Double): string;
begin
  if Frac(AValue) = 0 then
    Result := IntToStr(Round(AValue))
  else
    Result := FloatToStr(AValue);
end;

procedure TTableFrame.ApplyMathToSelection(const AOperand: Double; const AOp: Char);
var
  X, Y: Integer;
  V: Double;
begin
  PrepareHistoryAction('Math');
  CreateUndo(sgTable.Selection);

  for Y := sgTable.Selection.Top to sgTable.Selection.Bottom do
  begin
    for X := sgTable.Selection.Left to sgTable.Selection.Right do
    begin
      if X < sgTable.FixedCols then
        Continue;

      if not TryParseCellFloat(sgTable.Cells[X, Y], V) then
        Continue;

      case AOp of
        '*': V := V * AOperand;
        '/': if AOperand <> 0 then V := V / AOperand else Continue;
        '+': V := V + AOperand;
        '-': V := V - AOperand;
      end;

      sgTable.Cells[X, Y] := FormatCellNumber(V);
    end;
  end;

  SetModified(True);
end;

procedure TTableFrame.ZoomIn;
begin
  FZoomPercent := FZoomPercent + 10;
  ApplyZoom;
end;

procedure TTableFrame.ZoomOut;
begin
  FZoomPercent := Max(FZoomPercent - 10, 50);
  ApplyZoom;
end;

procedure TTableFrame.ZoomReset;
begin
  FZoomPercent := 100;
  ApplyZoom;
end;

function TTableFrame.HandleZoomShortcut(AKey: Word; Shift: TShiftState): Boolean;
begin
  Result := False;
  if not (ssCtrl in Shift) then
    Exit;

  case AKey of
    VK_OEM_PLUS, VK_ADD:
      begin
        ZoomIn;
        Result := True;
      end;
    VK_OEM_MINUS, VK_SUBTRACT:
      begin
        ZoomOut;
        Result := True;
      end;
    Ord('0'), VK_NUMPAD0:
      begin
        ZoomReset;
        Result := True;
      end;
  end;
end;

procedure TTableFrame.cbFixColumnsClick(Sender: TObject);
var selection: TGridRect;
    leftCol, topRow: Integer;
begin
  seFixedColumns.Enabled := cbFixColumns.Checked;

  selection := sgTable.Selection;
  leftCol := sgTable.LeftCol;
  topRow := sgTable.TopRow;
  try
    if cbFixColumns.Checked then
    begin
      sgTable.FixedCols := seFixedColumns.Value + 1
    end
    else
    begin
      if leftCol = sgTable.FixedCols then
        leftCol := 1;
      sgTable.FixedCols := 1;
    end;
  finally
    selection.Left := Max(selection.Left, sgTable.FixedCols);
    selection.Right := Max(selection.Right, sgTable.FixedCols);
    sgTable.Selection := selection;

    sgTable.LeftCol := Max(leftCol, sgTable.FixedCols);
    sgTable.TopRow := topRow;
  end;
end;

procedure TTableFrame.cbFixRowsClick(Sender: TObject);
var
  Selection: TGridRect;
  LeftCol, TopRow: Integer;
begin
  Selection := sgTable.Selection;
  LeftCol := sgTable.LeftCol;
  TopRow := sgTable.TopRow;

  try
    if cbFixRows.Checked then
      sgTable.FixedRows := 1
    else
    begin
      if TopRow = sgTable.FixedRows then
        TopRow := 0;
      sgTable.FixedRows := 0;
    end;
  finally
    Selection.Top := Max(Selection.Top, sgTable.FixedRows);
    Selection.Bottom := Max(Selection.Bottom, sgTable.FixedRows);
    sgTable.Selection := Selection;

    sgTable.LeftCol := Max(LeftCol, sgTable.FixedCols);
    sgTable.TopRow := Max(TopRow, sgTable.FixedRows);
  end;
end;

procedure TTableFrame.ClearStack(AStack: TStack<TEditorState>);
var state: TEditorState;
begin
  while AStack.Count > 0 do
  begin
    state := AStack.Pop;
    state.Free;
  end;
end;

constructor TTableFrame.Create(AOwner: TComponent; AFilename: String);
begin
  inherited Create(AOwner);
  FUndoStack := TStack<TEditorState>.Create();
  FRedoStack := TStack<TEditorState>.Create();
  FHistoryEntries := TObjectList<THistoryEntry>.Create(True);
  FHistoryNodeRefs := TObjectList<THistoryNodeRef>.Create(True);
  FHistoryIndex := -1;
  FNextHistoryID := 1;
  FHistoryWidth := plHistory.Width;
  FHistoryVisible := True;
  FDarkMode := False;
  LoadHistorySettings;
  FZoomPercent := 100;
  FPopupCol := -1;
  FResizingCol := -1;
  sgTable.OnMouseUp := sgTableMouseUp;

  sgTable.Font.Name := 'Arial';
  sgTable.Font.Size := 8;
  sgTable.Font.Style := [];
  sgTable.Font.Quality := fqDefault;
  sgTable.DefaultRowHeight := 24;

  cbFixColumns.Checked := False;
  seFixedColumns.Enabled := False;
  sgTable.FixedCols := 1;

  cbFixRows.Checked := False;
  sgTable.FixedRows := 0;

  sgTable.DefaultDrawing := False;
  sgTable.Options := sgTable.Options - [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine];

  Name := 'TableFrame' + THashSHA2.GetHashString(AFilename);
  LoadFile(AFilename);
  FInitialSnapshot := CaptureGridSnapshot;
  ApplyHistoryVisibility;
  RefreshHistoryView;
  EnsureToolbarLabels;
  UpdateToolbarLayout;
  ApplyTheme;
  UpdateGridScrollbars;
end;

function TTableFrame.CreateEditorState(AArea: TGridRect): TEditorState;
var
  X, Y: Integer;
  Line: string;
begin
  Result := TEditorState.Create;
  Result.Area := AArea;
  Result.RowCount := sgTable.RowCount;
  Result.ColCount := sgTable.ColCount;

  for Y := Result.Area.Top to Result.Area.Bottom do
  begin
    Line := '';
    for X := Result.Area.Left to Result.Area.Right do
    begin
      if (X >= 0) and (X < sgTable.ColCount) and (Y >= 0) and (Y < sgTable.RowCount) then
        Line := Line + sgTable.Cells[X, Y]
      else
        Line := Line + '';

      Line := Line + #9;
    end;
    Result.Values.Add(Line);
  end;
end;

procedure TTableFrame.CreateUndo(AArea: TGridRect);
var
  State: TEditorState;
begin
  if FPendingHistoryActive then
    CancelPendingHistory(False);

  ClearStack(FRedoStack);

  State := CreateEditorState(AArea);
  FUndoStack.Push(State);

  FPendingHistoryActive := True;
  FPendingHistoryArea := AArea;
  FPendingBeforeSnapshot := CaptureGridSnapshot;
  FPendingBeforeState := State.Clone;

  if Assigned(FOnModifiedChanged) then
    FOnModifiedChanged(Tabsheet);
end;

procedure TTableFrame.CreateUndo;
begin
  CreateUndo(sgTable.Selection);
end;


procedure TTableFrame.DeleteCurrentRow;
var
  currRow: Integer;
  i: Integer;
begin
  currRow := sgTable.Selection.Top;
  for i := currRow to sgTable.RowCount - 2 do
  begin
    sgTable.Rows[i].Assign(sgTable.Rows[i + 1]);
    sgTable.Cells[0, i] := IntToStr(i);
  end;
  sgTable.RowCount := sgTable.RowCount - 1;

  DeleteRowState(currRow, 1);
  SetModified(True);
end;

destructor TTableFrame.Destroy;
begin
  SaveHistorySettings;
  CancelPendingHistory(False);
  FreeAndNil(FInitialSnapshot);
  FreeAndNil(FHistoryNodeRefs);
  FreeAndNil(FHistoryEntries);
  FreeAndNil(FUndoStack);
  FreeAndNil(FRedoStack);
  inherited;
end;

function TTableFrame.Find(AText: String; AStartPos: TPoint; AOptions: TFindOptions; out ResultPosition: TPoint): Boolean;
var y: Integer;
    row: String;
    p: Integer;
    tmpSL: TStringList;
    ok: Boolean;
begin
  Result := false;

  if not (frMatchCase in AOptions) then
    AText := AnsiLowerCase(AText);

  tmpSL := TStringList.Create;
  try
    if frDown in AOptions then
    begin
      for y := AStartPos.Y to sgTable.RowCount-1 do
      begin
        tmpSL.Assign(sgTable.Rows[y]);
        tmpSL.Delete(0);

        row := StringReplace(tmpSL.Text, sLineBreak, #9, [rfReplaceAll]);
        if not (frMatchCase in AOptions) then
          row := AnsiLowerCase(row);

        p := pos(AText, row);
        while (p > 0) do
        begin
          ok := true;
          if frWholeWord in AOptions then
          begin
            if not (((p = 1) or ((p-1 > 1) and ((row[p-1] = ' ') or (row[p-1] = #9)))) and
               ((p+Length(AText) >= Length(row)) or ((p+Length(AText) <= Length(row)) and ((row[p+Length(AText)] = ' ') or (row[p+Length(AText)] = #9)))))
            then
              ok := false;
          end;
          
          if ok then
          begin
            tmpSL.Delimiter := #9;
            tmpSL.StrictDelimiter := true;
            tmpSL.DelimitedText := copy(row, 1, p);
            if (y = AStartPos.Y) then
            begin
              if tmpSL.Count > AStartPos.X then
              begin
                ResultPosition := Point(tmpSL.Count, y);
                Result := true;
                break;
              end;
            end
            else
            begin
              ResultPosition := Point(tmpSL.Count, y);
              Result := true;
              break;
            end;
          end;

          p := posEx(Atext, row, p+1)
        end;

        if Result then
          break;
      end;
    end
    else
    begin
      AText := ReverseString(AText);
      for y := AStartPos.Y downto 0 do
      begin
        tmpSL.Assign(sgTable.Rows[y]);
        tmpSL.Delete(0);

        row := ReverseString(StringReplace(tmpSL.Text, sLineBreak, #9, [rfReplaceAll]));
        if not (frMatchCase in AOptions) then
          row := AnsiLowerCase(row);

        p := pos(AText, row);
        while (p > 0) do
        begin
          ok := true;
          if frWholeWord in AOptions then
          begin
            if not (((p = 1) or ((p-1 > 1) and ((row[p-1] = ' ') or (row[p-1] = #9)))) and
               ((p+Length(AText) >= Length(row)) or ((p+Length(AText) <= Length(row)) and ((row[p+Length(AText)] = ' ') or (row[p+Length(AText)] = #9)))))
            then
              ok := false;
          end;

          if ok then
          begin
            tmpSL.Delimiter := #9;
            tmpSL.StrictDelimiter := true;
            tmpSL.DelimitedText := copy(row, 1, p);
            if (y = AStartPos.Y) then
            begin
              if (sgTable.ColCount - tmpSL.Count + 1) < AStartPos.X then
              begin
                ResultPosition := Point(sgTable.ColCount - tmpSL.Count + 1, y);
                Result := true;
                break;
              end;
            end
            else
            begin
              ResultPosition := Point(sgTable.ColCount - tmpSL.Count + 1, y);
              Result := true;
              break;
            end;
          end;

          p := posEx(Atext, row, p+1)
        end;

        if Result then
          break;
      end;
    end;
  finally
    tmpSL.Free;
  end;
end;

function TTableFrame.GetCanRedo: Boolean;
begin
  Result := FRedoStack.Count > 0;
end;

function TTableFrame.GetCanUndo: Boolean;
begin
  Result := FUndoStack.Count > 0;
end;

function TTableFrame.GetTabsheet: TTabsheet;
begin
  Result := TTabsheet(Parent);
end;


procedure TTableFrame.LoadFile(const AFilename: String);
var
  tabFile: TStringList;
  row: TArray<String>;
  i, j: Integer;
begin
  FFilename := AFilename;

  tabFile := TStringList.Create;
  try
    tabFile.LoadFromFile(AFilename, TEncoding.ANSI);

    if tabFile.Count > 0 then
    begin
      row := tabFile[0].Split([#9]);

      sgTable.RowCount := tabFile.Count;
      sgTable.ColCount := Length(row) + 1;
      sgTable.ColWidths[0] := 42;

      for i := 0 to tabFile.Count - 1 do
      begin
        row := tabFile[i].Split([#9]);

        if i > 0 then
          sgTable.Cells[0, i] := IntToStr(i);

        for j := 0 to High(row) do
          sgTable.Cells[j + 1, i] := row[j];
      end;
    end;

    for i := 0 to sgTable.RowCount - 1 do
      sgTable.RowHeights[i] := sgTable.DefaultRowHeight;

    AutoFitColumns(False);
    CaptureBaseMetrics;
    SyncHiddenArrays;
    ApplyZoom;

    FreeAndNil(FInitialSnapshot);
    FInitialSnapshot := CaptureGridSnapshot;
    FHistoryEntries.Clear;
    FHistoryIndex := -1;
    RefreshHistoryView;
  finally
    tabFile.Free;
  end;
end;

procedure TTableFrame.miGridAddCopyClick(Sender: TObject);
var
  currRow: Integer;
  NewRow: Integer;
begin
  PrepareHistoryAction('Clone Row');
  CreateUndo(GetWholeGridRect);
  currRow := sgTable.Selection.Top;
  sgTable.RowCount := sgTable.RowCount + 1;
  NewRow := sgTable.RowCount - 1;

  sgTable.Rows[NewRow].Assign(sgTable.Rows[currRow]);
  sgTable.Cells[0, NewRow] := IntToStr(NewRow);

  InsertRowState(NewRow, 1);
  sgTable.RowHeights[NewRow] := ScaleValue(FBaseRowHeight);

  SetModified(True);
end;

procedure TTableFrame.miGridAddNewClick(Sender: TObject);
var
  NewRow: Integer;
begin
  PrepareHistoryAction('Add Row');
  CreateUndo(GetWholeGridRect);
  sgTable.RowCount := sgTable.RowCount + 1;
  NewRow := sgTable.RowCount - 1;
  sgTable.Cells[0, NewRow] := IntToStr(NewRow);

  InsertRowState(NewRow, 1);
  sgTable.RowHeights[NewRow] := ScaleValue(FBaseRowHeight);

  SetModified(True);
end;

procedure TTableFrame.miGridDeleteRowClick(Sender: TObject);
begin
  PrepareHistoryAction('Delete Row');
  CreateUndo(GetWholeGridRect);
  DeleteCurrentRow;
end;

procedure TTableFrame.miGridInsertCopyClick(Sender: TObject);
var
  currRow: Integer;
  i: Integer;
begin
  PrepareHistoryAction('Insert Row Copy');
  CreateUndo(GetWholeGridRect);
  currRow := sgTable.Selection.Top;

  sgTable.RowCount := sgTable.RowCount + 1;
  for i := sgTable.RowCount - 1 downto currRow + 1 do
  begin
    sgTable.Rows[i].Assign(sgTable.Rows[i - 1]);
    sgTable.Cells[0, i] := IntToStr(i);
  end;

  sgTable.Rows[currRow].Assign(sgTable.Rows[currRow + 1]);
  sgTable.Cells[0, currRow] := IntToStr(currRow);

  InsertRowState(currRow, 1);
  sgTable.RowHeights[currRow] := ScaleValue(FBaseRowHeight);

  SetModified(True);
end;

procedure TTableFrame.miGridInsertNewClick(Sender: TObject);
var
  currRow: Integer;
  i: Integer;
begin
  PrepareHistoryAction('Insert Row');
  CreateUndo(GetWholeGridRect);
  currRow := sgTable.Selection.Top;

  sgTable.RowCount := sgTable.RowCount + 1;
  for i := sgTable.RowCount - 1 downto currRow + 1 do
  begin
    sgTable.Rows[i].Assign(sgTable.Rows[i - 1]);
    sgTable.Cells[0, i] := IntToStr(i);
  end;

  sgTable.Rows[currRow].Clear;
  sgTable.Cells[0, currRow] := IntToStr(currRow);

  InsertRowState(currRow, 1);
  sgTable.RowHeights[currRow] := ScaleValue(FBaseRowHeight);

  SetModified(True);
end;

procedure TTableFrame.miRowHideClick(Sender: TObject);
var
  Y: Integer;
begin
  for Y := Max(sgTable.Selection.Top, sgTable.FixedRows) to sgTable.Selection.Bottom do
    FHiddenRows[Y] := True;

  ApplyZoom;
end;

procedure TTableFrame.miColumnAddClick(Sender: TObject);
var
  Count, OldColCount, NewCol, Y: Integer;
  BaseWidth: Integer;
begin
  if not PromptForInteger('Add Columns', 'How many columns do you want to add?', 1, Count) then
    Exit;

  PrepareHistoryAction(Format('Add %d Column(s)', [Count]));
  CreateUndo(GetWholeGridRect);

  OldColCount := sgTable.ColCount;
  sgTable.ColCount := sgTable.ColCount + Count;

  BaseWidth := 80;
  InsertColState(OldColCount, Count, BaseWidth);

  for NewCol := OldColCount to sgTable.ColCount - 1 do
  begin
    sgTable.Cells[NewCol, 0] := 'Column' + IntToStr(NewCol);
    for Y := 1 to sgTable.RowCount - 1 do
      sgTable.Cells[NewCol, Y] := '';

    sgTable.ColWidths[NewCol] := ScaleValue(BaseWidth);
  end;

  SetModified(True);
  sgTable.Repaint;
end;

procedure TTableFrame.miColumnInsertClick(Sender: TObject);
var
  InsertAt, X, Y: Integer;
  BaseWidth: Integer;
begin
  PrepareHistoryAction('Insert Column');
  CreateUndo(GetWholeGridRect);
  InsertAt := Max(sgTable.Selection.Left, sgTable.FixedCols);
  BaseWidth := 80;

  sgTable.ColCount := sgTable.ColCount + 1;

  for X := sgTable.ColCount - 1 downto InsertAt + 1 do
    for Y := 0 to sgTable.RowCount - 1 do
      sgTable.Cells[X, Y] := sgTable.Cells[X - 1, Y];

  for Y := 0 to sgTable.RowCount - 1 do
    sgTable.Cells[InsertAt, Y] := '';

  sgTable.Cells[InsertAt, 0] := 'Column' + IntToStr(InsertAt);

  InsertColState(InsertAt, 1, BaseWidth);
  sgTable.ColWidths[InsertAt] := ScaleValue(BaseWidth);

  SetModified(True);
  sgTable.Repaint;
end;

procedure TTableFrame.miColumnHideClick(Sender: TObject);
var
  X: Integer;
begin
  for X := Max(sgTable.Selection.Left, sgTable.FixedCols) to sgTable.Selection.Right do
    FHiddenCols[X] := True;

  ApplyZoom;
end;

procedure TTableFrame.miColumnDeleteClick(Sender: TObject);
var
  StartCol, EndCol, DeleteCount, X, Y: Integer;
begin
  StartCol := Max(sgTable.Selection.Left, sgTable.FixedCols);
  EndCol := sgTable.Selection.Right;

  if EndCol < StartCol then
    Exit;

  DeleteCount := EndCol - StartCol + 1;
  if sgTable.ColCount - DeleteCount <= sgTable.FixedCols then
    Exit;

  PrepareHistoryAction('Delete Column(s)');
  CreateUndo(GetWholeGridRect);

  for X := StartCol to sgTable.ColCount - DeleteCount - 1 do
    for Y := 0 to sgTable.RowCount - 1 do
      sgTable.Cells[X, Y] := sgTable.Cells[X + DeleteCount, Y];

  sgTable.ColCount := sgTable.ColCount - DeleteCount;
  DeleteColState(StartCol, DeleteCount);

  SetModified(True);
  sgTable.Repaint;
end;

procedure TTableFrame.miResizeToFitClick(Sender: TObject);
var
  X: Integer;
begin
  AutoFitColumns(True);

  if Length(FBaseColWidths) <> sgTable.ColCount then
    SetLength(FBaseColWidths, sgTable.ColCount);

  for X := 0 to sgTable.ColCount - 1 do
    FBaseColWidths[X] := Max(1, MulDiv(sgTable.ColWidths[X], 100, FZoomPercent));

  sgTable.Repaint;
end;

procedure TTableFrame.miResizeToFitThisColumnClick(Sender: TObject);
begin
  if (FPopupCol < sgTable.FixedCols) or (FPopupCol >= sgTable.ColCount) then
    Exit;

  AutoFitColumn(FPopupCol, True);

  if Length(FBaseColWidths) <> sgTable.ColCount then
    SetLength(FBaseColWidths, sgTable.ColCount);

  FBaseColWidths[FPopupCol] := Max(1, MulDiv(sgTable.ColWidths[FPopupCol], 100, FZoomPercent));
  sgTable.Repaint;
end;

procedure TTableFrame.miUnhideAllClick(Sender: TObject);
var
  I: Integer;
begin
  for I := 0 to High(FHiddenRows) do
    FHiddenRows[I] := False;

  for I := 0 to High(FHiddenCols) do
    FHiddenCols[I] := False;

  ApplyZoom;
end;

procedure TTableFrame.miFillCellsClick(Sender: TObject);
var
  X, Y: Integer;
  V: string;
begin
  V := sgTable.Cells[sgTable.Selection.Left, sgTable.Selection.Top];
  PrepareHistoryAction('Fill Cells');
  CreateUndo(sgTable.Selection);

  for Y := sgTable.Selection.Top to sgTable.Selection.Bottom do
    for X := sgTable.Selection.Left to sgTable.Selection.Right do
      if X >= sgTable.FixedCols then
        sgTable.Cells[X, Y] := V;

  SetModified(True);
end;

procedure TTableFrame.miFillIncrementClick(Sender: TObject);
var
  X, Y: Integer;
  Seed, Prefix: string;
  StartValue, CurrentValue: Double;
  N: Integer;
begin
  Seed := Trim(sgTable.Cells[sgTable.Selection.Left, sgTable.Selection.Top]);
  if Seed = '' then
    Exit;

  PrepareHistoryAction('Increment Fill');
  CreateUndo(sgTable.Selection);

  if TryParseCellFloat(Seed, StartValue) then
  begin
    CurrentValue := StartValue;

    for Y := sgTable.Selection.Top to sgTable.Selection.Bottom do
    begin
      for X := sgTable.Selection.Left to sgTable.Selection.Right do
      begin
        if (X < sgTable.FixedCols) or (Y < sgTable.FixedRows) then
          Continue;

        sgTable.Cells[X, Y] := FormatCellNumber(CurrentValue);
        CurrentValue := CurrentValue + 1;
      end;
    end;
  end
  else
  begin
    Prefix := Seed;
    N := 1;

    for Y := sgTable.Selection.Top to sgTable.Selection.Bottom do
    begin
      for X := sgTable.Selection.Left to sgTable.Selection.Right do
      begin
        if (X < sgTable.FixedCols) or (Y < sgTable.FixedRows) then
          Continue;

        sgTable.Cells[X, Y] := Prefix + IntToStr(N);
        Inc(N);
      end;
    end;
  end;

  SetModified(True);
end;

procedure TTableFrame.miMathMultiplyClick(Sender: TObject);
var
  Operand: Double;
begin
  if PromptForFloat('Multiply', 'Multiply selected cells by:', 1, Operand) then
    ApplyMathToSelection(Operand, '*');
end;

procedure TTableFrame.miMathDivideClick(Sender: TObject);
var
  Operand: Double;
begin
  if PromptForFloat('Divide', 'Divide selected cells by:', 1, Operand) then
    ApplyMathToSelection(Operand, '/');
end;

procedure TTableFrame.miMathAddClick(Sender: TObject);
var
  Operand: Double;
begin
  if PromptForFloat('Add', 'Add this value to selected cells:', 1, Operand) then
    ApplyMathToSelection(Operand, '+');
end;

procedure TTableFrame.miMathSubtractClick(Sender: TObject);
var
  Operand: Double;
begin
  if PromptForFloat('Subtract', 'Subtract this value from selected cells:', 1, Operand) then
    ApplyMathToSelection(Operand, '-');
end;

procedure TTableFrame.ProcessEditorState(AFromStack,
  AToStack: TStack<TEditorState>);
var
  State, NewState: TEditorState;
  X, Y: Integer;
  LineValues: TStringList;
begin
  if AFromStack.Count > 0 then
  begin
    State := AFromStack.Pop;
    try
      NewState := CreateEditorState(State.Area);
      try
        sgTable.ColCount := Max(1, State.ColCount);
        sgTable.RowCount := Max(1, State.RowCount);

        LineValues := TStringList.Create;
        try
          LineValues.StrictDelimiter := True;
          LineValues.Delimiter := #9;

          for Y := State.Area.Top to State.Area.Bottom do
          begin
            if (Y < 0) or (Y >= sgTable.RowCount) then
              Continue;

            if (Y - State.Area.Top) >= State.Values.Count then
              Continue;

            LineValues.DelimitedText := State.Values[Y - State.Area.Top];
            for X := State.Area.Left to State.Area.Right do
            begin
              if (X < 0) or (X >= sgTable.ColCount) then
                Continue;

              if (X - State.Area.Left) < LineValues.Count then
                sgTable.Cells[X, Y] := LineValues[X - State.Area.Left]
              else
                sgTable.Cells[X, Y] := '';
            end;
          end;
        finally
          LineValues.Free;
        end;

        SyncHiddenArrays;
        if Length(FBaseColWidths) <> sgTable.ColCount then
        begin
          SetLength(FBaseColWidths, sgTable.ColCount);
          for X := 0 to sgTable.ColCount - 1 do
            if FBaseColWidths[X] <= 0 then
              FBaseColWidths[X] := Max(12, sgTable.ColWidths[X]);
        end;
        ApplyZoom;
      finally
        AToStack.Push(NewState);
      end;
    finally
      State.Free;
    end;

    if AFromStack = FUndoStack then
      Dec(FHistoryIndex)
    else if AFromStack = FRedoStack then
      Inc(FHistoryIndex);

    if FHistoryEntries.Count = 0 then
      FHistoryIndex := -1
    else
      FHistoryIndex := EnsureRange(FHistoryIndex, -1, FHistoryEntries.Count - 1);

    RefreshHistoryView;
    SetModified(FHistoryIndex >= 0);

    if Assigned(FOnModifiedChanged) then
      FOnModifiedChanged(Tabsheet);
  end;
end;

procedure TTableFrame.Redo;
begin
  ProcessEditorState(FRedoStack, FUndoStack);
end;

procedure TTableFrame.SaveFile;
var tabFile: TStringList;
    line: String;
    y,x: Integer;
begin
  tabFile := TStringList.Create;
  try
    for y := 0 to sgTable.RowCount-1 do
    begin
      line := '';
      for x := 1 to sgTable.ColCount-1 do
        line := line + sgTable.Cells[x,y] + #9;
      SetLength(line, Length(line)-1);
      tabFile.Add(line);
    end;

    tabFile.SaveToFile(FFilename, TEncoding.ANSI);
    SetModified(false);
  finally
    tabFile.Free;
  end;
end;

procedure TTableFrame.Select(ACol, ARow: Integer; AScrollTo: Boolean);
var sel: TGridRect;
    tmpLen: Integer;
    fixedWidth: Integer;
    visibleRowCount: Integer;
    i: Integer;
begin
  sel.Left := ACol;
  sel.Right := ACol;
  sel.Top := ARow;
  sel.Bottom := ARow;
  sgTable.Selection := sel;

  if AScrollTo then
  begin
    fixedWidth := 0;
    for i := 0 to sgTable.FixedCols-1 do
      fixedWidth := fixedWidth + sgTable.ColWidths[i];

    tmpLen := sgTable.ColWidths[ACol];
    while (tmpLen < (sgTable.Width-fixedWidth)) and (ACol >= sgTable.FixedCols) do
    begin
      dec(ACol);
      tmpLen := tmpLen + sgTable.ColWidths[ACol];
    end;
    inc(ACol);
    sgTable.LeftCol := Max(ACol, sgTable.FixedCols);

    visibleRowCount := (sgTable.ClientRect.Height div sgTable.RowHeights[0]);
    if sgTable.RowCount > visibleRowCount then
      visibleRowCount := ((sgTable.ClientRect.Height - 18) div sgTable.RowHeights[0]);

    if ARow < sgTable.TopRow then
      sgTable.TopRow := ARow
    else
    if ARow >= (sgTable.TopRow + visibleRowCount - sgTable.FixedRows) then
      sgTable.TopRow := ARow - visibleRowCount + sgTable.FixedRows + 1;
  end;
end;

procedure TTableFrame.SetModified(AModified: Boolean);
begin
  if AModified and FPendingHistoryActive and (not FApplyingHistorySnapshot) then
    CommitPendingHistory;

  if AModified <> FModified then
  begin
    FModified := AModified;
    if Modified then
      Tabsheet.Caption := Tabsheet.Caption + '*'
    else
      Tabsheet.Caption := String(Tabsheet.Caption).Trim(['*']);

    if Assigned(OnModifiedChanged) then
      OnModifiedChanged(Tabsheet);
  end;
end;

procedure TTableFrame.sgTableContextPopup(Sender: TObject; MousePos: TPoint;
  var Handled: Boolean);
var
  Row, Col: Integer;
  Sel: TGridRect;
begin
  sgTable.MouseToCell(MousePos.X, MousePos.Y, Col, Row);
  FPopupCol := -1;

  if (Col <> -1) and (Row <> -1) then
  begin
    Sel := sgTable.Selection;

    if (Col < sgTable.FixedCols) and (Row >= sgTable.FixedRows) then
    begin
      if not ((Row >= Sel.Top) and (Row <= Sel.Bottom)) then
      begin
        Sel.Left := sgTable.FixedCols;
        Sel.Right := sgTable.ColCount - 1;
        Sel.Top := Row;
        Sel.Bottom := Row;
        sgTable.Selection := Sel;
      end;
    end
    else if (Row < sgTable.FixedRows) and (Col >= sgTable.FixedCols) then
    begin
      if not ((Col >= Sel.Left) and (Col <= Sel.Right)) then
      begin
        Sel.Left := Col;
        Sel.Right := Col;
        Sel.Top := sgTable.FixedRows;
        Sel.Bottom := sgTable.RowCount - 1;
        sgTable.Selection := Sel;
      end;
      FPopupCol := Col;
    end
    else
    begin
      if not ((Col >= Sel.Left) and (Col <= Sel.Right) and
              (Row >= Sel.Top) and (Row <= Sel.Bottom)) then
        Select(Col, Row);

      if Col >= sgTable.FixedCols then
        FPopupCol := Col;
    end;
  end;

  miRowAdd.Enabled := True;
  miRowInsert.Enabled := True;
  miRowHide.Enabled := True;
  miRowDelete.Enabled := True;
  miRowClone.Enabled := True;

  miColumnAdd.Enabled := True;
  miColumnInsert.Enabled := True;
  miColumnHide.Enabled := True;
  miColumnDelete.Enabled := True;

  miResizeToFit.Enabled := True;
  miResizeToFitThisColumn.Enabled := (FPopupCol >= sgTable.FixedCols) and (FPopupCol < sgTable.ColCount);
  miUnhideAll.Enabled := True;

  miFillCells.Enabled := True;
  miFillIncrement.Enabled := True;

  miMathMultiply.Enabled := True;
  miMathDivide.Enabled := True;
  miMathAdd.Enabled := True;
  miMathSubtract.Enabled := True;

  Handled := False;
end;

procedure TTableFrame.sgTableDrawCell(Sender: TObject; ACol, ARow: Integer;
  Rect: TRect; State: TGridDrawState);
var
  TextRect: TRect;
  TextY: Integer;
  TextX: Integer;
  CellText: string;
  TextHeightPx: Integer;
  IsSelected: Boolean;
  IsFixedCell: Boolean;
  DefaultBorderColor: TColor;
  StrongBorderColor: TColor;
  SelectedBackColor: TColor;
  SelectedBorderColor: TColor;
begin
  if ((ARow >= 0) and (ARow < Length(FHiddenRows)) and FHiddenRows[ARow]) or
     ((ACol >= 0) and (ACol < Length(FHiddenCols)) and FHiddenCols[ACol]) then
    Exit;

  IsSelected :=
    (ACol >= sgTable.FixedCols) and
    (ACol >= sgTable.Selection.Left) and (ACol <= sgTable.Selection.Right) and
    (ARow >= sgTable.Selection.Top) and (ARow <= sgTable.Selection.Bottom);

  IsFixedCell := (gdFixed in State) or ((ACol = 0) and (ARow >= 0));

  if FDarkMode then
  begin
    DefaultBorderColor := RGB(70, 70, 74);
    StrongBorderColor := RGB(95, 95, 100);
    SelectedBackColor := RGB(0, 102, 204);
    SelectedBorderColor := RGB(64, 156, 255);
  end
  else
  begin
    DefaultBorderColor := RGB(180, 180, 180);
    StrongBorderColor := RGB(120, 120, 120);
    SelectedBackColor := RGB(0, 120, 215);
    SelectedBorderColor := RGB(0, 84, 180);
  end;

  sgTable.Canvas.Brush.Style := bsSolid;

  if IsSelected then
    sgTable.Canvas.Brush.Color := SelectedBackColor
  else if IsFixedCell then
  begin
    if FDarkMode then
      sgTable.Canvas.Brush.Color := RGB(56, 56, 58)
    else
      sgTable.Canvas.Brush.Color := clBtnFace;
  end
  else if (ARow mod 2 = 0) then
  begin
    if FDarkMode then
      sgTable.Canvas.Brush.Color := RGB(38, 38, 40)
    else
      sgTable.Canvas.Brush.Color := clAlternatingRow;
  end
  else
  begin
    if FDarkMode then
      sgTable.Canvas.Brush.Color := RGB(30, 30, 30)
    else
      sgTable.Canvas.Brush.Color := clWindow;
  end;

  sgTable.Canvas.FillRect(Rect);

  sgTable.Canvas.Font.Assign(sgTable.Font);
  sgTable.Canvas.Font.Quality := fqDefault;

  if IsSelected then
    sgTable.Canvas.Font.Color := clWhite
  else if FDarkMode then
    sgTable.Canvas.Font.Color := RGB(240, 240, 240)
  else
    sgTable.Canvas.Font.Color := clWindowText;

  if IsFixedCell then
    sgTable.Canvas.Font.Style := [fsBold]
  else
    sgTable.Canvas.Font.Style := [];

  CellText := sgTable.Cells[ACol, ARow];

  TextRect := Rect;
  if ACol = 0 then
  begin
    Inc(TextRect.Left, ScaleValue(2));
    Dec(TextRect.Right, ScaleValue(2));
    TextX := TextRect.Left + ((TextRect.Right - TextRect.Left) -
      sgTable.Canvas.TextWidth(CellText)) div 2;
  end
  else
  begin
    Inc(TextRect.Left, ScaleValue(6));
    Dec(TextRect.Right, ScaleValue(4));
    TextX := TextRect.Left;
  end;

  TextHeightPx := sgTable.Canvas.TextHeight('Wg');
  TextY := TextRect.Top + ((TextRect.Bottom - TextRect.Top) - TextHeightPx) div 2;

  if gdFixed in State then
    Dec(TextY, ScaleValue(1));

  sgTable.Canvas.TextRect(TextRect, TextX, TextY, CellText);

  sgTable.Canvas.Brush.Style := bsClear;
  sgTable.Canvas.Pen.Width := 1;

  if ACol = 0 then
  begin
    sgTable.Canvas.Pen.Color := DefaultBorderColor;
    sgTable.Canvas.MoveTo(Rect.Left, Rect.Top);
    sgTable.Canvas.LineTo(Rect.Left, Rect.Bottom);
  end;

  if ARow = 0 then
  begin
    sgTable.Canvas.Pen.Color := DefaultBorderColor;
    sgTable.Canvas.MoveTo(Rect.Left, Rect.Top);
    sgTable.Canvas.LineTo(Rect.Right, Rect.Top);
  end;

  if (ACol = 0) or (ARow = 0) then
    sgTable.Canvas.Pen.Color := StrongBorderColor
  else
    sgTable.Canvas.Pen.Color := DefaultBorderColor;

  sgTable.Canvas.MoveTo(Rect.Right - 1, Rect.Top);
  sgTable.Canvas.LineTo(Rect.Right - 1, Rect.Bottom);

  if (ACol = 0) or (ARow = 0) then
    sgTable.Canvas.Pen.Color := StrongBorderColor
  else
    sgTable.Canvas.Pen.Color := DefaultBorderColor;

  sgTable.Canvas.MoveTo(Rect.Left, Rect.Bottom - 1);
  sgTable.Canvas.LineTo(Rect.Right, Rect.Bottom - 1);

  sgTable.Canvas.Brush.Style := bsSolid;

  InflateRect(Rect, -1, -1);

  if IsSelected then
  begin
    sgTable.Canvas.Pen.Color := SelectedBorderColor;

    if (ACol = sgTable.Selection.Left) then
    begin
      sgTable.Canvas.MoveTo(Rect.Left, Rect.Top);
      sgTable.Canvas.LineTo(Rect.Left, Rect.Bottom);
    end;

    if (ACol = sgTable.Selection.Right) then
    begin
      sgTable.Canvas.MoveTo(Rect.Right - 1, Rect.Top);
      sgTable.Canvas.LineTo(Rect.Right - 1, Rect.Bottom);
    end;

    if (ARow = sgTable.Selection.Top) then
    begin
      sgTable.Canvas.MoveTo(Rect.Left, Rect.Top);
      sgTable.Canvas.LineTo(Rect.Right, Rect.Top);
    end;

    if (ARow = sgTable.Selection.Bottom) then
    begin
      sgTable.Canvas.MoveTo(Rect.Left, Rect.Bottom - 1);
      sgTable.Canvas.LineTo(Rect.Right, Rect.Bottom - 1);
    end;
  end;
end;

procedure TTableFrame.sgTableFixedCellClick(Sender: TObject; ACol, ARow: Integer);
var
  Gr: TGridRect;
begin
  if (ARow >= 0) and (ARow < sgTable.FixedRows) and (ACol >= sgTable.FixedCols) then
  begin
    if GetAsyncKeyState(VK_SHIFT) < 0 then
    begin
      Gr.Left := Min(ACol, sgTable.Selection.Left);
      Gr.Right := Max(ACol, sgTable.Selection.Right);
    end
    else
    begin
      Gr.Left := ACol;
      Gr.Right := ACol;
    end;

    Gr.Top := sgTable.FixedRows;
    Gr.Bottom := sgTable.RowCount - 1;
    sgTable.Selection := Gr;
    Exit;
  end;

  if (ACol >= 0) and (ACol < sgTable.FixedCols) and (ARow >= sgTable.FixedRows) then
  begin
    Gr.Left := sgTable.FixedCols;
    Gr.Right := sgTable.ColCount - 1;

    if GetAsyncKeyState(VK_SHIFT) < 0 then
    begin
      Gr.Top := Min(ARow, sgTable.Selection.Top);
      Gr.Bottom := Max(ARow, sgTable.Selection.Bottom);
    end
    else
    begin
      Gr.Top := ARow;
      Gr.Bottom := ARow;
    end;

    sgTable.Selection := Gr;
    Exit;
  end;
end;

procedure TTableFrame.sgTableGetEditText(Sender: TObject; ACol, ARow: Integer;
  var Value: string);
var
  CellArea: TGridRect;
begin
  if Value <> FOldValue then
    FOldValue := Value;

  if not FPendingHistoryActive then
  begin
    CellArea.Left := ACol;
    CellArea.Right := ACol;
    CellArea.Top := ARow;
    CellArea.Bottom := ARow;
    PrepareHistoryAction('Edit Cell');
    CreateUndo(CellArea);
  end;
end;

procedure TTableFrame.sgTableKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
var value, line: String;
    sl: TStringList;
    rowValues: TArray<String>;
    x,y: Integer;
    gr: TGridRect;
begin
  if Key = VK_CONTROL then
    Key := 0
  else
  if (Key = VK_DELETE) and (ssCtrl in Shift) then
  begin
    PrepareHistoryAction('Delete Row');
    CreateUndo(GetWholeGridRect);
    DeleteCurrentRow
  end
  else
  if (Key = VK_DELETE) and (Shift = []) and (sgTable.Cells[sgTable.Selection.Left, sgTable.Selection.Top] <> '') then
  begin
    PrepareHistoryAction('Clear Cells');
    CreateUndo();
    for y := sgTable.Selection.Top to sgTable.Selection.Bottom do
    begin
      for x := sgTable.Selection.Left to sgTable.Selection.Right do
        sgTable.Cells[x, y] := '';
    end;
    SetModified(true);
  end
  else  if ((Key = VK_OEM_PLUS) or (Key = VK_ADD)) and (ssCtrl in Shift) then
  begin
    ZoomIn;
    Key := 0;
  end
  else
  if ((Key = VK_OEM_MINUS) or (Key = VK_SUBTRACT)) and (ssCtrl in Shift) then
  begin
    ZoomOut;
    Key := 0;
  end
  else
  if ((Key = Ord('0')) or (Key = VK_NUMPAD0)) and (ssCtrl in Shift) then
  begin
    ZoomReset;
    Key := 0;
  end
  else
  if (Key = Ord('R')) and (ssCtrl in Shift) then
  begin
    FPopupCol := sgTable.Selection.Left;
    miResizeToFitThisColumnClick(miResizeToFitThisColumn);
    Key := 0;
  end
  else
  if (Key = VK_ESCAPE) and (sgTable.EditorMode) then
  begin
    sgTable.Cells[sgTable.Selection.Left, sgTable.Selection.Top] := FOldValue;
    CancelPendingHistory(True);
    sgTable.EditorMode := false;
  end
  else
  if (Key = VK_RETURN) and (sgTable.EditorMode) then
  begin
    value := sgTable.Cells[sgTable.Selection.Left, sgTable.Selection.Top];
    if Value <> FOldValue then
      SetModified(true)
    else
      CancelPendingHistory(True);
  end
  else
  if (Key = Word(VkKeyScan('c'))) and (ssCtrl in Shift) then
  begin
    sl := TStringList.Create;
    try
      for y := sgTable.Selection.Top to sgTable.Selection.Bottom do
      begin
        line := '';
        for x := sgTable.Selection.Left to sgTable.Selection.Right do
          line := line + sgTable.Cells[x, y] + #9;
        sl.Add(line.Trim([#9]))
      end;
      Clipboard.AsText := sl.Text.Trim([#13, #10, #9]);
    finally
      sl.Free;
    end;
  end
  else
  if (Key = Word(VkKeyScan('x'))) and (ssCtrl in Shift) then
  begin
    PrepareHistoryAction('Cut');
    CreateUndo;

    sl := TStringList.Create;
    try
      for y := sgTable.Selection.Top to sgTable.Selection.Bottom do
      begin
        line := '';
        for x := sgTable.Selection.Left to sgTable.Selection.Right do
        begin
          line := line + sgTable.Cells[x, y] + #9;
          sgTable.Cells[x, y] := '';
        end;
        sl.Add(line.Trim([#9]));
      end;
      Clipboard.AsText := sl.Text.Trim([#13, #10, #9]);
    finally
      sl.Free;
    end;

    SetModified(true);
  end
  else
  if (Key = Word(VkKeyScan('v'))) and (ssCtrl in Shift) then
  begin
    if Clipboard.HasFormat(CF_TEXT) then
    begin
      sl := TStringList.Create;
      try
        sl.Text := Clipboard.AsText;
        rowValues := sl[0].Split([#9]);

        if (((sgTable.Selection.Right - sgTable.Selection.Left)+1) mod Length(rowValues) = 0) and ((((sgTable.Selection.Bottom - sgTable.Selection.Top)+1) mod sl.Count) = 0) then
        begin
          PrepareHistoryAction('Paste');
          CreateUndo(sgTable.Selection);
          for y := sgTable.Selection.Top to sgTable.Selection.Bottom do
          begin
            line := sl[(y-sgTable.Selection.Top) mod sl.Count];
            rowValues := line.Split([#9]);
            for x := sgTable.Selection.Left to sgTable.Selection.Right do
            begin
              sgTable.Cells[x, y] := rowValues[(x-sgTable.Selection.Left) mod Length(rowValues)];
            end;
          end;
        end
        else
        begin
          gr.Left := sgTable.Selection.Left;
          gr.Top := sgTable.Selection.Top;
          gr.Right := gr.Left + Length(rowValues) - 1;
          gr.Bottom := gr.Top + sl.Count - 1;
          PrepareHistoryAction('Paste');
          CreateUndo(gr);

          for y := sgTable.Selection.Top to Min(sgTable.Selection.Top + sl.Count-1, sgTable.RowCount-1)  do
          begin
            line := sl[y-sgTable.Selection.Top];
            rowValues := line.Split([#9]);
            for x := sgTable.Selection.Left to Min(sgTable.Selection.Left + High(rowValues), sgTable.ColCount-1) do
              sgTable.Cells[x, y] := rowValues[x-sgTable.Selection.Left];
            sl.Add(line)
          end;
        end;
      finally
        sl.Free;
      end;

      SetModified(true);
    end;
  end;
end;

procedure TTableFrame.sgTableMouseDown(Sender: TObject; Button: TMouseButton;
  Shift: TShiftState; X, Y: Integer);
var
  Row, Col: Integer;
  Gr: TGridRect;
  HitCol: Integer;
begin
  HitCol := -1;
  if Button = mbLeft then
    HitCol := GetResizeHitCol(X, Y);

  if HitCol <> -1 then
  begin
    FResizingCol := HitCol;
    FResizeStartX := X;
    FResizeStartWidth := sgTable.ColWidths[HitCol];
    SetCapture(sgTable.Handle);
    sgTable.Cursor := crHSplit;
    Exit;
  end;

  sgTable.MouseToCell(X, Y, Col, Row);

  if (Col <> -1) and (Row <> -1) then
  begin
    Col := Max(Col, sgTable.FixedCols);
    Row := Max(Row, sgTable.FixedRows);
  end;

  if (ssShift in Shift) and (Col <> -1) and (Row <> -1) then
  begin
    Gr.Left := Min(sgTable.Selection.Left, Col);
    Gr.Right := Max(sgTable.Selection.Right, Col);
    Gr.Top := Min(sgTable.Selection.Top, Row);
    Gr.Bottom := Max(sgTable.Selection.Bottom, Row);
    sgTable.Selection := Gr;
    sgTable.Repaint;
    Exit;
  end;

  FSelStart := Point(Col, Row);
end;

procedure TTableFrame.sgTableMouseMove(Sender: TObject; Shift: TShiftState; X,
  Y: Integer);
var
  Col, Row: Integer;
  Gr: TGridRect;
  MaxRow, MaxCol: Boolean;
  HitCol: Integer;
  NewWidth: Integer;
begin
  if FResizingCol <> -1 then
  begin
    NewWidth := Max(12, FResizeStartWidth + (X - FResizeStartX));

    if sgTable.ColWidths[FResizingCol] <> NewWidth then
    begin
      sgTable.ColWidths[FResizingCol] := NewWidth;

      if Length(FBaseColWidths) <> sgTable.ColCount then
        SetLength(FBaseColWidths, sgTable.ColCount);

      FBaseColWidths[FResizingCol] := Max(1, MulDiv(NewWidth, 100, FZoomPercent));
    end;

    sgTable.Cursor := crHSplit;
    Exit;
  end;

  HitCol := GetResizeHitCol(X, Y);
  if HitCol <> -1 then
    sgTable.Cursor := crHSplit
  else
    sgTable.Cursor := crDefault;

  if (ssLeft in Shift) and (not sgTable.EditorMode) then
  begin
    MaxRow := sgTable.Selection.Bottom = sgTable.RowCount - 1;
    MaxCol := sgTable.Selection.Right = sgTable.ColCount - 1;

    sgTable.MouseToCell(X, Y, Col, Row);
    if (Col = -1) and (Row = -1) then
    begin
      if MaxRow then
      begin
        repeat
          Dec(Y);
          sgTable.MouseToCell(X, Y, Col, Row);
        until ((Col <> -1) and (Row <> -1)) or (Y < 0);
      end;

      if MaxCol then
      begin
        repeat
          Dec(X);
          sgTable.MouseToCell(X, Y, Col, Row);
        until ((Col <> -1) and (Row <> -1)) or (X < 0);
      end;
    end;

    if (X >= 0) and (Y >= 0) and (Col <> -1) and (Row <> -1) then
    begin
      Col := Max(Col, sgTable.FixedCols);
      Row := Max(Row, sgTable.FixedRows);

      Gr.Left := Min(Col, FSelStart.X);
      Gr.Right := Max(Col, FSelStart.X);
      Gr.Top := Min(Row, FSelStart.Y);
      Gr.Bottom := Max(Row, FSelStart.Y);

      sgTable.Selection := Gr;
      sgTable.Repaint;
    end;
  end;
end;

procedure TTableFrame.sgTableMouseUp(Sender: TObject; Button: TMouseButton;
  Shift: TShiftState; X, Y: Integer);
begin
  if FResizingCol <> -1 then
  begin
    FResizingCol := -1;
    ReleaseCapture;
    sgTable.Cursor := crDefault;
  end;
end;

procedure TTableFrame.sgTableMouseWheelDown(Sender: TObject; Shift: TShiftState;
  MousePos: TPoint; var Handled: Boolean);
var rowDisplayCount: Integer;
begin
  if ssShift in Shift then
    sgTable.LeftCol := sgTable.LeftCol + 1
  else
  begin
    rowDisplayCount := (sgTable.Height div sgTable.RowHeights[0]);
    if sgTable.RowCount > rowDisplayCount then
      rowDisplayCount := ((sgTable.Height - 18) div sgTable.RowHeights[0]);

    sgTable.TopRow := Min(sgTable.TopRow + 1, Max(sgTable.RowCount - rowDisplayCount + 3, 1));
  end;

  Handled := true;
end;

procedure TTableFrame.sgTableMouseWheelUp(Sender: TObject; Shift: TShiftState;
  MousePos: TPoint; var Handled: Boolean);
begin
  if ssShift in Shift then
    sgTable.LeftCol := Max(sgTable.LeftCol - 1, sgTable.FixedCols)
  else
    sgTable.TopRow := Max(sgTable.TopRow - 1, sgTable.FixedRows);

  Handled := true;
end;

procedure TTableFrame.sgTableSelectCell(Sender: TObject; ACol, ARow: Integer;
  var CanSelect: Boolean);
var
  Value: string;
  NewSelection: TGridRect;
begin
  if sgTable.EditorMode then
  begin
    Value := sgTable.Cells[sgTable.Selection.Right, sgTable.Selection.Bottom];
    if Value <> FOldValue then
      SetModified(True)
    else
      CancelPendingHistory(True);
  end
  else if GetAsyncKeyState(VK_SHIFT) < 0 then
  begin
    ACol := Max(ACol, sgTable.FixedCols);
    ARow := Max(ARow, sgTable.FixedRows);

    NewSelection.Left := Max(sgTable.Selection.Left, sgTable.FixedCols);
    NewSelection.Top := Max(sgTable.Selection.Top, sgTable.FixedRows);
    NewSelection.Right := ACol;
    NewSelection.Bottom := ARow;
    sgTable.Selection := NewSelection;
    sgTable.Repaint;
    CanSelect := False;
  end;
end;

procedure TTableFrame.Undo;
begin
  ProcessEditorState(FUndoStack, FRedoStack);
end;

{ TEditorState }

constructor TEditorState.Create;
begin
  FValues := TStringList.Create;
end;

destructor TEditorState.Destroy;
begin
  FreeAndNil(FValues);
  inherited;
end;

function TEditorState.Clone: TEditorState;
begin
  Result := TEditorState.Create;
  Result.Area := Area;
  Result.RowCount := RowCount;
  Result.ColCount := ColCount;
  Result.Values.Assign(Values);
end;

end.
