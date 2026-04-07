unit uTableFrame;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes,
  Vcl.Graphics, Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.Grids, Vcl.ComCtrls,
  Vcl.ExtCtrls, Vcl.StdCtrls, Vcl.Samples.Spin, Vcl.Menus, Clipbrd, Generics.Collections;

type
  TEditorState = class
  strict private
    FArea: TGridREct;
    FValues: TStrings;
  public
    constructor Create;
    destructor Destroy; override;
    property Area: TGridREct read FArea write FArea;
    property Values: TStrings read FValues;
  end;

  TTableFrame = class(TFrame)
    {$region 'Components'}
    sgTable: TStringGrid;
    Panel1: TPanel;
    cbFixColumns: TCheckBox;
    seFixedColumns: TSpinEdit;
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
    cbFixRows: TCheckBox;
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

    procedure DeleteCurrentRow;
    procedure SetModified(AModified: Boolean);
    function GetTabsheet: TTabsheet;
    function GetCanUndo: Boolean;
    function GetCanRedo: Boolean;
    function ScaleValue(AValue: Integer): Integer;
    procedure CaptureBaseMetrics;
    procedure ApplyZoom;
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
  Math, StrUtils, System.Types, System.IOUtils, System.Hash;

const
  clAlternatingRow = $00FAFAFA;

{ TTableFrame }

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

    for i := 0 to High(FBaseColWidths) do
    begin
      if (i < Length(FHiddenCols)) and FHiddenCols[i] then
        NewColWidth := 0
      else
        NewColWidth := ScaleValue(FBaseColWidths[i]);

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
    sgTable.Invalidate;
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
  FZoomPercent := 100;
  FPopupCol := -1;

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

  Name := 'TableFrame' + THashSHA2.GetHashString(AFilename);
  LoadFile(AFilename);
end;

function TTableFrame.CreateEditorState(AArea: TGridRect): TEditorState;
var x,y: Integer;
    line: String;
begin
  Result := TEditorState.Create;
  Result.Area := AArea;
  for y := Result.Area.Top to Result.Area.Bottom do
  begin
    line := '';
    for x := Result.Area.Left to Result.Area.Right do
      line := line + sgTable.Cells[x, y] + #9;
    Result.Values.Add(line)
  end;
end;

procedure TTableFrame.CreateUndo(AArea: TGridRect);
var state: TEditorState;
begin
  ClearStack(FRedoStack);

  state := CreateEditorState(AArea);
  FUndoStack.Push(state);
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
  finally
    tabFile.Free;
  end;
end;

procedure TTableFrame.miGridAddCopyClick(Sender: TObject);
var
  currRow: Integer;
  NewRow: Integer;
begin
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
  sgTable.RowCount := sgTable.RowCount + 1;
  NewRow := sgTable.RowCount - 1;
  sgTable.Cells[0, NewRow] := IntToStr(NewRow);

  InsertRowState(NewRow, 1);
  sgTable.RowHeights[NewRow] := ScaleValue(FBaseRowHeight);

  SetModified(True);
end;

procedure TTableFrame.miGridDeleteRowClick(Sender: TObject);
begin
  DeleteCurrentRow;
end;

procedure TTableFrame.miGridInsertCopyClick(Sender: TObject);
var
  currRow: Integer;
  i: Integer;
begin
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

  CreateUndo(sgTable.Selection);

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
var state, newState: TEditorState;
    x,y: Integer;
    lineValues: TStringList;
begin
  if AFromStack.Count > 0 then
  begin
    state := AFromStack.Pop;
    try
      newState := CreateEditorState(state.Area);
      try
        lineValues := TStringList.Create;
        try
          lineValues.StrictDelimiter := true;
          lineValues.Delimiter := #9;
          for y := state.Area.Top to state.Area.Bottom do
          begin
            lineValues.DelimitedText := state.Values[y-state.Area.Top];
            for x := state.Area.Left to state.Area.Right do
              sgTable.Cells[x,y] := lineValues[x-state.Area.Left];
          end;
        finally
          lineValues.Free;
        end;
      finally
        AToStack.Push(newState);
      end;
    finally
      state.Free;
    end;

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
  FPopupCol := Col;

  if (Col <> -1) and (Row <> -1) then
  begin
    Sel := sgTable.Selection;
    if not ((Col >= Sel.Left) and (Col <= Sel.Right) and
            (Row >= Sel.Top) and (Row <= Sel.Bottom)) then
      Select(Col, Row);
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
begin
  InflateRect(Rect, 1, 1);

  IsSelected :=
    (ACol >= sgTable.Selection.Left) and (ACol <= sgTable.Selection.Right) and
    (ARow >= sgTable.Selection.Top) and (ARow <= sgTable.Selection.Bottom);

  IsFixedCell := (gdFixed in State);

  sgTable.Canvas.Brush.Style := bsSolid;

  if IsSelected then
  begin
    sgTable.Canvas.Brush.Color := RGB(0, 120, 215);
  end
  else if IsFixedCell then
  begin
    sgTable.Canvas.Brush.Color := clBtnFace;
  end
  else if (ARow mod 2 = 0) then
  begin
    sgTable.Canvas.Brush.Color := clAlternatingRow;
  end
  else
  begin
    sgTable.Canvas.Brush.Color := clWindow;
  end;

  sgTable.Canvas.FillRect(Rect);

  sgTable.Canvas.Font.Assign(sgTable.Font);
  sgTable.Canvas.Font.Quality := fqDefault;

  if IsSelected then
    sgTable.Canvas.Font.Color := clWhite
  else
    sgTable.Canvas.Font.Color := clWindowText;

  if IsFixedCell or ((ACol = 0) and (ARow > 0)) then
    sgTable.Canvas.Font.Style := [fsBold]
  else
    sgTable.Canvas.Font.Style := [];

  CellText := sgTable.Cells[ACol, ARow];

  TextRect := Rect;
  if (ACol = 0) and (ARow > 0) then
  begin
    Inc(TextRect.Left, ScaleValue(2));
    Dec(TextRect.Right, ScaleValue(2));
    TextX := TextRect.Left + ((TextRect.Right - TextRect.Left) - sgTable.Canvas.TextWidth(CellText)) div 2;
  end
  else
  begin
    Inc(TextRect.Left, ScaleValue(6));
    Dec(TextRect.Right, ScaleValue(4));
    TextX := TextRect.Left;
  end;

  TextHeightPx := sgTable.Canvas.TextHeight('Wg');
  TextY := TextRect.Top + ((TextRect.Bottom - TextRect.Top) - TextHeightPx) div 2;

  if IsFixedCell then
    Dec(TextY, ScaleValue(1));

  sgTable.Canvas.TextRect(TextRect, TextX, TextY, CellText);

  sgTable.Canvas.Pen.Color := RGB(180, 180, 180); // 선도 약간 진하게
  sgTable.Canvas.Brush.Style := bsClear;
  sgTable.Canvas.Rectangle(Rect);
  sgTable.Canvas.Brush.Style := bsSolid;

  InflateRect(Rect, -1, -1);

  if IsSelected then
  begin
    sgTable.Canvas.Pen.Color := RGB(0, 84, 180);

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
begin
  if Value <> FOldValue then
    FOldValue := Value;
  CreateUndo();
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
    DeleteCurrentRow
  else
  if (Key = VK_DELETE) and (Shift = []) and (sgTable.Cells[sgTable.Selection.Left, sgTable.Selection.Top] <> '') then
  begin
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
    sgTable.EditorMode := false;
  end
  else
  if (Key = VK_RETURN) and (sgTable.EditorMode) then
  begin
    value := sgTable.Cells[sgTable.Selection.Left, sgTable.Selection.Top];
    if Value <> FOldValue then
      SetModified(true);
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
begin
  sgTable.MouseToCell(X, Y, Col, Row);

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
var col, row: Integer;
    gr: TGridRect;
    maxRow, maxCol: Boolean;
begin
  if (ssLeft in Shift) and (not sgTable.EditorMode) then
  begin
    maxRow := sgTable.Selection.Bottom = sgTable.RowCount-1;
    maxCol := sgTable.Selection.Right = sgTable.ColCount-1;

    sgTable.MouseToCell(X, Y, col, row);
    if (col = -1) and (row = -1) then
    begin
      if maxRow then
      begin
        repeat
          dec(Y);
          sgTable.MouseToCell(X, Y, col, row);
        until ((col <> -1) and (row <> -1)) or (Y < 0);
      end;

      if maxCol then
      begin
        repeat
          dec(X);
          sgTable.MouseToCell(X, Y, col, row);
        until ((col <> -1) and (row <> -1)) or (X < 0);
      end;
    end;

    if (X >= 0) and (Y >= 0) then
    begin
      gr.Left := Min(col, FSelStart.X);
      gr.Right := Max(col, FSelStart.X);
      gr.Top := Min(row, FSelStart.Y);
      gr.Bottom := Max(row, FSelStart.Y);

      sgTable.Selection := gr;
      sgTable.Repaint;
    end;
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
var value: String;
    newSelection: TGridRect;
begin
  if sgTable.EditorMode then
  begin
    value := sgTable.Cells[sgTable.Selection.Right, sgTable.Selection.Bottom];
    if Value <> FOldValue then
      SetModified(true)
    else
      FUndoStack.Pop.Free;
  end
  else if GetAsyncKeyState(VK_SHIFT) < 0 then
  begin
    newSelection.Left := sgTable.Selection.Left;
    newSelection.Top := sgTable.Selection.Top;
    newSelection.Right := ACol;
    newSelection.Bottom := ARow;
    sgTable.Selection := newSelection;
    sgTable.Repaint;
    CanSelect := false;
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

end.
