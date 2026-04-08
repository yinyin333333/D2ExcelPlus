object TableFrame: TTableFrame
  Left = 0
  Top = 0
  Width = 674
  Height = 526
  ParentShowHint = False
  ShowHint = False
  TabOrder = 0
  object Panel1: TPanel
    Left = 0
    Top = 0
    Width = 594
    Height = 31
    Align = alTop
    BevelOuter = bvNone
    TabOrder = 0
    object sbToggleHistory: TSpeedButton
      AlignWithMargins = True
      Left = 499
      Top = 3
      Width = 92
      Height = 25
      Align = alRight
      Caption = 'Hide History'
      Flat = True
      OnClick = sbToggleHistoryClick
      ExplicitLeft = 499
    end
    object cbFixColumns: TCheckBox
      Left = 8
      Top = 6
      Width = 177
      Height = 17
      Caption = 'Lock first              Columns'
      Checked = True
      State = cbChecked
      TabOrder = 0
      OnClick = cbFixColumnsClick
    end
    object seFixedColumns: TSpinEdit
      Left = 72
      Top = 4
      Width = 33
      Height = 24
      Enabled = False
      MaxValue = 9
      MinValue = 1
      TabOrder = 1
      Value = 1
      OnChange = cbFixColumnsClick
    end
    object cbFixRows: TCheckBox
      Left = 191
      Top = 8
      Width = 97
      Height = 17
      Caption = 'Lock first row'
      Checked = True
      State = cbChecked
      TabOrder = 2
      OnClick = cbFixRowsClick
    end
  end
  object plHistory: TPanel
    Left = 434
    Top = 31
    Width = 240
    Height = 495
    Align = alRight
    BevelOuter = bvNone
    TabOrder = 1
    object pnlHistoryHeader: TPanel
      Left = 0
      Top = 0
      Width = 240
      Height = 30
      Align = alTop
      BevelOuter = bvNone
      Caption = ''
      TabOrder = 0
      object lblHistoryTitle: TLabel
        Left = 8
        Top = 8
        Width = 42
        Height = 13
        Caption = 'History'
      end
    end
    object tvHistory: TListBox
      AlignWithMargins = True
      Left = 3
      Top = 33
      Width = 237
      Height = 459
      Margins.Right = 0
      Align = alClient
      BorderStyle = bsNone
      ExtendedSelect = False
      IntegralHeight = False
      ItemHeight = 18
      Style = lbOwnerDrawVariable
      TabOrder = 1
      OnDblClick = tvHistoryDblClick
      OnDrawItem = tvHistoryDrawItem
      OnMeasureItem = tvHistoryMeasureItem
    end
  end
  object spHistory: TSplitter
    Left = 431
    Top = 31
    Width = 3
    Height = 495
    Align = alRight
    ExplicitLeft = 431
    ExplicitTop = 31
    ExplicitHeight = 495
  end
  object sgTable: TStringGrid
    Left = 0
    Top = 31
    Width = 351
    Height = 495
    Align = alClient
    DefaultColWidth = 10
    DefaultDrawing = False
    DoubleBuffered = True
    DrawingStyle = gdsGradient
    RowCount = 3
    GradientEndColor = 15790320
    GradientStartColor = 15790320
    Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goRangeSelect, goDrawFocusSelected, goColSizing, goEditing, goThumbTracking, goFixedColClick, goFixedRowClick]
    ParentDoubleBuffered = False
    PopupMenu = pmGrid
    TabOrder = 2
    OnContextPopup = sgTableContextPopup
    OnDrawCell = sgTableDrawCell
    OnFixedCellClick = sgTableFixedCellClick
    OnGetEditText = sgTableGetEditText
    OnKeyDown = sgTableKeyDown
    OnMouseDown = sgTableMouseDown
    OnMouseMove = sgTableMouseMove
    OnMouseWheelDown = sgTableMouseWheelDown
    OnMouseWheelUp = sgTableMouseWheelUp
    OnSelectCell = sgTableSelectCell
  end
  object pmGrid: TPopupMenu
    Left = 392
    Top = 160
    object miColumnOps: TMenuItem
      Caption = 'Column Operations'
      object miColumnAdd: TMenuItem
        Caption = 'Add Columns...'
        OnClick = miColumnAddClick
      end
      object miColumnInsert: TMenuItem
        Caption = 'Insert Column'
        OnClick = miColumnInsertClick
      end
      object miColumnHide: TMenuItem
        Caption = 'Hide Column(s)'
        OnClick = miColumnHideClick
      end
      object miColumnDelete: TMenuItem
        Caption = 'Delete Column(s)'
        OnClick = miColumnDeleteClick
      end
    end
    object miRowOps: TMenuItem
      Caption = 'Row Operations'
      object miRowAdd: TMenuItem
        Caption = 'Add Rows...'
        OnClick = miGridAddNewClick
      end
      object miRowInsert: TMenuItem
        Caption = 'Insert Row'
        OnClick = miGridInsertNewClick
      end
      object miRowHide: TMenuItem
        Caption = 'Hide Row(s)'
        OnClick = miRowHideClick
      end
      object miRowDelete: TMenuItem
        Caption = 'Delete Row(s)'
        OnClick = miGridDeleteRowClick
      end
      object miRowClone: TMenuItem
        Caption = 'Clone Row...'
        OnClick = miGridAddCopyClick
      end
    end
    object miResizeToFit: TMenuItem
      Caption = 'Resize To Fit'
      OnClick = miResizeToFitClick
    end
    object miResizeToFitThisColumn: TMenuItem
      Caption = 'Resize To Fit This Column'
      OnClick = miResizeToFitThisColumnClick
    end
    object miUnhideAll: TMenuItem
      Caption = 'Unhide All'
      OnClick = miUnhideAllClick
    end
    object miFill: TMenuItem
      Caption = 'Fill'
      object miFillCells: TMenuItem
        Caption = 'Fill Cells'
        OnClick = miFillCellsClick
      end
      object miFillIncrement: TMenuItem
        Caption = 'Increment Fill'
        OnClick = miFillIncrementClick
      end
    end
    object miMath: TMenuItem
      Caption = 'Math'
      object miMathMultiply: TMenuItem
        Caption = '*   Multiply'
        OnClick = miMathMultiplyClick
      end
      object miMathDivide: TMenuItem
        Caption = '\   Divide'
        OnClick = miMathDivideClick
      end
      object miMathAdd: TMenuItem
        Caption = '+    Add'
        OnClick = miMathAddClick
      end
      object miMathSubtract: TMenuItem
        Caption = '-   Subtract'
        OnClick = miMathSubtractClick
      end
    end
  end
end
