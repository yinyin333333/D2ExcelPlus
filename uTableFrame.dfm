object TableFrame: TTableFrame
  Left = 0
  Top = 0
  Width = 674
  Height = 526
  ParentShowHint = False
  ShowHint = False
  TabOrder = 0
  object sgTable: TStringGrid
    Left = 0
    Top = 31
    Width = 674
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
    TabOrder = 0
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
  object Panel1: TPanel
    Left = 0
    Top = 0
    Width = 674
    Height = 31
    Align = alTop
    BevelOuter = bvNone
    TabOrder = 1
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
      Height = 23
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
