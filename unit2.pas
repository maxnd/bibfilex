// ***********************************************************************
// ***********************************************************************
// Bibfilex 1.2
// Author and copyright: Massimo Nardello, Modena (Italy) 2013-2016.
// Free software released under GPL licence vers. 3 or following.

// This program is free software: you can redistribute it and/or modify
// it under the terms of the GNU General Public License as published by
// the Free Software Foundation, either version 3 of the License, or
// (at your option) any later version. You can read the version 3
// of the Licence in http://www.gnu.org/licenses/gpl-3.0.txt
// or in the file Licence.txt included in the files of the
// source code of this software.

// This program is distributed in the hope that it will be useful,
// but WITHOUT ANY WARRANTY; without even the implied warranty of
// MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the
// GNU General Public License for more details.
// ***********************************************************************
// ***********************************************************************
unit Unit2;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, FileUtil, Forms, Controls, Graphics, Dialogs, StdCtrls,
  ExtCtrls, DB, LCLProc, LCLType;

type

  { TfmFilters }

  TfmFilters = class(TForm)
    bnFiltersApply: TButton;
    bnFiltersClear: TButton;
    bnFiltersRemove: TButton;
    bnFiltersClose: TButton;
    cbFilterCond1: TComboBox;
    cbFilterCond2: TComboBox;
    cbFilterCond3: TComboBox;
    cbFilterField1: TComboBox;
    cbFilterField2: TComboBox;
    cbFilterField3: TComboBox;
    edFilterValue1: TEdit;
    edFilterValue2: TEdit;
    edFilterValue3: TEdit;
    lbFilterCond1: TLabel;
    lbFilterCond2: TLabel;
    lbFilterCond3: TLabel;
    lbFilterField1: TLabel;
    lbFilterField2: TLabel;
    lbFilterField3: TLabel;
    lbFilterValue1: TLabel;
    lbFilterValue2: TLabel;
    lbFilterValue3: TLabel;
    meFilterWhere: TMemo;
    rgFilterCond: TRadioGroup;
    rgFilterUser: TRadioGroup;
    procedure bnFiltersApplyClick(Sender: TObject);
    procedure bnFiltersClearClick(Sender: TObject);
    procedure bnFiltersCloseClick(Sender: TObject);
    procedure bnFiltersRemoveClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure MoveFocus(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure rgFilterUserClick(Sender: TObject);
    procedure CompileOnFilterFieldsChange(Sender: TObject);
  private
    procedure CompileFilterBox;
    function GetRealConditionName(IDComboField: integer): string;
    { private declarations }
  public
    { public declarations }
  end;

var
  fmFilters: TfmFilters;

implementation

uses Unit1;

{$R *.lfm}

{ TfmFilters }

procedure TfmFilters.bnFiltersApplyClick(Sender: TObject);
begin
  // Set filter
  if meFilterWhere.Text = '' then
  begin
    MessageDlg('No condition has been specified.', mtWarning, [mbOK], 0);
    Abort;
  end;
  if ((cbFilterField1.Text <> '') and
    (cbFilterField1.Items.IndexOf(cbFilterField1.Text) < 0)) then
  begin
    MessageDlg('The first field name is wrong.', mtWarning, [mbOK], 0);
    Abort;
  end;
  if ((cbFilterField2.Text <> '') and
    (cbFilterField2.Items.IndexOf(cbFilterField2.Text) < 0)) then
  begin
    MessageDlg('The second field name is wrong.', mtWarning, [mbOK], 0);
    Abort;
  end;
  if ((cbFilterField3.Text <> '') and
    (cbFilterField3.Items.IndexOf(cbFilterField3.Text) < 0)) then
  begin
    MessageDlg('The third field name is wrong.', mtWarning, [mbOK], 0);
    Abort;
  end;
  // No use of " and ' together
  if ((edFilterValue1.Text <> '') and
    (UTF8Pos(#39, edFilterValue1.Text) > 0) and
    (UTF8Pos('"', edFilterValue1.Text) > 0)) then
  begin
    MessageDlg('The characters " and ' + #39 + ' cannot be searched together.',
      mtWarning, [mbOK], 0);
    Abort;
  end;
  if ((edFilterValue2.Text <> '') and
    (UTF8Pos(#39, edFilterValue2.Text) > 0) and
    (UTF8Pos('"', edFilterValue2.Text) > 0)) then
  begin
    MessageDlg('The characters " and ' + #39 + ' cannot be searched together.',
      mtWarning, [mbOK], 0);
    Abort;
  end;
  if ((edFilterValue3.Text <> '') and
    (UTF8Pos(#39, edFilterValue3.Text) > 0) and
    (UTF8Pos('"', edFilterValue3.Text) > 0)) then
  begin
    MessageDlg('The characters " and ' + #39 + ' cannot be searched together.',
      mtWarning, [mbOK], 0);
    Abort;
  end;
  Screen.Cursor := crHourGlass;
  Application.ProcessMessages;
  Unit1.flFilterActive := 2;
  fmMain.sqItems.Close;
  fmMain.sqItems.SQL := 'Select * from Items where ' + meFilterWhere.Text +
    fmMain.CreateOrderBy;
  try
    try
      fmMain.sqItems.Open;
    except
      Screen.Cursor := crDefault;
      MessageDlg('The filter is not correct; all the items' +
        LineEnding + 'will be selected.',
        mtWarning, [mbOK], 0);
      Screen.Cursor := crHourGlass;
      Application.ProcessMessages;
      Unit1.flFilterActive := 0;
      fmMain.sqItems.Close;
      fmMain.sqItems.SQL := 'Select * from Items ' + fmMain.CreateOrderBy;
      fmMain.sqItems.Open;
    end;
  finally
    Screen.Cursor := crDefault;
  end;
end;

procedure TfmFilters.bnFiltersClearClick(Sender: TObject);
begin
  // Clear filter
  cbFilterField1.ItemIndex := -1;
  cbFilterField2.ItemIndex := -1;
  cbFilterField3.ItemIndex := -1;
  cbFilterCond1.ItemIndex := -1;
  cbFilterCond2.ItemIndex := -1;
  cbFilterCond3.ItemIndex := -1;
  edFilterValue1.Clear;
  edFilterValue2.Clear;
  edFilterValue3.Clear;
  meFilterWhere.Clear;
  meFilterWhere.ReadOnly := True;
  rgFilterUser.ItemIndex := 0;
end;

procedure TfmFilters.bnFiltersRemoveClick(Sender: TObject);
begin
  // Remove filter
  try
    Screen.Cursor := crHourGlass;
    Application.ProcessMessages;
    Unit1.flFilterActive := 0;
    fmMain.sqItems.Close;
    fmMain.sqItems.SQL := 'Select * from Items ' + fmMain.CreateOrderBy;
    fmMain.sqItems.Open;
  finally
    Screen.Cursor := crDefault;
  end;
end;

procedure TfmFilters.FormCreate(Sender: TObject);
begin
  // Set color
  {$ifdef Win32}
  fmFilters.Color := clWhite;
  {$endif}
  {$ifdef Darwin}
  fmFilters.Color := clWhite;
  {$endif}
end;

procedure TfmFilters.FormKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  // Close on ESC
  if key = 27 then
    Close;
end;

procedure TfmFilters.MoveFocus(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  // Move focus with Return or Tab
  if ((Key = VK_TAB) and (Shift = [ssShift])) then
  begin
    Key := 0;
    SelectNext(ActiveControl,False,True);
  end
  else
  if ((Key = VK_RETURN) or (Key = VK_TAB)) then
  begin
    Key := 0;
    SelectNext(ActiveControl,True,True);
  end;
end;

procedure TfmFilters.bnFiltersCloseClick(Sender: TObject);
begin
  // Close
  Close;
end;

procedure TfmFilters.CompileOnFilterFieldsChange(Sender: TObject);
begin
  // Compile the memo filter
  CompileFilterBox;
end;

//***************************************************************
//********************** COMMON PROCEDURES **********************
//***************************************************************

procedure TfmFilters.rgFilterUserClick(Sender: TObject);
begin
  // Set default or user defined filter
  meFilterWhere.ReadOnly := rgFilterUser.ItemIndex = 0;
  if rgFilterUser.ItemIndex = 1 then
  begin
    meFilterWhere.Color := clWhite;
  end
  else
  begin
    meFilterWhere.Color := clForm;
  end;
end;

procedure TfmFilters.CompileFilterBox;
var
  mySQL, wildchst, wildchend, stDelimiter: string;
begin
  // Compile the filter memo
  // Delimiter for strings is " unless this character is contained
  // in the text to search
  if ((Pos('"', edFilterValue1.Text) > 0) or (Pos('"', edFilterValue2.Text) > 0) or
    (Pos('"', edFilterValue3.Text) > 0)) then
    stDelimiter := #39
  else
    stDelimiter := '"';
  // First condition
  if cbFilterField1.Items.IndexOf(cbFilterField1.Text) > -1 then
  begin
    wildchst := '';
    wildchend := '';
    if cbFilterCond1.ItemIndex > 2 then
      wildchend := '%';
    if cbFilterCond1.ItemIndex = 4 then
      wildchst := '%';
    if cbFilterCond1.Text <> '' then
    begin
      if ((fmMain.sqItems.FieldByName(cbFilterField1.Text).DataType = ftString) or
        (fmMain.sqItems.FieldByName(cbFilterField1.Text).DataType = ftMemo)) then
        mySQL := mySQL + 'Lower(' + cbFilterField1.Text + ')' +
          GetRealConditionName(cbFilterCond1.ItemIndex) + stDelimiter +
          wildchst + UTF8LowerCase(edFilterValue1.Text) +
          wildchend + stDelimiter + ' '
      else
      begin
        edFilterValue1.Text :=
          StringReplace(edFilterValue1.Text, ',', '.', [rfReplaceAll]);
        mySQL := mySQL + 'Lower(' + cbFilterField1.Text + ')' +
          GetRealConditionName(cbFilterCond1.ItemIndex) + wildchst +
          UTF8LowerCase(edFilterValue1.Text) + wildchend + ' ';
      end;
    end
    else
      mySQL := mySQL + cbFilterField1.Text;
  end;
  // Second condition
  if cbFilterField2.Items.IndexOf(cbFilterField2.Text) > -1 then
  begin
    wildchst := '';
    wildchend := '';
    if cbFilterCond2.ItemIndex > 2 then
      wildchend := '%';
    if cbFilterCond2.ItemIndex = 4 then
      wildchst := '%';
    if rgFilterCond.ItemIndex = 0 then
      mySQL := mySQL + ' AND '
    else
      mySQL := mySQL + ' OR ';
    if cbFilterCond2.Text <> '' then
    begin
      if ((fmMain.sqItems.FieldByName(cbFilterField2.Text).DataType = ftString) or
        (fmMain.sqItems.FieldByName(cbFilterField2.Text).DataType = ftMemo)) then
        mySQL := mySQL + 'Lower(' + cbFilterField2.Text + ')' +
          GetRealConditionName(cbFilterCond2.ItemIndex) + stDelimiter +
          wildchst + UTF8LowerCase(edFilterValue2.Text) + wildchend + stDelimiter + ' '
      else
      begin
        edFilterValue2.Text :=
          StringReplace(edFilterValue2.Text, ',', '.', [rfReplaceAll]);
        mySQL := mySQL + cbFilterField2.Text +
          GetRealConditionName(cbFilterCond2.ItemIndex) + wildchst +
          UTF8LowerCase(edFilterValue2.Text) + wildchend + ' ';
      end;
    end
    else
      mySQL := mySQL + cbFilterField2.Text;
  end;
  // Third condition
  if cbFilterField3.Items.IndexOf(cbFilterField3.Text) > -1 then
  begin
    wildchst := '';
    wildchend := '';
    if cbFilterCond3.ItemIndex > 2 then
      wildchend := '%';
    if cbFilterCond3.ItemIndex = 4 then
      wildchst := '%';
    if rgFilterCond.ItemIndex = 0 then
      mySQL := mySQL + ' AND '
    else
      mySQL := mySQL + ' OR ';
    if cbFilterCond3.Text <> '' then
    begin
      if ((fmMain.sqItems.FieldByName(cbFilterField3.Text).DataType = ftString) or
        (fmMain.sqItems.FieldByName(cbFilterField3.Text).DataType = ftMemo)) then
        mySQL := mySQL + 'Lower(' + cbFilterField3.Text + ')' +
          GetRealConditionName(cbFilterCond3.ItemIndex) + stDelimiter +
          wildchst + UTF8LowerCase(edFilterValue3.Text) + wildchend + stDelimiter + ' '
      else
      begin
        edFilterValue3.Text :=
          StringReplace(edFilterValue3.Text, ',', '.', [rfReplaceAll]);
        mySQL := mySQL + cbFilterField3.Text +
          GetRealConditionName(cbFilterCond3.ItemIndex) + wildchst +
          UTF8LowerCase(edFilterValue3.Text) + wildchend + ' ';
      end;
    end
    else
      mySQL := mySQL + cbFilterField3.Text;
  end;
  meFilterWhere.Text := mySQL;
end;

function TfmFilters.GetRealConditionName(IDComboField: integer): string;
begin
  Result := '';
  case IDComboField of
    0: Result := ' = ';
    1: Result := ' >= ';
    2: Result := ' <= ';
    3: Result := ' like ';
    4: Result := ' like ';
  end;
end;

end.
