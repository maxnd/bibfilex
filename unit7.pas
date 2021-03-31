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
unit Unit7;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, FileUtil, Forms, Controls, Graphics, Dialogs, ExtCtrls,
  StdCtrls, Grids, Sqlite3DS;

type

  { TfmCheckDoubles }

  TfmCheckDoubles = class(TForm)
    bnCheck: TButton;
    bnFiltersClose: TButton;
    pnCheckDoubles: TPanel;
    sgCheckDoubles: TStringGrid;
    procedure bnCheckClick(Sender: TObject);
    procedure bnFiltersCloseClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormKeyPress(Sender: TObject; var Key: char);
    procedure sgCheckDoublesDblClick(Sender: TObject);
    procedure sgCheckDoublesKeyPress(Sender: TObject; var Key: char);
    procedure sgCheckDoublesPrepareCanvas(Sender: TObject; aCol, aRow: integer;
      aState: TGridDrawState);
  private
    procedure CheckDoubles;
    procedure GoToItem;
    { private declarations }
  public
    { public declarations }
  end;

var
  fmCheckDoubles: TfmCheckDoubles;

implementation

{$R *.lfm}

uses Unit1;

{ TfmCheckDoubles }

procedure TfmCheckDoubles.FormCreate(Sender: TObject);
begin
  // Set color
  {$ifdef Win32}
  fmCheckDoubles.Color := clWhite;
  {$endif}
  {$ifdef Darwin}
  fmCheckDoubles.Color := clWhite;
  {$endif}
  // Look of the grid
  sgCheckDoubles.FocusRectVisible := False;
end;

procedure TfmCheckDoubles.bnCheckClick(Sender: TObject);
begin
  // Check doubles
  CheckDoubles;
end;

procedure TfmCheckDoubles.bnFiltersCloseClick(Sender: TObject);
begin
  // Close
  Close;
end;

procedure TfmCheckDoubles.FormKeyPress(Sender: TObject; var Key: char);
begin
  // Close
  if key = #27 then
  begin
    Close;
  end;
end;

procedure TfmCheckDoubles.sgCheckDoublesDblClick(Sender: TObject);
begin
  // Go to selected item
  GoToItem;
end;

procedure TfmCheckDoubles.sgCheckDoublesKeyPress(Sender: TObject; var Key: char);
begin
  // Go to selected item
  if key = #13 then
  begin
    GoToItem;
  end;
end;

procedure TfmCheckDoubles.sgCheckDoublesPrepareCanvas(Sender: TObject;
  aCol, aRow: integer; aState: TGridDrawState);
begin
  // Main item in bold
  if sgCheckDoubles.Cells[0, aRow] = '*' then
  begin
    sgCheckDoubles.Canvas.Font.Style := [fsBold];
  end;
end;

procedure TfmCheckDoubles.CheckDoubles;
var
  sqStartItem, sqCheckItem: TSqlite3Dataset;
  flQuoted: boolean;
  iRow: integer;
begin
  // Check double items
  try
    Screen.Cursor := crHourGlass;
    sqStartItem := TSqlite3Dataset.Create(Self);
    sqCheckItem := TSqlite3Dataset.Create(Self);
    with sqStartItem do
    begin
      FileName := fmMain.sqItems.FileName;
      TableName := 'Items';
      SQL := fmMain.sqItems.SQL;
      Open;
    end;
    with sqCheckItem do
    begin
      FileName := fmMain.sqItems.FileName;
      TableName := 'Items';
      SQL := fmMain.sqItems.SQL;
      Open;
    end;
    iRow := 0;
    sgCheckDoubles.RowCount := 1;
    Application.ProcessMessages;
    while not sqStartItem.EOF do
    begin
      flQuoted := False;
      sqCheckItem.RecNo := sqStartItem.RecNo;
      sqCheckItem.Next;
      while not sqCheckItem.EOF do
      begin
        // Same author(s), title, year
        if LowerCase(sqStartItem.FieldByName('author').AsString) =
          LowerCase(sqCheckItem.FieldByName('author').AsString) then
          if LowerCase(sqStartItem.FieldByName('title').AsString) =
            LowerCase(sqCheckItem.FieldByName('title').AsString) then
            if LowerCase(sqStartItem.FieldByName('year').AsString) =
              LowerCase(sqCheckItem.FieldByName('year').AsString) then
            begin
              if flQuoted = False then
              begin
                sgCheckDoubles.RowCount := sgCheckDoubles.RowCount + 1;
                Inc(iRow);
                sgCheckDoubles.Cells[0, iRow] := '*';
                sgCheckDoubles.Cells[1, iRow] :=
                  sqStartItem.FieldByName('IDItems').AsString;
                sgCheckDoubles.Cells[2, iRow] :=
                  sqStartItem.FieldByName('author').AsString;
                sgCheckDoubles.Cells[3, iRow] :=
                  sqStartItem.FieldByName('title').AsString;
                sgCheckDoubles.Cells[4, iRow] :=
                  sqStartItem.FieldByName('year').AsString;
                flQuoted := True;
              end;
              sgCheckDoubles.RowCount := sgCheckDoubles.RowCount + 1;
              Inc(iRow);
              sgCheckDoubles.Cells[1, iRow] :=
                sqCheckItem.FieldByName('IDItems').AsString;
              sgCheckDoubles.Cells[2, iRow] :=
                sqCheckItem.FieldByName('author').AsString;
              sgCheckDoubles.Cells[3, iRow] :=
                sqCheckItem.FieldByName('title').AsString;
              sgCheckDoubles.Cells[4, iRow] :=
                sqCheckItem.FieldByName('year').AsString;
            end;
        sqCheckItem.Next;
        Application.ProcessMessages;
      end;
      sqStartItem.Next;
      Application.ProcessMessages;
    end;
    if iRow = 0 then
    begin
      MessageDlg('No double items were found.', mtInformation, [mbOK], 0);
    end;
  finally
    sqStartItem.Free;
    sqCheckItem.Free;
    Screen.Cursor := crDefault;
  end;
end;

procedure TfmCheckDoubles.GoToItem;
begin
  // Go to the selected item
  fmMain.sqItems.Locate('IDItems', sgCheckDoubles.Cells[1, sgCheckDoubles.Row], []);
  fmMain.Show;
end;

end.



