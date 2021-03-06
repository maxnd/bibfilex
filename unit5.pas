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
unit Unit5;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, FileUtil, Forms, Controls, Graphics, Dialogs, Grids,
  ExtCtrls, StdCtrls, Clipbrd;

type

  { TfmChar }

  TfmChar = class(TForm)
    bnCharOK: TButton;
    bnCharCancel: TButton;
    pnBottom: TPanel;
    sgChar: TStringGrid;
    procedure bnCharOKClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure sgCharDblClick(Sender: TObject);
    procedure sgCharKeyDown(Sender: TObject; var Key: word; Shift: TShiftState);
    procedure sgCharKeyUp(Sender: TObject; var Key: word; Shift: TShiftState);
  private
    { private declarations }
  public
    { public declarations }
  end;

var
  fmChar: TfmChar;

implementation

{$R *.lfm}

{ TfmChar }

procedure TfmChar.FormCreate(Sender: TObject);
begin
  {$ifdef Win32}
  fmChar.Color := clWhite;
  {$endif}
  {$ifdef Darwin}
  fmChar.Color := clWhite;
  {$endif}
  with sgChar do
  begin
    Cells[0, 0] := '??';
    Cells[1, 0] := '??';
    Cells[2, 0] := '??';
    Cells[3, 0] := '??';
    Cells[4, 0] := '??';
    Cells[5, 0] := '??';
    Cells[6, 0] := '??';
    Cells[0, 1] := '??';
    Cells[1, 1] := '??';
    Cells[2, 1] := '??';
    Cells[3, 1] := '??';
    Cells[0, 2] := '??';
    Cells[1, 2] := '??';
    Cells[2, 2] := '??';
    Cells[3, 2] := '??';
    Cells[0, 3] := '??';
    Cells[1, 3] := '??';
    Cells[2, 3] := '??';
    Cells[3, 3] := '??';
    Cells[4, 3] := '??';
    Cells[0, 4] := '??';
    Cells[1, 4] := '??';
    Cells[2, 4] := '??';
    Cells[3, 4] := '??';
    Cells[0, 5] := '??';
    Cells[1, 5] := '??';
    Cells[2, 5] := '??';
    Cells[3, 5] := '??';
    Cells[4, 5] := '??';
    Cells[5, 5] := '??';
    Cells[6, 5] := '??';
    Cells[0, 6] := '??';
    Cells[1, 6] := '??';
    Cells[2, 6] := '??';
    Cells[3, 6] := '??';
    Cells[0, 7] := '??';
    Cells[1, 7] := '??';
    Cells[2, 7] := '??';
    Cells[3, 7] := '??';
    Cells[0, 8] := '??';
    Cells[1, 8] := '??';
    Cells[2, 8] := '??';
    Cells[3, 8] := '??';
    Cells[4, 8] := '??';
    Cells[0, 9] := '??';
    Cells[1, 9] := '??';
    Cells[2, 9] := '??';
    Cells[3, 9] := '??';
  end;
end;

procedure TfmChar.bnCharOKClick(Sender: TObject);
begin
  // OK button
  sgCharDblClick(nil);
end;

procedure TfmChar.sgCharKeyDown(Sender: TObject; var Key: word; Shift: TShiftState);
begin
  // Close on Esc
  if key = 27 then
    fmChar.ModalResult := mrCancel;
end;

procedure TfmChar.sgCharKeyUp(Sender: TObject; var Key: word; Shift: TShiftState);
begin
  // Do not insert what follows in KeyDown:
  // a Return message is sent to the component (grid, memo) in the main form
  // Copy the symbol and quit
  if key = 13 then
  begin
    key := 0;
    sgCharDblClick(nil);
  end;
end;

procedure TfmChar.sgCharDblClick(Sender: TObject);
begin
  // Quit with mrOK
  fmChar.ModalResult := mrOk;
end;

end.
