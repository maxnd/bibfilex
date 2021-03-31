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
    Cells[0, 0] := 'À';
    Cells[1, 0] := 'Á';
    Cells[2, 0] := 'Â';
    Cells[3, 0] := 'Ä';
    Cells[4, 0] := 'Ä';
    Cells[5, 0] := 'Å';
    Cells[6, 0] := 'Æ';
    Cells[0, 1] := 'È';
    Cells[1, 1] := 'É';
    Cells[2, 1] := 'Ê';
    Cells[3, 1] := 'Ë';
    Cells[0, 2] := 'Ì';
    Cells[1, 2] := 'Í';
    Cells[2, 2] := 'Î';
    Cells[3, 2] := 'Ï';
    Cells[0, 3] := 'Ò';
    Cells[1, 3] := 'Ó';
    Cells[2, 3] := 'Ô';
    Cells[3, 3] := 'Õ';
    Cells[4, 3] := 'Ö';
    Cells[0, 4] := 'Ù';
    Cells[1, 4] := 'Ú';
    Cells[2, 4] := 'Û';
    Cells[3, 4] := 'Ü';
    Cells[0, 5] := 'à';
    Cells[1, 5] := 'á';
    Cells[2, 5] := 'â';
    Cells[3, 5] := 'ã';
    Cells[4, 5] := 'ä';
    Cells[5, 5] := 'å';
    Cells[6, 5] := 'æ';
    Cells[0, 6] := 'è';
    Cells[1, 6] := 'é';
    Cells[2, 6] := 'ê';
    Cells[3, 6] := 'ë';
    Cells[0, 7] := 'ì';
    Cells[1, 7] := 'í';
    Cells[2, 7] := 'î';
    Cells[3, 7] := 'ï';
    Cells[0, 8] := 'ò';
    Cells[1, 8] := 'ó';
    Cells[2, 8] := 'ô';
    Cells[3, 8] := 'õ';
    Cells[4, 8] := 'ö';
    Cells[0, 9] := 'ù';
    Cells[1, 9] := 'ú';
    Cells[2, 9] := 'û';
    Cells[3, 9] := 'ü';
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
