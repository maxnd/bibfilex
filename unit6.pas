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
unit Unit6;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, FileUtil, Forms, Controls, Graphics, Dialogs, StdCtrls,
  ExtCtrls, LCLProc, Clipbrd, Grids;

type

  { TfmKeywords }

  TfmKeywords = class(TForm)
    bnCharOK: TButton;
    bnCharInsert: TButton;
    edFind: TEdit;
    lbKeyFind: TLabel;
    lbKeyword: TListBox;
    pnBottom: TPanel;
    procedure bnCharInsertClick(Sender: TObject);
    procedure bnCharOKClick(Sender: TObject);
    procedure edFindChange(Sender: TObject);
    procedure edFindKeyDown(Sender: TObject; var Key: word; Shift: TShiftState);
    procedure FormCreate(Sender: TObject);
    procedure FormKeyDown(Sender: TObject; var Key: word; Shift: TShiftState);
    procedure FormShow(Sender: TObject);
    procedure lbKeywordDblClick(Sender: TObject);
    procedure lbKeywordKeyDown(Sender: TObject; var Key: word; Shift: TShiftState);
  private
    procedure InsertKeyword;
    { private declarations }
  public
    { public declarations }
  end;

var
  fmKeywords: TfmKeywords;

implementation

{$R *.lfm}

uses Unit1;

{ TfmKeywords }

procedure TfmKeywords.FormShow(Sender: TObject);
begin
  // Initialize the form
  if lbKeyword.Items.Count > 0 then
    lbKeyword.ItemIndex := 0;
  edFind.Clear;
  edFind.SetFocus;
end;

procedure TfmKeywords.edFindChange(Sender: TObject);
var
  i: integer;
begin
  // Find keyword
  if lbKeyword.Items.Count > 0 then
  begin
    if edFind.Text <> '' then
    begin
      for i := 0 to lbKeyword.Items.Count - 1 do
      begin
        if UTF8Copy(lbKeyword.Items[i], 1, UTF8Length(edFind.Text)) = edFind.Text then
        begin
          lbKeyword.ItemIndex := i;
          Break;
        end;
      end;
    end
    else
    begin
      lbKeyword.ItemIndex := 0;
    end;
  end;
end;

procedure TfmKeywords.bnCharOKClick(Sender: TObject);
begin
  // Close button
  fmKeywords.ModalResult := mrOk;
end;

procedure TfmKeywords.bnCharInsertClick(Sender: TObject);
begin
  // Insert with button
  InsertKeyword;
end;

procedure TfmKeywords.lbKeywordKeyDown(Sender: TObject; var Key: word;
  Shift: TShiftState);
begin
  // Insert on Return
  if key = 13 then
  begin
    key := 0;
    InsertKeyword;
  end;
end;

procedure TfmKeywords.edFindKeyDown(Sender: TObject; var Key: word; Shift: TShiftState);
begin
  // Insert on Return
  if key = 13 then
  begin
    key := 0;
    InsertKeyword;
  end
  // Arrow up
  else if ((Key = 38) or (Key = 40)) then
  begin
    lbKeyword.SetFocus;
  end;
end;

procedure TfmKeywords.FormCreate(Sender: TObject);
begin
  // Set color
  {$ifdef Win32}
  fmKeywords.Color := clWhite;
  {$endif}
  {$ifdef Darwin}
  fmKeywords.Color := clWhite;
  {$endif}
end;

procedure TfmKeywords.FormKeyDown(Sender: TObject; var Key: word; Shift: TShiftState);
begin
  // Close on Esc
  if key = 27 then
    fmKeywords.ModalResult := mrOk;
end;

procedure TfmKeywords.lbKeywordDblClick(Sender: TObject);
begin
  // Copy the keyword
  InsertKeyword;
end;

procedure TfmKeywords.InsertKeyword;
var
  Editor: TStringCellEditor;
begin
  // Insert the keywords in the sgRequired grid
  if lbKeyword.ItemIndex > 0 then
  begin
    Editor := TStringCellEditor(fmMain.sgRequiredFields.Editor);
    Editor.EditText := Editor.EditText + ', ' +
      lbKeyword.Items[lbKeyword.ItemIndex];
  end;
end;

end.




