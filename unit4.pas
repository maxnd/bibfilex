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
unit Unit4;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, FileUtil, Forms, Controls, Graphics, Dialogs, ExtCtrls,
  StdCtrls;

type

  { TfmStopList }

  TfmStopList = class(TForm)
    bnOK: TButton;
    bnSort: TButton;
    meStopList: TMemo;
    pnBottom: TPanel;
    procedure bnOKClick(Sender: TObject);
    procedure bnSortClick(Sender: TObject);
    procedure FormClose(Sender: TObject; var CloseAction: TCloseAction);
    procedure FormCreate(Sender: TObject);
  private
    { private declarations }
  public
    { public declarations }
  end;

var
  fmStopList: TfmStopList;

implementation

uses Unit1;

{$R *.lfm}

{ TfmStopList }

procedure TfmStopList.FormCreate(Sender: TObject);
begin
  // Set color
  {$ifdef Win32}
  fmStopList.Color := clWhite;
  {$endif}
  {$ifdef Darwin}
  fmStopList.Color := clWhite;
  {$endif}
  // Load list
  if FileExistsUTF8(Unit1.myHomeDir + 'stoplist-bibfilex') = True then
    try
      meStopList.Lines.LoadFromFile(Unit1.myHomeDir + 'stoplist-bibfilex');
    except
      MessageDlg('Error on loading the stop list.', mtWarning, [mbOK], 0);
    end;
end;

procedure TfmStopList.FormClose(Sender: TObject; var CloseAction: TCloseAction);
begin
  // Save list
  try
    meStopList.Lines.SaveToFile(Unit1.myHomeDir + 'stoplist-bibfilex');
  except
    MessageDlg('Error on saving the stop list.', mtError, [mbOK], 0);
  end;
end;

procedure TfmStopList.bnOKClick(Sender: TObject);
begin
  // Save list and quit
  try
    meStopList.Lines.SaveToFile(Unit1.myHomeDir + 'stoplist-bibfilex');
  except
    MessageDlg('Error on saving the stop list.', mtError, [mbOK], 0);
  end;
  Close;
end;

procedure TfmStopList.bnSortClick(Sender: TObject);
var
  myList: TStringList;
  i: integer;
begin
  // Sort the items
  try
    myList := TStringList.Create;
    myList.Text := meStopList.Text;
    myList.Sort;
    meStopList.Text := myList.Text;
  finally
    myList.Free;
  end;
end;

end.
