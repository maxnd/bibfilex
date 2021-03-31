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
unit Copyright;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, FileUtil, Forms, Controls, Graphics, Dialogs, StdCtrls,
  ExtCtrls, LCLIntf;

type

  { TfmCopyright }

  TfmCopyright = class(TForm)
    bnCopyright: TButton;
    imImage: TImage;
    lbCopyrightSubTitle: TLabel;
    lbGroup: TLabel;
    lbSite: TLabel;
    lbCopyrightAuthor1: TLabel;
    lbCopyrightAuthor2: TLabel;
    lbCopyrightName: TLabel;
    meCopyrightText: TMemo;
    procedure bnCopyrightClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure lbGroupClick(Sender: TObject);
    procedure lbSiteClick(Sender: TObject);
  private
    { private declarations }
  public
    { public declarations }
  end;

var
  fmCopyright: TfmCopyright;

implementation

{$R *.lfm}

uses Unit1;

{ TfmCopyright }

procedure TfmCopyright.bnCopyrightClick(Sender: TObject);
begin
  // Quit
  Close;
end;

procedure TfmCopyright.FormCreate(Sender: TObject);
begin
  // Set color
  {$ifdef Win32}
  fmCopyright.Color := clWhite;
  {$endif}
  {$ifdef Darwin}
  fmCopyright.Color := clWhite;
  {$endif}
end;

procedure TfmCopyright.FormKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  // Quit
  if key = 27 then
    Close;
end;

procedure TfmCopyright.lbSiteClick(Sender: TObject);
begin
  // Open the website
  OpenURL('http://sites.google.com/site/bibfilex/');
end;

procedure TfmCopyright.lbGroupClick(Sender: TObject);
begin
  // Open the group
  OpenURL('https://groups.google.com/forum/?fromgroups#!forum/bibfilex/');
end;

end.
