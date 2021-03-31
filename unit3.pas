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
unit Unit3;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, FileUtil, Forms, Controls, Graphics, Dialogs, StdCtrls,
  CheckLst, ExtCtrls, Grids, ComCtrls, IniFiles, LCLProc;

type

  { TfmOptions }

  TfmOptions = class(TForm)
    bvBevelOptions4b: TBevel;
    bnCitLoad: TButton;
    bnCitSave: TButton;
    bnOK: TButton;
    bvBevelOptions2: TBevel;
    bvBevelOptions1: TBevel;
    bvBevelOptions4a: TBevel;
    bvBevelOptions3: TBevel;
    bvBevelOptions5: TBevel;
    bxFontSize: TListBox;
    cbAuthorAsIdem: TCheckBox;
    cbAllAuthors: TCheckBox;
    cbCitAsIbidem: TCheckBox;
    cbAutoSaveBibLatex: TCheckBox;
    cbBibTexLowercase: TCheckBox;
    cbDotAftertFirstName: TCheckBox;
    cbEditorAsAuthor: TCheckBox;
    cbNotExpAbsRew: TCheckBox;
    cbExpDoubleBrakets: TCheckBox;
    cbGridRows: TCheckBox;
    cbOpenLast: TCheckBox;
    cbExpAbsRew: TCheckBox;
    cbShortenAuthor: TCheckBox;
    cbShortenTitle: TCheckBox;
    cbEscDot: TCheckBox;
    clFieldsShown: TCheckListBox;
    edAuthSep: TEdit;
    edCitAsIbidem: TEdit;
    edPostnoteOnePage: TEdit;
    edPostnoteMorePages: TEdit;
    edMoreAuth: TEdit;
    edBibTexKey: TEdit;
    edEditorMention: TEdit;
    edIdemBiblio: TEdit;
    edIdemNotes: TEdit;
    edPatternName: TEdit;
    edUserDefCom: TEdit;
    lbCitAsIbidem: TLabel;
    lbPostnotePag: TLabel;
    lbMoreAuth: TLabel;
    lbInfoUserCmd: TLabel;
    lbIdemBiblio: TLabel;
    lbIdemNotes: TLabel;
    lbAuthSep: TLabel;
    lbConvWpFormat: TLabel;
    lbEditorMention: TLabel;
    lbExpOptions: TLabel;
    lbGenOptions: TLabel;
    lbBibTexKey: TLabel;
    lbFieldsShown: TLabel;
    lbFontSize: TLabel;
    lbPatternName: TLabel;
    lbUserDefCom: TLabel;
    pcOptions: TPageControl;
    pnOptPattern: TPanel;
    pnOptBottom: TPanel;
    rgIdemCmd: TRadioGroup;
    rgConvWpFormat: TRadioGroup;
    rgFormatDate: TRadioGroup;
    sdPanDir: TSelectDirectoryDialog;
    sgCitation: TStringGrid;
    tsCitationOptions: TTabSheet;
    tsConversion: TTabSheet;
    tsGeneralOptions: TTabSheet;
    tsCitationsPatterns: TTabSheet;
    procedure bnCitLoadClick(Sender: TObject);
    procedure bnCitSaveClick(Sender: TObject);
    procedure bnOKClick(Sender: TObject);
    procedure edBibTexKeyExit(Sender: TObject);
    procedure FormClose(Sender: TObject; var CloseAction: TCloseAction);
    procedure FormCreate(Sender: TObject);
    procedure rgConvWpFormatClick(Sender: TObject);
  private
    procedure ShowHideGridRows(ShowRows: boolean);
    { private declarations }
  public
    { public declarations }
  end;

var
  fmOptions: TfmOptions;
  // Home directory to store configuration file
  myHomeDir: string;
  // Name of configuration file
  myConfigFile: string;
  // Standard model of BibTex
  myBibTexKey: string;

implementation

uses Unit1;

{$R *.lfm}

{ TfmOptions }

procedure TfmOptions.FormCreate(Sender: TObject);
var
  MyIni: TIniFile;
  i: integer;
  SelItems: TStringList;
  stList: string;
begin
  // Set color and other stuff
  {$ifdef Win32}
  bxFontSize.ScrollWidth := 0;
  fmOptions.Color := clWhite;
  {$endif}
  {$ifdef Darwin}
  fmOptions.Color := clWhite;
  {$endif}
  // Set default values in lists
  myBibTexKey := '|author|[00]|year|[04]';
  edBibTexKey.Text := myBibTexKey;
  for i := 0 to 4 do
    clFieldsShown.Checked[i] := True;
  bxFontSize.ItemIndex := 0;
  // Load ini
  if FileExists(Unit1.myHomeDir + Unit1.myConfigFile) then
  begin
    MyIni := TIniFile.Create(Unit1.myHomeDir + Unit1.myConfigFile);
    try
      bxFontSize.ItemIndex := MyIni.ReadInteger('bibfilex', 'fontsize', 0);
      rgFormatDate.ItemIndex := MyIni.ReadInteger('bibfilex', 'dateformat', 0);
      if MyIni.ReadInteger('bibfilex', 'openfileonstart', 0) = 1 then
        cbOpenLast.Checked := True
      else
        cbOpenLast.Checked := False;
      if MyIni.ReadInteger('bibfilex', 'expdoublebrakets', 0) = 1 then
        cbExpDoubleBrakets.Checked := True
      else
        cbExpDoubleBrakets.Checked := False;
      if MyIni.ReadInteger('bibfilex', 'mergeabsrew', 0) = 1 then
        cbExpAbsRew.Checked := True
      else
        cbExpAbsRew.Checked := False;
      if MyIni.ReadInteger('bibfilex', 'notexpabsrev', 0) = 1 then
        cbNotExpAbsRew.Checked := True
      else
        cbNotExpAbsRew.Checked := False;
      if MyIni.ReadInteger('bibfilex', 'autosavebiblatex', 0) = 1 then
        cbAutoSaveBibLatex.Checked := True
      else
        cbAutoSaveBibLatex.Checked := False;
      if MyIni.ReadInteger('bibfilex', 'gridrows', 0) = 1 then
      begin
        cbGridRows.Checked := True;
        ShowHideGridRows(True);
      end
      else
      begin
        cbGridRows.Checked := False;
        ShowHideGridRows(False);
      end;
      if MyIni.ReadInteger('bibfilex', 'editorasauthor', 0) = 1 then
        cbEditorAsAuthor.Checked := True
      else
        cbEditorAsAuthor.Checked := False;
      if MyIni.ReadInteger('bibfilex', 'dotaftername', 0) = 1 then
        cbDotAftertFirstName.Checked := True
      else
        cbDotAftertFirstName.Checked := False;
      if MyIni.ReadInteger('bibfilex', 'shortentitle', 0) = 1 then
        cbShortenTitle.Checked := True
      else
        cbShortenTitle.Checked := False;
      if MyIni.ReadInteger('bibfilex', 'shortenauthor', 0) = 1 then
        cbShortenAuthor.Checked := True
      else
        cbShortenAuthor.Checked := False;
      if MyIni.ReadInteger('bibfilex', 'citasibidem', 0) = 1 then
        cbCitAsIbidem.Checked := True
      else
        cbCitAsIbidem.Checked := False;
      if MyIni.ReadInteger('bibfilex', 'authorasidem', 0) = 1 then
        cbAuthorAsIdem.Checked := True
      else
        cbAuthorAsIdem.Checked := False;
      if MyIni.ReadInteger('bibfilex', 'escdot', 0) = 1 then
        cbEscDot.Checked := True
      else
        cbEscDot.Checked := False;
      if MyIni.ReadInteger('bibfilex', 'allauth', 0) = 1 then
        cbAllAuthors.Checked := True
      else
        cbAllAuthors.Checked := False;
      if MyIni.ReadInteger('bibfilex', 'bibtexlower', 0) = 1 then
        cbBibTexLowercase.Checked := True
      else
        cbBibTexLowercase.Checked := False;
      rgIdemCmd.ItemIndex := MyIni.ReadInteger('bibfilex', 'idemcmd', 0);
      rgConvWpFormat.ItemIndex := MyIni.ReadInteger('bibfilex', 'convwp', 0);
      edAuthSep.Text := MyIni.ReadString('bibfilex', 'authsep', '--');
      edMoreAuth.Text := MyIni.ReadString('bibfilex', 'moreauth', 'et al.');
      edEditorMention.Text := MyIni.ReadString('bibfilex', 'edmention', 'ed.');
      edPostnoteOnePage.Text := MyIni.ReadString('bibfilex', 'onepage', 'p.');
      edPostnoteMorePages.Text := MyIni.ReadString('bibfilex', 'morepages', 'pp.');
      edCitAsIbidem.Text := MyIni.ReadString('bibfilex', 'ibidemmention', 'Ibid.');
      edIdemNotes.Text := MyIni.ReadString('bibfilex', 'authoridemnotes', 'Id.');
      edIdemBiblio.Text := MyIni.ReadString('bibfilex', 'authoridembiblio', '____');
      edBibTexKey.Text := MyIni.ReadString('bibfilex', 'bibtexkeyptrn', '');
      edPatternName.Text := MyIni.ReadString('bibfilex', 'patternname', '');
      edUserDefCom.Text := MyIni.ReadString('bibfilex', 'usrdefcmd', '');
      if edBibTexKey.Text = '' then
        edBibTexKey.Text := myBibTexKey;
      SelItems := TStringList.Create;
      stList := MyIni.ReadString('bibfilex', 'fieldsvisingrid', '');
      SelItems.Text := StringReplace(stList, '|', LineEnding, [rfReplaceAll]);
      for i := 0 to clFieldsShown.Count - 1 do
        clFieldsShown.Checked[i] := False;
      if SelItems.Count > 0 then
        for i := 0 to SelItems.Count - 1 do
          clFieldsShown.Checked[StrToInt(SelItems[i])] := True
      else
        for i := 0 to 4 do
          clFieldsShown.Checked[i] := True;
      if rgConvWpFormat.ItemIndex < rgConvWpFormat.Items.Count - 1 then
      begin
        edUserDefCom.Enabled := False;
      end
      else
      begin
        edUserDefCom.Enabled := True;
      end;
    finally
      MyIni.Free;
      SelItems.Free;
    end;
  end;
  if FileExists(Unit1.myHomeDir + 'citations2-' + Unit1.myConfigFile) then
    try
      sgCitation.Clean;
      sgCitation.LoadFromFile(Unit1.myHomeDir + 'citations2-' +
        Unit1.myConfigFile);
    except
      MessageDlg('The file of the citations seems to be corrupted.',
        mtError, [mbOK], 0);
    end;
end;

procedure TfmOptions.FormClose(Sender: TObject; var CloseAction: TCloseAction);
var
  MyIni: TIniFile;
  SelItems: TStringList;
  i: integer;
  stList: string;
  flChecked: boolean = False;
begin
  // Set options
  // Check that at least 1 col is checked
  for i := 0 to clFieldsShown.Count - 1 do
  begin
    if clFieldsShown.Checked[i] = True then
      flChecked := True;
  end;
  if flChecked = False then
  begin
    for i := 0 to 4 do
      clFieldsShown.Checked[i] := True;
  end;
  fmMain.SetOptions;
  fmMain.grItems.AutoSizeColumns;
  // Save elements of the form
  try
    MyIni := TIniFile.Create(Unit1.myHomeDir + Unit1.myConfigFile);
    SelItems := TStringList.Create;
    for i := 0 to clFieldsShown.Items.Count - 1 do
      if clFieldsShown.Checked[i] = True then
        SelItems.Add(IntToStr(i));
    stList := StringReplace(SelItems.Text, LineEnding, '|', [rfReplaceAll]);
    MyIni.WriteString('bibfilex', 'fieldsvisingrid', stList);
    MyIni.WriteInteger('bibfilex', 'fontsize', bxFontSize.ItemIndex);
    MyIni.WriteInteger('bibfilex', 'dateformat', rgFormatDate.ItemIndex);
    if cbOpenLast.Checked = True then
      MyIni.WriteInteger('bibfilex', 'openfileonstart', 1)
    else
      MyIni.WriteInteger('bibfilex', 'openfileonstart', 0);
    if cbExpDoubleBrakets.Checked = True then
      MyIni.WriteInteger('bibfilex', 'expdoublebrakets', 1)
    else
      MyIni.WriteInteger('bibfilex', 'expdoublebrakets', 0);
    if cbExpAbsRew.Checked = True then
      MyIni.WriteInteger('bibfilex', 'mergeabsrew', 1)
    else
      MyIni.WriteInteger('bibfilex', 'mergeabsrew', 0);
    if cbNotExpAbsRew.Checked = True then
      MyIni.WriteInteger('bibfilex', 'notexpabsrev', 1)
    else
      MyIni.WriteInteger('bibfilex', 'notexpabsrev', 0);
    if cbAutoSaveBibLatex.Checked = True then
      MyIni.WriteInteger('bibfilex', 'autosavebiblatex', 1)
    else
      MyIni.WriteInteger('bibfilex', 'autosavebiblatex', 0);
    if cbGridRows.Checked = True then
    begin
      MyIni.WriteInteger('bibfilex', 'gridrows', 1);
      ShowHideGridRows(True);
    end
    else
    begin
      MyIni.WriteInteger('bibfilex', 'gridrows', 0);
      ShowHideGridRows(False);
    end;
    if cbEditorAsAuthor.Checked = True then
      MyIni.WriteInteger('bibfilex', 'editorasauthor', 1)
    else
      MyIni.WriteInteger('bibfilex', 'editorasauthor', 0);
    if cbDotAftertFirstName.Checked = True then
      MyIni.WriteInteger('bibfilex', 'dotaftername', 1)
    else
      MyIni.WriteInteger('bibfilex', 'dotaftername', 0);
    if cbShortenTitle.Checked = True then
      MyIni.WriteInteger('bibfilex', 'shortentitle', 1)
    else
      MyIni.WriteInteger('bibfilex', 'shortentitle', 0);
    if cbShortenAuthor.Checked = True then
      MyIni.WriteInteger('bibfilex', 'shortenauthor', 1)
    else
      MyIni.WriteInteger('bibfilex', 'shortenauthor', 0);
    if cbCitAsIbidem.Checked = True then
      MyIni.WriteInteger('bibfilex', 'citasibidem', 1)
    else
      MyIni.WriteInteger('bibfilex', 'citasibidem', 0);
    if cbAuthorAsIdem.Checked = True then
      MyIni.WriteInteger('bibfilex', 'authorasidem', 1)
    else
      MyIni.WriteInteger('bibfilex', 'authorasidem', 0);
    if cbEscDot.Checked = True then
      MyIni.WriteInteger('bibfilex', 'escdot', 1)
    else
      MyIni.WriteInteger('bibfilex', 'escdot', 0);
    if cbAllAuthors.Checked = True then
      MyIni.WriteInteger('bibfilex', 'allauth', 1)
    else
      MyIni.WriteInteger('bibfilex', 'allauth', 0);
    if cbBibTexLowercase.Checked = True then
      MyIni.WriteInteger('bibfilex', 'bibtexlower', 1)
    else
      MyIni.WriteInteger('bibfilex', 'bibtexlower', 0);
    MyIni.WriteInteger('bibfilex', 'idemcmd', rgIdemCmd.ItemIndex);
    MyIni.WriteInteger('bibfilex', 'convwp', rgConvWpFormat.ItemIndex);
    MyIni.WriteString('bibfilex', 'authsep', edAuthSep.Text);
    MyIni.WriteString('bibfilex', 'moreauth', edMoreAuth.Text);
    MyIni.WriteString('bibfilex', 'edmention', edEditorMention.Text);
    MyIni.WriteString('bibfilex', 'onepage', edPostnoteOnePage.Text);
    MyIni.WriteString('bibfilex', 'morepages', edPostnoteMorePages.Text);
    MyIni.WriteString('bibfilex', 'ibidemmention', edCitAsIbidem.Text);
    MyIni.WriteString('bibfilex', 'authoridemnotes', edIdemNotes.Text);
    MyIni.WriteString('bibfilex', 'authoridembiblio', edIdemBiblio.Text);
    MyIni.WriteString('bibfilex', 'bibtexkeyptrn', edBibTexKey.Text);
    MyIni.WriteString('bibfilex', 'patternname', edPatternName.Text);
    MyIni.WriteString('bibfilex', 'usrdefcmd', edUserDefCom.Text);
  finally
    SelItems.Free;
    MyIni.Free;
  end;
  try
    sgCitation.SaveToFile(Unit1.myHomeDir + 'citations2-' + Unit1.myConfigFile);
  except
    MessageDlg('It was not possibile to save the citations file.',
      mtError, [mbOK], 0);
  end;
end;

procedure TfmOptions.edBibTexKeyExit(Sender: TObject);
var
  stModel: string;
  i: integer;
begin
  // Check if the model of the BibTex key is correct
  edBibTexKey.Text := StringReplace(edBibTexKey.Text, ' ', '', [rfReplaceAll]);
  stModel := edBibTexKey.Text;
  // Check for delimiters
  if ((UTF8Pos('[', stModel) = 0) or (UTF8Pos(']', stModel) = 0)) then
  begin
    MessageDlg('The BibTex key composition' + LineEnding +
      'is not typed correctely.', mtWarning, [mbOK], 0);
    edBibTexKey.Text := myBibTexKey;
    Exit;
  end;
  while UTF8Pos('[', stModel) > 0 do
  begin
    // Check fields names
    if clFieldsShown.Items.IndexOf(UTF8Copy(stModel, UTF8Pos('|', stModel) +
      1, UTF8Pos('[', stModel) - UTF8Pos('|', stModel) - 2)) < 0 then
    begin
      MessageDlg('One of the fields in the BibTex key composition' +
        LineEnding + 'is not typed correctely.', mtWarning, [mbOK], 0);
      edBibTexKey.Text := myBibTexKey;
      Exit;
    end;
    // Check field length
    if UTF8Length(UTF8Copy(stModel, UTF8Pos('[', stModel) + 1,
      UTF8Pos(']', stModel) - UTF8Pos('[', stModel) - 1)) <> 2 then
    begin
      MessageDlg('One of the field length in the BibTex key composition' +
        LineEnding + 'has more or less that two digits.', mtWarning, [mbOK], 0);
      edBibTexKey.Text := myBibTexKey;
      Exit;
    end;
    // Check it's numeric
    try
      i := StrToInt(UTF8Copy(stModel, UTF8Pos('[', stModel) + 1,
        UTF8Pos(']', stModel) - UTF8Pos('[', stModel) - 1));
    except
      MessageDlg('One of the field length in the BibTex key composition' +
        LineEnding + 'is not typed correctely.', mtWarning, [mbOK], 0);
      edBibTexKey.Text := myBibTexKey;
      Exit;
    end;
    stModel := UTF8Copy(stModel, UTF8Pos(']', stModel) + 1, UTF8Length(stModel));
  end;
end;

procedure TfmOptions.bnOKClick(Sender: TObject);
begin
  // Close
  Close;
end;

procedure TfmOptions.bnCitLoadClick(Sender: TObject);
begin
  // Open a citation structure
  fmMain.odOpenFile.Filter := 'All files|*';
  fmMain.odOpenFile.DefaultExt := '';
  fmMain.odOpenFile.FileName := '';
  if fmMain.odOpenFile.Execute = True then
    try
      sgCitation.Clean;
      sgCitation.LoadFromFile(fmMain.odOpenFile.FileName);
      edPatternName.Text :=
        ExtractFileNameOnly(fmMain.odOpenFile.FileName);
    except
      MessageDlg('Error on loading citations pattern.', mtError, [mbOK], 0);
    end;
end;

procedure TfmOptions.bnCitSaveClick(Sender: TObject);
begin
  // Save a citation structure
  fmMain.sdSaveFile.Filter := 'All files|*';
  fmMain.sdSaveFile.DefaultExt := '';
  fmMain.sdSaveFile.FileName := '';
  if fmMain.sdSaveFile.Execute = True then
    try
      sgCitation.SaveToFile(fmMain.sdSaveFile.FileName);
    except
      MessageDlg('Error on saving citations pattern.', mtError, [mbOK], 0);
    end;
end;

procedure TfmOptions.rgConvWpFormatClick(Sender: TObject);
begin
  // Enable user command field
  if rgConvWpFormat.ItemIndex < rgConvWpFormat.Items.Count - 1 then
  begin
    edUserDefCom.Enabled := False;
  end
  else
  begin
    edUserDefCom.Enabled := True;
  end;
end;

procedure TfmOptions.ShowHideGridRows(ShowRows: boolean);
begin
  // Show or hide row lines in the entry grids
  with fmMain do
  begin
    if ShowRows = True then
    begin
      sgRequiredFields.Options := sgRequiredFields.Options + [goHorzLine];
      sgOptionalFields1.Options := sgRequiredFields.Options + [goHorzLine];
      sgOptionalFields2.Options := sgRequiredFields.Options + [goHorzLine];
      sgOptionalFields3.Options := sgRequiredFields.Options + [goHorzLine];
    end
    else
    begin
      sgRequiredFields.Options := sgRequiredFields.Options - [goHorzLine];
      sgOptionalFields1.Options := sgRequiredFields.Options - [goHorzLine];
      sgOptionalFields2.Options := sgRequiredFields.Options - [goHorzLine];
      sgOptionalFields3.Options := sgRequiredFields.Options - [goHorzLine];
    end;
  end;
end;

end.
