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
unit Unit1;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, Sqlite3DS, DB, FileUtil, Forms, Controls, Graphics,
  Dialogs, Menus, ComCtrls, DBGrids, ExtCtrls, Grids, StdCtrls, DBCtrls,
  LCLProc, IniFiles, Zipper, Process, Clipbrd, LCLType, LCLIntf, types,
  LazUTF8, LazFileUtils;

type

  { TfmMain }

  TfmMain = class(TForm)
    apAppProp: TApplicationProperties;
    bnAddKey: TButton;
    bnRemoveKey: TButton;
    bvLine1: TBevel;
    bnAutoSize: TButton;
    bvBevelFind: TBevel;
    cbItemsKind: TComboBox;
    cbFieldSmartFilter: TComboBox;
    cbMultiselect: TCheckBox;
    dsToolsTables: TDataSource;
    dbOwned: TDBCheckBox;
    cbFilterKey: TComboBox;
    ilNavigator: TImageList;
    lbAuthorTitle: TLabel;
    lbOwned: TLabel;
    meAbstract: TMemo;
    miToolsDecrypt: TMenuItem;
    miToolsEncrypt: TMenuItem;
    miLine3GPG: TMenuItem;
    miToolCheckDoubles: TMenuItem;
    miLine3b: TMenuItem;
    miItemsStoreKeyword: TMenuItem;
    miItemsApplyKeyword: TMenuItem;
    miToolsReplaceKeys: TMenuItem;
    miLine3d: TMenuItem;
    miFileCopyAs: TMenuItem;
    miLIne0: TMenuItem;
    miLine1c: TMenuItem;
    miFileImportFromBibfilex: TMenuItem;
    miFileExportToBifilex: TMenuItem;
    miToolsConvert: TMenuItem;
    miToolKeyList: TMenuItem;
    miItemsModifyKey: TMenuItem;
    meReview: TMemo;
    miItemsCopyCrossref: TMenuItem;
    miToolsUpgrade: TMenuItem;
    miItemsOrderByBookTitle: TMenuItem;
    miItemsOrderByJournalTitle: TMenuItem;
    pnReview: TPanel;
    pnAbstract: TPanel;
    pmItemsOpenLink: TMenuItem;
    pmItemsLine2: TMenuItem;
    pmItemsLine1: TMenuItem;
    miToolChar: TMenuItem;
    miLine3c: TMenuItem;
    miShowManual: TMenuItem;
    miLine4: TMenuItem;
    nvNavigator: TDBNavigator;
    dsItems: TDatasource;
    edSmartFilter: TEdit;
    grItems: TDBGrid;
    ilImageListToolbar: TImageList;
    imgImageBack: TImage;
    lbSmartFilter: TLabel;
    lbFilterKeyword: TLabel;
    lbAttNames: TListBox;
    lbLastEdited: TLabel;
    lbListAttach: TLabel;
    lbEntryTypes: TLabel;
    miLine2i: TMenuItem;
    miItemsCopyCitationHtml: TMenuItem;
    miToolsStopList: TMenuItem;
    miToolsReplaceCitations: TMenuItem;
    miLine3a: TMenuItem;
    miItemsCopyCitationTxt: TMenuItem;
    miLine2c: TMenuItem;
    miLine2d: TMenuItem;
    miItemsFilterFromLatex: TMenuItem;
    miItemsRenameKeyword: TMenuItem;
    pnMultiselect: TPanel;
    pnGrid: TPanel;
    pmItemsNew: TMenuItem;
    pmItemsUndo: TMenuItem;
    pmItemsDelete: TMenuItem;
    pmItemsSave: TMenuItem;
    miItemsCopyKey: TMenuItem;
    miItemsUndo: TMenuItem;
    miItemsRemoveFilter: TMenuItem;
    miItemsCreateKeywordList: TMenuItem;
    miAttachNew: TMenuItem;
    miAttachDelete: TMenuItem;
    miLine2g: TMenuItem;
    miAttachSaveAs: TMenuItem;
    miAttachOpen: TMenuItem;
    miLine2a: TMenuItem;
    miItemsAttachments: TMenuItem;
    miFileCreateBibLatex: TMenuItem;
    miLine2b: TMenuItem;
    miItemsCreateKey: TMenuItem;
    miItemsOrderByAuthor: TMenuItem;
    miItemsOrderByTitle: TMenuItem;
    miItemsOrderByDate: TMenuItem;
    miItemsOrderByID: TMenuItem;
    miItemsOrderByYear: TMenuItem;
    miLine2e: TMenuItem;
    miItemsOrderBy: TMenuItem;
    miLIne1b: TMenuItem;
    miLineAtt1: TMenuItem;
    miLineLastFile: TMenuItem;
    miFileOpenLast4: TMenuItem;
    miFileOpenLast2: TMenuItem;
    miFileOpenLast1: TMenuItem;
    miFileOpenLast3: TMenuItem;
    miLine1d: TMenuItem;
    miFileImportFromBiblatex: TMenuItem;
    miFileExportToBiblatex: TMenuItem;
    miCopyrightShow: TMenuItem;
    miToolsOptions: TMenuItem;
    miToolsCompact: TMenuItem;
    miLine3: TMenuItem;
    miItemsFilter: TMenuItem;
    miLine2: TMenuItem;
    miItemsNew: TMenuItem;
    miItemsDelete: TMenuItem;
    miItemsSave: TMenuItem;
    miCopyright: TMenuItem;
    miTools: TMenuItem;
    miItems: TMenuItem;
    miFileNew: TMenuItem;
    miFileOpen: TMenuItem;
    miFileClose: TMenuItem;
    miLIne1: TMenuItem;
    miFileExit: TMenuItem;
    mmMainMenu: TMainMenu;
    miFile: TMenuItem;
    odOpenFile: TOpenDialog;
    pmAttachments: TPopupMenu;
    pmAttDelete: TMenuItem;
    pmAttNew: TMenuItem;
    pmAttOpen: TMenuItem;
    pmAttSaveAs: TMenuItem;
    pnKindSearch: TPanel;
    pcPageControl: TPageControl;
    pnFields: TPanel;
    pnMain: TPanel;
    pmMain: TPopupMenu;
    sdSaveFile: TSaveDialog;
    sbStatusBar: TStatusBar;
    sdSelDirDialog: TSelectDirectoryDialog;
    sgOptionalFields1: TStringGrid;
    sgOptionalFields2: TStringGrid;
    sgOptionalFields3: TStringGrid;
    sgRequiredFields: TStringGrid;
    spSplitterItems: TSplitter;
    sqItems: TSqlite3Dataset;
    sqToolsTables: TSqlite3Dataset;
    tbFileNew: TToolButton;
    tbFileOpen: TToolButton;
    tbItemsSave: TToolButton;
    tbDiv1: TToolButton;
    tbItemsNew: TToolButton;
    ToolButton6: TToolButton;
    tbItemsFilter: TToolButton;
    tbDiv2: TToolButton;
    tbFileCreateBibLatex: TToolButton;
    tsRequiredFields: TTabSheet;
    tsOptionalFields1: TTabSheet;
    tsOptionalFields2: TTabSheet;
    tsOptionalFields3: TTabSheet;
    tsAbstract: TTabSheet;
    tsReview: TTabSheet;
    tbToolBar: TToolBar;
    procedure apAppPropException(Sender: TObject; E: Exception);
    procedure bnAddKeyClick(Sender: TObject);
    procedure bnAutoSizeClick(Sender: TObject);
    procedure bnRemoveKeyClick(Sender: TObject);
    procedure cbFilterKeySelect(Sender: TObject);
    procedure dbOwnedChange(Sender: TObject);
    procedure cbFilterKeyKeyDown(Sender: TObject; var Key: word;
      Shift: TShiftState);
    procedure cbItemsKindChange(Sender: TObject);
    procedure Abstract_ReviewChange(Sender: TObject);
    procedure cbItemsKindKeyDown(Sender: TObject; var Key: word;
      Shift: TShiftState);
    procedure cbItemsKindSelect(Sender: TObject);
    procedure cbMultiselectChange(Sender: TObject);
    procedure dsItemsDataChange(Sender: TObject; Field: TField);
    procedure edSmartFilterKeyDown(Sender: TObject; var Key: word; Shift: TShiftState);
    procedure FormResize(Sender: TObject);
    procedure grItemsPrepareCanvas(Sender: TObject; DataCol: integer;
      Column: TColumn; AState: TGridDrawState);
    procedure meAbstractMouseWheel(Sender: TObject; Shift: TShiftState;
      WheelDelta: integer; MousePos: TPoint; var Handled: boolean);
    procedure meReviewMouseWheel(Sender: TObject; Shift: TShiftState;
      WheelDelta: integer; MousePos: TPoint; var Handled: boolean);
    procedure miItemsApplyKeywordClick(Sender: TObject);
    procedure miItemsStoreKeywordClick(Sender: TObject);
    procedure miToolCheckDoublesClick(Sender: TObject);
    procedure miToolsDecryptClick(Sender: TObject);
    procedure miToolsEncryptClick(Sender: TObject);
    procedure miToolsReplaceKeysClick(Sender: TObject);
    procedure miFileCopyAsClick(Sender: TObject);
    procedure miFileExportToBifilexClick(Sender: TObject);
    procedure miFileImportFromBibfilexClick(Sender: TObject);
    procedure miItemsCopyCrossrefClick(Sender: TObject);
    procedure miItemsModifyKeyClick(Sender: TObject);
    procedure miToolCharClick(Sender: TObject);
    procedure miToolKeyListClick(Sender: TObject);
    procedure miToolsUpgradeClick(Sender: TObject);
    procedure pmItemsOpenLinkClick(Sender: TObject);
    procedure pmMainPopup(Sender: TObject);
    procedure sgOptionalFields1KeyUp(Sender: TObject; var Key: word;
      Shift: TShiftState);
    procedure sgOptionalFields2KeyUp(Sender: TObject; var Key: word;
      Shift: TShiftState);
    procedure sgOptionalFields3KeyUp(Sender: TObject; var Key: word;
      Shift: TShiftState);
    procedure sgRequiredFieldsKeyUp(Sender: TObject; var Key: word;
      Shift: TShiftState);
    procedure sgRequiredFieldsSelectCell(Sender: TObject; aCol, aRow: integer;
      var CanSelect: boolean);
    procedure SmartFilter;
    procedure FormActivate(Sender: TObject);
    procedure FormClose(Sender: TObject; var CloseAction: TCloseAction);
    procedure FormCreate(Sender: TObject);
    procedure FormDropFiles(Sender: TObject; const FileNames: array of string);
    procedure FormKeyDown(Sender: TObject; var Key: word; Shift: TShiftState);
    procedure grItemsDblClick(Sender: TObject);
    procedure lbAttNamesDblClick(Sender: TObject);
    procedure meAbstractKeyDown(Sender: TObject; var Key: word; Shift: TShiftState);
    procedure meReviewKeyDown(Sender: TObject; var Key: word; Shift: TShiftState);
    procedure miItemsCopyCitationHtmlClick(Sender: TObject);
    procedure miItemsCopyCitationTxtClick(Sender: TObject);
    procedure miItemsFilterFromLatexClick(Sender: TObject);
    procedure miItemsRenameKeywordClick(Sender: TObject);
    procedure miAttachDeleteClick(Sender: TObject);
    procedure miAttachNewClick(Sender: TObject);
    procedure miAttachOpenClick(Sender: TObject);
    procedure miCopyrightShowClick(Sender: TObject);
    procedure miFileCloseClick(Sender: TObject);
    procedure miFileCreateBibLatexClick(Sender: TObject);
    procedure miFileExitClick(Sender: TObject);
    procedure miFileExportToBiblatexClick(Sender: TObject);
    procedure miFileNewClick(Sender: TObject);
    procedure miFileOpenClick(Sender: TObject);
    procedure miFileImportFromBiblatexClick(Sender: TObject);
    procedure miItemsCopyKeyClick(Sender: TObject);
    procedure miItemsCreateKeyClick(Sender: TObject);
    procedure miItemsCreateKeywordListClick(Sender: TObject);
    procedure miItemsDeleteClick(Sender: TObject);
    procedure miItemsNewClick(Sender: TObject);
    procedure miItemsOrderByIDClick(Sender: TObject);
    procedure miItemsRemoveFilterClick(Sender: TObject);
    procedure miItemsSaveClick(Sender: TObject);
    procedure miFileOpenLast1Click(Sender: TObject);
    procedure miFileOpenLast2Click(Sender: TObject);
    procedure miFileOpenLast3Click(Sender: TObject);
    procedure miFileOpenLast4Click(Sender: TObject);
    procedure miItemsFilterClick(Sender: TObject);
    procedure miItemsUndoClick(Sender: TObject);
    procedure miShowManualClick(Sender: TObject);
    procedure miToolsCompactClick(Sender: TObject);
    procedure miToolsOptionsClick(Sender: TObject);
    procedure miToolsReplaceCitationsClick(Sender: TObject);
    procedure miToolsStopListClick(Sender: TObject);
    procedure sgOptionalFields1KeyDown(Sender: TObject; var Key: word;
      Shift: TShiftState);
    procedure sgOptionalFields1Resize(Sender: TObject);
    procedure sgOptionalFields1ValidateEntry(Sender: TObject;
      aCol, aRow: integer; const OldValue: string; var NewValue: string);
    procedure sgOptionalFields2KeyDown(Sender: TObject; var Key: word;
      Shift: TShiftState);
    procedure sgOptionalFields2Resize(Sender: TObject);
    procedure sgOptionalFields2ValidateEntry(Sender: TObject;
      aCol, aRow: integer; const OldValue: string; var NewValue: string);
    procedure sgOptionalFields3KeyDown(Sender: TObject; var Key: word;
      Shift: TShiftState);
    procedure sgOptionalFields3Resize(Sender: TObject);
    procedure sgOptionalFields3ValidateEntry(Sender: TObject;
      aCol, aRow: integer; const OldValue: string; var NewValue: string);
    procedure sgRequiredFieldsKeyDown(Sender: TObject; var Key: word;
      Shift: TShiftState);
    procedure sgRequiredFieldsPrepareCanvas(Sender: TObject;
      aCol, aRow: integer; aState: TGridDrawState);
    procedure sgRequiredFieldsResize(Sender: TObject);
    procedure DataGridsSetEditText(Sender: TObject; ACol, ARow: integer;
      const Value: string);
    procedure sqItemsAfterPost(DataSet: TDataSet);
    procedure sqItemsAfterScroll(DataSet: TDataSet);
    procedure sqItemsBeforeDelete(DataSet: TDataSet);
    procedure sqItemsBeforeScroll(DataSet: TDataSet);
    procedure SetItemsGridColumns;
    procedure SetOptions;
    procedure MoveFocus(Sender: TObject; var Key: word; Shift: TShiftState);
    function CreateOrderBy: string;
  private
    procedure ActivateSaving;
    procedure ActivateSuggestion(grGrid: TStringGrid);
    function CheckUniqueBibTexKey(BibTexKey: string; IDItems: integer): string;
    function CleanAbstractReview(stText: string): string;
    function CleanKeywordField(myField: string): string;
    procedure CloseDataTables;
    procedure CompileSummary;
    procedure CreateAttachment(FileNames: array of string);
    function CreateBibTexKey(IsItemsDataSet: boolean; SaveBefore: boolean): string;
    function CreateCitation(BibTexKey: string; AuthorNameBefore: boolean;
      AuthorAsIdem: boolean; IsHTML: boolean; SkipPages: boolean;
      CitationKind: smallint; MarkAroundAuthor: boolean; TitPages: string): string;
    procedure CreateDataTables(DataFileName: string);
    procedure CreateKeywordList;
    procedure ExportToBibfilex(stFileName: string);
    procedure ExportToBibLatex(stFileName: string; Filtered: boolean);
    procedure FilterFromLatex(FileName: string);
    procedure FilterOnKeyword;
    function FormatAuthorInCitation(stAuthor: string; NameBefore: boolean;
      OnlyFamilyName: boolean): string;
    procedure ImportBibLatex(stFileName: string);
    procedure ImportFromBibfilex(stFileName: string);
    procedure InsertKeyAbstRew(meMemo: TMemo);
    function IsDirectoryEmpty(const myDir: string): boolean;
    procedure LoadItems;
    function MakeTableUpgrade(DataFileName: string): boolean;
    procedure NewItem(iKind: Short);
    procedure OpenDataTables(stFileName: string);
    procedure RenameKeyword;
    function ReplaceCiteInLatex(FileName: string): boolean;
    procedure SaveItems;
    procedure SetAttachmentMenuItems;
    procedure SetMarkAbstractReviewTab;
    procedure SetStatusBar;
    function ShortenTitle(stTitle: string): string;
    function SuggestValue(StartText: string; Field: string): string;
    { private declarations }
  public
    { public declarations }
  end;

var
  fmMain: TfmMain;
  // Home directory to store configuration file
  myHomeDir: string;
  // Name of configuration file
  myConfigFile: string;
  // First activation of the main form
  blFirstAct: boolean = True;
  // Last Database used
  LastDatabase1, LastDatabase2, LastDatabase3, LastDatabase4: string;
  // Data modified or not
  flDataEdit: boolean = False;
  // Export to Latex on exit or not
  flExportOnExit: boolean = False;
  // Filter active or not
  flFilterActive: short = 0;
  //List of Bookmarks
  BookmarkList: TStringList;
  // Keyword in buffer
  stBuffKeyword: string;
  // Current and Previous Author in citations, with full name
  CitAutCurrent: string;
  CitAutPrevious: string;
  // Mail for pgp encryption
  stRecipient: string = '';

implementation

{$R *.lfm}

uses UnitSetFields, Unit2, Unit3, Unit4, Unit5, Unit6, Unit7, Copyright;

//***************************************************************
//**********************  FORM PROCEDURES  **********************
//***************************************************************

procedure TfmMain.FormCreate(Sender: TObject);
var
  MyIni: TIniFile;
  i: integer;
  stKeywordsList: string;
begin
  // Set home directory, data directories and CR
  {$ifdef Linux}
  myHomeDir := GetEnvironmentVariable('HOME') + DirectorySeparator +
    '.config' + DirectorySeparator + 'bibfilex' + DirectorySeparator;
  myConfigFile := 'bibfilex';
  {$endif}
  {$ifdef Win32}
  myHomeDir := GetAppConfigDir(False);
  myConfigFile := 'bibfilex.ini';
  lbAttNames.ScrollWidth := 0;
  miLine3GPG.Visible := False;
  miToolsEncrypt.Visible := False;
  miToolsDecrypt.Visible := False;
  fmMain.Color := clWhite;
  pnKindSearch.Color := clWhite;
  // Copy in HTML format is not working in Windows
  miItemsCopyCitationHtml.Visible := False;
  {$endif}
  {$ifdef Darwin}
  myHomeDir := GetEnvironmentVariable('HOME') + DirectorySeparator +
    'Library' + DirectorySeparator + 'Preferences' + DirectorySeparator;
  myConfigFile := 'bibfilex.plist';
  imgImageBack.Visible := False;
  miLine3GPG.Visible := False;
  miToolsEncrypt.Visible := False;
  miToolsDecrypt.Visible := False;
  fmMain.Color := clWhite;
  pnKindSearch.Color := clWhite;
  miFileOpen.ShortCut := miFileOpen.ShortCut - 12288;
  miFileCreateBibLatex.ShortCut := miFileCreateBibLatex.ShortCut - 12288;
  miFileImportFromBiblatex.ShortCut := miFileImportFromBiblatex.ShortCut - 12288;
  miFileExit.ShortCut := miFileExit.ShortCut - 12288;
  miItemsNew.ShortCut := miItemsNew.ShortCut - 12288;
  miItemsSave.ShortCut := miItemsSave.ShortCut - 12288;
  miItemsUndo.ShortCut := miItemsUndo.ShortCut - 12288;
  miItemsCreateKey.ShortCut := miItemsCreateKey.ShortCut - 12288;
  miItemsModifyKey.ShortCut := miItemsModifyKey.ShortCut - 12288;
  miItemsCopyKey.ShortCut := miItemsCopyKey.ShortCut - 12288;
  miItemsCopyCitationTxt.ShortCut := miItemsCopyCitationTxt.ShortCut - 12288;
  miItemsCopyCitationHtml.ShortCut := miItemsCopyCitationHtml.ShortCut - 12288;
  miItemsCreateKeywordList.ShortCut := miItemsCreateKeywordList.ShortCut - 12288;
  miItemsStoreKeyword.ShortCut := miItemsStoreKeyword.ShortCut - 12288;
  miItemsApplyKeyword.ShortCut := miItemsApplyKeyword.ShortCut - 12288;
  miItemsFilter.ShortCut := miItemsFilter.ShortCut - 12288;
  miItemsRemoveFilter.ShortCut := miItemsRemoveFilter.ShortCut - 12288;
  miItemsCopyCrossref.ShortCut := miItemsCopyCrossref.ShortCut - 12288;
  miItemsOrderByID.ShortCut := miItemsOrderByID.ShortCut - 12288;
  miItemsOrderByAuthor.ShortCut := miItemsOrderByAuthor.ShortCut - 12288;
  miItemsOrderByTitle.ShortCut := miItemsOrderByTitle.ShortCut - 12288;
  miItemsOrderByJournalTitle.ShortCut := miItemsOrderByJournalTitle.ShortCut - 12288;
  miItemsOrderByBookTitle.ShortCut := miItemsOrderByBookTitle.ShortCut - 12288;
  miItemsOrderByYear.ShortCut := miItemsOrderByYear.ShortCut - 12288;
  miItemsOrderByDate.ShortCut := miItemsOrderByDate.ShortCut - 12288;
  miToolChar.ShortCut := miToolChar.ShortCut - 12288;
  miToolKeyList.ShortCut := miToolKeyList.ShortCut - 12288;
  {$endif}
  if DirectoryExists(myHomeDir) = False then
  begin
    CreateDirUTF8(myHomeDir);
  end;
  // Load main form dimensions from ini file
  if FileExists(myHomeDir + myConfigFile) then
  begin
    MyIni := TIniFile.Create(myHomeDir + myConfigFile);
    try
      if MyIni.ReadString('bibfilex', 'maximize', '') = 'true' then
      begin
        fmMain.WindowState := wsMaximized;
      end
      else
      begin
        fmMain.Top := MyIni.ReadInteger('bibfilex', 'top', 0);
        fmMain.Left := MyIni.ReadInteger('bibfilex', 'left', 0);
        if MyIni.ReadInteger('bibfilex', 'width', 0) > 50 then
          fmMain.Width := MyIni.ReadInteger('bibfilex', 'width', 0);
        if MyIni.ReadInteger('bibfilex', 'heigth', 0) > 50 then
          fmMain.Height := MyIni.ReadInteger('bibfilex', 'heigth', 0);
      end;
      // Items grid width
      if MyIni.ReadInteger('bibfilex', 'itemsgridwidth', 200) > 10 then
        pnGrid.Width := MyIni.ReadInteger('bibfilex', 'itemsgridwidth', 200);
      // Order by
      if MyIni.ReadInteger('bibfilex', 'orderby', 0) = 0 then
        miItemsOrderByID.Checked := True
      else if MyIni.ReadInteger('bibfilex', 'orderby', 0) = 1 then
        miItemsOrderByAuthor.Checked := True
      else if MyIni.ReadInteger('bibfilex', 'orderby', 0) = 2 then
        miItemsOrderByTitle.Checked := True
      else if MyIni.ReadInteger('bibfilex', 'orderby', 0) = 3 then
        miItemsOrderByYear.Checked := True
      else if MyIni.ReadInteger('bibfilex', 'orderby', 0) = 4 then
        miItemsOrderByDate.Checked := True;
      if MyIni.ReadString('bibfilex', 'buffkeyword', '') <> '' then
        stBuffKeyword := MyIni.ReadString('bibfilex', 'buffkeyword', '');
      if MyIni.ReadString('bibfilex', 'savedkeywords', '') <> '' then
      begin
        stKeywordsList := MyIni.ReadString('bibfilex', 'savedkeywords', '');
        while UTF8Pos(#9, stKeywordsList) > 0 do
        begin
          cbFilterKey.Items.Add(UTF8Copy(stKeywordsList, 1,
            UTF8Pos(#9, stKeywordsList) - 1));
          stKeywordsList := UTF8Copy(stKeywordsList,
            UTF8Pos(#9, stKeywordsList) + 1, UTF8Length(stKeywordsList));
        end;
        cbFilterKey.Items.Add(stKeywordsList);
      end;
      // Field of smart filter
      cbFieldSmartFilter.ItemIndex :=
        MyIni.ReadInteger('bibfilex', 'fieldsmartflt', 0);
      // Menu of opened databases
      if MyIni.ReadString('bibfilex', 'lastfile1', '') <> '' then
      begin
        LastDatabase1 := MyIni.ReadString('bibfilex', 'lastfile1', '');
        miFileOpenLast1.Caption := ExtractFileNameOnly(LastDatabase1);
      end
      else
      begin
        miFileOpenLast1.Visible := False;
      end;
      if MyIni.ReadString('bibfilex', 'lastfile2', '') <> '' then
      begin
        LastDatabase2 := MyIni.ReadString('bibfilex', 'lastfile2', '');
        miFileOpenLast2.Caption := ExtractFileNameOnly(LastDatabase2);
      end
      else
      begin
        miFileOpenLast2.Visible := False;
      end;
      if MyIni.ReadString('bibfilex', 'lastfile3', '') <> '' then
      begin
        LastDatabase3 := MyIni.ReadString('bibfilex', 'lastfile3', '');
        miFileOpenLast3.Caption := ExtractFileNameOnly(LastDatabase3);
      end
      else
      begin
        miFileOpenLast3.Visible := False;
      end;
      if MyIni.ReadString('bibfilex', 'lastfile4', '') <> '' then
      begin
        LastDatabase4 := MyIni.ReadString('bibfilex', 'lastfile4', '');
        miFileOpenLast4.Caption := ExtractFileNameOnly(LastDatabase4);
      end
      else
      begin
        miFileOpenLast4.Visible := False;
      end;
      if ((miFileOpenLast1.Visible = False) and
        (miFileOpenLast2.Visible = False) and (miFileOpenLast3.Visible = False) and
        (miFileOpenLast4.Visible = False)) then
      begin
        miLineLastFile.Visible := False;
      end;
    finally
      MyIni.Free;
    end;
  end
  else
  begin
    miFileOpenLast1.Visible := False;
    miFileOpenLast2.Visible := False;
    miFileOpenLast3.Visible := False;
    miFileOpenLast4.Visible := False;
    miLineLastFile.Visible := False;
  end;
  // Start settings
  sgRequiredFields.ColWidths[0] := 120;
  sgRequiredFields.ColWidths[1] := sgRequiredFields.Width - 120;
  sgOptionalFields1.ColWidths[0] := 120;
  sgOptionalFields1.ColWidths[1] := sgOptionalFields1.Width - 120;
  sgOptionalFields2.ColWidths[0] := 120;
  sgOptionalFields2.ColWidths[1] := sgOptionalFields2.Width - 120;
  sgOptionalFields3.ColWidths[0] := 120;
  sgOptionalFields3.ColWidths[1] := sgOptionalFields3.Width - 120;
  // Initialize bookmarks list
  BookmarkList := TStringList.Create;
  for i := 0 to 9 do
    BookmarkList.Add('');
  // To disable proper menu items
  CloseDataTables;
  grItems.Color := $F7F7F7; //Very light gray
  grItems.SelectedColor := clGray;
  // Align dbOwned with its external label on various platforms
  lbOwned.Top := dbOwned.Top + dbOwned.Height - lbOwned.Height - 3;
end;

procedure TfmMain.FormActivate(Sender: TObject);
begin
  // On first activation of the form (see notes at the bottom)
  if blFirstAct = True then
  begin
    // Set options (not on Create: the fmOption form has not been created)
    SetOptions;
    // Open last database
    if fmOptions.cbOpenLast.Checked = True then
    begin
      if FileExistsUTF8(LastDatabase1) then
        OpenDataTables(LastDatabase1);
    end;
    blFirstAct := False;
  end;
end;

procedure TfmMain.FormClose(Sender: TObject; var CloseAction: TCloseAction);
var
  MyIni: TIniFile;
  i: integer;
  stKeywordsList: string;
begin
  // Restore normalm selection, to avoid error on exit
  grItems.Options := grItems.Options - [dgMultiselect, dgRowSelect];
  // Save all data
  SaveItems;
  if fmOptions.cbAutoSaveBibLatex.Checked = True then
    if flExportOnExit = True then
      ExportToBibLatex(ExtractFilePath(sqItems.FileName) +
        ExtractFileNameOnly(sqItems.FileName) + '.bib', False);
  try
    MyIni := TIniFile.Create(myHomeDir + myConfigFile);
    if fmMain.WindowState = wsMaximized then
    begin
      MyIni.WriteString('bibfilex', 'maximize', 'true');
    end
    else
    begin
      MyIni.WriteString('bibfilex', 'maximize', 'false');
      MyIni.WriteInteger('bibfilex', 'top', fmMain.Top);
      MyIni.WriteInteger('bibfilex', 'left', fmMain.Left);
      MyIni.WriteInteger('bibfilex', 'width', fmMain.Width);
      MyIni.WriteInteger('bibfilex', 'heigth', fmMain.Height);
    end;
    // Main grid width
    if pnGrid.Width > 10 then
      MyIni.WriteInteger('bibfilex', 'itemsgridwidth', pnGrid.Width)
    else
      MyIni.WriteInteger('bibfilex', 'itemsgridwidth', 200);
    // Order by
    if miItemsOrderByID.Checked = True then
      MyIni.WriteInteger('bibfilex', 'orderby', 0)
    else if miItemsOrderByAuthor.Checked = True then
      MyIni.WriteInteger('bibfilex', 'orderby', 1)
    else if miItemsOrderByTitle.Checked = True then
      MyIni.WriteInteger('bibfilex', 'orderby', 2)
    else if miItemsOrderByYear.Checked = True then
      MyIni.WriteInteger('bibfilex', 'orderby', 3)
    else if miItemsOrderByDate.Checked = True then
      MyIni.WriteInteger('bibfilex', 'orderby', 4);
    MyIni.WriteString('bibfilex', 'buffkeyword', stBuffKeyword);
    if cbFilterKey.Items.Count > 0 then
    begin
      stKeywordsList := '';
      for i := 0 to cbFilterKey.Items.Count - 1 do
      begin
        stKeywordsList := stKeywordsList + #9 + cbFilterKey.Items[i];
      end;
    end
    else
    begin
      stKeywordsList := '';
    end;
    MyIni.WriteString('bibfilex', 'savedkeywords', stKeywordsList);
    // Field of smart filter
    MyIni.WriteInteger('bibfilex', 'fieldsmartflt', cbFieldSmartFilter.ItemIndex);
    // Last databases
    if LastDatabase1 <> '' then
    begin
      MyIni.WriteString('bibfilex', 'lastfile1', LastDatabase1);
    end;
    if LastDatabase2 <> '' then
    begin
      MyIni.WriteString('bibfilex', 'lastfile2', LastDatabase2);
    end;
    if LastDatabase3 <> '' then
    begin
      MyIni.WriteString('bibfilex', 'lastfile3', LastDatabase3);
    end;
    if LastDatabase4 <> '' then
    begin
      MyIni.WriteString('bibfilex', 'lastfile4', LastDatabase4);
    end;
  finally
    MyIni.Free;
  end;
  // Destroy bookmarks list
  BookmarkList.Free;
end;

procedure TfmMain.FormDropFiles(Sender: TObject; const FileNames: array of string);
begin
  // Attach dropped files
  if miAttachNew.Enabled = True then
  begin
    CreateAttachment(FileNames);
    SetAttachmentMenuItems;
  end;
end;

procedure TfmMain.FormKeyDown(Sender: TObject; var Key: word; Shift: TShiftState);
begin
  // Select the tabs
  if ((sqItems.Active = True) and (pcPageControl.Visible = True)) then
  begin
    if ((Key = Ord('1')) and (Shift = [ssCtrl])) then
    begin
      pcPageControl.ActivePage := tsRequiredFields;
      sgRequiredFields.SetFocus;
      key := 0;
    end
    else if ((Key = Ord('2')) and (Shift = [ssCtrl])) then
    begin
      pcPageControl.ActivePage := tsOptionalFields1;
      sgOptionalFields1.SetFocus;
      key := 0;
    end
    else if ((Key = Ord('3')) and (Shift = [ssCtrl])) then
    begin
      pcPageControl.ActivePage := tsOptionalFields2;
      sgOptionalFields2.SetFocus;
      key := 0;
    end
    else if ((Key = Ord('4')) and (Shift = [ssCtrl])) then
    begin
      pcPageControl.ActivePage := tsOptionalFields3;
      sgOptionalFields3.SetFocus;
      key := 0;
    end
    else if ((Key = Ord('5')) and (Shift = [ssCtrl])) then
    begin
      pcPageControl.ActivePage := tsAbstract;
      meAbstract.SetFocus;
      key := 0;
    end
    else if ((Key = Ord('6')) and (Shift = [ssCtrl])) then
    begin
      pcPageControl.ActivePage := tsReview;
      meReview.SetFocus;
      key := 0;
    end;
    if Key = VK_F12 then
    begin
      edSmartFilter.SetFocus;
      key := 0;
    end;
    if ((Key = Ord('N')) and (Shift = [ssCtrl, ssShift])) then
      // New article
    begin
      if miItemsNew.Enabled = True then
      begin
        NewItem(0);
        key := 0;
      end;
    end;
    if ((grItems.Focused = False) and (meAbstract.Focused = False) and
      (meReview.Focused = False)) then
    begin
      // Ctrl + PgUp to move backward
      if ((Key = 33) and (Shift = [ssCtrl])) then
      begin
        sqItems.Prior;
        key := 0;
      end;
      // Ctrl + PgDn to move forward
      if ((Key = 34) and (Shift = [ssCtrl])) then
      begin
        sqItems.Next;
        key := 0;
      end;
      // Ctrl + Home to move to the first
      if ((Key = 36) and (Shift = [ssCtrl])) then
      begin
        sqItems.First;
        key := 0;
      end;
      // Ctrl + End to move to the last
      if ((Key = 35) and (Shift = [ssCtrl])) then
      begin
        sqItems.Last;
        key := 0;
      end;
    end;
    // Set bookmark (0 Key=48; 9 Key=57)
    if ((Key >= Ord('0')) and (Key <= Ord('9')) and
      (Shift = [ssShift, ssCtrl])) then
    begin
      if sqItems.RecordCount > 0 then
      begin
        BookmarkList[Key - 48] := sqItems.FieldByName('IDItems').AsString;
        sbStatusBar.SimpleText :=
          'Set bookmark n. ' + IntToStr(Key - 48) + ' to item n. ' +
          BookmarkList[Key - 48] + '.';
        key := 0;
      end;
    end
    // Go to bookmark
    else if ((Key >= Ord('0')) and (Key <= Ord('9')) and
      (Shift = [ssCtrl, ssAlt])) then
    begin
      if sqItems.Locate('IDItems', BookmarkList[Key - 48], []) = False then
        sbStatusBar.SimpleText :=
          'Item n. ' + BookmarkList[Key - 48] + ' not found.';
      key := 0;
    end;
  end;
end;

procedure TfmMain.FormResize(Sender: TObject);
begin
  // Save on resize, to fix a bug with DBGrid
  SaveItems;
end;

procedure TfmMain.grItemsDblClick(Sender: TObject);
var
  stFieldName: string;
  i: integer;
begin
  // Select the field of the read only grid in compilation area
  if ((pcPageControl.Visible = True) and (sqItems.RecordCount > 0)) then
  begin
    stFieldName := LowerCase(grItems.SelectedColumn.FieldName);
    for i := 0 to sgRequiredFields.RowCount - 1 do
    begin
      if LowerCase(sgRequiredFields.Cells[0, i]) = stFieldName then
      begin
        pcPageControl.ActivePage := tsRequiredFields;
        sgRequiredFields.SetFocus;
        sgRequiredFields.Row := i;
        Exit;
      end;
    end;
    for i := 0 to sgOptionalFields1.RowCount - 1 do
    begin
      if LowerCase(sgOptionalFields1.Cells[0, i]) = stFieldName then
      begin
        pcPageControl.ActivePage := tsOptionalFields1;
        sgOptionalFields1.SetFocus;
        sgOptionalFields1.Row := i;
        Exit;
      end;
    end;
    for i := 0 to sgOptionalFields2.RowCount - 1 do
    begin
      if LowerCase(sgOptionalFields2.Cells[0, i]) = stFieldName then
      begin
        pcPageControl.ActivePage := tsOptionalFields2;
        sgOptionalFields2.SetFocus;
        sgOptionalFields2.Row := i;
        Exit;
      end;
    end;
    for i := 0 to sgOptionalFields3.RowCount - 1 do
    begin
      if LowerCase(sgOptionalFields3.Cells[0, i]) = stFieldName then
      begin
        pcPageControl.ActivePage := tsOptionalFields3;
        sgOptionalFields3.SetFocus;
        sgOptionalFields3.Row := i;
        Exit;
      end;
    end;
  end;
end;

procedure TfmMain.grItemsPrepareCanvas(Sender: TObject; DataCol: integer;
  Column: TColumn; AState: TGridDrawState);
begin
  // Italics for items with attachments
  if Column.FieldName = 'IDItems' then
    if sqItems.FieldByName('AttName').AsString <> '' then
      grItems.Canvas.Font.Bold := True;
end;

procedure TfmMain.cbMultiselectChange(Sender: TObject);
begin
  // Activate multiselect in grid
  if cbMultiselect.Checked = True then
    grItems.Options := grItems.Options + [dgMultiselect, dgRowSelect]
  else
    grItems.Options := grItems.Options - [dgMultiselect, dgRowSelect];
end;

procedure TfmMain.bnAutoSizeClick(Sender: TObject);
begin
  // Resize the columns
  grItems.Options := grItems.Options - [dgAutoSizeColumns];
  grItems.ResetColWidths;
  grItems.Options := grItems.Options + [dgAutoSizeColumns];
end;

procedure TfmMain.cbItemsKindChange(Sender: TObject);
begin
  // Set fields in grids
  SaveItems;
  UnitSetFields.SetFieldsInGrid;
  flDataEdit := True;
  ActivateSaving;
end;

procedure TfmMain.cbItemsKindSelect(Sender: TObject);
begin
  // Change kind of item saving old data
  sqItems.Edit;
  sqItems.FieldByName('entrytype').AsString := LowerCase(cbItemsKind.Text);
  sqItems.Post;
  sqItems.ApplyUpdates;
  LoadItems;
end;

procedure TfmMain.dbOwnedChange(Sender: TObject);
begin
  // Activate edit in all the grids
  flDataEdit := True;
  ActivateSaving;
end;

procedure TfmMain.DataGridsSetEditText(Sender: TObject; ACol, ARow: integer;
  const Value: string);
begin
  // Activate edit in all the grids
  flDataEdit := True;
  ActivateSaving;
end;

procedure TfmMain.Abstract_ReviewChange(Sender: TObject);
begin
  // Set edit flag on true when a memo is edited
  flDataEdit := True;
  ActivateSaving;
end;

procedure TfmMain.cbItemsKindKeyDown(Sender: TObject; var Key: word;
  Shift: TShiftState);
begin
  // Set focus in compilation grid
  if ((Key = 13) or (Key = 9)) then
  begin
    pcPageControl.TabIndex := 0;
    sgRequiredFields.SetFocus;
    key := 0;
  end;
end;

procedure TfmMain.sgOptionalFields1ValidateEntry(Sender: TObject;
  aCol, aRow: integer; const OldValue: string; var NewValue: string);
var
  myDate: TDate;
begin
  // Date validation in sgOptionalFields1
  if sgOptionalFields1.Cells[1, aRow] <> '' then
  begin
    if ((sgOptionalFields1.Cells[0, aRow] = 'Date') or
      (sgOptionalFields1.Cells[0, aRow] = 'Urldate') or
      (sgOptionalFields1.Cells[0, aRow] = 'Eventdate')) then
      try
        myDate := StrToDate(sgOptionalFields1.Cells[1, aRow]);
      except
        MessageDlg('The date is not valid.', mtWarning, [mbOK], 0);
        sgOptionalFields1.Cells[1, aRow] := '';
      end;
  end;
end;

procedure TfmMain.sgOptionalFields2ValidateEntry(Sender: TObject;
  aCol, aRow: integer; const OldValue: string; var NewValue: string);
var
  myDate: TDate;
begin
  // Date validation in sgOptionalFields2
  if sgOptionalFields2.Cells[1, aRow] <> '' then
  begin
    if ((sgOptionalFields2.Cells[0, aRow] = 'Date') or
      (sgOptionalFields2.Cells[0, aRow] = 'Urldate') or
      (sgOptionalFields2.Cells[0, aRow] = 'Eventdate')) then
      try
        myDate := StrToDate(sgOptionalFields2.Cells[1, aRow]);
      except
        MessageDlg('The date is not valid.', mtWarning, [mbOK], 0);
        sgOptionalFields2.Cells[1, aRow] := '';
      end;
  end;
end;

procedure TfmMain.sgOptionalFields3ValidateEntry(Sender: TObject;
  aCol, aRow: integer; const OldValue: string; var NewValue: string);
var
  myDate: TDate;
begin
  // Date validation in sgOptionalFields3
  if sgOptionalFields3.Cells[1, aRow] <> '' then
  begin
    if ((sgOptionalFields3.Cells[0, aRow] = 'Date') or
      (sgOptionalFields3.Cells[0, aRow] = 'Urldate') or
      (sgOptionalFields3.Cells[0, aRow] = 'Eventdate')) then
      try
        myDate := StrToDate(sgOptionalFields3.Cells[1, aRow]);
      except
        MessageDlg('The date is not valid.', mtWarning, [mbOK], 0);
        sgOptionalFields3.Cells[1, aRow] := '';
      end;
  end;
end;

procedure TfmMain.sgRequiredFieldsKeyDown(Sender: TObject; var Key: word;
  Shift: TShiftState);
begin
  // Text suggestion
  if ((Key = 32) and (ssCtrl in Shift)) then
  begin
    ActivateSuggestion(sgRequiredFields);
    key := 0;
  end
  // Activate save on paste
  else if ((Key = Ord('V')) and (ssCtrl in Shift)) then
  begin
    flDataEdit := True;
    ActivateSaving;
  end;
end;

procedure TfmMain.sgRequiredFieldsKeyUp(Sender: TObject; var Key: word;
  Shift: TShiftState);
var
  Editor: TStringCellEditor;
begin
  if key = 13 then
  begin
    if sgRequiredFields.EditorMode = True then
    begin
      Editor := TStringCellEditor(sgRequiredFields.Editor);
      Editor.SelStart := UTF8Length(Editor.Text);
      key := 0;
    end;
  end;
end;

procedure TfmMain.sgRequiredFieldsPrepareCanvas(Sender: TObject;
  aCol, aRow: integer; aState: TGridDrawState);
begin
  // Title in bold
  if ((aCol = 1) and (LowerCase(sgRequiredFields.Cells[0, aRow]) = 'title')) then
  begin
    sgRequiredFields.Canvas.Font.Style := [fsBold];
  end
  else if ((aCol = 1) and (LowerCase(sgRequiredFields.Cells[0, aRow]) =
    'bibtexkey')) then
  begin
    sgRequiredFields.Canvas.Brush.Color := clBtnFace;
  end;
end;

procedure TfmMain.sgOptionalFields1KeyDown(Sender: TObject; var Key: word;
  Shift: TShiftState);
begin
  // Text suggestion
  if ((Key = 32) and (ssCtrl in Shift)) then
  begin
    ActivateSuggestion(sgOptionalFields1);
    key := 0;
  end
  // Activate save on paste
  else if ((Key = Ord('V')) and (ssCtrl in Shift)) then
  begin
    flDataEdit := True;
    ActivateSaving;
  end;
end;

procedure TfmMain.sgOptionalFields1KeyUp(Sender: TObject; var Key: word;
  Shift: TShiftState);
var
  Editor: TStringCellEditor;
begin
  if key = 13 then
  begin
    if sgOptionalFields1.EditorMode = True then
    begin
      Editor := TStringCellEditor(sgOptionalFields1.Editor);
      Editor.SelStart := UTF8Length(Editor.Text);
      key := 0;
    end;
  end;
end;

procedure TfmMain.sgOptionalFields2KeyDown(Sender: TObject; var Key: word;
  Shift: TShiftState);
begin
  // Text suggestion
  if ((Key = 32) and (ssCtrl in Shift)) then
  begin
    ActivateSuggestion(sgOptionalFields2);
    key := 0;
  end
  // Activate save on paste
  else if ((Key = Ord('V')) and (ssCtrl in Shift)) then
  begin
    flDataEdit := True;
    ActivateSaving;
  end;
end;

procedure TfmMain.sgOptionalFields2KeyUp(Sender: TObject; var Key: word;
  Shift: TShiftState);
var
  Editor: TStringCellEditor;
begin
  if key = 13 then
  begin
    if sgOptionalFields2.EditorMode = True then
    begin
      Editor := TStringCellEditor(sgOptionalFields2.Editor);
      Editor.SelStart := UTF8Length(Editor.Text);
      key := 0;
    end;
  end;
end;

procedure TfmMain.sgOptionalFields3KeyDown(Sender: TObject; var Key: word;
  Shift: TShiftState);
begin
  // Text suggestion
  if ((Key = 32) and (ssCtrl in Shift)) then
  begin
    ActivateSuggestion(sgOptionalFields3);
    key := 0;
  end
  // Activate save on paste
  else if ((Key = Ord('V')) and (ssCtrl in Shift)) then
  begin
    flDataEdit := True;
    ActivateSaving;
  end;
end;

procedure TfmMain.sgOptionalFields3KeyUp(Sender: TObject; var Key: word;
  Shift: TShiftState);
var
  Editor: TStringCellEditor;
begin
  if key = 13 then
  begin
    if sgOptionalFields3.EditorMode = True then
    begin
      Editor := TStringCellEditor(sgOptionalFields3.Editor);
      Editor.SelStart := UTF8Length(Editor.Text);
      key := 0;
    end;
  end;
end;

procedure TfmMain.sgRequiredFieldsResize(Sender: TObject);
begin
  // Resize grid column
  sgRequiredFields.ColWidths[0] := 120;
  sgRequiredFields.ColWidths[1] := sgRequiredFields.Width - 120;
end;

procedure TfmMain.sgRequiredFieldsSelectCell(Sender: TObject;
  aCol, aRow: integer; var CanSelect: boolean);
begin
  // bibtex key read only
  if LowerCase(sgRequiredFields.Cells[0, aRow]) = 'bibtexkey' then
    CanSelect := False;
end;

procedure TfmMain.sgOptionalFields1Resize(Sender: TObject);
begin
  // Resize grid column
  sgOptionalFields1.ColWidths[0] := 120;
  sgOptionalFields1.ColWidths[1] := sgOptionalFields1.Width - 120;
end;

procedure TfmMain.sgOptionalFields2Resize(Sender: TObject);
begin
  // Resize grid column
  sgOptionalFields2.ColWidths[0] := 120;
  sgOptionalFields2.ColWidths[1] := sgOptionalFields2.Width - 120;
end;

procedure TfmMain.sgOptionalFields3Resize(Sender: TObject);
begin
  // Resize grid column
  sgOptionalFields3.ColWidths[0] := 120;
  sgOptionalFields3.ColWidths[1] := sgOptionalFields3.Width - 120;
end;

procedure TfmMain.lbAttNamesDblClick(Sender: TObject);
begin
  // New or open attachment
  if lbAttNames.Items.Count > 0 then
  begin
    if miAttachOpen.Enabled = True then
      miAttachOpenClick(nil);
  end
  else
  begin
    if miAttachNew.Enabled = True then
      miAttachNewClick(nil);
  end;
end;

procedure TfmMain.meAbstractKeyDown(Sender: TObject; var Key: word; Shift: TShiftState);
begin
  // Zoom in Abstract
  if meAbstract.Font.Size = 0 then
    meAbstract.Font.Size := 10;
  if ((Key = 187) and (Shift = [ssCtrl])) then
  begin
    meAbstract.Font.Size := meAbstract.Font.Size + 1;
    key := 0;
  end
  else if ((Key = 189) and (Shift = [ssCtrl])) then
  begin
    if meAbstract.Font.Size > 6 then
      meAbstract.Font.Size := meAbstract.Font.Size - 1;
    key := 0;
  end
  // Clean single CR in Abstract
  else if ((Key = Ord('H')) and (Shift = [ssCtrl, ssShift])) then
  begin
    if ((sqItems.RecordCount > 0) and (meAbstract.Text <> '')) then
    begin
      sqItems.Edit;
      meAbstract.Text := CleanAbstractReview(meAbstract.Text);
    end;
  end
  // Insert the bibtex key
  else if ((Key = Ord('Q')) and (Shift = [ssCtrl, ssShift])) then
  begin
    InsertKeyAbstRew(meAbstract);
  end;
end;

procedure TfmMain.meAbstractMouseWheel(Sender: TObject; Shift: TShiftState;
  WheelDelta: integer; MousePos: TPoint; var Handled: boolean);
begin
  // Change Abstract Zoom with mouse weel
  if meAbstract.Font.Size = 0 then
    meAbstract.Font.Size := 10;
  if Shift = [ssCtrl] then
  begin
    if WheelDelta > 0 then
    begin
      meAbstract.Font.Size := meAbstract.Font.Size + 1;
    end
    else
    begin
      if meAbstract.Font.Size > 6 then
      begin
        meAbstract.Font.Size := meAbstract.Font.Size - 1;
      end;
    end;
  end;
end;

procedure TfmMain.meReviewKeyDown(Sender: TObject; var Key: word; Shift: TShiftState);
begin
  // Zoom in Review
  if meReview.Font.Size = 0 then
    meReview.Font.Size := 10;
  if ((Key = 187) and (Shift = [ssCtrl])) then
  begin
    meReview.Font.Size := meReview.Font.Size + 1;
    key := 0;
  end
  else if ((Key = 189) and (Shift = [ssCtrl])) then
  begin
    if meReview.Font.Size > 6 then
      meReview.Font.Size := meReview.Font.Size - 1;
    key := 0;
  end
  // Clean single CR in Review
  else if ((Key = Ord('H')) and (Shift = [ssCtrl, ssShift])) then
  begin
    if ((sqItems.RecordCount > 0) and (meReview.Text <> '')) then
    begin
      sqItems.Edit;
      meReview.Text := CleanAbstractReview(meReview.Text);
    end;
  end
  // Insert the bibtex key
  else if ((Key = Ord('Q')) and (Shift = [ssCtrl, ssShift])) then
  begin
    InsertKeyAbstRew(meReview);
  end;
end;

procedure TfmMain.meReviewMouseWheel(Sender: TObject; Shift: TShiftState;
  WheelDelta: integer; MousePos: TPoint; var Handled: boolean);
begin
  // Change Review Zoom with mouse weel
  if meReview.Font.Size = 0 then
    meReview.Font.Size := 10;
  if Shift = [ssCtrl] then
  begin
    if WheelDelta > 0 then
    begin
      meReview.Font.Size := meReview.Font.Size + 1;
    end
    else
    begin
      if meReview.Font.Size > 6 then
      begin
        meReview.Font.Size := meReview.Font.Size - 1;
      end;
    end;
  end;
end;

//***************************************************************
//**********************  MENU PROCEDURES  **********************
//***************************************************************

procedure TfmMain.miFileNewClick(Sender: TObject);
var
  stFileName: string;
begin
  // Create a file
  sdSaveFile.Title := 'Save Bibfilex file';
  sdSaveFile.Filter := 'Bibfilex files|*.bfx|All files|*';
  sdSaveFile.DefaultExt := '.bfx';
  sdSaveFile.FileName := '';
  if sdSaveFile.Execute = True then
  begin
    stFileName := sdSaveFile.FileName;
    {$ifdef Darwin}
    // On OSX, the component TSaveDialog indicates the path of a file
    // in the home directory as /. So...
    if ExtractFilePath(stFileName) = '/' then
      stFileName := GetEnvironmentVariable('HOME') + stFileName;
    {$endif}
    CreateDataTables(stFileName);
    OpenDataTables(stFileName);
  end;
end;

procedure TfmMain.miFileOpenClick(Sender: TObject);
begin
  // Open a file
  odOpenFile.Title := 'Open Bibfilex file';
  odOpenFile.Filter := 'Bibfilex files|*.bfx|All files|*';
  odOpenFile.Options := odOpenFile.Options - [ofAllowMultiSelect];
  odOpenFile.DefaultExt := '.bfx';
  odOpenFile.FileName := '';
  if odOpenFile.Execute = True then
  begin
    OpenDataTables(odOpenFile.FileName);
  end;
end;

procedure TfmMain.miFileCloseClick(Sender: TObject);
begin
  // Close file
  CloseDataTables;
end;

procedure TfmMain.miFileCopyAsClick(Sender: TObject);
var
  SearchRec: TSearchRec;
  AttOrigDir, AttDestDir: string;
begin
  // Copy current file with its possibile attachments
  sdSaveFile.Title := 'Copy Bibfilex file';
  sdSaveFile.Filter := '*.bfx|*';
  sdSaveFile.DefaultExt := '.bfx';
  sdSaveFile.FileName := '';
  if sdSaveFile.Execute then
    try
      Screen.Cursor := crHourGlass;
      Application.ProcessMessages;
      try
        CopyFile(sqItems.FileName, sdSaveFile.FileName);
        AttOrigDir := ExtractFileNameWithoutExt(sqItems.FileName);
        if DirectoryExistsUTF8(AttOrigDir) = True then
          try
            AttDestDir := ExtractFileNameWithoutExt(sdSaveFile.FileName);
            CreateDirUTF8(AttDestDir);
            // faSysFile (= normal file) to avoid that also the directory is found
            if FindFirst(AttOrigDir + DirectorySeparator + '*',
              faSysFile, SearchRec) = 0 then
              repeat
                CopyFile(AttOrigDir + DirectorySeparator + SearchRec.Name,
                  AttDestDir + DirectorySeparator + SearchRec.Name);
              until
                FindNext(SearchRec) <> 0;
          finally
            FindClose(SearchRec);
          end;
      except
        MessageDlg('Error in copying the file in use.',
          mtError, [mbOK], 0);
      end;
    finally
      Screen.Cursor := crDefault;
    end;
end;

procedure TfmMain.miFileCreateBibLatexClick(Sender: TObject);
begin
  // Create Biblatex file
  SaveItems;
  ExportToBiblatex(ExtractFilePath(sqItems.FileName) +
    ExtractFileNameOnly(sqItems.FileName) + '.bib', False);
  flExportOnExit := False;
end;

procedure TfmMain.miFileExportToBiblatexClick(Sender: TObject);
var
  stFileName: string;
begin
  // Export to Biblatex file
  SaveItems;
  sdSaveFile.Title := 'Export to Biblatex file';
  sdSaveFile.Filter := 'Biblatex files|*.bib*|All files|*';
  sdSaveFile.DefaultExt := '.bib';
  sdSaveFile.FileName := '';
  if sdSaveFile.Execute = True then
  begin
    stFileName := sdSaveFile.FileName;
    {$ifdef Darwin}
    // On OSX, the component TSaveDialog indicates the path of a file
    // in the home directory as /. So...
    if ExtractFilePath(stFileName) = '/' then
      stFileName := GetEnvironmentVariable('HOME') + stFileName;
    {$endif}
    if FileExistsUTF8(ExtractFileNameWithoutExt(stFileName) + '.bfx') = True then
    begin
      MessageDlg('There''s a Bibfilex file with the same name; ' +
        'change the name of the Biblatex file ' +
        'to avoid possible attachments overlap.',
        mtWarning, [mbOK], 0);
      Abort;
    end;
    ExportToBibLatex(stFileName, True);
  end;
end;

procedure TfmMain.miFileImportFromBiblatexClick(Sender: TObject);
var
  i: integer;
begin
  // Import from Biblatex file
  SaveItems;
  odOpenFile.Title := 'Import from Biblatex file';
  odOpenFile.Filter := 'Biblatex files|*.bib*|All files|*';
  odOpenFile.Options := odOpenFile.Options + [ofAllowMultiSelect];
  odOpenFile.DefaultExt := '.bib';
  odOpenFile.FileName := '';
  if odOpenFile.Execute = True then
  begin
    for i := 0 to odOpenFile.Files.Count - 1 do
    begin
      ImportBibLatex(odOpenFile.Files[i]);
    end;
  end;
end;

procedure TfmMain.miFileExportToBifilexClick(Sender: TObject);
begin
  // Export to a Bibfilex file
  SaveItems;
  odOpenFile.Title := 'Export to Bibfilex file';
  odOpenFile.Filter := 'Bibfilex files|*.bfx|All files|*';
  odOpenFile.Options := odOpenFile.Options - [ofAllowMultiSelect];
  odOpenFile.DefaultExt := '.bfx';
  odOpenFile.FileName := '';
  if odOpenFile.Execute = True then
  begin
    ExportToBibfilex(odOpenFile.FileName);
  end;
end;

procedure TfmMain.miFileImportFromBibfilexClick(Sender: TObject);
begin
  // Import a Bibfilex file
  SaveItems;
  odOpenFile.Title := 'Import from Bibfilex file';
  odOpenFile.Filter := 'Bibfilex files|*.bfx|All files|*';
  odOpenFile.Options := odOpenFile.Options - [ofAllowMultiSelect];
  odOpenFile.DefaultExt := '.bfx';
  odOpenFile.FileName := '';
  if odOpenFile.Execute = True then
  begin
    ImportFromBibfilex(odOpenFile.FileName);
  end;
end;

procedure TfmMain.miFileOpenLast1Click(Sender: TObject);
begin
  // Open last1 file
  if FileExistsUTF8(LastDatabase1) then
  begin
    OpenDataTables(LastDatabase1);
  end
  else
  begin
    MessageDlg('The file is not available.', mtError, [mbOK], 0);
  end;
end;

procedure TfmMain.miFileOpenLast2Click(Sender: TObject);
begin
  // Open last2 file
  if FileExistsUTF8(LastDatabase2) then
  begin
    OpenDataTables(LastDatabase2);
  end
  else
  begin
    MessageDlg('The file is not available.', mtError, [mbOK], 0);
  end;
end;

procedure TfmMain.miFileOpenLast3Click(Sender: TObject);
begin
  // Open last3 file
  if FileExistsUTF8(LastDatabase3) then
  begin
    OpenDataTables(LastDatabase3);
  end
  else
  begin
    MessageDlg('The file is not available.', mtError, [mbOK], 0);
  end;
end;

procedure TfmMain.miFileOpenLast4Click(Sender: TObject);
begin
  // Open last4 file
  if FileExistsUTF8(LastDatabase4) then
  begin
    OpenDataTables(LastDatabase4);
  end
  else
  begin
    MessageDlg('The file is not available.', mtError, [mbOK], 0);
  end;
end;

procedure TfmMain.miFileExitClick(Sender: TObject);
begin
  // Quit
  SaveItems;
  Close;
end;

procedure TfmMain.miItemsNewClick(Sender: TObject);
begin
  // New item: book; the new article is in the shortcuts
  NewItem(1);
end;

procedure TfmMain.miItemsSaveClick(Sender: TObject);
begin
  // Save item
  SaveItems;
end;

procedure TfmMain.miItemsDeleteClick(Sender: TObject);
var
  i: integer;
begin
  // Delete item
  if cbMultiselect.Checked = False then
  begin
    if MessageDlg('Delete the current item?', mtConfirmation, [mbOK, mbCancel], 0) =
      mrCancel then
      Abort;
    sqItems.Delete;
    sqItems.ApplyUpdates;
  end
  else
  begin
    if MessageDlg('Delete the selected items?', mtConfirmation,
      [mbOK, mbCancel], 0) = mrCancel then
      Abort;
    if grItems.SelectedRows.Count > 0 then
      try
        Screen.Cursor := crHourGlass;
        Application.ProcessMessages;
        with grItems.DataSource.DataSet do
        begin
          for i := 0 to grItems.SelectedRows.Count - 1 do
          begin
            GotoBookmark(Pointer(grItems.SelectedRows.Items[i]));
            sqItems.Delete;
          end;
          sqItems.ApplyUpdates;
        end;
        grItems.SelectedRows.Clear;
      finally
        Screen.Cursor := crDefault;
      end;
  end;
end;

procedure TfmMain.miItemsUndoClick(Sender: TObject);
begin
  // Undo editing
  if MessageDlg('Undo changes to the current item?', mtConfirmation,
    [mbOK, mbCancel], 0) = mrCancel then
    Abort;
  if dsItems.State in [dsEdit] then
    sqItems.Cancel;
  LoadItems;
end;

procedure TfmMain.miAttachNewClick(Sender: TObject);
begin
  // Insert attachment in the current item
  SaveItems;
  odOpenFile.Title := 'Open attachment';
  odOpenFile.Filter := 'All files|*';
  odOpenFile.Options := odOpenFile.Options - [ofAllowMultiSelect];
  odOpenFile.DefaultExt := '';
  odOpenFile.FileName := '';
  if odOpenFile.Execute = True then
  begin
    CreateAttachment(odOpenFile.FileName);
    SetAttachmentMenuItems;
  end;
end;

procedure TfmMain.miAttachOpenClick(Sender: TObject);
var
  myUnZip: TUnZipper;
  myList: TStringList;
  AttDir, OrigFileName, OutDir: string;
begin
  // Open and Save As attachment (miAttachOpen and miAttachSaveAs)
  if lbAttNames.ItemIndex < 0 then
  begin
    MessageDlg('No attachment is selected.', mtWarning, [mbOK], 0);
    Abort;
  end;
  AttDir := ExtractFileNameWithoutExt(sqItems.FileName);
  if DirectoryExistsUTF8(AttDir) = False then
  begin
    MessageDlg('The attachment directory is not available.',
      mtWarning, [mbOK], 0);
    Abort;
  end;
  if ((Sender = miAttachSaveAs) or (Sender = pmAttSaveAs)) then
  begin
    sdSelDirDialog.Title := 'Save attachment';
    if sdSelDirDialog.Execute then
      OutDir := sdSelDirDialog.FileName
    else
      Abort;
  end
  // The other Sender could be miAttachOpen or lbAttNames (double clic)
  else
  begin
    OutDir := GetTempDir;
  end;
  try
    Screen.Cursor := crHourGlass;
    Application.ProcessMessages;
    try
      myUnZip := TUnZipper.Create;
      myList := TStringList.Create;
      AttDir := ExtractFileNameWithoutExt(sqItems.FileName);
      OrigFileName := lbAttNames.Items.ValueFromIndex[lbAttNames.ItemIndex];
      myList.Add(OrigFileName);
      myUnZip.OutputPath := OutDir;
      myUnZip.FileName := AttDir + DirectorySeparator +
        sqItems.FieldByName('IDItems').AsString + '-' +
        ExtractFileName(OrigFileName) + '.zip';
      myUnZip.UnZipFiles(myList);
      if ((Sender <> miAttachSaveAs) and (Sender <> pmAttSaveAs)) then
        OpenDocument(OutDir + DirectorySeparator + OrigFileName);
    except
      MessageDlg('It was not possible to open the file.', mtError, [mbOK], 0);
    end;
  finally
    myUnZip.Free;
    myList.Free;
    Screen.Cursor := crDefault;
  end;
end;

procedure TfmMain.miAttachDeleteClick(Sender: TObject);
var
  AttDir, OrigFileName: string;
begin
  // Delete attachment
  if lbAttNames.ItemIndex < 0 then
  begin
    MessageDlg('No attachment is selected.', mtWarning, [mbOK], 0);
    Abort;
  end;
  if MessageDlg('Delete the selected attachment?', mtConfirmation,
    [mbOK, mbCancel], 0) = mrCancel then
    Abort;
  OrigFileName := lbAttNames.Items.ValueFromIndex[lbAttNames.ItemIndex];
  lbAttNames.Items.Delete(lbAttNames.ItemIndex);
  sqItems.Edit;
  sqItems.FieldByName('AttName').AsString := lbAttNames.Items.Text;
  sqItems.Post;
  sqItems.ApplyUpdates;
  AttDir := ExtractFileNameWithoutExt(sqItems.FileName);
  if DirectoryExistsUTF8(AttDir) = False then
  begin
    MessageDlg('Attachment directory is not available.', mtWarning, [mbOK], 0);
    Abort;
  end;
  try
    DeleteFileUTF8(AttDir + DirectorySeparator +
      sqItems.FieldByName('IDItems').AsString + '-' +
      ExtractFileName(OrigFileName) + '.zip');
    if IsDirectoryEmpty(AttDir) = True then
      DeleteDirectory(AttDir, False);
    SetAttachmentMenuItems;
  except
    MessageDlg('It is not possible to delete the file.', mtError, [mbOK], 0);
  end;
end;

procedure TfmMain.miItemsCreateKeyClick(Sender: TObject);
var
  i: integer;
  oldBibTex: string;
begin
  // Create BibTex key
  pcPageControl.ActivePageIndex := 0;
  i := 0;
  while LowerCase(sgRequiredFields.Cells[0, i]) <> 'bibtexkey' do
    Inc(i);
  if sgRequiredFields.Cells[1, i] <> '' then
  begin
    oldBibTex := sgRequiredFields.Cells[1, i];
    if MessageDlg('Overwrite the existing bibtex key?', mtWarning,
      [mbOK, mbCancel], 0) = mrCancel then
      Abort;
  end;
  SaveItems;
  sqItems.Edit;
  sgRequiredFields.Cells[1, i] := CreateBibTexKey(True, True);
  flDataEdit := True;
  ActivateSaving;
  // Replace the possibile bibtex keys in Abstract and Reviews
  if UTF8Pos(oldBibTex, meAbstract.Text) > 0 then
  begin
    meAbstract.Text := StringReplace(meAbstract.Text, '{' + oldBibTex +
      '}', '{' + sgRequiredFields.Cells[1, i] + '}', [rfReplaceAll, rfIgnoreCase]);
  end;
  if UTF8Pos(oldBibTex, meReview.Text) > 0 then
  begin
    meReview.Text := StringReplace(meReview.Text, '{' + oldBibTex +
      '}', '{' + sgRequiredFields.Cells[1, i] + '}', [rfReplaceAll, rfIgnoreCase]);
  end;
end;

procedure TfmMain.miItemsModifyKeyClick(Sender: TObject);
var
  stKey, oldBibTex: string;
  i: integer;
begin
  // Modify BibTex key
  SaveItems;
  pcPageControl.ActivePageIndex := 0;
  stKey := sqItems.FieldByName('bibtexkey').AsString;
  if InputQuery('Modify BibTex key',
    'Insert the new BibTex key for the current item.', stKey) = True then
  begin
    if stKey = '' then
    begin
      MessageDlg('An empty key is not allowed.', mtWarning, [mbOK], 0);
      Abort;
    end;
    i := 0;
    while LowerCase(sgRequiredFields.Cells[0, i]) <> 'bibtexkey' do
      Inc(i);
    oldBibTex := sgRequiredFields.Cells[1, i];
    sqItems.Edit;
    sgRequiredFields.Cells[1, i] :=
      CheckUniqueBibTexKey(stKey, sqItems.FieldByName('IDItems').AsInteger);
    flDataEdit := True;
    ActivateSaving;
    // Replace the possibile bibtex keys in Abstract and Reviews
    if UTF8Pos(oldBibTex, meAbstract.Text) > 0 then
    begin
      meAbstract.Text := StringReplace(meAbstract.Text, '{' + oldBibTex +
        '}', '{' + sgRequiredFields.Cells[1, i] + '}', [rfReplaceAll, rfIgnoreCase]);
    end;
    if UTF8Pos(oldBibTex, meReview.Text) > 0 then
    begin
      meReview.Text := StringReplace(meReview.Text, '{' + oldBibTex +
        '}', '{' + sgRequiredFields.Cells[1, i] + '}', [rfReplaceAll, rfIgnoreCase]);
    end;
  end;
end;

procedure TfmMain.miItemsCopyKeyClick(Sender: TObject);
var
  i: integer;
  stData, stPages: string;
begin
  // Copy BibTex key
  SaveItems;
  pcPageControl.ActivePageIndex := 0;
  if cbMultiselect.Checked = True then
  begin
    if fmOptions.clFieldsShown.Checked[fmOptions.clFieldsShown.Items.IndexOf(
      'bibtexkey')] = False then
    begin
      MessageDlg('It is not possible to copy in the clipboard ' +
        'the citations since the bibtexkey column ' +
        'is not shown in the items left grid; ' +
        'check its name in the options to make it visible.',
        mtWarning, [mbOK], 0);
      Exit;
    end
    else
    begin
      if grItems.SelectedRows.Count > 0 then
      begin
        stData := '';
        with grItems.DataSource.DataSet do
        begin
          for i := 0 to grItems.SelectedRows.Count - 1 do
          begin
            GotoBookmark(Pointer(grItems.SelectedRows.Items[i]));
            if stData <> '' then
              stData := stData + LineEnding;
            stData := stData + '\cite{' + sqItems.FieldByName(
              'bibtexkey').AsString + '}.';
          end;
        end;
        Clipboard.AsText := stData;
      end;
    end;
  end
  else
  begin
    for i := 0 to sgRequiredFields.RowCount - 1 do
      if LowerCase(sgRequiredFields.Cells[0, i]) = 'bibtexkey' then
      begin
        if InputQuery('Postnote', 'Insert postnote (like pages).', stPages) = True then
        begin
          if stPages <> '' then
            Clipboard.AsText :=
              '\cite[' + stPages + ']{' + sgRequiredFields.Cells[1, i] + '}.'
          else
            Clipboard.AsText := '\cite{' + sgRequiredFields.Cells[1, i] + '}.';
        end
        else
          Clipboard.AsText := sgRequiredFields.Cells[1, i];
        Exit;
      end;
  end;
end;

procedure TfmMain.miItemsCopyCitationTxtClick(Sender: TObject);
var
  i: integer;
  stData: string;
begin
  // Copy citation as LaTex text
  SaveItems;
  if cbMultiselect.Checked = True then
  begin
    if fmOptions.clFieldsShown.Checked[fmOptions.clFieldsShown.Items.IndexOf(
      'bibtexkey')] = False then
    begin
      MessageDlg('It is not possible to copy in the clipboard ' +
        'the citations since the bibtexkey column ' +
        'is not shown in the items left grid; ' +
        'check its name in the options to make it visible.',
        mtWarning, [mbOK], 0);
      Exit;
    end
    else
    begin
      if grItems.SelectedRows.Count > 0 then
      begin
        stData := '';
        with grItems.DataSource.DataSet do
        begin
          for i := 0 to grItems.SelectedRows.Count - 1 do
          begin
            GotoBookmark(Pointer(grItems.SelectedRows.Items[i]));
            stData := stData + CreateCitation(
              sqItems.FieldByName('bibtexkey').AsString, True, False,
              False, False, 0, False, '') + '.' + LineEnding;
          end;
        end;
        Clipboard.AsText := stData;
      end;
    end;
  end
  else
  begin
    Clipboard.AsText := CreateCitation(sqItems.FieldByName('bibtexkey').AsString,
      True, False, False, False, 0, False, '') + '.';
  end;
end;

procedure TfmMain.miItemsCopyCitationHtmlClick(Sender: TObject);
var
  Fid: TClipboardFormat;
  Strm: TMemoryStream;
  stHTML: string;
  i: integer;
begin
  // Copy citation as HTML
  // Only for Linux and MacOS
  SaveItems;
  if cbMultiselect.Checked = True then
  begin
    if fmOptions.clFieldsShown.Checked[fmOptions.clFieldsShown.Items.IndexOf(
      'bibtexkey')] = False then
    begin
      MessageDlg('It is not possible to copy in the clipboard ' +
        'the citations since the bibtexkey column ' +
        'is not shown in the items left grid; ' +
        'check its name in the options to make it visible.',
        mtWarning, [mbOK], 0);
      Exit;
    end
    else
    begin
      if grItems.SelectedRows.Count > 0 then
      begin
        stHTML := '';
        with grItems.DataSource.DataSet do
        begin
          for i := 0 to grItems.SelectedRows.Count - 1 do
          begin
            GotoBookmark(Pointer(grItems.SelectedRows.Items[i]));
            stHTML := stHTML + CreateCitation(
              sqItems.FieldByName('bibtexkey').AsString, True, False,
              True, False, 0, False, '') + '.<p>';
          end;
        end;
      end;
    end;
  end
  else
  begin
    stHTML := CreateCitation(sqItems.FieldByName('bibtexkey').AsString,
      True, False, True, False, 0, False, '') + '.';
  end;
  try
    Strm := TMemoryStream.Create;
    Strm.WriteBuffer(Pointer(stHTML)^, Length(stHTML));
    Strm.Position := 0;
    Clipboard.Clear;
    Fid := Clipboard.FindFormatID('text/html');
    if Fid = 0 then
      Fid := RegisterClipboardFormat('text/html');
    Clipboard.AddFormat(Fid, Strm);
  finally
    Strm.Free;
  end;
end;

procedure TfmMain.miItemsCreateKeywordListClick(Sender: TObject);
begin
  // Create keyword list
  SaveItems;
  CreateKeywordList;
end;

procedure TfmMain.miItemsRenameKeywordClick(Sender: TObject);
begin
  // Rename a keyword
  SaveItems;
  RenameKeyword;
end;

procedure TfmMain.miItemsStoreKeywordClick(Sender: TObject);
begin
  // Store keyword in the buffer
  stBuffKeyword := InputBox('Keyword buffer',
    'Type the keyword(s) to be saved in the buffer.', stBuffKeyword);
  stBuffKeyword := CleanKeywordField(stBuffKeyword);
end;

procedure TfmMain.miItemsApplyKeywordClick(Sender: TObject);
begin
  // Apply stored keyword to the current item
  if ((stBuffKeyword <> '') and (sqItems.RecordCount > 0)) then
  begin
    SaveItems;
    sqItems.Edit;
    if UTF8Pos(stBuffKeyword, sqItems.FieldByName('keywords').AsString) < 1 then
    begin
      if sqItems.FieldByName('keywords').AsString <> '' then
      begin
        sqItems.FieldByName('keywords').AsString :=
          sqItems.FieldByName('keywords').AsString + ', ' + stBuffKeyword;
      end
      else
      begin
        sqItems.FieldByName('keywords').AsString := stBuffKeyword;
      end;
    end
    else
    begin
      sqItems.FieldByName('keywords').AsString :=
        StringReplace(sqItems.FieldByName('keywords').AsString,
        stBuffKeyword + ', ', '', [rfIgnoreCase]);
      sqItems.FieldByName('keywords').AsString :=
        StringReplace(sqItems.FieldByName('keywords').AsString,
        stBuffKeyword, '', [rfIgnoreCase]);
    end;
    sqItems.FieldByName('keywords').AsString :=
      CleanKeywordField(sqItems.FieldByName('keywords').AsString);
    sqItems.Post;
    sqItems.ApplyUpdates;
    LoadItems;
  end;
end;

procedure TfmMain.miItemsFilterFromLatexClick(Sender: TObject);
begin
  // Filter from Latex file
  SaveItems;
  odOpenFile.Title := 'Open LaTex file to filter from';
  odOpenFile.Filter := 'Latex files|*.tex';
  odOpenFile.Options := odOpenFile.Options - [ofAllowMultiSelect];
  odOpenFile.DefaultExt := 'tex';
  odOpenFile.FileName := '';
  if odOpenFile.Execute = True then
  begin
    FilterFromLatex(odOpenFile.FileName);
  end;
end;

procedure TfmMain.miItemsFilterClick(Sender: TObject);
begin
  // Standard filter
  SaveItems;
  fmFilters.ShowModal;
end;

procedure TfmMain.miItemsRemoveFilterClick(Sender: TObject);
var
  iIDItem: integer;
begin
  // Remove filter
  SaveItems;
  try
    Screen.Cursor := crHourGlass;
    Application.ProcessMessages;
    iIDItem := -1;
    if sqItems.RecordCount > 0 then
    begin
      iIDItem := sqItems.FieldByName('IDItems').AsInteger;
    end;
    flFilterActive := 0;
    sqItems.Close;
    sqItems.SQL := 'Select * from Items ' + CreateOrderBy;
    sqItems.Open;
    if iIDItem > -1 then
    begin
      sqItems.Locate('IDItems', iIDItem, []);
    end;
  finally
    Screen.Cursor := crDefault;
  end;
end;

procedure TfmMain.miItemsCopyCrossrefClick(Sender: TObject);
var
  sqMasterItem: TSqlite3Dataset;
  stMasterType, stCurrType: string;
begin
  // Copy crossref
  if sqItems.FieldByName('crossref').AsString = '' then
  begin
    MessageDlg('The current item has no crossref content.', mtWarning, [mbOK], 0);
    Exit;
  end;
  if MessageDlg('Copy the proper fields from the cross referenced item?',
    mtConfirmation, [mbOK, mbCancel], 0) = mrCancel then
  begin
    Exit;
  end;
  SaveItems;
  sqMasterItem := TSqlite3Dataset.Create(Self);
  with sqMasterItem do
    try
      FileName := sqItems.FileName;
      TableName := 'Items';
      SQL := 'Select * from Items where Lower(bibtexkey) = "' +
        UTF8LowerCase(sqItems.FieldByName('crossref').AsString) + '"';
      Open;
      if RecordCount = 0 then
      begin
        MessageDlg('No item has the bibtex key ' +
          'specified in the crossref field.', mtWarning, [mbOK], 0);
        Exit;
      end
      else
      begin
        stMasterType := LowerCase(sqMasterItem.FieldByName(
          'entrytype').AsString);
        stCurrType := LowerCase(sqItems.FieldByName('entrytype').AsString);
        sqItems.Edit;
        // Cf. The Biblatex Guide Version 2.9a, p. 252-253.
        // Source: mvbook, book. Target: inbook, bookinbook, suppbook
        if (((stMasterType = 'mvbook') or (stMasterType = 'book')) and
          ((stCurrType = 'inbook') or (stCurrType = 'bookinbook') or
          (stCurrType = 'suppbook'))) then
        begin
          sqItems.FieldByName('bookauthor').AsString :=
            sqMasterItem.FieldByName('author').AsString;
          // Added by Max
          sqItems.FieldByName('editor').AsString :=
            sqMasterItem.FieldByName('editor').AsString;
          sqItems.FieldByName('editora').AsString :=
            sqMasterItem.FieldByName('editora').AsString;
          sqItems.FieldByName('editorb').AsString :=
            sqMasterItem.FieldByName('editorb').AsString;
          sqItems.FieldByName('editorc').AsString :=
            sqMasterItem.FieldByName('editorc').AsString;
        end;
        // Source: mvbook. Target: book, inbook, bookinbook, suppbook
        if ((stMasterType = 'mvbook') and ((stCurrType = 'book') or
          (stCurrType = 'inbook') or (stCurrType = 'bookinbook') or
          (stCurrType = 'suppbook'))) then
        begin
          sqItems.FieldByName('maintitle').AsString :=
            sqMasterItem.FieldByName('title').AsString;
          sqItems.FieldByName('mainsubtitle').AsString :=
            sqMasterItem.FieldByName('subtitle').AsString;
          sqItems.FieldByName('maintitleaddon').AsString :=
            sqMasterItem.FieldByName('titleaddon').AsString;
        end;
        // Source: mvcollection, mvreference. Target: collection, reference,
        // incollection, inreference, suppcollection
        if (((stMasterType = 'mvcollection') or
          (stMasterType = 'mvreference')) and
          ((stCurrType = 'collection') or (stCurrType = 'reference') or
          (stCurrType = 'incollection') or (stCurrType = 'inreference') or
          (stCurrType = 'suppcollection'))) then
        begin
          sqItems.FieldByName('maintitle').AsString :=
            sqMasterItem.FieldByName('title').AsString;
          sqItems.FieldByName('mainsubtitle').AsString :=
            sqMasterItem.FieldByName('subtitle').AsString;
          sqItems.FieldByName('maintitleaddon').AsString :=
            sqMasterItem.FieldByName('titleaddon').AsString;
        end;
        // Source: mvproceedings. Target: proceedings, inproceedings.
        if ((stMasterType = 'mvproceedings') and
          ((stCurrType = 'proceedings') or (stCurrType = 'inproceedings'))) then
        begin
          sqItems.FieldByName('maintitle').AsString :=
            sqMasterItem.FieldByName('title').AsString;
          sqItems.FieldByName('mainsubtitle').AsString :=
            sqMasterItem.FieldByName('subtitle').AsString;
          sqItems.FieldByName('maintitleaddon').AsString :=
            sqMasterItem.FieldByName('titleaddon').AsString;
        end;
        // Source: book. Target: inbook, bookinbook, suppbook
        if ((stMasterType = 'book') and ((stCurrType = 'inbook') or
          (stCurrType = 'bookinbook') or (stCurrType = 'suppbook'))) then
        begin
          sqItems.FieldByName('booktitle').AsString :=
            sqMasterItem.FieldByName('title').AsString;
          sqItems.FieldByName('booksubtitle').AsString :=
            sqMasterItem.FieldByName('subtitle').AsString;
          sqItems.FieldByName('booktitleaddon').AsString :=
            sqMasterItem.FieldByName('titleaddon').AsString;
          // Added by Max
          sqItems.FieldByName('series').AsString :=
            sqMasterItem.FieldByName('series').AsString;
          sqItems.FieldByName('number').AsString :=
            sqMasterItem.FieldByName('number').AsString;
          sqItems.FieldByName('publisher').AsString :=
            sqMasterItem.FieldByName('publisher').AsString;
          sqItems.FieldByName('location').AsString :=
            sqMasterItem.FieldByName('location').AsString;
          sqItems.FieldByName('year').AsString :=
            sqMasterItem.FieldByName('year').AsString;
        end;
        // Source: collection, reference. Target: incollection,
        // inreference, suppcollection
        if (((stMasterType = 'collection') or (stMasterType = 'reference')) and
          ((stCurrType = 'incollection') or (stCurrType = 'inreference') or
          (stCurrType = 'suppcollection'))) then
        begin
          sqItems.FieldByName('booktitle').AsString :=
            sqMasterItem.FieldByName('title').AsString;
          sqItems.FieldByName('booksubtitle').AsString :=
            sqMasterItem.FieldByName('subtitle').AsString;
          sqItems.FieldByName('booktitleaddon').AsString :=
            sqMasterItem.FieldByName('titleaddon').AsString;
        end;
        // Source: proceedings. Target: inproceedings
        if ((stMasterType = 'proceedings') and
          (stCurrType = 'inproceedings')) then

        begin
          sqItems.FieldByName('booktitle').AsString :=
            sqMasterItem.FieldByName('title').AsString;
          sqItems.FieldByName('booksubtitle').AsString :=
            sqMasterItem.FieldByName('subtitle').AsString;
          sqItems.FieldByName('booktitleaddon').AsString :=
            sqMasterItem.FieldByName('titleaddon').AsString;
        end;
        // Source: periodical. Target: article, suppperiodical
        if ((stMasterType = 'periodical') and
          ((stCurrType = 'article') or (stCurrType = 'suppperiodical'))) then
        begin
          sqItems.FieldByName('journaltitle').AsString :=
            sqMasterItem.FieldByName('title').AsString;
          sqItems.FieldByName('journalsubtitle').AsString :=
            sqMasterItem.FieldByName('subtitle').AsString;
        end;
      end;
      sqItems.Post;
      sqItems.ApplyUpdates;
      LoadItems;
    finally
      sqMasterItem.Free;
    end;
end;

procedure TfmMain.miItemsOrderByIDClick(Sender: TObject);
var
  IDItem: integer;
begin
  // Set order by in the sqItems SQL (valid for all the Order by items group)
  SaveItems;
  IDItem := sqItems.FieldByName('IDItems').AsInteger;
  sqItems.Close;
  sqItems.SQL := Copy(sqItems.SQL, 1, Pos(' order by ', sqItems.SQL) - 1) +
    CreateOrderBy;
  sqItems.Open;
  sqItems.Locate('IDItems', IDItem, []);
  grItems.Options := grItems.Options - [dgAutoSizeColumns];
  grItems.ResetColWidths;
  grItems.Options := grItems.Options + [dgAutoSizeColumns];
end;

procedure TfmMain.miToolsOptionsClick(Sender: TObject);
begin
  // Options
  SaveItems;
  // Do not show the form in modal mode because on OSX it's not possible
  // to open an OpenFile dialog over it
  fmOptions.Show;
end;

procedure TfmMain.miToolsStopListClick(Sender: TObject);
begin
  // Stop list
  SaveItems;
  fmStopList.ShowModal;
end;

procedure TfmMain.miToolCharClick(Sender: TObject);
var
  Editor: TStringCellEditor;
  iStart: integer = -1;
  iStartBox: smallint = -1;
begin
  // Show char maps
  if ((pcPageControl.PageIndex = 0) and (sgRequiredFields.EditorMode = True)) then
  begin
    Editor := TStringCellEditor(sgRequiredFields.Editor);
    iStart := Editor.SelStart;
  end
  else if ((pcPageControl.PageIndex = 1) and
    (sgOptionalFields1.EditorMode = True)) then
  begin
    Editor := TStringCellEditor(sgOptionalFields1.Editor);
    iStart := Editor.SelStart;
  end
  else if ((pcPageControl.PageIndex = 2) and
    (sgOptionalFields2.EditorMode = True)) then
  begin
    Editor := TStringCellEditor(sgOptionalFields2.Editor);
    iStart := Editor.SelStart;
  end
  else if ((pcPageControl.PageIndex = 3) and
    (sgOptionalFields3.EditorMode = True)) then
  begin
    Editor := TStringCellEditor(sgOptionalFields3.Editor);
    iStart := Editor.SelStart;
  end
  else if ((pcPageControl.PageIndex = 4) and (meAbstract.Focused = True)) then
  begin
    iStart := meAbstract.SelStart;
  end
  else if ((pcPageControl.PageIndex = 5) and (meReview.Focused = True)) then
  begin
    iStart := meReview.SelStart;
  end
  else if edSmartFilter.Focused = True then
  begin
    iStartBox := 0;
    iStart := edSmartFilter.SelStart;
  end
  else if cbFilterKey.Focused = True then
  begin
    iStartBox := 1;
    iStart := cbFilterKey.SelStart;
  end;
  if iStart = -1 then
  begin
    MessageDlg('No field is in editing to insert a special character.',
      mtWarning, [mbOK], 0);
    Exit;
  end;
  if fmChar.ShowModal = mrOk then
  begin
    if iStartBox = 0 then
    begin
      edSmartFilter.Text := UTF8Copy(edSmartFilter.Text, 1, iStart) +
        fmChar.sgChar.Cells[fmChar.sgChar.Col, fmChar.sgChar.Row] +
        UTF8Copy(edSmartFilter.Text, iStart + 1, UTF8Length(edSmartFilter.Text));
      edSmartFilter.SelStart := iStart + 1;
    end
    else if iStartBox = 1 then
    begin
      cbFilterKey.Text := UTF8Copy(cbFilterKey.Text, 1, iStart) +
        fmChar.sgChar.Cells[fmChar.sgChar.Col, fmChar.sgChar.Row] +
        UTF8Copy(cbFilterKey.Text, iStart + 1, UTF8Length(cbFilterKey.Text));
      cbFilterKey.SelStart := iStart + 1;
    end
    else if pcPageControl.PageIndex < 4 then
    begin
      Editor.SelStart := iStart;
      Editor.EditText := UTF8Copy(Editor.EditText, 1, iStart) +
        fmChar.sgChar.Cells[fmChar.sgChar.Col, fmChar.sgChar.Row] +
        UTF8Copy(Editor.EditText, iStart + 1, UTF8Length(Editor.EditText));
      Editor.SelStart := iStart + 1;
    end
    else if pcPageControl.PageIndex = 4 then
    begin
      meAbstract.Text := UTF8Copy(meAbstract.Text, 1, iStart) +
        fmChar.sgChar.Cells[fmChar.sgChar.Col, fmChar.sgChar.Row] +
        UTF8Copy(meAbstract.Text, iStart + 1, UTF8Length(meAbstract.Text));
      meAbstract.SelStart := iStart + 1;
    end
    else if pcPageControl.PageIndex = 5 then
    begin
      meReview.Text := UTF8Copy(meReview.Text, 1, iStart) +
        fmChar.sgChar.Cells[fmChar.sgChar.Col, fmChar.sgChar.Row] +
        UTF8Copy(meReview.Text, iStart + 1, UTF8Length(meReview.Text));
      meReview.SelStart := iStart + 1;
    end;
  end;
end;

procedure TfmMain.miToolKeyListClick(Sender: TObject);
var
  Editor: TStringCellEditor;
begin
  // Keywords list
  if ((pcPageControl.PageIndex > 0) or
    (LowerCase(sgRequiredFields.Cells[0, sgRequiredFields.Row]) <> 'keywords') or
    (sgRequiredFields.EditorMode = False)) then
  begin
    MessageDlg('The keywords field is not in editing.',
      mtWarning, [mbOK], 0);
    Exit;
  end;
  if fmKeywords.ShowModal = mrOk then
  begin
    Editor := TStringCellEditor(sgRequiredFields.Editor);
    Editor.SelStart := UTF8Length(Editor.EditText);
    Editor.SelLength := 0;
  end;
end;

procedure TfmMain.miToolCheckDoublesClick(Sender: TObject);
begin
  // Check for double items
  if miItemsRemoveFilter.Enabled = True then
  begin
    miItemsRemoveFilterClick(nil);
  end;
  fmCheckDoubles.Show;
end;

procedure TfmMain.miToolsReplaceKeysClick(Sender: TObject);
begin
  // Replace all BibTex keys
  SaveItems;
  if flFilterActive > 0 then
  begin
    if MessageDlg('Replace the BibTex keys of the filtered items ' +
      'with new ones formatted according to the defined pattern?',
      mtConfirmation, [mbOK, mbCancel], 0) = mrCancel then
      Exit;
  end
  else
  begin
    if MessageDlg('Replace the BibTex keys of all the items ' +
      'with new ones formatted according to the defined pattern?',
      mtConfirmation, [mbOK, mbCancel], 0) = mrCancel then
      Exit;
  end;
  if MessageDlg('Are you really sure you want to replace ' +
    'the existing BibTex keys with new ones?', mtWarning, [mbOK, mbCancel], 0) =
    mrCancel then
    Exit;
  Screen.Cursor := crHourGlass;
  try
    sqToolsTables.Close;
    sqToolsTables.FileName := sqItems.FileName;
    sqToolsTables.SQL := sqItems.SQL;
    sqToolsTables.TableName := 'Items';
    sqToolsTables.PrimaryKey := 'IDItems';
    sqToolsTables.Open;
    while not sqToolsTables.EOF do
    begin
      sqToolsTables.Edit;
      sqToolsTables.FieldByName('bibtexkey').AsString :=
        CreateBibTexKey(False, False);
      sqToolsTables.Post;
      // Apply is here to store the BibTex value before creating new keys:
      // otherwise the sqKey in the check unique procedure could not read it
      sqToolsTables.ApplyUpdates;
      sbStatusBar.SimpleText :=
        'Replaced key of IDItem ' + sqToolsTables.FieldByName('IDItems').AsString + '.';
      Application.ProcessMessages;
      sqToolsTables.Next;
    end;
    sqToolsTables.Close;
    sqItems.Close;
    sqItems.Open;
  finally
    Screen.Cursor := crDefault;
  end;
end;

procedure TfmMain.miToolsReplaceCitationsClick(Sender: TObject);
var
  Proc: TProcess;
  stOutputFile, stInputFile, stOutputFileNoSpace, stInputFileNoSpace,
  stUserCmd, stExt, stDelimiter: string;
begin
  // Replace /cita command with citations in a LaTex file
  SaveItems;
  odOpenFile.Title := 'Open file to replace citations';
  odOpenFile.Filter := 'Latex files|*.tex';
  odOpenFile.Options := odOpenFile.Options - [ofAllowMultiSelect];
  odOpenFile.DefaultExt := 'tex';
  odOpenFile.FileName := '';
  if odOpenFile.Execute = True then
  begin
    if UTF8Pos(#39, odOpenFile.FileName) > 0 then
    begin
      MessageDlg('The file name cannot contain an apostrophe.',
        mtError, [mbOK], 0);
    end
    else if ReplaceCiteInLatex(odOpenFile.FileName) = False then
    begin
      MessageDlg('Error on replacing \cite and similar ' +
        'commands in the selected file.',
        mtError, [mbOK], 0);
    end
    else if Sender = miToolsConvert then
      try
        try
          Screen.Cursor := crHourGlass;
          Application.ProcessMessages;
          Proc := TProcess.Create(nil);
          Proc.Options := Proc.Options + [poWaitOnExit, poUsePipes];
          stInputFile := ExtractFilePath(odOpenFile.FileName) +
            ExtractFileNameOnly(odOpenFile.FileName) + '-fullcitations.tex';
          {$ifdef Linux}
          stInputFileNoSpace :=
            StringReplace(stInputFile, ' ', '\ ', [rfReplaceAll]);
          {$endif}
          {$ifdef Win32}
          stInputFileNoSpace := '"' + stInputFile + '"';
          {$endif}
          {$ifdef Darwin}
          stInputFileNoSpace :=
            StringReplace(stInputFile, ' ', '\ ', [rfReplaceAll]);
          {$endif}
          stOutputFile := ExtractFilePath(odOpenFile.FileName) +
            ExtractFileNameOnly(odOpenFile.FileName);
          case fmOptions.rgConvWpFormat.ItemIndex of
            0: stOutputFile := stOutputFile + '.odt';
            1: stOutputFile := stOutputFile + '.docx';
            2: stOutputFile := stOutputFile + '.doc';
            3: stOutputFile := stOutputFile + '.htm';
          end;
          if fmOptions.rgConvWpFormat.ItemIndex =
            fmOptions.rgConvWpFormat.Items.Count - 1 then
          begin
            stExt := UTF8Copy(fmOptions.edUserDefCom.Text,
              UTF8Pos('%o', fmOptions.edUserDefCom.Text) + 2,
              UTF8Length(fmOptions.edUserDefCom.Text));
            if UTF8Pos(' ', stExt) > 0 then
              stExt := UTF8Copy(stExt, 1, UTF8Pos(' ', stExt));
            stOutputFile := stOutputFile + stExt;
          end;
          {$ifdef Linux}
          stOutputFileNoSpace :=
            StringReplace(stOutputFile, ' ', '\ ', [rfReplaceAll]);
          {$endif}
          {$ifdef Win32}
          stOutputFileNoSpace := '"' + stOutputFile + '"';
          {$endif}
          {$ifdef Darwin}
          stOutputFileNoSpace :=
            StringReplace(stOutputFile, ' ', '\ ', [rfReplaceAll]);
          {$endif}
          if fmOptions.rgConvWpFormat.ItemIndex =
            fmOptions.rgConvWpFormat.Items.Count - 1 then
          begin
            stUserCmd := fmOptions.edUserDefCom.Text;
            stUserCmd := StringReplace(stUserCmd, '%i', stInputFileNoSpace, []);
            stUserCmd := StringReplace(stUserCmd, '%o' + stExt,
              stOutputFileNoSpace, []);
            {$ifdef Linux}
            Proc.CommandLine := 'sh -c "' + stUserCmd + '"';
            {$endif}
            {$ifdef Win32}
            Proc.CommandLine := stUserCmd;
            {$endif}
            {$ifdef Darwin}
            Proc.CommandLine := 'sh -c "' + stUserCmd + '"';
            {$endif}
          end
          else
          begin
            {$ifdef Linux}
            Proc.CommandLine :=
              'sh -c "pandoc ' + stInputFileNoSpace + ' -o ' + stOutputFileNoSpace + '"';
            {$endif}
            {$ifdef Win32}
            Proc.CommandLine :=
              'pandoc ' + stInputFileNoSpace + ' -o ' + stOutputFileNoSpace;
            {$endif}
            {$ifdef Darwin}
            Proc.CommandLine :=
              'sh -c "/usr/local/bin/pandoc ' + stInputFileNoSpace +
              ' -o ' + stOutputFileNoSpace + '"';
            {$endif}
          end;
          Proc.Execute;
          if fmOptions.rgConvWpFormat.ItemIndex <
            fmOptions.rgConvWpFormat.Items.Count - 1 then
          begin
            if FileExistsUTF8(stOutputFile) then
            begin
              if OpenDocument(stOutputFile) = False then
              begin
                MessageDlg('It''s not possible to open the converted file.',
                  mtWarning, [mbOK], 0);
              end;
            end
            else
            begin
              MessageDlg('Error on converting the selected .tex file ' +
                'to another format: check the file ' +
                'and Pandoc installation.',
                mtError, [mbOK], 0);
            end;
          end
          else
          begin
            MessageDlg('The conversion is finished; ' +
              'you may open the converted file.',
              mtInformation, [mbOK], 0);
          end;
        except
          MessageDlg('Error on converting the selected .tex file ' +
            'to another format: check the file and Pandoc installation.',
            mtError, [mbOK], 0);
        end;
      finally
        Proc.Free;
        Screen.Cursor := crDefault;
      end;
  end;
end;

procedure TfmMain.miToolsCompactClick(Sender: TObject);
var
  IDItem: integer;
begin
  // Compact the database
  odOpenFile.Title := 'Open Bibfilex file to compact';
  odOpenFile.Filter := 'Bibfilex files|*.bfx';
  odOpenFile.Options := odOpenFile.Options - [ofAllowMultiSelect];
  odOpenFile.DefaultExt := '.bfx';
  odOpenFile.FileName := '';
  if odOpenFile.Execute = True then
    try
      if odOpenFile.FileName = sqItems.FileName then
      begin
        if sqItems.Active = True then
        begin
          MessageDlg('It is not possible to compact the file in use.',
            mtError, [mbOK], 0);
          Abort;
        end;
      end;
      Screen.Cursor := crHourGlass;
      Application.ProcessMessages;
      try
        if FileExistsUTF8(ExtractFileNameWithoutExt(odOpenFile.FileName) + '.bak') then
        begin
          DeleteFileUTF8(ExtractFileNameWithoutExt(odOpenFile.FileName) + '.bak');
        end;
        CopyFile(odOpenFile.FileName,
          ExtractFileNameWithoutExt(odOpenFile.FileName) + '.bak');
        sqToolsTables.Close;
        sqToolsTables.FileName := odOpenFile.FileName;
        sqToolsTables.ExecSQL('Vacuum');
        Screen.Cursor := crDefault;
        MessageDlg('Compact was successful; ' +
          'anyway, the original file is available as ' +
          ExtractFileNameWithoutExt(odOpenFile.FileName) + '.bak.',
          mtInformation, [mbOK], 0);
      except
        MessageDlg('Compact was NOT successful; ' +
          'the original file is available as ' +
          ExtractFileNameWithoutExt(odOpenFile.FileName) + '.bak.',
          mtError, [mbOK], 0);
      end;
    finally
      Screen.Cursor := crDefault;
    end;
end;

procedure TfmMain.miToolsUpgradeClick(Sender: TObject);
begin
  // Upgrade data
  odOpenFile.Title := 'Open Bibfilex file to be upgraded';
  odOpenFile.Filter := 'Bibfilex files|*.bfx|All files|*';
  odOpenFile.Options := odOpenFile.Options - [ofAllowMultiSelect];
  odOpenFile.DefaultExt := '.bfx';
  odOpenFile.FileName := '';
  if odOpenFile.Execute = True then
    try
      if ((odOpenFile.FileName = sqItems.FileName) and (sqItems.Active = True)) then
      begin
        MessageDlg('It is not possible to upgrade the file in use.',
          mtWarning, [mbOK], 0);
      end
      else if MakeTableUpgrade(odOpenFile.FileName) = False then
      begin
        MessageDlg('The structure of the selected file ' +
          'is already at the last version.', mtInformation, [mbOK], 0);
      end
      else
      begin
        MessageDlg('The structure of the selected file ' +
          'has been upgraded to the last version. ' +
          'The original file has been renamed as ' +
          ExtractFileNameWithoutExt(odOpenFile.FileName) + '.bak',
          mtInformation, [mbOK], 0);
      end
    except
      Screen.Cursor := crDefault;
      MessageDlg('It was not possible to upgrade data file.',
        mtError, [mbOK], 0);
      Abort;
    end;
end;

procedure TfmMain.miToolsDecryptClick(Sender: TObject);
var
  Proc: TProcess;
  stFileNameNoSpace: string;
begin
  // Decrypt with GPG
  odOpenFile.Title := 'Open encrypted file';
  odOpenFile.Filter := 'Encrypted file|*.pgp';
  odOpenFile.Options := odOpenFile.Options - [ofAllowMultiSelect];
  odOpenFile.DefaultExt := '.pgp';
  odOpenFile.FileName := '';
  if odOpenFile.Execute = True then
  begin
    if FileExistsUTF8(ExtractFileNameWithoutExt(odOpenFile.FileName)) = True then
    begin
      MessageDlg('Existing decrypted file cannot be overwritten.', mtWarning, [mbOK], 0);
      Abort;
    end;
    stFileNameNoSpace :=
      StringReplace(odOpenFile.FileName, ' ', '\ ', [rfReplaceAll]);
    try
      Proc := TProcess.Create(nil);
      Proc.Options := Proc.Options + [poWaitOnExit, poUsePipes];
      Proc.CommandLine := 'sh -c "gpg2 --batch ' + ' --armor --output ' +
        ExtractFileNameWithoutExt(stFileNameNoSpace) + ' --decrypt ' +
        stFileNameNoSpace + '"';
      Proc.Execute;
      Proc.Free;
      Screen.Cursor := crDefault;
      if FileExistsUTF8(ExtractFileNameWithoutExt(odOpenFile.FileName)) =
        False then
        MessageDlg('It was not possible to decrypt the file.', mtWarning, [mbOK], 0);
    except
      Screen.Cursor := crDefault;
      MessageDlg('It was not possible to decrypt the file.', mtWarning, [mbOK], 0);
    end;
  end;
end;

procedure TfmMain.miToolsEncryptClick(Sender: TObject);
var
  Proc: TProcess;
  stFileNameNoSpace: string;
begin
  // Encrypt with GPG
  odOpenFile.Title := 'Open a file';
  odOpenFile.Filter := 'All files|*.*';
  odOpenFile.Options := odOpenFile.Options - [ofAllowMultiSelect];
  odOpenFile.DefaultExt := '*';
  odOpenFile.FileName := '';
  if odOpenFile.Execute = True then
  begin
    if FileExistsUTF8(odOpenFile.FileName + '.pgp') = True then
    begin
      MessageDlg('Existing encrypted file cannot be overwritten.', mtWarning, [mbOK], 0);
      Abort;
    end;
    InputQuery('Recipient', 'Type the recipient email or ID.', False, stRecipient);
    if stRecipient = '' then
      Abort;
    stFileNameNoSpace :=
      StringReplace(odOpenFile.FileName, ' ', '\ ', [rfReplaceAll]);
    try
      Proc := TProcess.Create(nil);
      Proc.Options := Proc.Options + [poWaitOnExit, poUsePipes];
      Proc.CommandLine := 'sh -c "gpg2 --batch ' + ' --armor --output ' +
        stFileNameNoSpace + '.pgp' + ' --recipient ' + stRecipient +
        ' --encrypt --sign ' + stFileNameNoSpace + '"';
      Proc.Execute;
      Proc.Free;
      Screen.Cursor := crDefault;
      if FileExistsUTF8(odOpenFile.FileName + '.pgp') = False then
        MessageDlg('It was not possible to encrypt the file.', mtWarning, [mbOK], 0);
    except
      Screen.Cursor := crDefault;
      MessageDlg('It was not possible to encrypt the file.', mtWarning, [mbOK], 0);
    end;
  end;
end;

procedure TfmMain.miShowManualClick(Sender: TObject);
begin
  // Show manual
  if OpenDocument(ExtractFileDir(Application.ExeName) + DirectorySeparator +
    'bibfilex-manual.pdf') = False then
  begin
    MessageDlg('The user manual is not installed; ' +
      'download it from Bibfilex website, ' +
      'which should be opened now in the default browser.',
      mtInformation, [mbOK], 0);
    OpenURL('http://sites.google.com/site/bibfilex/');
  end;
end;

procedure TfmMain.miCopyrightShowClick(Sender: TObject);
begin
  // Show copyright
  fmCopyright.ShowModal;
end;

procedure TfmMain.pmMainPopup(Sender: TObject);
begin
  // Activate or deactivate open item
  pmItemsOpenLink.Enabled := False;
  if pcPageControl.ActivePageIndex = 0 then
  begin
    if UTF8LowerCase(UTF8Copy(sgRequiredFields.Cells[sgRequiredFields.Col,
      sgRequiredFields.Row], 1, 4)) = 'http' then
      pmItemsOpenLink.Enabled := True;
  end
  else if pcPageControl.ActivePageIndex = 1 then
  begin
    if UTF8LowerCase(UTF8Copy(sgOptionalFields1.Cells[sgOptionalFields1.Col,
      sgOptionalFields1.Row], 1, 4)) = 'http' then
      pmItemsOpenLink.Enabled := True;
  end
  else if pcPageControl.ActivePageIndex = 2 then
  begin
    if UTF8LowerCase(UTF8Copy(sgOptionalFields2.Cells[sgOptionalFields2.Col,
      sgOptionalFields2.Row], 1, 4)) = 'http' then
      pmItemsOpenLink.Enabled := True;
  end
  else if pcPageControl.ActivePageIndex = 3 then
  begin
    if UTF8LowerCase(UTF8Copy(sgOptionalFields3.Cells[sgOptionalFields3.Col,
      sgOptionalFields3.Row], 1, 4)) = 'http' then
      pmItemsOpenLink.Enabled := True;
  end;
end;

procedure TfmMain.pmItemsOpenLinkClick(Sender: TObject);
begin
  // Open link
  if pcPageControl.ActivePageIndex = 0 then
    OpenURL(sgRequiredFields.Cells[sgRequiredFields.Col, sgRequiredFields.Row])
  else if pcPageControl.ActivePageIndex = 1 then
    OpenURL(sgOptionalFields1.Cells[sgOptionalFields1.Col, sgOptionalFields1.Row])
  else if pcPageControl.ActivePageIndex = 2 then
    OpenURL(sgOptionalFields2.Cells[sgOptionalFields2.Col, sgOptionalFields2.Row])
  else if pcPageControl.ActivePageIndex = 3 then
    OpenURL(sgOptionalFields3.Cells[sgOptionalFields3.Col, sgOptionalFields3.Row]);
end;

//***************************************************************
//********************** DATASET PROCEDURES *********************
//***************************************************************

procedure TfmMain.sqItemsAfterPost(DataSet: TDataSet);
begin
  // Last change date
  if sqItems.FieldByName('timestamp').AsDateTime > 0 then
    lbLastEdited.Caption := 'Last modified on ' +
      FormatDateTime('ddddd', sqItems.FieldByName('timestamp').AsDateTime) +
      ' at ' + FormatDateTime('t', sqItems.FieldByName('timestamp').AsDateTime) + '.'
  else
    lbLastEdited.Caption := '';
  CompileSummary;
  SetMarkAbstractReviewTab;
end;

procedure TfmMain.sqItemsAfterScroll(DataSet: TDataSet);
begin
  // Load item
  LoadItems;
end;

procedure TfmMain.sqItemsBeforeDelete(DataSet: TDataSet);
var
  AttDir: string;
  myNamesList: TStringList;
  i: integer;
begin
  // Delete possible attachments
  if sqItems.FieldByName('AttName').AsString <> '' then
    try
      Screen.Cursor := crHourGlass;
      AttDir := ExtractFileNameWithoutExt(sqItems.FileName);
      myNamesList := TStringList.Create;
      if DirectoryExistsUTF8(AttDir) = False then
      begin
        MessageDlg('The attachment directory is not available.',
          mtError, [mbOK], 0);
        Abort;
      end;
      myNamesList.Text := sqItems.FieldByName('AttName').AsString;
      for i := 0 to myNamesList.Count - 1 do
        try
          DeleteFileUTF8(AttDir + DirectorySeparator +
            sqItems.FieldByName('IDItems').AsString + '-' +
            ExtractFileName(myNamesList[i]) + '.zip');
        except
          MessageDlg('It is not possible to delete the attachment.',
            mtError, [mbOK], 0);
        end;
      if IsDirectoryEmpty(AttDir) = True then
        try
          DeleteDirectory(AttDir, False);
        except
          MessageDlg('It is not possible to delete the attachments directory.',
            mtError, [mbOK], 0);
        end;
    finally
      myNamesList.Free;
      Screen.Cursor := crDefault;
    end;
end;

procedure TfmMain.sqItemsBeforeScroll(DataSet: TDataSet);
begin
  // Save item
  SaveItems;
end;

procedure TfmMain.dsItemsDataChange(Sender: TObject; Field: TField);
begin
  // Activate and deactivate grids, memos and menu items
  if sqItems.RecordCount = 0 then
  begin
    cbItemsKind.ItemIndex := -1;
    cbItemsKind.Enabled := False;
    dbOwned.Enabled := False;
    // RowCount = 0 raises an error just on opening a file
    sgRequiredFields.RowCount := 1;
    // Clear remove also rows
    sgRequiredFields.Cells[0, 0] := '';
    sgRequiredFields.Cells[1, 0] := '';
    sgRequiredFields.Options := sgRequiredFields.Options - [goEditing];
    sgOptionalFields1.RowCount := 1;
    sgOptionalFields1.Cells[0, 0] := '';
    sgOptionalFields1.Cells[1, 0] := '';
    sgOptionalFields1.Options := sgOptionalFields1.Options - [goEditing];
    sgOptionalFields2.RowCount := 1;
    sgOptionalFields2.Cells[0, 0] := '';
    sgOptionalFields2.Cells[1, 0] := '';
    sgOptionalFields2.Options := sgOptionalFields2.Options - [goEditing];
    sgOptionalFields3.RowCount := 1;
    sgOptionalFields3.Cells[0, 0] := '';
    sgOptionalFields3.Cells[1, 0] := '';
    sgOptionalFields3.Options := sgOptionalFields3.Options - [goEditing];
    meAbstract.Clear;
    meAbstract.Enabled := False;
    meReview.Clear;
    meReview.Enabled := False;
    miItemsDelete.Enabled := False;
    pmItemsDelete.Enabled := False;
    miItemsAttachments.Enabled := False;
    miAttachNew.Enabled := False;
    miAttachOpen.Enabled := False;
    miAttachSaveAs.Enabled := False;
    miAttachDelete.Enabled := False;
    pmAttNew.Enabled := False;
    pmAttOpen.Enabled := False;
    pmAttSaveAs.Enabled := False;
    pmAttDelete.Enabled := False;
    miItemsCreateKey.Enabled := False;
    miItemsModifyKey.Enabled := False;
    miItemsCopyKey.Enabled := False;
    miItemsCopyCitationTxt.Enabled := False;
    miItemsCopyCitationHtml.Enabled := False;
    miItemsSave.Enabled := False;
    pmItemsSave.Enabled := False;
    tbItemsSave.Enabled := False;
    miItemsUndo.Enabled := False;
    pmItemsUndo.Enabled := False;
    miItemsCopyCrossref.Enabled := False;
    lbLastEdited.Caption := '';
    lbAttNames.Items.Text := '';
    lbAuthorTitle.Caption := '';
  end
  else
  begin
    cbItemsKind.Enabled := True;
    dbOwned.Enabled := True;
    sgRequiredFields.Options := sgRequiredFields.Options + [goEditing];
    sgOptionalFields1.Options := sgOptionalFields1.Options + [goEditing];
    sgOptionalFields2.Options := sgOptionalFields2.Options + [goEditing];
    sgOptionalFields3.Options := sgOptionalFields3.Options + [goEditing];
    meAbstract.Enabled := True;
    meReview.Enabled := True;
    miItemsDelete.Enabled := True;
    pmItemsDelete.Enabled := True;
    miItemsAttachments.Enabled := True;
    miAttachNew.Enabled := True;
    miAttachOpen.Enabled := True;
    miAttachSaveAs.Enabled := True;
    miAttachDelete.Enabled := True;
    pmAttNew.Enabled := True;
    pmAttOpen.Enabled := True;
    pmAttSaveAs.Enabled := True;
    pmAttDelete.Enabled := True;
    miItemsCreateKey.Enabled := True;
    miItemsModifyKey.Enabled := True;
    miItemsCopyKey.Enabled := True;
    miItemsCopyCitationTxt.Enabled := True;
    miItemsCopyCitationHtml.Enabled := True;
    miItemsSave.Enabled := flDataEdit;
    pmItemsSave.Enabled := flDataEdit;
    tbItemsSave.Enabled := flDataEdit;
    miItemsUndo.Enabled := flDataEdit;
    pmItemsUndo.Enabled := flDataEdit;
    miItemsCopyCrossref.Enabled := True;
  end;
  SetStatusBar;
end;

procedure TfmMain.edSmartFilterKeyDown(Sender: TObject; var Key: word;
  Shift: TShiftState);
begin
  // Fast filter
  if key = 13 then
  begin
    SaveItems;
    SmartFilter;
    key := 0;
  end;
end;

procedure TfmMain.cbFilterKeyKeyDown(Sender: TObject; var Key: word;
  Shift: TShiftState);
var
  stKeyword: string;
  iLenText: integer;
begin
  // Filter on keyword
  if key = 13 then
  begin
    SaveItems;
    cbFilterKey.Text := CleanKeywordField(cbFilterKey.Text);
    FilterOnKeyword;
    key := 0;
  end
  // Autocompletion
  else if ((key = 32) and (ssCtrl in Shift)) then
  begin
    iLenText := UTF8Length(cbFilterKey.Text);
    if UTF8Pos(',', cbFilterKey.Text) = 0 then
      cbFilterKey.Text :=
        SuggestValue(cbFilterKey.Text, 'keywords')
    // More keywords
    else
    begin
      stKeyword := cbFilterKey.Text;
      while UTF8Pos(',', stKeyword) > 0 do
        stKeyword := UTF8Copy(stKeyword, UTF8Pos(',', stKeyword) +
          1, UTF8Length(stKeyword));
      while UTF8Copy(stKeyword, 1, 1) = ' ' do
        stKeyword := UTF8Copy(stKeyword, 2, UTF8Length(stKeyword));
      cbFilterKey.Text :=
        UTF8Copy(cbFilterKey.Text, 1, UTF8Length(cbFilterKey.Text) -
        UTF8Length(stKeyword)) + SuggestValue(stKeyword, 'keywords');
    end;
    cbFilterKey.SelStart := iLenText;
    cbFilterKey.SelLength := UTF8Length(cbFilterKey.Text);
    key := 0;
  end;
end;

procedure TfmMain.cbFilterKeySelect(Sender: TObject);
begin
  // Filter on keyword by list item
  SaveItems;
  FilterOnKeyword;
end;

procedure TfmMain.bnAddKeyClick(Sender: TObject);
begin
  // Add keyword in the list
  if cbFilterKey.Text <> '' then
  begin
    if cbFilterKey.Items.IndexOf(cbFilterKey.Text) < 0 then
    begin
      cbFilterKey.Items.Add(cbFilterKey.Text);
    end;
  end;
end;

procedure TfmMain.bnRemoveKeyClick(Sender: TObject);
begin
  // Remove keyword in the list
  if cbFilterKey.Text <> '' then
  begin
    if cbFilterKey.Items.IndexOf(cbFilterKey.Text) > -1 then
    begin
      cbFilterKey.Items.Delete(cbFilterKey.Items.IndexOf(cbFilterKey.Text));
    end;
  end;
end;

procedure TfmMain.apAppPropException(Sender: TObject; E: Exception);
begin
  // Error handling
  if sqItems.Active = True then
  begin
    if dsItems.State in [dsEdit, dsBrowse] then
    begin
      if MessageDlg('An error has occured; ' +
        'undo changes to the current item?', mtError, [mbOK, mbCancel], 0) =
        mrOk then
      begin
        sqItems.Cancel;
        LoadItems;
      end;
    end;
  end
  else
  begin
    MessageDlg('An error has occured.', mtError, [mbOK], 0);
  end;
end;

//***************************************************************
//********************** COMMON PROCEDURES **********************
//***************************************************************

procedure TfmMain.CreateDataTables(DataFileName: string);
var
  mySQL: string;
begin
  // Create new archive
  // Overwriting not allowed: it could make troubles
  if FileExistsUTF8(DataFileName) = True then
  begin
    MessageDlg('A file with the same name is already existing ' +
      'and cannot be overwritten.', mtWarning, [mbOK], 0);
    Abort;
  end;
  // Create Data Table and indexes
  Screen.Cursor := crHourGlass;
  Application.ProcessMessages;
  try
    with sqToolsTables do
      try
        FileName := DataFileName;
        // Create Subjects Table and indexes
        mySQL := 'CREATE TABLE Items (';
        // System and main fields
        mySQL := MySQL + 'IDItems INTEGER PRIMARY KEY, ';
        mySQL := MySQL + 'bibtexkey VARCHAR(80) collate nocase, ';
        mySQL := MySQL + 'entrytype VARCHAR(30), ';
        mySQL := MySQL + 'author VARCHAR(200), ';
        mySQL := MySQL + 'title VARCHAR(200), ';
        mySQL := MySQL + 'journaltitle VARCHAR(200), ';
        mySQL := MySQL + 'bookauthor VARCHAR(200), ';
        mySQL := MySQL + 'booktitle VARCHAR(200), ';
        mySQL := MySQL + 'year VARCHAR(15), ';
        mySQL := MySQL + 'keywords VARCHAR(200), ';
        mySQL := MySQL + 'crossref VARCHAR(200), ';
        // All other fields sorted by name
        mySQL := MySQL + 'abstract TEXT, ';
        mySQL := MySQL + 'addendum VARCHAR(200), ';
        mySQL := MySQL + 'afterword VARCHAR(200), ';
        mySQL := MySQL + 'annotation VARCHAR(200), ';
        mySQL := MySQL + 'annotator VARCHAR(200), ';
        mySQL := MySQL + 'authortype VARCHAR(200), ';
        mySQL := MySQL + 'bookpagination VARCHAR(200), ';
        mySQL := MySQL + 'booksubtitle VARCHAR(200), ';
        mySQL := MySQL + 'booktitleaddon VARCHAR(200), ';
        mySQL := MySQL + 'chapter VARCHAR(200), ';
        mySQL := MySQL + 'commentator VARCHAR(200), ';
        mySQL := MySQL + 'date DATE, ';
        mySQL := MySQL + 'doi VARCHAR(100), ';
        mySQL := MySQL + 'edition VARCHAR(10), ';
        mySQL := MySQL + 'editor VARCHAR(200), ';
        mySQL := MySQL + 'editora VARCHAR(200), ';
        mySQL := MySQL + 'editorb VARCHAR(200), ';
        mySQL := MySQL + 'editorc VARCHAR(200), ';
        mySQL := MySQL + 'editortype VARCHAR(200), ';
        mySQL := MySQL + 'editoratype VARCHAR(200), ';
        mySQL := MySQL + 'editorbtype VARCHAR(200), ';
        mySQL := MySQL + 'editorctype VARCHAR(200), ';
        mySQL := MySQL + 'eid VARCHAR(100), ';
        mySQL := MySQL + 'eprint VARCHAR(200), ';
        mySQL := MySQL + 'eprintclass VARCHAR(200), ';
        mySQL := MySQL + 'eprinttype VARCHAR(200), ';
        mySQL := MySQL + 'eventdate DATE, ';
        mySQL := MySQL + 'eventtitle VARCHAR(200), ';
        mySQL := MySQL + 'eventtitleaddon VARCHAR(200), ';
        mySQL := MySQL + 'file TEXT, ';
        mySQL := MySQL + 'foreword VARCHAR(200), ';
        mySQL := MySQL + 'holder VARCHAR(200), ';
        mySQL := MySQL + 'howpublished VARCHAR(200), ';
        mySQL := MySQL + 'indextitle VARCHAR(200), ';
        mySQL := MySQL + 'institution VARCHAR(200), ';
        mySQL := MySQL + 'introduction VARCHAR(200), ';
        mySQL := MySQL + 'isan VARCHAR(100), ';
        mySQL := MySQL + 'isbn VARCHAR(100), ';
        mySQL := MySQL + 'ismn VARCHAR(100), ';
        mySQL := MySQL + 'isrn VARCHAR(100), ';
        mySQL := MySQL + 'issn VARCHAR(100), ';
        mySQL := MySQL + 'issue VARCHAR(200), ';
        mySQL := MySQL + 'issuesubtitle VARCHAR(200), ';
        mySQL := MySQL + 'issuetitle VARCHAR(200), ';
        mySQL := MySQL + 'iswc VARCHAR(100), ';
        mySQL := MySQL + 'journalsubtitle VARCHAR(200), ';
        mySQL := MySQL + 'label VARCHAR(200), ';
        mySQL := MySQL + 'language VARCHAR(50), ';
        mySQL := MySQL + 'library VARCHAR(200), ';
        mySQL := MySQL + 'location VARCHAR(200), ';
        mySQL := MySQL + 'mainsubtitle VARCHAR(200), ';
        mySQL := MySQL + 'maintitle VARCHAR(200), ';
        mySQL := MySQL + 'maintitleaddon VARCHAR(200), ';
        mySQL := MySQL + 'month VARCHAR(50), ';
        mySQL := MySQL + 'nameaddon VARCHAR(200), ';
        mySQL := MySQL + 'note VARCHAR(200), ';
        mySQL := MySQL + 'number VARCHAR(50), ';
        mySQL := MySQL + 'organization VARCHAR(200), ';
        mySQL := MySQL + 'origdate DATE, ';
        mySQL := MySQL + 'origlanguage VARCHAR(50), ';
        mySQL := MySQL + 'origlocation VARCHAR(200), ';
        mySQL := MySQL + 'origpublisher VARCHAR(200), ';
        mySQL := MySQL + 'origtitle VARCHAR(200), ';
        mySQL := MySQL + 'pages VARCHAR(50), ';
        mySQL := MySQL + 'pagetotal VARCHAR(50), ';
        mySQL := MySQL + 'pagination VARCHAR(200), ';
        mySQL := MySQL + 'part VARCHAR(200), ';
        mySQL := MySQL + 'publisher VARCHAR(200), ';
        mySQL := MySQL + 'pubstate VARCHAR(200), ';
        mySQL := MySQL + 'reprinttitle VARCHAR(200), ';
        mySQL := MySQL + 'review TEXT, '; // added by JabRef
        mySQL := MySQL + 'series VARCHAR(200), ';
        mySQL := MySQL + 'shortauthor VARCHAR(200), ';
        mySQL := MySQL + 'shorteditor VARCHAR(200), ';
        mySQL := MySQL + 'shorthand VARCHAR(200), ';
        mySQL := MySQL + 'shorthandintro VARCHAR(200), ';
        mySQL := MySQL + 'shortjournal VARCHAR(200), ';
        mySQL := MySQL + 'shortseries VARCHAR(200), ';
        mySQL := MySQL + 'shorttitle VARCHAR(200), ';
        mySQL := MySQL + 'subtitle VARCHAR(200), ';
        mySQL := MySQL + 'titleaddon VARCHAR(200), ';
        mySQL := MySQL + 'translator VARCHAR(200), ';
        mySQL := MySQL + 'type VARCHAR(200), ';
        mySQL := MySQL + 'url VARCHAR(200), ';
        mySQL := MySQL + 'urldate DATE, ';
        mySQL := MySQL + 'venue VARCHAR(200), ';
        mySQL := MySQL + 'version VARCHAR(200), ';
        mySQL := MySQL + 'volume VARCHAR(50), ';
        mySQL := MySQL + 'volumes VARCHAR(50), ';
        mySQL := MySQL + 'timestamp DATE, ';
        mySQL := MySQL + 'AttName TEXT, ';
        mySQL := MySQL + 'owned VARCHAR(1)); ';
        ExecSQL(MySQL);
        mySQL := 'CREATE UNIQUE INDEX idxIDItems ON Items (IDItems);';
        ExecSQL(MySQL);
        mySQL := 'CREATE UNIQUE INDEX idxbibtexkey ON Items (bibtexkey);';
        ExecSQL(MySQL);
        mySQL := 'CREATE INDEX idxauthor ON Items (author);';
        ExecSQL(MySQL);
        mySQL := 'CREATE INDEX idxtitle ON Items (title);';
        ExecSQL(MySQL);
        mySQL := 'CREATE INDEX idxyear ON Items (year);';
        ExecSQL(MySQL);
        mySQL := 'CREATE INDEX idxbooktitle ON Items (booktitle);';
        ExecSQL(MySQL);
        mySQL := 'CREATE INDEX idxjournaltitle ON Items (journaltitle);';
        ExecSQL(MySQL);
        mySQL := 'CREATE INDEX idxtimestamp ON Items (timestamp);';
        ExecSQL(MySQL);
      except
        MessageDlg('It is not possible to create the file.', mtError, [mbOK], 0);
      end;
  finally
    Screen.Cursor := crDefault;
  end;
end;

function TfmMain.MakeTableUpgrade(DataFileName: string): boolean;
var
  mySQL: string;
  sqUpgradeTables: TSqlite3Dataset;
begin
  // Update database structure
  Result := False;
  sqUpgradeTables := TSqlite3Dataset.Create(Self);
  with sqUpgradeTables do
    try
      FileName := DataFileName;
      TableName := 'Items';
      SQL := 'Select * from Items where IDItems = -1';
      Open;
      // Upgrade to 1.0.0
      if FieldCount < 101 then
      begin
        Screen.Cursor := crHourGlass;
        Application.ProcessMessages;
        Close;
        CopyFile(DataFileName, ExtractFileNameWithoutExt(DataFileName) +
          '.bak', [cffOverwriteFile]);
        mySQL := 'ALTER TABLE Items ADD crossref VARCHAR(200);';
        ExecSQL(MySQL);
        mySQL := 'ALTER TABLE Items ADD owned VARCHAR(1);';
        ExecSQL(MySQL);
        mySQL := 'ALTER TABLE Items ADD eventtitleaddon VARCHAR(200);';
        ExecSQL(MySQL);
        mySQL := 'CREATE INDEX idxbooktitle ON Items (booktitle);';
        ExecSQL(MySQL);
        mySQL := 'CREATE INDEX idxjournaltitle ON Items (journaltitle);';
        ExecSQL(MySQL);
        mySQL := 'CREATE INDEX idxtimestamp ON Items (timestamp);';
        ExecSQL(MySQL);
        Screen.Cursor := crDefault;
        Result := True;
        // Error Message is in OpenTable procedure
      end
    finally
      sqUpgradeTables.Free;
    end;
end;

procedure TfmMain.OpenDataTables(stFileName: string);
begin
  // Open a file
  // File name could be the last opened or specified program name, so...
  Screen.Cursor := crHourGlass;
  Application.ProcessMessages;
  if FileExistsUTF8(stFileName) then
  begin
    // Close tables
    CloseDataTables;
    // A file may be already opened, so...
    sqItems.Close;
    sqItems.FileName := stFileName;
    sqItems.TableName := 'Items';
    sqItems.SQL := 'Select * from Items ' + CreateOrderBy;
    sqItems.PrimaryKey := 'IDItems';
    try
      try
        sqItems.Open;
        // Column of ID, aligned to right
        grItems.Columns.Items[0].Alignment := taLeftJustify;
        grItems.Columns.Items[0].Title.Alignment := taLeftJustify;
        pnMain.Visible := True;
        sqItems.Active := True;
        if miItemsOrderByID.Checked = True then
        begin
          sqItems.Last;
        end;
        miFileClose.Enabled := True;
        miFileCopyAs.Enabled := True;
        miFileCreateBibLatex.Enabled := True;
        tbFileCreateBibLatex.Enabled := True;
        miFileExportToBiblatex.Enabled := True;
        miFileImportFromBiblatex.Enabled := True;
        miFileExportToBifilex.Enabled := True;
        miFileImportFromBibfilex.Enabled := True;
        miItemsNew.Enabled := True;
        pmItemsNew.Enabled := True;
        tbItemsNew.Enabled := True;
        miItemsCreateKeywordList.Enabled := True;
        miItemsRenameKeyword.Enabled := True;
        miItemsStoreKeyword.Enabled := True;
        miItemsApplyKeyword.Enabled := True;
        miItemsFilterFromLatex.Enabled := True;
        miItemsRemoveFilter.Enabled := True;
        miItemsOrderBy.Enabled := True;
        miItemsCopyCrossref.Enabled := True;
        miItemsFilter.Enabled := True;
        tbItemsFilter.Enabled := True;
        miToolChar.Enabled := True;
        miToolKeyList.Enabled := True;
        miToolsReplaceCitations.Enabled := True;
        miToolsConvert.Enabled := True;
        miToolsReplaceKeys.Enabled := True;
        // Fill keyboard filter combo
        CreateKeywordList;
        // Update last archives menu
        if sqItems.FileName = LastDatabase2 then
        begin
          LastDatabase2 := LastDatabase1;
          LastDatabase1 := sqItems.FileName;
        end
        else if sqItems.FileName = LastDatabase3 then
        begin
          LastDatabase3 := LastDatabase2;
          LastDatabase2 := LastDatabase1;
          LastDatabase1 := sqItems.FileName;
        end
        else if sqItems.FileName <> LastDatabase1 then
        begin
          LastDatabase4 := LastDatabase3;
          LastDatabase3 := LastDatabase2;
          LastDatabase2 := LastDatabase1;
          LastDatabase1 := sqItems.FileName;
        end;
        miLineLastFile.Visible := False;
        if LastDatabase1 <> '' then
        begin
          miFileOpenLast1.Caption := ExtractFileNameOnly(LastDatabase1);
          miFileOpenLast1.Visible := True;
          miLineLastFile.Visible := True;
        end;
        if LastDatabase2 <> '' then
        begin
          miFileOpenLast2.Caption := ExtractFileNameOnly(LastDatabase2);
          miFileOpenLast2.Visible := True;
          miLineLastFile.Visible := True;
        end;
        if LastDatabase3 <> '' then
        begin
          miFileOpenLast3.Caption := ExtractFileNameOnly(LastDatabase3);
          miFileOpenLast3.Visible := True;
          miLineLastFile.Visible := True;
        end;
        if LastDatabase4 <> '' then
        begin
          miFileOpenLast4.Caption := ExtractFileNameOnly(LastDatabase4);
          miFileOpenLast4.Visible := True;
          miLineLastFile.Visible := True;
        end;
        // ExtractFileName does not show the path
        fmMain.Caption := 'Bibfilex - ' + ExtractFilePath(sqItems.FileName) +
          ExtractFileNameOnly(sqItems.FileName);
        if edSmartFilter.Visible = True then
          edSmartFilter.SetFocus;
      except
        // See notes at the bottom for this message
        MessageDlg('It is not possibile to open the file. ' +
          'If a new version of Bibfilex has been installed, ' +
          'try to upgrade the structure of the file ' +
          'with the menu item Tools - Upgrade data, ' +
          'then quit the software, run it again ' +
          'and open the new upgraded file.',
          mtError, [mbOK], 0);
        // Dataset remains open, so...
        sqItems.Active := False;
        pnMain.Visible := False;
        fmMain.Caption := 'Bibfilex';
      end;
    finally
      Screen.Cursor := crDefault;
    end;
  end;
end;

procedure TfmMain.CloseDataTables;
var
  i: integer;
begin
  // Close the file
  SaveItems;
  pnMain.Visible := False;
  sqItems.Active := False;
  miFileClose.Enabled := False;
  miFileCopyAs.Enabled := False;
  miFileCreateBibLatex.Enabled := False;
  tbFileCreateBibLatex.Enabled := False;
  miFileExportToBiblatex.Enabled := False;
  miFileImportFromBiblatex.Enabled := False;
  miFileExportToBifilex.Enabled := False;
  miFileImportFromBibfilex.Enabled := False;
  miItemsNew.Enabled := False;
  pmItemsNew.Enabled := False;
  tbItemsNew.Enabled := False;
  miItemsSave.Enabled := False;
  pmItemsSave.Enabled := False;
  tbItemsSave.Enabled := False;
  miItemsUndo.Enabled := False;
  pmItemsUndo.Enabled := False;
  miItemsDelete.Enabled := False;
  pmItemsDelete.Enabled := False;
  miItemsAttachments.Enabled := False;
  miAttachNew.Enabled := False;
  miAttachOpen.Enabled := False;
  miAttachSaveAs.Enabled := False;
  miAttachDelete.Enabled := False;
  pmAttNew.Enabled := False;
  pmAttOpen.Enabled := False;
  pmAttSaveAs.Enabled := False;
  pmAttDelete.Enabled := False;
  miItemsCreateKey.Enabled := False;
  miItemsModifyKey.Enabled := False;
  miItemsCopyKey.Enabled := False;
  miItemsCopyCitationTxt.Enabled := False;
  miItemsCopyCitationHtml.Enabled := False;
  miItemsCreateKeywordList.Enabled := False;
  miItemsRenameKeyword.Enabled := False;
  miItemsStoreKeyword.Enabled := False;
  miItemsApplyKeyword.Enabled := False;
  miItemsFilterFromLatex.Enabled := False;
  miItemsRemoveFilter.Enabled := False;
  miItemsOrderBy.Enabled := False;
  miItemsCopyCrossref.Enabled := False;
  miItemsFilter.Enabled := False;
  tbItemsFilter.Enabled := False;
  miToolChar.Enabled := False;
  miToolKeyList.Enabled := False;
  miToolsReplaceCitations.Enabled := False;
  miToolsConvert.Enabled := False;
  miToolsReplaceKeys.Enabled := False;
  cbFilterKey.Text := '';
  flFilterActive := 0;
  // Clear the bookmarks
  for i := 0 to 9 do
    BookmarkList[i] := '';
  sbStatusBar.SimpleText := 'No file in use.';
  fmMain.Caption := 'Bibfilex';
end;

procedure TfmMain.LoadItems;
var
  i: integer;
begin
  // Load items from DB
  // If there are no records, see dsItemsDataChange event
  if sqItems.RecordCount > 0 then
  begin
    // Entry type data
    cbItemsKind.ItemIndex :=
      cbItemsKind.Items.IndexOf(sqItems.FieldByName('entrytype').AsString);
    // Set fields in grids
    UnitSetFields.SetFieldsInGrid;
    // Required fields
    for i := 0 to sgRequiredFields.RowCount - 1 do
      sgRequiredFields.Cells[1, i] :=
        sqItems.FieldByName(LowerCase(sgRequiredFields.Cells[0, i])).AsString;
    // Optional fields 1
    for i := 0 to sgOptionalFields1.RowCount - 1 do
      sgOptionalFields1.Cells[1, i] :=
        sqItems.FieldByName(LowerCase(sgOptionalFields1.Cells[0, i])).AsString;
    // Optional fields 2
    for i := 0 to sgOptionalFields2.RowCount - 1 do
      sgOptionalFields2.Cells[1, i] :=
        sqItems.FieldByName(LowerCase(sgOptionalFields2.Cells[0, i])).AsString;
    // Optional fields 3
    for i := 0 to sgOptionalFields3.RowCount - 1 do
      sgOptionalFields3.Cells[1, i] :=
        sqItems.FieldByName(LowerCase(sgOptionalFields3.Cells[0, i])).AsString;
    // Abstract and Review
    meAbstract.Text := sqItems.FieldByName('abstract').AsString;
    meReview.Text := sqItems.FieldByName('review').AsString;
    CompileSummary;
    // Last change date
    if sqItems.FieldByName('timestamp').AsDateTime > 0 then
      lbLastEdited.Caption := 'Last modified on ' +
        FormatDateTime('ddddd', sqItems.FieldByName('timestamp').AsDateTime) +
        ' at ' + FormatDateTime('t', sqItems.FieldByName(
        'timestamp').AsDateTime) + '.'
    else
      lbLastEdited.Caption := 'No last change date available.';
    // To update the attachment list
    lbAttNames.Items.Text := sqItems.FieldByName('AttName').AsString;
    SetAttachmentMenuItems;
    // Changes has set edit flag to True, so...
    flDataEdit := False;
    miItemsSave.Enabled := False;
    pmItemsSave.Enabled := False;
    tbItemsSave.Enabled := False;
    miItemsUndo.Enabled := False;
    pmItemsUndo.Enabled := False;
  end;
  SetMarkAbstractReviewTab;
end;

procedure TfmMain.NewItem(iKind: Short);
begin
  // New item: 0=Article; 1=Book
  sqItems.Append;
  // Set book as the default type
  if iKind = 0 then
    sqItems.FieldByName('entrytype').AsString := 'article'
  else
    sqItems.FieldByName('entrytype').AsString := 'book';
  sqItems.FieldByName('timestamp').AsDateTime := Now;
  sqItems.Post;
  sqItems.ApplyUpdates;
  if iKind = 0 then
    cbItemsKind.ItemIndex := 0
  else
    cbItemsKind.ItemIndex := 1;
  UnitSetFields.SetFieldsInGrid;
  pcPageControl.ActivePage := tsRequiredFields;
  sgRequiredFields.Row := 0;
  if cbItemsKind.Visible = True then
    cbItemsKind.SetFocus;
end;

procedure TfmMain.SaveItems;
var
  i: integer;
begin
  if ((sqItems.Active = True) and (flDataEdit = True) and
    (sqItems.RecordCount > 0)) then
    try
      // Save items to DB
      if cbItemsKind.ItemIndex = -1 then
      begin
        MessageDlg('The entry type has not been defined.', mtWarning, [mbOK], 0);
        Abort;
      end;
      Screen.Cursor := crHourGlass;
      Application.ProcessMessages;
      sqItems.Edit;
      // Entry type data
      if sqItems.FieldByName('entrytype').AsString <>
        LowerCase(cbItemsKind.Text) then
        sqItems.FieldByName('entrytype').AsString :=
          LowerCase(cbItemsKind.Text);
      // Required fields
      for i := 0 to sgRequiredFields.RowCount - 1 do
      begin
        // Clean keyword field
        if LowerCase(sgRequiredFields.Cells[0, i]) = 'keywords' then
          sgRequiredFields.Cells[1, i] :=
            CleanKeywordField(sgRequiredFields.Cells[1, i]);
        // Check that BibText key is existing, unique or modify it
        if LowerCase(sgRequiredFields.Cells[0, i]) = 'bibtexkey' then
        begin
          if sgRequiredFields.Cells[1, i] = '' then
          begin
            sgRequiredFields.Cells[1, i] := CreateBibTexKey(True, False);
          end
          else
          begin
            sgRequiredFields.Cells[1, i] :=
              CheckUniqueBibTexKey(sgRequiredFields.Cells[1, i],
              sqItems.FieldByName('IDItems').AsInteger);
          end;
        end;
        // Other fields
        if sqItems.FieldByName(LowerCase(sgRequiredFields.Cells[0, i])).AsString <>
          sgRequiredFields.Cells[1, i] then
          sqItems.FieldByName(LowerCase(sgRequiredFields.Cells[0, i])).AsString :=
            sgRequiredFields.Cells[1, i];
      end;
      // Optional fields 1
      for i := 0 to sgOptionalFields1.RowCount - 1 do
        if sqItems.FieldByName(LowerCase(sgOptionalFields1.Cells[0, i])).AsString <>
          sgOptionalFields1.Cells[1, i] then
          sqItems.FieldByName(LowerCase(sgOptionalFields1.Cells[0, i])).AsString :=
            sgOptionalFields1.Cells[1, i];
      // Optional fields 2
      for i := 0 to sgOptionalFields2.RowCount - 1 do
        if sqItems.FieldByName(LowerCase(sgOptionalFields2.Cells[0, i])).AsString <>
          sgOptionalFields2.Cells[1, i] then
          sqItems.FieldByName(LowerCase(sgOptionalFields2.Cells[0, i])).AsString :=
            sgOptionalFields2.Cells[1, i];
      // Optional fields 3
      for i := 0 to sgOptionalFields3.RowCount - 1 do
        if sqItems.FieldByName(LowerCase(sgOptionalFields3.Cells[0, i])).AsString <>
          sgOptionalFields3.Cells[1, i] then
          sqItems.FieldByName(LowerCase(sgOptionalFields3.Cells[0, i])).AsString :=
            sgOptionalFields3.Cells[1, i];
      // Abstract and Review
      if sqItems.FieldByName('abstract').AsString <> meAbstract.Text then
        sqItems.FieldByName('abstract').AsString := meAbstract.Text;
      if sqItems.FieldByName('review').AsString <> meReview.Text then
        sqItems.FieldByName('review').AsString := meReview.Text;
      // Set date of last change
      sqItems.FieldByName('timestamp').AsDateTime := Now;
      try
        sqItems.Post;
        sqItems.ApplyUpdates;
        flDataEdit := False;
        flExportOnExit := True;
        miItemsSave.Enabled := False;
        pmItemsSave.Enabled := False;
        tbItemsSave.Enabled := False;
        miItemsUndo.Enabled := False;
        pmItemsUndo.Enabled := False;
      except
        MessageDlg('It is not possibile to save the item; ' +
          'check that the BibTex key is unique ' + 'within the file in use.',
          mtWarning, [mbOK], 0);
      end;
    finally
      Screen.Cursor := crDefault;
    end;
end;

procedure TfmMain.ActivateSaving;
begin
  // Activate save buttons
  miItemsSave.Enabled := flDataEdit;
  tbItemsSave.Enabled := flDataEdit;
  pmItemsSave.Enabled := flDataEdit;
  miItemsUndo.Enabled := flDataEdit;
  pmItemsUndo.Enabled := flDataEdit;
end;

procedure TfmMain.ImportBibLatex(stFileName: string);
var
  FileBibLatex: TextFile;
  myZipper: TZipper;
  stLine, stFieldName, stFiles, stImpFile: string;
  NumImp: integer = 0;
  flInField: boolean = False;
begin
  // Import Biblatex file
  AssignFile(FileBibLatex, stFileName);
  Reset(FileBibLatex);
  sqToolsTables.Close;
  sqToolsTables.FileName := sqItems.FileName;
  sqToolsTables.TableName := 'Items';
  sqToolsTables.PrimaryKey := 'IDItems';
  sqToolsTables.Open;
  Screen.Cursor := crHourGlass;
  Application.ProcessMessages;
  try
    while not EOF(FileBibLatex) do
      try
        ReadLn(FileBibLatex, stLine);
        if stLine = '' then
        begin
          // No empty lines, if not within a field
          if flInField = False then
          begin
            Continue;
          end;
        end;
        // Do not clear the double brackets, since they may surround a word
        // Remove tabs and spaces at the beginning, at the end and double spaces
        stLine := StringReplace(stLine, #9, ' ', [rfReplaceAll]);
        while UTF8Pos('  ', stLine) > 0 do
          stLine := StringReplace(stLine, '  ', ' ', []);
        while UTF8Copy(stLine, 1, 1) = ' ' do
          stLine := UTF8Copy(stLine, 2, UTF8Length(stLine));
        while UTF8Copy(stLine, UTF8Length(stLine), 1) = ' ' do
          stLine := UTF8Copy(stLine, 1, UTF8Length(stLine) - 1);
        // Replace the \& values
        stLine := StringReplace(stLine, '\&', '&', [rfReplaceAll]);
        // Replace the '' as "
        stLine := StringReplace(stLine, #39#39, '"', [rfReplaceAll]);
        // Lower case
        // Mendeley escape
        stLine := StringReplace(stLine, '\`{a}', 'à', [rfReplaceAll]);
        stLine := StringReplace(stLine, '\''{a}', 'á', [rfReplaceAll]);
        stLine := StringReplace(stLine, '\^{a}', 'â', [rfReplaceAll]);
        stLine := StringReplace(stLine, '\~{a}', 'ã', [rfReplaceAll]);
        stLine := StringReplace(stLine, '\"{a}', 'ä', [rfReplaceAll]);
        stLine := StringReplace(stLine, '\`{e}', 'è', [rfReplaceAll]);
        stLine := StringReplace(stLine, '\''{e}', 'é', [rfReplaceAll]);
        stLine := StringReplace(stLine, '\^{e}', 'ê', [rfReplaceAll]);
        stLine := StringReplace(stLine, '\"{e}', 'ë', [rfReplaceAll]);
        stLine := StringReplace(stLine, '\`{i}', 'ì', [rfReplaceAll]);
        stLine := StringReplace(stLine, '\''{i}', 'í', [rfReplaceAll]);
        stLine := StringReplace(stLine, '\^{i}', 'î', [rfReplaceAll]);
        stLine := StringReplace(stLine, '\"{i}', 'ï', [rfReplaceAll]);
        stLine := StringReplace(stLine, '\`{o}', 'ò', [rfReplaceAll]);
        stLine := StringReplace(stLine, '\''{o}', 'ó', [rfReplaceAll]);
        stLine := StringReplace(stLine, '\^{o}', 'ô', [rfReplaceAll]);
        stLine := StringReplace(stLine, '\~{o}', 'õ', [rfReplaceAll]);
        stLine := StringReplace(stLine, '\"{o}', 'ö', [rfReplaceAll]);
        stLine := StringReplace(stLine, '\`{u}', 'ù', [rfReplaceAll]);
        stLine := StringReplace(stLine, '\''{u}', 'ú', [rfReplaceAll]);
        stLine := StringReplace(stLine, '\^{u}', 'û', [rfReplaceAll]);
        stLine := StringReplace(stLine, '\"{u}', 'ü', [rfReplaceAll]);
        // Google escape
        stLine := StringReplace(stLine, '{\`a}', 'à', [rfReplaceAll]);
        stLine := StringReplace(stLine, '{\''a}', 'á', [rfReplaceAll]);
        stLine := StringReplace(stLine, '{\^a}', 'â', [rfReplaceAll]);
        stLine := StringReplace(stLine, '{\~a}', 'ã', [rfReplaceAll]);
        stLine := StringReplace(stLine, '{\"a}', 'ä', [rfReplaceAll]);
        stLine := StringReplace(stLine, '{\`e}', 'è', [rfReplaceAll]);
        stLine := StringReplace(stLine, '{\''e}', 'é', [rfReplaceAll]);
        stLine := StringReplace(stLine, '{\^e}', 'ê', [rfReplaceAll]);
        stLine := StringReplace(stLine, '{\"e}', 'ë', [rfReplaceAll]);
        stLine := StringReplace(stLine, '{\`i}', 'ì', [rfReplaceAll]);
        stLine := StringReplace(stLine, '{\''i}', 'í', [rfReplaceAll]);
        stLine := StringReplace(stLine, '{\^i}', 'î', [rfReplaceAll]);
        stLine := StringReplace(stLine, '{\"i}', 'ï', [rfReplaceAll]);
        stLine := StringReplace(stLine, '{\`o}', 'ò', [rfReplaceAll]);
        stLine := StringReplace(stLine, '{\''o}', 'ó', [rfReplaceAll]);
        stLine := StringReplace(stLine, '{\^o}', 'ô', [rfReplaceAll]);
        stLine := StringReplace(stLine, '{\~o}', 'õ', [rfReplaceAll]);
        stLine := StringReplace(stLine, '{\"o}', 'ö', [rfReplaceAll]);
        stLine := StringReplace(stLine, '{\`u}', 'ù', [rfReplaceAll]);
        stLine := StringReplace(stLine, '{\''u}', 'ú', [rfReplaceAll]);
        stLine := StringReplace(stLine, '{\^u}', 'û', [rfReplaceAll]);
        stLine := StringReplace(stLine, '{\"u}', 'ü', [rfReplaceAll]);
        // Upper case
        // Mendeley escape
        stLine := StringReplace(stLine, '\`{A}', UTF8UpperCase('à'), [rfReplaceAll]);
        stLine := StringReplace(stLine, '\''{A}', UTF8UpperCase('á'), [rfReplaceAll]);
        stLine := StringReplace(stLine, '\^{A}', UTF8UpperCase('â'), [rfReplaceAll]);
        stLine := StringReplace(stLine, '\~{A}', UTF8UpperCase('ã'), [rfReplaceAll]);
        stLine := StringReplace(stLine, '\"{A}', UTF8UpperCase('ä'), [rfReplaceAll]);
        stLine := StringReplace(stLine, '\`{E}', UTF8UpperCase('è'), [rfReplaceAll]);
        stLine := StringReplace(stLine, '\''{E}', UTF8UpperCase('é'), [rfReplaceAll]);
        stLine := StringReplace(stLine, '\^{E}', UTF8UpperCase('ê'), [rfReplaceAll]);
        stLine := StringReplace(stLine, '\"{E}', UTF8UpperCase('ë'), [rfReplaceAll]);
        stLine := StringReplace(stLine, '\`{I}', UTF8UpperCase('ì'), [rfReplaceAll]);
        stLine := StringReplace(stLine, '\''{I}', UTF8UpperCase('í'), [rfReplaceAll]);
        stLine := StringReplace(stLine, '\^{I}', UTF8UpperCase('î'), [rfReplaceAll]);
        stLine := StringReplace(stLine, '\"{I}', UTF8UpperCase('ï'), [rfReplaceAll]);
        stLine := StringReplace(stLine, '\`{O}', UTF8UpperCase('ò'), [rfReplaceAll]);
        stLine := StringReplace(stLine, '\''{O}', UTF8UpperCase('ó'), [rfReplaceAll]);
        stLine := StringReplace(stLine, '\^{O}', UTF8UpperCase('ô'), [rfReplaceAll]);
        stLine := StringReplace(stLine, '\~{O}', UTF8UpperCase('õ'), [rfReplaceAll]);
        stLine := StringReplace(stLine, '\"{O}', UTF8UpperCase('ö'), [rfReplaceAll]);
        stLine := StringReplace(stLine, '\`{U}', UTF8UpperCase('ù'), [rfReplaceAll]);
        stLine := StringReplace(stLine, '\''{U}', UTF8UpperCase('ú'), [rfReplaceAll]);
        stLine := StringReplace(stLine, '\^{U}', UTF8UpperCase('û'), [rfReplaceAll]);
        stLine := StringReplace(stLine, '\"{U}', UTF8UpperCase('ü'), [rfReplaceAll]);
        // Google escape
        stLine := StringReplace(stLine, '{\`A}', UTF8UpperCase('à'), [rfReplaceAll]);
        stLine := StringReplace(stLine, '{\''A}', UTF8UpperCase('á'), [rfReplaceAll]);
        stLine := StringReplace(stLine, '{\^A}', UTF8UpperCase('â'), [rfReplaceAll]);
        stLine := StringReplace(stLine, '{\~A}', UTF8UpperCase('ã'), [rfReplaceAll]);
        stLine := StringReplace(stLine, '{\"A}', UTF8UpperCase('ä'), [rfReplaceAll]);
        stLine := StringReplace(stLine, '{\`E}', UTF8UpperCase('è'), [rfReplaceAll]);
        stLine := StringReplace(stLine, '{\''E}', UTF8UpperCase('é'), [rfReplaceAll]);
        stLine := StringReplace(stLine, '{\^E}', UTF8UpperCase('ê'), [rfReplaceAll]);
        stLine := StringReplace(stLine, '{\"E}', UTF8UpperCase('ë'), [rfReplaceAll]);
        stLine := StringReplace(stLine, '{\`I}', UTF8UpperCase('ì'), [rfReplaceAll]);
        stLine := StringReplace(stLine, '{\''I}', UTF8UpperCase('í'), [rfReplaceAll]);
        stLine := StringReplace(stLine, '{\^I}', UTF8UpperCase('î'), [rfReplaceAll]);
        stLine := StringReplace(stLine, '{\"I}', UTF8UpperCase('ï'), [rfReplaceAll]);
        stLine := StringReplace(stLine, '{\`O}', UTF8UpperCase('ò'), [rfReplaceAll]);
        stLine := StringReplace(stLine, '{\''O}', UTF8UpperCase('ó'), [rfReplaceAll]);
        stLine := StringReplace(stLine, '{\^O}', UTF8UpperCase('ô'), [rfReplaceAll]);
        stLine := StringReplace(stLine, '{\~O}', UTF8UpperCase('õ'), [rfReplaceAll]);
        stLine := StringReplace(stLine, '{\"O}', UTF8UpperCase('ö'), [rfReplaceAll]);
        stLine := StringReplace(stLine, '{\`U}', UTF8UpperCase('ù'), [rfReplaceAll]);
        stLine := StringReplace(stLine, '{\''U}', UTF8UpperCase('ú'), [rfReplaceAll]);
        stLine := StringReplace(stLine, '{\^U}', UTF8UpperCase('û'), [rfReplaceAll]);
        stLine := StringReplace(stLine, '{\"U}', UTF8UpperCase('ü'), [rfReplaceAll]);
        // Beginning of (JabRef) comments
        if UTF8Copy(stLine, 1, 9) = '@comment{' then
          Continue
        else if UTF8Copy(stLine, 1, 1) = '%' then
          Continue
        // Beginning of an item
        else if UTF8Copy(stLine, 1, 1) = '@' then
        begin
          sqToolsTables.Append;
          sqToolsTables.FieldByName('entrytype').AsString :=
            LowerCase(UTF8Copy(stLine, 2, UTF8Pos('{', stLine) - 2));
          sqToolsTables.FieldByName('bibtexkey').AsString :=
            UTF8Copy(stLine, UTF8Pos('{', stLine) + 1, UTF8Pos(',', stLine) -
            UTF8Pos('{', stLine) - 1);
        end
        // End of an item
        else if stLine = '}' then
        begin
          if dsToolsTables.State in [dsEdit, dsInsert] then
          begin
            // Change used BibTex keys
            if sqToolsTables.FieldByName('bibtexkey').AsString <> '' then
            begin
              sqToolsTables.FieldByName('bibtexkey').AsString :=
                CheckUniqueBibTexKey(
                sqToolsTables.FieldByName('bibtexkey').AsString, -1);
            end;
            // Clean keywords
            if sqToolsTables.FieldByName('keywords').AsString <> '' then
            begin
              sqToolsTables.FieldByName('keywords').AsString :=
                CleanKeywordField(
                sqToolsTables.FieldByName('keywords').AsString);
            end;
            if sqToolsTables.FieldByName('file').AsString <> '' then
            begin
              stFiles := sqToolsTables.FieldByName('file').AsString;
              sqToolsTables.FieldByName('file').AsString := '';
              stFiles := StringReplace(stFiles, ';', ':', [rfReplaceAll]);
              if UTF8Copy(stFiles, 1, 1) = ':' then
              begin
                stFiles := UTF8Copy(stFiles, 2, UTF8Length(stFiles));
              end;
              stFiles := stFiles + ':';
              while UTF8Pos(':', stFiles) > 1 do
              begin
                stImpFile := '';
                if FileExistsUTF8(UTF8Copy(stFiles, 1,
                  UTF8Pos(':', stFiles) - 1)) = True then
                begin
                  stImpFile :=
                    UTF8Copy(stFiles, 1, UTF8Pos(':', stFiles) - 1);
                end
                else
                if FileExistsUTF8(DirectorySeparator +
                  UTF8Copy(stFiles, 1, UTF8Pos(':', stFiles) - 1)) = True then
                begin
                  stImpFile :=
                    DirectorySeparator + UTF8Copy(stFiles, 1, UTF8Pos(':', stFiles) - 1);
                end
                else
                if FileExistsUTF8(ExtractFileDir(stFileName) +
                  DirectorySeparator + UTF8Copy(stFiles, 1,
                  UTF8Pos(':', stFiles) - 1)) = True then
                begin
                  stImpFile :=
                    ExtractFileDir(stFileName) + DirectorySeparator +
                    UTF8Copy(stFiles, 1, UTF8Pos(':', stFiles) - 1);
                end;
                if stImpFile <> '' then
                  try
                    if DirectoryExistsUTF8(
                      ExtractFileNameWithoutExt(sqToolsTables.FileName)) = False then
                    begin
                      CreateDirUTF8(ExtractFileNameWithoutExt(sqToolsTables.FileName));
                    end;
                    myZipper := TZipper.Create;
                    myZipper.FileName :=
                      ExtractFileNameWithoutExt(sqToolsTables.FileName) +
                      DirectorySeparator +
                      sqToolsTables.FieldByName('IDItems').AsString +
                      '-' + ExtractFileName(stImpFile) + '.zip';
                    myZipper.Entries.AddFileEntry(stImpFile,
                      ExtractFileName(stImpFile));
                    myZipper.ZipAllFiles;
                    if sqToolsTables.FieldByName('AttName').AsString = '' then
                    begin
                      sqToolsTables.FieldByName('AttName').AsString :=
                        ExtractFileName(stImpFile);
                    end
                    else
                    begin
                      sqToolsTables.FieldByName('AttName').AsString :=
                        sqToolsTables.FieldByName('AttName').AsString +
                        LineEnding + ExtractFileName(stImpFile);
                    end;
                  finally
                    myZipper.Free;
                  end;
                stFiles := UTF8Copy(stFiles, UTF8Pos(':', stFiles) +
                  1, UTF8Length(stFiles));
              end;
            end;
            sqToolsTables.Post;
            Inc(NumImp);
            sbStatusBar.SimpleText := 'Items imported: ' + IntToStr(NumImp) + '.';
            Application.ProcessMessages;
          end;
        end
        else if ((stLine = '') and (flInField = True)) then
        begin
          if ((LowerCase(stFieldName) = 'abstract') or
            (LowerCase(stFieldName) = 'review')) then
          begin
            sqToolsTables.FieldByName(stFieldName).AsString :=
              sqToolsTables.FieldByName(stFieldName).AsString +
              LineEnding + LineEnding;
          end
          else
          begin
            Continue;
          end;
        end
        // New field
        else if ((UTF8Pos('{', stLine) > 0) and (flInField = False)) then
        begin
          // The = may have no space before and after
          stLine := StringReplace(stLine, '=', ' = ', []);
          while UTF8Pos('  ', stLine) > 0 do
            stLine := StringReplace(stLine, '  ', ' ', []);
          stFieldName := UTF8Copy(stLine, 1, UTF8Pos('=', stLine) - 2);
          // In Biblatex comment is not used
          if LowerCase(stFieldName) = 'comment' then
            stFieldName := 'note';
          // In Mendeley address is used
          if LowerCase(stFieldName) = 'address' then
            stFieldName := 'location';
          // In Mendeley the keywords are exported with commas without spaces,
          // but this is fixed just before saving the item
          // Check if the field is existing
          if ((sqToolsTables.FindField(stFieldName) <> nil) and
            (LowerCase(stFieldName) <> 'attname')) then
          begin
            // In JabRef, etc. the date may be separated by points, etc.
            if LowerCase(stFieldName) = 'timestamp' then
            begin
              stLine := StringReplace(stLine, '.', DateSeparator, [rfReplaceAll]);
              stLine := StringReplace(stLine, '/', DateSeparator, [rfReplaceAll]);
              stLine := StringReplace(stLine, '-', DateSeparator, [rfReplaceAll]);
            end;
            // Only line, non last field
            if UTF8Copy(stLine, UTF8Length(stLine) - 1, 2) = '},' then
            begin
              sqToolsTables.FieldByName(stFieldName).AsString :=
                UTF8Copy(stLine, UTF8Pos('{', stLine) + 1,
                UTF8Length(stLine) - UTF8Pos('{', stLine) - 2);
            end
            // Only line, last field
            else
            if UTF8Copy(stLine, UTF8Length(stLine), 1) = '}' then
            begin
              sqToolsTables.FieldByName(stFieldName).AsString :=
                UTF8Copy(stLine, UTF8Pos('{', stLine) + 1,
                UTF8Length(stLine) - UTF8Pos('{', stLine) - 1);
            end
            // First line of a multiline field
            else
            begin
              sqToolsTables.FieldByName(stFieldName).AsString :=
                UTF8Copy(stLine, UTF8Pos('{', stLine) + 1,
                UTF8Length(stLine) - UTF8Pos('{', stLine)) + ' ';
              flInField := True;
            end;
          end;
        end
        // Line following the first of a multiline field
        else if flInField = True then
        begin
          // In Mendeley the keywords are exported with commas without spaces
          if LowerCase(stFieldName) = 'keywords' then
          begin
            stLine := StringReplace(stLine, ',', ', ', [rfReplaceAll]);
            while UTF8Pos('  ', stLine) > 0 do
              stLine := StringReplace(stLine, '  ', ' ', []);
          end;
          // Last line
          if ((UTF8Copy(stLine, UTF8Length(stLine), 1) = '}') or
            (UTF8Copy(stLine, UTF8Length(stLine) - 1, 2) = '},')) then
          begin
            sqToolsTables.FieldByName(stFieldName).AsString :=
              sqToolsTables.FieldByName(stFieldName).AsString +
              UTF8Copy(stLine, 1, UTF8Pos('}', stLine) - 1);
            // Spaces may come
            while UTF8Pos('  ', sqToolsTables.FieldByName(
                stFieldName).AsString) > 0 do
              sqToolsTables.FieldByName(stFieldName).AsString :=
                StringReplace(sqToolsTables.FieldByName(stFieldName).AsString,
                '  ', ' ', []);
            flInField := False;
          end
          // Middle line
          else
          begin
            if ((LowerCase(stFieldName) = 'abstract') or
              (LowerCase(stFieldName) = 'review')) then
            begin
              sqToolsTables.FieldByName(stFieldName).AsString :=
                sqToolsTables.FieldByName(stFieldName).AsString +
                LineEnding + UTF8Copy(stLine, 1, UTF8Length(stLine));
            end
            else
            begin
              sqToolsTables.FieldByName(stFieldName).AsString :=
                sqToolsTables.FieldByName(stFieldName).AsString +
                UTF8Copy(stLine, 1, UTF8Length(stLine)) + ' ';
            end;
          end;
        end;
      except
        if MessageDlg('Importation error on item n. ' + IntToStr(NumImp) +
          '. ' + 'Check that the date format of the items to be imported ' +
          'is the same that is specified in the options of the software ' +
          'and that the date fields are not among double curly brackets. ' +
          'Continue or abort the procedure?', mtError, [mbAbort, mbIgnore], 0) =
          mrAbort then
        begin
          sqToolsTables.Cancel;
          Exit;
        end
        else
        begin
          Continue;
        end;
      end;
  finally
    sqToolsTables.ApplyUpdates;
    sqToolsTables.Close;
    sqItems.Close;
    sqItems.Open;
    CreateKeywordList;
    CloseFile(FileBibLatex);
    Screen.Cursor := crDefault;
  end;
end;

procedure TfmMain.ExportToBibLatex(stFileName: string; Filtered: boolean);
var
  myUnZip: TUnZipper;
  FileBibLatex: TextFile;
  i, n, NumExp: integer;
  stItem, stOpenBrakets, stCloseBrackets, stAuthEdit: string;
  slAtt: TStringList;
begin
  //Export to Biblatex file
  Screen.Cursor := crHourGlass;
  Application.ProcessMessages;
  try
    AssignFile(FileBibLatex, stFileName);
    Rewrite(FileBibLatex);
    WriteLn(FileBibLatex, '% This file was created with Bibfilex');
    WriteLn(FileBibLatex, '');
    NumExp := 0;
    if fmOptions.cbExpDoubleBrakets.Checked = True then
    begin
      stOpenBrakets := '{{';
      stCloseBrackets := '}}';
    end
    else
    begin
      stOpenBrakets := '{';
      stCloseBrackets := '}';
    end;
    sqToolsTables.Close;
    sqToolsTables.FileName := sqItems.FileName;
    sqToolsTables.TableName := 'Items';
    sqToolsTables.PrimaryKey := 'IDItems';
    if Filtered = True then
      sqToolsTables.SQL := sqItems.SQL
    else
      sqToolsTables.SQL := 'Select * from Items ' + CreateOrderBy;
    sqToolsTables.Open;
    while not sqToolsTables.EOF do
    begin
      stItem := '';
      stItem := '@' + UTF8UpperCase(sqToolsTables.FieldByName('entrytype').AsString) +
        '{' + sqToolsTables.FieldByName('bibtexkey').AsString;
      // From author to the last field
      for i := 3 to sqToolsTables.Fields.Count - 2 do
      begin
        if sqToolsTables.Fields[i].AsString <> '' then
        begin
          if ((LowerCase(sqToolsTables.Fields[i].FieldName) = 'author') or
            (LowerCase(sqToolsTables.Fields[i].FieldName) = 'editor') or
            (LowerCase(sqToolsTables.Fields[i].FieldName) = 'editora') or
            (LowerCase(sqToolsTables.Fields[i].FieldName) = 'editorb') or
            (LowerCase(sqToolsTables.Fields[i].FieldName) = 'editorc')) then
          begin
            stAuthEdit := sqToolsTables.Fields[i].AsString;
            // One or more institutional authors
            if ((UTF8Pos(',', stAuthEdit) < 1) and
              (UTF8Pos(' ', stAuthEdit) > 0)) then
            begin
              stAuthEdit := '{' + stAuthEdit + '}';
              if UTF8Pos(' and ', stAuthEdit) > 0 then
              begin
                stAuthEdit :=
                  StringReplace(stAuthEdit, ' and ', '} and {',
                  [rfReplaceAll, rfIgnoreCase]);
              end;
            end;
            stItem := stItem + ',' + LineEnding + '  ' +
              LowerCase(sqToolsTables.Fields[i].FieldName) + ' = ' +
              stOpenBrakets + stAuthEdit + stCloseBrackets;
          end
          else if LowerCase(sqToolsTables.Fields[i].FieldName) = 'owned' then
          begin
            Continue;
          end
          else if LowerCase(sqToolsTables.Fields[i].FieldName) = 'abstract' then
          begin
            if fmOptions.cbNotExpAbsRew.Checked = True then
            begin
              Continue;
            end
            else if fmOptions.cbExpAbsRew.Checked = True then
            begin
              stItem := stItem + ',' + LineEnding + '  ' + 'review = ' +
                stOpenBrakets + sqToolsTables.Fields[i].AsString + stCloseBrackets;
            end
            else
            begin
              stItem := stItem + ',' + LineEnding + '  ' + 'abstract = ' +
                stOpenBrakets + sqToolsTables.Fields[i].AsString + stCloseBrackets;
            end;
          end
          else if LowerCase(sqToolsTables.Fields[i].FieldName) = 'review' then
          begin
            if fmOptions.cbNotExpAbsRew.Checked = True then
            begin
              Continue;
            end
            else
            begin
              stItem := stItem + ',' + LineEnding + '  ' + 'review = ' +
                stOpenBrakets + sqToolsTables.Fields[i].AsString + stCloseBrackets;
            end;
          end
          else if LowerCase(sqToolsTables.Fields[i].FieldName) = 'title' then
          begin
            stItem := stItem + ',' + LineEnding + '  ' +
              LowerCase(sqToolsTables.Fields[i].FieldName) + ' = ' +
              stOpenBrakets + StringReplace(sqToolsTables.Fields[i].AsString,
              '*', '', [rfReplaceAll]) + stCloseBrackets;
          end
          else if LowerCase(sqToolsTables.Fields[i].FieldName) <> 'attname' then
          begin
            stItem := stItem + ',' + LineEnding + '  ' +
              LowerCase(sqToolsTables.Fields[i].FieldName) + ' = ' +
              stOpenBrakets + sqToolsTables.Fields[i].AsString + stCloseBrackets;
          end;
        end;
      end;
      if Filtered = True then
      begin
        if DirectoryExistsUTF8(ExtractFileNameWithoutExt(sqItems.FileName)) =
          True then
          if DirectoryExistsUTF8(ExtractFileNameWithoutExt(stFileName)) =
            False then
            try
              CreateDirUTF8(ExtractFileNameWithoutExt(stFileName));
            except
              MessageDlg('It is not possibile to create the folder attachments.',
                mtError, [mbOK], 0);
              Exit;
            end;
        if sqToolsTables.FieldByName('AttName').AsString <> '' then
          try
            myUnZip := TUnZipper.Create;
            myUnZip.OutputPath := ExtractFileNameWithoutExt(stFileName);
            slAtt := TStringList.Create;
            slAtt.Text := sqToolsTables.FieldByName('AttName').AsString;
            stItem := stItem + ',' + LineEnding + '  file = ' + stOpenBrakets;
            for n := 0 to slAtt.Count - 1 do
            begin
              stItem := stItem + slAtt[n] + ':' +
                ExtractFileNameOnly(stFileName) + DirectorySeparator +
                sqToolsTables.FieldByName('IDItems').AsString + '-' +
                slAtt[n];
              if n < slAtt.Count - 1 then
                stItem := stItem + ';'
              else
                stItem := stItem + stCloseBrackets;
              if FileExistsUTF8(ExtractFileNameWithoutExt(sqItems.FileName) +
                DirectorySeparator + sqToolsTables.FieldByName(
                'IDItems').AsString + '-' + slAtt[n] + '.zip') = True then
              begin
                myUnZip.UnZipAllFiles(ExtractFileNameWithoutExt(sqItems.FileName) +
                  DirectorySeparator +
                  sqToolsTables.FieldByName('IDItems').AsString +
                  '-' + slAtt[n] + '.zip');
                // Rename the original file name adding the ID
                RenameFileUTF8(ExtractFileNameWithoutExt(stFileName) +
                  DirectorySeparator + slAtt[n],
                  ExtractFileNameWithoutExt(stFileName) +
                  DirectorySeparator +
                  sqToolsTables.FieldByName('IDItems').AsString +
                  '-' + slAtt[n]);
              end;
            end;
          finally
            myUnZip.Free;
            slAtt.Free;
          end;
      end;
      stItem := StringReplace(stItem, '&', '\&', [rfReplaceAll]);
      stItem := StringReplace(stItem, '"', #39#39, [rfReplaceAll]);
      // This is not the beginning of a LaTex command
      stItem := StringReplace(stItem, '\ ', '/ ', [rfReplaceAll]);
      stItem := stItem + LineEnding + '}' + LineEnding;
      WriteLn(FileBibLatex, stItem);
      sqToolsTables.Next;
      Inc(NumExp);
      sbStatusBar.SimpleText := 'Items exported: ' + IntToStr(NumExp) + '.';
      Application.ProcessMessages;
    end;
  finally
    sqToolsTables.Close;
    CloseFile(FileBibLatex);
    Screen.Cursor := crDefault;
  end;
end;

procedure TfmMain.ImportFromBibfilex(stFileName: string);
var
  i, n, ImpNum: integer;
  AttReadDir, AttWriteDir: string;
  sqInUse, sqExternal: TSqlite3Dataset;
  slAttachList: TStringList;
  blNoDir: boolean = False;
begin
  ;
  // Import from Bibfilex
  Screen.Cursor := crHourGlass;
  Application.ProcessMessages;
  sqInUse := TSqlite3Dataset.Create(Self);
  sqInUse.FileName := sqItems.FileName;
  sqInUse.TableName := 'Items';
  sqInUse.PrimaryKey := 'IDItems';
  sqInUse.AutoIncrementKey := True;
  sqInUse.Open;
  sqExternal := TSqlite3Dataset.Create(Self);
  sqExternal.FileName := stFileName;
  sqExternal.TableName := 'Items';
  sqExternal.PrimaryKey := 'IDItems';
  sqExternal.AutoIncrementKey := True;
  sqExternal.Open;
  slAttachList := TStringList.Create;
  ImpNum := 0;
  try
    try
      while not sqExternal.EOF do
      begin
        sqInUse.Append;
        for i := 1 to sqExternal.FieldCount - 1 do
        begin
          if sqExternal.Fields[i].AsString <> '' then
          begin
            if UTF8LowerCase(sqExternal.Fields[i].FieldName) = 'bibtexkey' then
            begin
              sqInUse.FieldByName('bibtexkey').AsString :=
                CheckUniqueBibTexKey(sqExternal.FieldByName('bibtexkey').AsString,
                sqInUse.FieldByName('IDItems').AsInteger);
            end
            else
              // Do not use field number: the position of fields might be different
              sqInUse.FieldByName(sqExternal.Fields[i].FieldName).Value :=
                sqExternal.FieldByName(sqExternal.Fields[i].FieldName).Value;
          end;
        end;
        // Copy possible attachments
        if sqExternal.FieldByName('AttName').AsString <> '' then
        begin
          AttReadDir := ExtractFileNameWithoutExt(sqExternal.FileName);
          if DirectoryExistsUTF8(AttReadDir) = False then
          begin
            if blNoDir = False then
            begin
              MessageDlg('The attachments directory of the ' +
                'file to be imported is not available: ' +
                'attachments will not be imported.',
                mtWarning, [mbOK], 0);
              blNoDir := True;
            end;
            sqInUse.FieldByName('AttName').AsString := '';
          end;
          AttWriteDir := ExtractFileNameWithoutExt(sqInUse.FileName);
          if DirectoryExistsUTF8(AttWriteDir) = False then
            try
              CreateDir(AttWriteDir);
            except
              MessageDlg('It''s not possible to create the attachments ' +
                'directory of the file in use.',
                mtWarning, [mbOK], 0);
              Screen.Cursor := crDefault;
              Abort;
            end;
          slAttachList.Clear;
          slAttachList.Text := sqExternal.FieldByName('AttName').AsString;
          for n := 0 to slAttachList.Count - 1 do
          begin
            if FileExistsUTF8(AttReadDir + DirectorySeparator +
              sqExternal.FieldByName('IDItems').AsString + '-' +
              slAttachList[n] + '.zip') then
            begin
              CopyFile(AttReadDir + DirectorySeparator +
                sqExternal.FieldByName('IDItems').AsString +
                '-' + slAttachList[n] + '.zip',
                AttWriteDir + DirectorySeparator +
                sqInUse.FieldByName('IDItems').AsString + '-' +
                slAttachList[n] + '.zip');
            end;
          end;
        end;
        sqInUse.Post;
        sqExternal.Next;
        Inc(ImpNum);
        sbStatusBar.SimpleText := 'Imported items: ' + IntToStr(ImpNum);
        Application.ProcessMessages;
      end;
      sqInUse.ApplyUpdates;
      if IsDirectoryEmpty(AttWriteDir) = True then
      begin
        RemoveDirUTF8(AttWriteDir);
      end;
      sqItems.Close;
      sqItems.Open;
    except
      MessageDlg('Error in importing data from Bibfilex file.',
        mtWarning, [mbOK], 0);
    end;
  finally
    slAttachList.Free;
    sqInUse.Free;
    sqExternal.Free;
    Screen.Cursor := crDefault;
  end;
end;

procedure TfmMain.ExportToBibfilex(stFileName: string);
var
  i, n, ExpNum: integer;
  AttReadDir, AttWriteDir: string;
  sqInUse, sqExternal: TSqlite3Dataset;
  slAttachList: TStringList;
  blNoDir: boolean = False;
begin
  // Export to Bibfilex
  Screen.Cursor := crHourGlass;
  Application.ProcessMessages;
  sqInUse := TSqlite3Dataset.Create(Self);
  sqInUse.FileName := sqItems.FileName;
  sqInUse.TableName := 'Items';
  sqInUse.PrimaryKey := 'IDItems';
  sqInUse.AutoIncrementKey := True;
  sqInUse.SQL := sqItems.SQL;
  sqInUse.Open;
  sqExternal := TSqlite3Dataset.Create(Self);
  sqExternal.FileName := stFileName;
  sqExternal.TableName := 'Items';
  sqExternal.PrimaryKey := 'IDItems';
  sqExternal.AutoIncrementKey := True;
  sqExternal.Open;
  slAttachList := TStringList.Create;
  ExpNum := 0;
  try
    try
      while not sqInUse.EOF do
      begin
        sqExternal.Append;
        for i := 1 to sqInUse.FieldCount - 1 do
        begin
          if sqInUse.Fields[i].AsString <> '' then
          begin
            ;
            if UTF8LowerCase(sqInUse.Fields[i].FieldName) = 'bibtexkey' then
            begin
              sqExternal.FieldByName('bibtexkey').AsString :=
                CheckUniqueBibTexKey(sqInUse.FieldByName('bibtexkey').AsString,
                sqExternal.FieldByName('IDItems').AsInteger);
            end
            else
            begin
              // Do not use field number: the position of fields might be different
              sqExternal.FieldByName(sqInUse.Fields[i].FieldName).Value :=
                sqInUse.FieldByName(sqInUse.Fields[i].FieldName).Value;
            end;
          end;
        end;
        // Copy possible attachments
        if sqInUse.FieldByName('AttName').AsString <> '' then
        begin
          AttReadDir := ExtractFileNameWithoutExt(sqInUse.FileName);
          if DirectoryExistsUTF8(AttReadDir) = False then
          begin
            if blNoDir = False then
            begin
              MessageDlg('The attachments directory of the ' +
                'file in use is not available: ' +
                'the attachments will not be exported.',
                mtWarning, [mbOK], 0);
              blNoDir := True;
            end;
            sqExternal.FieldByName('AttName').AsString := '';
          end;
          AttWriteDir := ExtractFileNameWithoutExt(sqExternal.FileName);
          if DirectoryExistsUTF8(AttWriteDir) = False then
            try
              CreateDir(AttWriteDir);
            except
              MessageDlg('It''s not possible to create the attachments ' +
                'directory of the file of destination.',
                mtWarning, [mbOK], 0);
              Screen.Cursor := crDefault;
              Abort;
            end;
          slAttachList.Clear;
          slAttachList.Text := sqInUse.FieldByName('AttName').AsString;
          for n := 0 to slAttachList.Count - 1 do
          begin
            if FileExistsUTF8(AttReadDir + DirectorySeparator +
              sqInUse.FieldByName('IDItems').AsString + '-' +
              slAttachList[n] + '.zip') then
            begin
              CopyFile(AttReadDir + DirectorySeparator +
                sqInUse.FieldByName('IDItems').AsString + '-' +
                slAttachList[n] + '.zip',
                AttWriteDir + DirectorySeparator +
                sqExternal.FieldByName('IDItems').AsString +
                '-' + slAttachList[n] + '.zip');
            end;
          end;
        end;
        sqExternal.Post;
        sqInUse.Next;
        Inc(ExpNum);
        sbStatusBar.SimpleText := 'Exported items: ' + IntToStr(ExpNum);
        Application.ProcessMessages;
      end;
      sqExternal.ApplyUpdates;
      if IsDirectoryEmpty(AttWriteDir) = True then
      begin
        RemoveDirUTF8(AttWriteDir);
      end;
    except
      MessageDlg('Error in exporting data to Bibfilex file.',
        mtWarning, [mbOK], 0);
    end;
  finally
    sqInUse.Free;
    sqExternal.Free;
    slAttachList.Free;
    Screen.Cursor := crDefault;
  end;
end;

procedure TfmMain.SetStatusBar;
var
  stStBar: string;
begin
  // Set the values in the status bar
  if sqItems.RecordCount = 0 then
    stStBar := 'No items'
  else
    stStBar := 'Item n.' + ' ' + IntToStr(sqItems.RecNo) + ' ' +
      'of' + ' ' + IntToStr(sqItems.RecordCount);
  if flFilterActive = 1 then
    stStBar := stStBar + ' (smart filter active)'
  else if flFilterActive = 2 then
    stStBar := stStBar + ' (standard filter active)'
  else if flFilterActive = 3 then
    stStBar := stStBar + ' (filter from keywords active)'
  else if flFilterActive = 4 then
    stStBar := stStBar + ' (filter from LaTex document active)';
  sbStatusBar.SimpleText := stStBar + '.';
end;

function TfmMain.CreateOrderBy: string;
begin
  // Create order by statement in sqItems SQL
  Result := ' order by IDItems ';
  if miItemsOrderByID.Checked = True then
    Result := ' order by IDItems '
  else if miItemsOrderByAuthor.Checked = True then
    Result := ' order by author, title '
  else if miItemsOrderByTitle.Checked = True then
    Result := ' order by title, author '
  else if miItemsOrderByJournalTitle.Checked = True then
    Result := ' order by journaltitle, author '
  else if miItemsOrderByBookTitle.Checked = True then
    Result := ' order by booktitle, author '
  else if miItemsOrderByYear.Checked = True then
    Result := ' order by year, author, title '
  else if miItemsOrderByDate.Checked = True then
    Result := ' order by timestamp, author, title ';
end;

procedure TfmMain.SetItemsGridColumns;
var
  i, n: integer;
begin
  // Set the columns in the items grid
  with grItems do
  begin
    Columns.Clear;
    n := 0;
    for i := 0 to fmOptions.clFieldsShown.Items.Count - 1 do
      if fmOptions.clFieldsShown.Checked[i] = True then
      begin
        Columns.Add;
        Columns[n].FieldName := fmOptions.clFieldsShown.Items[i];
        Columns[n].Width := 150;
        Columns[n].Alignment := taLeftJustify;
        Columns[n].Title.Alignment := taLeftJustify;
        Inc(n);
      end;
  end;
end;

procedure TfmMain.SetOptions;
begin
  // Set options according to fmOptions
  SetItemsGridColumns;
  if fmOptions.bxFontSize.ItemIndex > 0 then
  begin
    fmMain.Font.Size := fmOptions.bxFontSize.ItemIndex + 9;
    // No parent font
    lbAuthorTitle.Font.Size := fmMain.Font.Size;
    meAbstract.Font.Size := fmMain.Font.Size + 5;
    meReview.Font.Size := fmMain.Font.Size + 5;
    fmFilters.Font.Size := fmOptions.bxFontSize.ItemIndex + 9;
    fmOptions.Font.Size := fmOptions.bxFontSize.ItemIndex + 9;
    fmKeywords.Font.Size := fmOptions.bxFontSize.ItemIndex + 9;
    fmChar.Font.Size := fmOptions.bxFontSize.ItemIndex + 9;
    fmStopList.Font.Size := fmOptions.bxFontSize.ItemIndex + 9;
  end
  else
  begin
    fmMain.Font.Size := 0;
    // No parent font
    lbAuthorTitle.Font.Size := fmMain.Font.Size;
    meAbstract.Font.Size := 14;
    meReview.Font.Size := 14;
    fmFilters.Font.Size := 0;
    fmOptions.Font.Size := 0;
    fmKeywords.Font.Size := 0;
    fmChar.Font.Size := 0;
    fmStopList.Font.Size := 0;
  end;
  if fmOptions.rgFormatDate.ItemIndex = 0 then
    ShortDateFormat := 'mm' + DateSeparator + 'dd' + DateSeparator + 'yyyy'
  else if fmOptions.rgFormatDate.ItemIndex = 1 then
    ShortDateFormat := 'dd' + DateSeparator + 'mm' + DateSeparator + 'yyyy'
  else if fmOptions.rgFormatDate.ItemIndex = 2 then
    ShortDateFormat := 'yyyy' + DateSeparator + 'mm' + DateSeparator + 'dd'
  else if fmOptions.rgFormatDate.ItemIndex = 3 then
    ShortDateFormat := 'yyyy' + DateSeparator + 'dd' + DateSeparator + 'mm';
end;

function TfmMain.CreateBibTexKey(IsItemsDataSet: boolean; SaveBefore: boolean): string;
var
  stModel, stField, stValue, stSize, stStopWord, stAuth: string;
  iSize, i, n: integer;
  slAuthList: TStringList;
begin
  // Create the BibText key
  // Data must be saved to let the function know the current bibtex key.
  // If it's empty, it does not matter. False value is useful
  // to avoid loop with SaveItems procedure
  Result := '';
  if SaveBefore = True then
    SaveItems;
  stModel := fmOptions.edBibTexKey.Text;
  while UTF8Pos('|', stModel) > 0 do
  begin
    stAuth := '';
    stField := UTF8Copy(stModel, UTF8Pos('|', stModel) + 1,
      UTF8Pos('[', stModel) - UTF8Pos('|', stModel) - 2);
    if IsItemsDataSet = True then
    begin
      if ((LowerCase(stField) = 'author') and
        (sqItems.FieldByName('author').AsString = '')) then
      begin
        stValue := sqItems.FieldByName('editor').AsString;
      end
      else
      begin
        stValue := sqItems.FieldByName(stField).AsString;
      end;
    end
    else
    begin
      if ((LowerCase(stField) = 'author') and
        (sqToolsTables.FieldByName('author').AsString = '')) then
      begin
        stValue := sqToolsTables.FieldByName('editor').AsString;
      end
      else
      begin
        stValue := sqToolsTables.FieldByName(stField).AsString;
      end;
    end;
    // Remove signs
    if ((LowerCase(stField) <> 'author') and (LowerCase(stField) <> 'editor')) then
      stValue := StringReplace(stValue, ',', '', [rfReplaceAll]);
    stValue := StringReplace(stValue, '´', '', [rfReplaceAll]);
    stValue := StringReplace(stValue, '*', '', [rfReplaceAll]);
    stValue := StringReplace(stValue, '"', '', [rfReplaceAll]);
    stValue := StringReplace(stValue, '“', '', [rfReplaceAll]);
    stValue := StringReplace(stValue, '”', '', [rfReplaceAll]);
    stValue := StringReplace(stValue, '«', '', [rfReplaceAll]);
    stValue := StringReplace(stValue, '»', '', [rfReplaceAll]);
    stValue := StringReplace(stValue, '!', '', [rfReplaceAll]);
    stValue := StringReplace(stValue, '?', '', [rfReplaceAll]);
    stValue := StringReplace(stValue, '¿', '', [rfReplaceAll]);
    stValue := StringReplace(stValue, '&', '', [rfReplaceAll]);
    stValue := StringReplace(stValue, '@', '', [rfReplaceAll]);
    stValue := StringReplace(stValue, ':', '', [rfReplaceAll]);
    stValue := StringReplace(stValue, ';', '', [rfReplaceAll]);
    stValue := StringReplace(stValue, '.', '', [rfReplaceAll]);
    stValue := StringReplace(stValue, '-', '', [rfReplaceAll]);
    stValue := StringReplace(stValue, '–', '', [rfReplaceAll]);
    stValue := StringReplace(stValue, '\emph{', '', [rfReplaceAll]);
    stValue := StringReplace(stValue, '\textit{', '', [rfReplaceAll]);
    stValue := StringReplace(stValue, '\textsc{', '', [rfReplaceAll]);
    stValue := StringReplace(stValue, '\textbf{', '', [rfReplaceAll]);
    stValue := StringReplace(stValue, '\', '', [rfReplaceAll]);
    stValue := StringReplace(stValue, '{', '', [rfReplaceAll]);
    stValue := StringReplace(stValue, '}', '', [rfReplaceAll]);
    stValue := StringReplace(stValue, '[', '', [rfReplaceAll]);
    stValue := StringReplace(stValue, ']', '', [rfReplaceAll]);
    // Size
    stSize := UTF8Copy(stModel, UTF8Pos('[', stModel) + 1, 2);
    iSize := StrToInt(stSize);
    // Delete words in stop list
    if ((LowerCase(stField) <> 'author') and (LowerCase(stField) <> 'editor')) then
    begin
      for i := 0 to fmStopList.meStopList.Lines.Count - 1 do
      begin
        // If there is a ' after the word do no add a space
        if UTF8Pos('''', fmStopList.meStopList.Lines[i]) < 1 then
          stStopWord := fmStopList.meStopList.Lines[i] + ' '
        else
          stStopWord := fmStopList.meStopList.Lines[i];
        if UTF8LowerCase(stStopWord) = UTF8LowerCase(
          UTF8Copy(stValue, 1, UTF8Length(stStopWord))) then
        begin
          stValue := UTF8Copy(stValue, UTF8Length(stStopWord) + 1,
            UTF8Length(stValue));
          // So that the first proper letter in the word may be capitalized...
          stValue := UTF8UpperCase(UTF8Copy(stValue, 1, 1)) +
            UTF8Copy(stValue, 2, UTF8Length(stValue));
          Break;
        end;
      end;
    end;
    if iSize = 0 then
    begin
      if ((LowerCase(stField) = 'author') or (LowerCase(stField) = 'editor')) then
        try
          slAuthList := TStringList.Create;
          slAuthList.Text := StringReplace(stValue, ' and ', LineEnding,
            [rfReplaceAll]);
          for n := 0 to slAuthList.Count - 1 do
          begin
            if UTF8LowerCase(slAuthList[n]) = 'others' then
              Break;
            // Remove the first space if after one or two letters
            // for authors like S. Francis
            if UTF8Pos(' ', slAuthList[n]) < 4 then
              slAuthList[n] := StringReplace(slAuthList[n], ' ', '', []);
            // Author or editor: all the family name
            if UTF8Pos(',', slAuthList[n]) > 0 then
              iSize := UTF8Pos(',', slAuthList[n]) - 1
            else if UTF8Pos(' ', slAuthList[n]) > 0 then
              iSize := UTF8Pos(' ', slAuthList[n]) - 1
            else
              iSize := UTF8Length(slAuthList[n]);
            stAuth := stAuth + UTF8Copy(slAuthList[n], 1, iSize);
            if fmOptions.cbAllAuthors.Checked = False then
              Break;
          end;
        finally
          slAuthList.Free;
        end
      else
      begin
        if UTF8Pos(' ', stValue) > 0 then
          iSize := UTF8Pos(' ', stValue) - 1
        else
          iSize := UTF8Length(stValue);
      end;
    end;
    // No author
    if stAuth = '' then
    begin
      stModel := StringReplace(stModel, '|' + stField + '|[' +
        stSize + ']', UTF8Copy(stValue, 1, iSize), [rfIgnoreCase]);
    end
    else
      // Author
    begin
      stModel := StringReplace(stModel, '|' + stField + '|[' +
        stSize + ']', stAuth, [rfIgnoreCase]);
    end;
  end;
  // Replace unvalid characters
  // The ' is here and not above because it's useful to cut article
  // at the beginning of a title
  stModel := StringReplace(stModel, '''', '', [rfReplaceAll]);
  stModel := StringReplace(stModel, ' ', '', [rfReplaceAll]);
  stModel := StringReplace(stModel, 'À', 'A', [rfReplaceAll]);
  stModel := StringReplace(stModel, 'Á', 'A', [rfReplaceAll]);
  stModel := StringReplace(stModel, 'Â', 'A', [rfReplaceAll]);
  stModel := StringReplace(stModel, 'Ä', 'A', [rfReplaceAll]);
  stModel := StringReplace(stModel, 'Ä', 'A', [rfReplaceAll]);
  stModel := StringReplace(stModel, 'Å', 'A', [rfReplaceAll]);
  stModel := StringReplace(stModel, 'Æ', 'AE', [rfReplaceAll]);
  stModel := StringReplace(stModel, 'È', 'E', [rfReplaceAll]);
  stModel := StringReplace(stModel, 'É', 'E', [rfReplaceAll]);
  stModel := StringReplace(stModel, 'Ê', 'E', [rfReplaceAll]);
  stModel := StringReplace(stModel, 'Ë', 'E', [rfReplaceAll]);
  stModel := StringReplace(stModel, 'Ì', 'I', [rfReplaceAll]);
  stModel := StringReplace(stModel, 'Í', 'I', [rfReplaceAll]);
  stModel := StringReplace(stModel, 'Î', 'I', [rfReplaceAll]);
  stModel := StringReplace(stModel, 'Ï', 'I', [rfReplaceAll]);
  stModel := StringReplace(stModel, 'Ò', 'O', [rfReplaceAll]);
  stModel := StringReplace(stModel, 'Ó', 'O', [rfReplaceAll]);
  stModel := StringReplace(stModel, 'Ô', 'O', [rfReplaceAll]);
  stModel := StringReplace(stModel, 'Õ', 'O', [rfReplaceAll]);
  stModel := StringReplace(stModel, 'Ö', 'O', [rfReplaceAll]);
  stModel := StringReplace(stModel, 'Ù', 'U', [rfReplaceAll]);
  stModel := StringReplace(stModel, 'Ú', 'U', [rfReplaceAll]);
  stModel := StringReplace(stModel, 'Û', 'U', [rfReplaceAll]);
  stModel := StringReplace(stModel, 'Ü', 'U', [rfReplaceAll]);
  stModel := StringReplace(stModel, 'à', 'a', [rfReplaceAll]);
  stModel := StringReplace(stModel, 'á', 'a', [rfReplaceAll]);
  stModel := StringReplace(stModel, 'â', 'a', [rfReplaceAll]);
  stModel := StringReplace(stModel, 'ã', 'a', [rfReplaceAll]);
  stModel := StringReplace(stModel, 'ä', 'a', [rfReplaceAll]);
  stModel := StringReplace(stModel, 'å', 'a', [rfReplaceAll]);
  stModel := StringReplace(stModel, 'æ', 'ae', [rfReplaceAll]);
  stModel := StringReplace(stModel, 'è', 'e', [rfReplaceAll]);
  stModel := StringReplace(stModel, 'é', 'e', [rfReplaceAll]);
  stModel := StringReplace(stModel, 'ê', 'e', [rfReplaceAll]);
  stModel := StringReplace(stModel, 'ë', 'e', [rfReplaceAll]);
  stModel := StringReplace(stModel, 'ì', 'i', [rfReplaceAll]);
  stModel := StringReplace(stModel, 'í', 'i', [rfReplaceAll]);
  stModel := StringReplace(stModel, 'î', 'i', [rfReplaceAll]);
  stModel := StringReplace(stModel, 'ï', 'i', [rfReplaceAll]);
  stModel := StringReplace(stModel, 'ò', 'o', [rfReplaceAll]);
  stModel := StringReplace(stModel, 'ó', 'o', [rfReplaceAll]);
  stModel := StringReplace(stModel, 'ô', 'o', [rfReplaceAll]);
  stModel := StringReplace(stModel, 'õ', 'o', [rfReplaceAll]);
  stModel := StringReplace(stModel, 'ö', 'o', [rfReplaceAll]);
  stModel := StringReplace(stModel, 'ù', 'u', [rfReplaceAll]);
  stModel := StringReplace(stModel, 'ú', 'u', [rfReplaceAll]);
  stModel := StringReplace(stModel, 'û', 'u', [rfReplaceAll]);
  stModel := StringReplace(stModel, 'ü', 'u', [rfReplaceAll]);
  stModel := StringReplace(stModel, 'ß', 'ss', [rfReplaceAll]);
  stModel := StringReplace(stModel, 'Ç', 'C', [rfReplaceAll]);
  stModel := StringReplace(stModel, 'Ð', 'Edh', [rfReplaceAll]);
  stModel := StringReplace(stModel, 'Ñ', 'N', [rfReplaceAll]);
  stModel := StringReplace(stModel, 'Ø', 'Oe', [rfReplaceAll]);
  stModel := StringReplace(stModel, 'Ý', 'Y', [rfReplaceAll]);
  stModel := StringReplace(stModel, 'Þ', 'th', [rfReplaceAll]);
  stModel := StringReplace(stModel, 'ç', 'c', [rfReplaceAll]);
  stModel := StringReplace(stModel, 'ð', 'edh', [rfReplaceAll]);
  stModel := StringReplace(stModel, 'ñ', 'n', [rfReplaceAll]);
  stModel := StringReplace(stModel, 'ø', 'oe', [rfReplaceAll]);
  stModel := StringReplace(stModel, 'ÿ', 'y', [rfReplaceAll]);
  stModel := StringReplace(stModel, 'þ', 'th', [rfReplaceAll]);
  // Set eventually lowercase
  if fmOptions.cbBibTexLowercase.Checked = True then
    stModel := UTF8LowerCase(stModel);
  //Check the value is unique
  if IsItemsDataSet = True then
  begin
    Result := CheckUniqueBibTexKey(stModel, sqItems.FieldByName(
      'IDItems').AsInteger);
  end
  else
  begin
    Result := CheckUniqueBibTexKey(stModel, sqToolsTables.FieldByName(
      'IDItems').AsInteger);
  end;
end;

function TfmMain.CheckUniqueBibTexKey(BibTexKey: string; IDItems: integer): string;
var
  iLetter: integer;
  sqKey: TSqlite3Dataset;
begin
  // Check that the key is unique or add letters to it
  sqKey := TSqlite3Dataset.Create(Self);
  sqKey.FileName := sqItems.FileName;
  sqKey.TableName := 'Items';
  sqKey.PrimaryKey := 'IDItems';
  sqKey.SQL := 'Select IDItems, bibtexkey from Items ' +
    'where Lower(bibtexkey) = "' + UTF8LowerCase(BibTexKey) +
    '" and IDItems <> ' + IntToStr(IDItems) + ' collate nocase';
  try
    sqKey.Open;
    // Char before 'a'
    iLetter := 96;
    while sqKey.RecordCount > 0 do
    begin
      // Select all letters up to 'z'
      if iLetter < 122 then
        Inc(iLetter)
      // Also the BibTexKey + 'z' key exists: restart from 'a'
      else
      begin
        BibTexKey := BibTexKey + '-z';
        iLetter := 97;
      end;
      sqKey.Close;
      sqKey.SQL := 'Select IDItems, bibtexkey from Items ' +
        'where Lower(bibtexkey) = "' + UTF8LowerCase(BibTexKey) +
        '-' + char(iLetter) + '" and IDItems <> ' + IntToStr(IDItems) +
        ' collate nocase';
      sqKey.Open;
    end;
    if iLetter > 96 then
      Result := BibTexKey + '-' + char(iLetter)
    else
      Result := BibTexKey;
  finally
    sqKey.Free;
  end;
end;

procedure TfmMain.SetAttachmentMenuItems;
begin
  // Set attachment menu items
  if lbAttNames.Items.Count > 0 then
  begin
    miAttachOpen.Enabled := True;
    pmAttOpen.Enabled := True;
    miAttachSaveAs.Enabled := True;
    pmAttSaveAs.Enabled := True;
    miAttachDelete.Enabled := True;
    pmAttDelete.Enabled := True;
  end
  else
  begin
    miAttachOpen.Enabled := False;
    pmAttOpen.Enabled := False;
    miAttachSaveAs.Enabled := False;
    pmAttSaveAs.Enabled := False;
    miAttachDelete.Enabled := False;
    pmAttDelete.Enabled := False;
  end;
end;

procedure TfmMain.ActivateSuggestion(grGrid: TStringGrid);
var
  stKeyword: string;
  iLenText: integer;
begin
  // Text suggestion
  if (grGrid.Editor is TStringCellEditor) then
  begin
    if LowerCase(grGrid.Cells[0, grGrid.Row]) = 'bibtexkey' then
    begin
      Exit;
    end;
    iLenText := UTF8Length(grGrid.Cells[1, grGrid.Row]);
    // No keywords
    if LowerCase(grGrid.Cells[0, grGrid.Row]) <> 'keywords' then
    begin
      grGrid.Cells[1, grGrid.Row] :=
        SuggestValue(grGrid.Cells[1, grGrid.Row], grGrid.Cells[0, grGrid.Row]);
    end
    // Keywords
    else
    begin
      // One keyword
      if UTF8Pos(',', grGrid.Cells[1, grGrid.Row]) = 0 then
        grGrid.Cells[1, grGrid.Row] :=
          SuggestValue(grGrid.Cells[1, grGrid.Row], grGrid.Cells[0, grGrid.Row])
      // More keywords
      else
      begin
        stKeyword := grGrid.Cells[1, grGrid.Row];
        while UTF8Pos(',', stKeyword) > 0 do
          stKeyword := UTF8Copy(stKeyword, UTF8Pos(',', stKeyword) +
            1, UTF8Length(stKeyword));
        while UTF8Copy(stKeyword, 1, 1) = ' ' do
          stKeyword := UTF8Copy(stKeyword, 2, UTF8Length(stKeyword));
        grGrid.Cells[1, grGrid.Row] :=
          UTF8Copy(grGrid.Cells[1, grGrid.Row], 1,
          UTF8Length(grGrid.Cells[1, grGrid.Row]) - UTF8Length(stKeyword)) +
          SuggestValue(stKeyword, grGrid.Cells[0, grGrid.Row]);
      end;
    end;
    TStringCellEditor(grGrid.Editor).SelStart := iLenText;
    TStringCellEditor(grGrid.Editor).SelLength :=
      UTF8Length(grGrid.Cells[1, grGrid.Row]);
  end;
end;

function TfmMain.SuggestValue(StartText: string; Field: string): string;
var
  sqSuggValue: TSqlite3Dataset;
  i: integer;
begin
  // Get the first value in the field beginning with a defined text
  Result := StartText;
  if StartText = '' then
    Exit;
  // Crossref field autocompleted with bibtex keys
  if LowerCase(Field) = 'crossref' then
    try
      sqSuggValue := TSqlite3Dataset.Create(Self);
      sqSuggValue.FileName := sqItems.FileName;
      sqSuggValue.PrimaryKey := sqItems.PrimaryKey;
      sqSuggValue.SQL := 'Select IDItems, bibtexkey from Items ' +
        'where Lower(bibtexkey) like "' + UTF8LowerCase(StartText) +
        '%" order by bibtexkey';
      sqSuggValue.Open;
      if sqSuggValue.RecordCount > 0 then
        Result := sqSuggValue.FieldByName('bibtexkey').AsString;
    finally
      sqSuggValue.Free;
    end
  // No keyword field
  else if LowerCase(Field) <> 'keywords' then
    try
      sqSuggValue := TSqlite3Dataset.Create(Self);
      sqSuggValue.FileName := sqItems.FileName;
      sqSuggValue.PrimaryKey := sqItems.PrimaryKey;
      sqSuggValue.SQL := 'Select IDItems, ' + Field + ' from Items ' +
        'where Lower(' + Field + ') like "' + UTF8LowerCase(StartText) +
        '%" order by ' + Field;
      sqSuggValue.Open;
      if sqSuggValue.RecordCount > 0 then
        Result := sqSuggValue.FieldByName(Field).AsString;
    finally
      sqSuggValue.Free;
    end
  // Keyword field
  else
  begin
    for i := 0 to fmKeywords.lbKeyword.Items.Count - 1 do
      if UTF8LowerCase(UTF8Copy(fmKeywords.lbKeyword.Items[i], 1,
        UTF8Length(StartText))) = UTF8LowerCase(StartText) then
      begin
        Result := fmKeywords.lbKeyword.Items[i] + ', ';
        Break;
      end;
  end;
end;

procedure TfmMain.CreateKeywordList;
var
  TagsList: TStringList;
  sqKeywords: TSqlite3Dataset;
  i: integer;
begin
  // Create keyword list
  Screen.Cursor := crHourGlass;
  Application.ProcessMessages;
  sqKeywords := TSqlite3Dataset.Create(Self);
  sqKeywords.FileName := sqItems.FileName;
  sqKeywords.PrimaryKey := sqItems.PrimaryKey;
  sqKeywords.SQL :=
    'Select IDItems, keywords from Items where keywords not null';
  sqKeywords.Open;
  try
    TagsList := TStringList.Create;
    TagsList.Capacity := 5000;
    fmKeywords.lbKeyword.Items.Clear;
    fmKeywords.lbKeyword.Sorted := False;
    while not sqKeywords.EOF do
    begin
      TagsList.Text := StringReplace(sqKeywords.FieldByName('keywords').AsString,
        ', ', LineEnding, [rfReplaceAll]);
      for i := 0 to TagsList.Count - 1 do
      begin
        if fmKeywords.lbKeyword.Items.IndexOf(TagsList[i]) < 0 then
          fmKeywords.lbKeyword.Items.Add(TagsList[i]);
      end;
      sqKeywords.Next;
    end;
    fmKeywords.lbKeyword.Sorted := True;
  finally
    TagsList.Free;
    sqKeywords.Free;
    Screen.Cursor := crDefault;
  end;
end;

function TfmMain.CleanKeywordField(myField: string): string;
var
  myText: string;
  i: integer;
  TagsList, TagsCurr: TStringList;
begin
  // Clean the keyword field from wrong inputs
  if myField <> '' then
  begin
    myText := myField;
    myText := StringReplace(myText, ',', ', ', [rfReplaceAll]);
    myText := StringReplace(myText, ' ,', ',', [rfReplaceAll]);
    while UTF8Pos('  ', myText) > 0 do
      myText := StringReplace(myText, '  ', ' ', []);
    if UTF8Copy(myText, UTF8Length(myText), 1) = ',' then
      myText := UTF8Copy(myText, 1, UTF8Length(myText) - 1);
    if UTF8Copy(myText, UTF8Length(myText), 1) = ' ' then
      myText := UTF8Copy(myText, 1, UTF8Length(myText) - 1);
    // To be repeted
    if UTF8Copy(myText, UTF8Length(myText), 1) = ',' then
      myText := UTF8Copy(myText, 1, UTF8Length(myText) - 1);
    if UTF8Copy(myText, 1, 1) = ' ' then
      myText := UTF8Copy(myText, 2, UTF8Length(myText));
    // If something wrong happens later...
    Result := myText;
    // Remove duplicates
    try
      TagsList := TStringList.Create;
      TagsList.Capacity := 500;
      TagsCurr := TStringList.Create;
      TagsCurr.Capacity := 500;
      TagsCurr.Text := StringReplace(myText, ', ', LineEnding, [rfReplaceAll]);
      for i := 0 to TagsCurr.Count - 1 do
      begin
        if TagsList.IndexOf(TagsCurr[i]) < 0 then
          TagsList.Add(TagsCurr[i]);
      end;
      myText := TagsList.Text;
    finally
      TagsList.Free;
      TagsCurr.Free;
    end;
    Result := StringReplace(myText, LineEnding, ', ', [rfReplaceAll]);
    Result := UTF8Copy(Result, 1, UTF8Length(Result) - 2);
  end
  else
    Result := '';
end;

procedure TfmMain.SmartFilter;
var
  stDelimiter: string;
begin
  // Smart filter
  if edSmartFilter.Text = '' then
  begin
    MessageDlg('No text has been specified.',
      mtWarning, [mbOK], 0);
    Abort;
  end;
  if ((UTF8Pos(#39, edSmartFilter.Text) > 0) and
    (UTF8Pos('"', edSmartFilter.Text) > 0)) then
  begin
    MessageDlg('The characters " and ' + #39 + ' cannot be searched together.',
      mtWarning, [mbOK], 0);
    Abort;
  end;
  if UTF8Pos('"', edSmartFilter.Text) > 0 then
    stDelimiter := #39
  else
    stDelimiter := '"';
  if edSmartFilter.Text <> '' then
  begin
    Screen.Cursor := crHourGlass;
    Application.ProcessMessages;
    flFilterActive := 1;
    sqItems.Close;
    if LowerCase(cbFieldSmartFilter.Text) = 'iditems' then
    begin
      if UTF8Pos('-', edSmartFilter.Text) > 0 then
      begin
        sqItems.SQL := 'Select * from Items where ' +
          UTF8LowerCase(cbFieldSmartFilter.Text) + ' >= ' +
          UTF8Copy(edSmartFilter.Text, 1, UTF8Pos('-', edSmartFilter.Text) -
          1) + ' and ' + UTF8LowerCase(cbFieldSmartFilter.Text) +
          ' <= ' + UTF8Copy(edSmartFilter.Text, UTF8Pos('-', edSmartFilter.Text) +
          1, UTF8Length(edSmartFilter.Text)) + CreateOrderBy;
      end
      else
      if UTF8Copy(edSmartFilter.Text, 1, 1) = '>' then
      begin
        sqItems.SQL := 'Select * from Items where ' +
          UTF8LowerCase(cbFieldSmartFilter.Text) + ' >= ' +
          UTF8Copy(edSmartFilter.Text, 2, UTF8Length(edSmartFilter.Text)) +
          CreateOrderBy;
      end
      else
      if UTF8Copy(edSmartFilter.Text, 1, 1) = '<' then
      begin
        sqItems.SQL := 'Select * from Items where ' +
          UTF8LowerCase(cbFieldSmartFilter.Text) + ' <= ' +
          UTF8Copy(edSmartFilter.Text, 2, UTF8Length(edSmartFilter.Text)) +
          CreateOrderBy;
      end
      else
      begin
        sqItems.SQL := 'Select * from Items where ' +
          UTF8LowerCase(cbFieldSmartFilter.Text) + ' = ' + edSmartFilter.Text +
          ' ' + CreateOrderBy;
      end;
    end
    else if LowerCase(cbFieldSmartFilter.Text) = 'author' then
    begin
      sqItems.SQL := 'Select * from Items where Lower(' +
        cbFieldSmartFilter.Text + ') like ' + stDelimiter + '%' +
        UTF8LowerCase(edSmartFilter.Text) + '%' + stDelimiter + ' ' +
        CreateOrderBy;
    end
    else
      sqItems.SQL := 'Select * from Items where Lower(' +
        cbFieldSmartFilter.Text + ') like ' + stDelimiter +
        UTF8LowerCase(edSmartFilter.Text) + '%' + stDelimiter + ' ' +
        CreateOrderBy;
    try
      try
        sqItems.Open;
      except
        Screen.Cursor := crDefault;
        MessageDlg('The filter is not correct; all the items will be selected.',
          mtWarning, [mbOK], 0);
        Screen.Cursor := crHourGlass;
        Application.ProcessMessages;
        flFilterActive := 0;
        sqItems.Close;
        sqItems.SQL := 'Select * from Items ' + CreateOrderBy;
        sqItems.Open;
      end;
    finally
      Screen.Cursor := crDefault;
    end;
  end;
end;

procedure TfmMain.FilterOnKeyword;
var
  KeywordList: TStringList;
  i: integer;
begin
  // Filter items on keyword
  if cbFilterKey.Text = '' then
  begin
    MessageDlg('No keyword has been specified.', mtWarning, [mbOK], 0);
    Abort;
  end;
  Screen.Cursor := crHourGlass;
  Application.ProcessMessages;
  flFilterActive := 3;
  sqItems.Close;
  try
    KeywordList := TStringList.Create;
    KeywordList.Text := StringReplace(cbFilterKey.Text, ', ',
      LineEnding, [rfReplaceAll]);
    sqItems.SQL := 'Select * from Items where ';
    for i := 0 to KeywordList.Count - 1 do
    begin
      sqItems.SQL := sqItems.SQL + ' ((Lower(keywords) = "' +
        UTF8LowerCase(KeywordList[i]) + '") ' +
        // Beginning tag
        'or (Lower(keywords) like "' + UTF8LowerCase(KeywordList[i]) + ', %") ' +
        // Ending tag
        'or (Lower(keywords) like "%, ' + UTF8LowerCase(KeywordList[i]) + '") ' +
        // Middle tag
        'or (Lower(keywords) like "%, ' + UTF8LowerCase(KeywordList[i]) + ', %")) ';
      if i < KeywordList.Count - 1 then
        sqItems.SQL := sqItems.SQL + ' or ';
    end;
    sqItems.SQL := sqItems.SQL + CreateOrderBy;
  finally
    KeywordList.Free;
  end;
  try
    try
      sqItems.Open;
    except
      Screen.Cursor := crDefault;
      MessageDlg('The filter is not correct; all the items will be selected.',
        mtWarning, [mbOK], 0);
      Screen.Cursor := crHourGlass;
      Application.ProcessMessages;
      flFilterActive := 0;
      sqItems.Close;
      sqItems.SQL := 'Select * from Items ' + CreateOrderBy;
      sqItems.Open;
    end;
  finally
    Screen.Cursor := crDefault;
  end;
end;

procedure TfmMain.RenameKeyword;
var
  OldValue, NewValue: string;
  sqRenKeywords: TSqlite3Dataset;
  IDItems, NumChanged: integer;
  ResultInput: boolean;
begin
  // Rename a keyword in the file in use
  SaveItems;
  ResultInput := InputQuery('Old keyword', 'Insert the old keyword to be replaced.',
    OldValue);
  if ResultInput = False then
    Exit;
  ResultInput := InputQuery('New keyword',
    'Insert the new keyword or leave empty to delete the old one.', NewValue);
  if ResultInput = False then
    Exit;
  try
    Screen.Cursor := crHourGlass;
    Application.ProcessMessages;
    NumChanged := 0;
    sqRenKeywords := TSqlite3Dataset.Create(Self);
    sqRenKeywords.FileName := sqItems.FileName;
    sqRenKeywords.TableName := 'Items';
    sqRenKeywords.PrimaryKey := 'IDItems';
    sqRenKeywords.SQL := 'Select IDItems, keywords from Items ' +
      'where Lower(keywords) like "%' + UTF8LowerCase(OldValue) + '%"';
    sqRenKeywords.Open;
    while not sqRenKeywords.EOF do
    begin
      if ((UTF8LowerCase(sqRenKeywords.FieldByName('keywords').AsString) =
        UTF8LowerCase(OldValue)) or
        (UTF8Pos(UTF8LowerCase(OldValue) + ', ',
        UTF8LowerCase(sqRenKeywords.FieldByName('keywords').AsString)) > 0) or
        (UTF8Pos(', ' + UTF8LowerCase(OldValue),
        UTF8LowerCase(sqRenKeywords.FieldByName('keywords').AsString)) > 0)) then
      begin
        sqRenKeywords.Edit;
        // Replace
        if NewValue <> '' then
          sqRenKeywords.FieldByName('keywords').AsString :=
            StringReplace(sqRenKeywords.FieldByName('keywords').AsString,
            OldValue, NewValue, [rfIgnoreCase, rfReplaceAll])
        // Delete
        else
        begin
          sqRenKeywords.FieldByName('keywords').AsString :=
            StringReplace(sqRenKeywords.FieldByName('keywords').AsString,
            OldValue + ', ', '', [rfIgnoreCase, rfReplaceAll]);
          sqRenKeywords.FieldByName('keywords').AsString :=
            StringReplace(sqRenKeywords.FieldByName('keywords').AsString,
            ', ' + OldValue, '', [rfIgnoreCase, rfReplaceAll]);
          sqRenKeywords.FieldByName('keywords').AsString :=
            StringReplace(sqRenKeywords.FieldByName('keywords').AsString,
            OldValue, '', [rfIgnoreCase, rfReplaceAll]);
        end;
        sqRenKeywords.Post;
        sqRenKeywords.ApplyUpdates;
        Inc(NumChanged);
        sbStatusBar.SimpleText :=
          'Updated items: ' + IntToStr(NumChanged) + '.';
      end;
      Application.ProcessMessages;
      sqRenKeywords.Next;
    end;
    if sqItems.RecordCount > 0 then
    begin
      IDItems := sqItems.FieldByName('IDItems').AsInteger;
      sqItems.Close;
      sqItems.Open;
      sqItems.Locate('IDItems', IDItems, []);
    end;
    CreateKeywordList;
  finally
    Screen.Cursor := crDefault;
    sqRenKeywords.Free;
  end;
end;

procedure TfmMain.FilterFromLatex(FileName: string);
var
  slCiteList, slCiteCmd: TStringList;
  meText: TMemo;
  stCite, stText: string;
  i, n: integer;
begin
  // Create filter from the \cite commands in a LaTex file
  try
    Screen.Cursor := crHourGlass;
    Application.ProcessMessages;
    meText := TMemo.Create(Self);
    slCiteList := TStringList.Create;
    slCiteCmd := TStringList.Create;
    meText.Lines.LoadFromFile(FileName);
    slCiteCmd.Add('\cite');
    slCiteCmd.Add('\textcite');
    slCiteCmd.Add('\parencite');
    slCiteCmd.Add('\smartcite');
    for n := 0 to slCiteCmd.Count - 1 do
    begin
      stText := meText.Text;
      while UTF8Pos(slCiteCmd[n], stText) > 0 do
      begin
        stCite := '';
        i := UTF8Pos(slCiteCmd[n], stText);
        while UTF8Copy(stText, i, 1) <> '{' do
        begin
          Inc(i);
        end;
        Inc(i);
        while UTF8Copy(stText, i, 1) <> '}' do
        begin
          stCite := stCite + UTF8Copy(stText, i, 1);
          Inc(i);
        end;
        stCite := UTF8Copy(stCite, 1, UTF8Length(stCite));
        slCiteList.Add(stCite);
        stText := UTF8Copy(stText, i, UTF8Length(stText));
      end;
    end;
    if slCiteList.Count = 0 then
    begin
      MessageDlg('No \cite or similar command was found ' +
        'in the selected file.', mtWarning, [mbOK], 0);
      Exit;
    end
    else
    begin
      flFilterActive := 4;
      sqItems.Close;
      sqItems.SQL := 'Select * from Items where Lower(bibtexkey) in (';
      for i := 0 to slCiteList.Count - 1 do
        sqItems.SQL := sqItems.SQL + '"' + UTF8LowerCase(slCiteList[i]) + '", ';
      sqItems.SQL := Copy(sqItems.SQL, 1, UTF8Length(sqItems.SQL) - 2);
      sqItems.SQL := sqItems.SQL + ')' + CreateOrderBy;
      try
        sqItems.Open;
        Screen.Cursor := crDefault;
      except
        MessageDlg('The filter is not correct; all the items ' +
          'will be selected.',
          mtWarning, [mbOK], 0);
        Screen.Cursor := crHourGlass;
        Application.ProcessMessages;
        flFilterActive := 0;
        sqItems.Close;
        sqItems.SQL := 'Select * from Items ' + CreateOrderBy;
        sqItems.Open;
      end;
    end;
  finally
    meText.Free;
    slCiteList.Free;
    slCiteCmd.Free;
    Screen.Cursor := crDefault;
  end;
end;

function TfmMain.ReplaceCiteInLatex(FileName: string): boolean;
var
  sqGetCitations: TSqlite3Dataset;
  slCiteList, slUniqueCiteList, slCiteCmd, slBiblioList: TStringList;
  meText: TMemo;
  stCite, stPreNote, stPostNote, stTitPages, stText, stPrevAuthor: string;
  i, n, y: integer;
  flFirstCitation, flAuthorAsIdem, flIbidem, flNoKeys: boolean;
begin
  // Replace \cite in Latex file with citations
  Result := True;
  try
    Screen.Cursor := crHourGlass;
    Application.ProcessMessages;
    meText := TMemo.Create(Self);
    slCiteList := TStringList.Create;
    slUniqueCiteList := TStringList.Create;
    slCiteCmd := TStringList.Create;
    slBiblioList := TStringList.Create;
    sqGetCitations := TSqlite3Dataset.Create(Self);
    meText.Lines.LoadFromFile(FileName);
    slCiteCmd.Add('\cite');
    slCiteCmd.Add('\textcite');
    slCiteCmd.Add('\parencite');
    slCiteCmd.Add('\smartcite');
    for n := 0 to slCiteCmd.Count - 1 do
    begin
      stText := meText.Text;
      while UTF8Pos(slCiteCmd[n], stText) > 0 do
      begin
        stCite := '';
        i := UTF8Pos(slCiteCmd[n], stText);
        while UTF8Copy(stText, i, 1) <> '{' do
        begin
          Inc(i);
        end;
        Inc(i);
        while UTF8Copy(stText, i, 1) <> '}' do
        begin
          stCite := stCite + UTF8Copy(stText, i, 1);
          Inc(i);
        end;
        // If not footnotes before i, we are in the same footnote, etc.
        // or the citation is like the previous
        if ((UTF8Pos(fmOptions.rgIdemCmd.Items[fmOptions.rgIdemCmd.ItemIndex],
          stText) > i) or (UTF8Pos(
          fmOptions.rgIdemCmd.Items[fmOptions.rgIdemCmd.ItemIndex], stText) < 1)) then
        begin
          // The same citation as before
          if ((slCiteList.Count > 0) and
            (stCite = UTF8Copy(slCiteList[slCiteList.Count - 1], 2,
            UTF8Length(slCiteList[slCiteList.Count - 1])))) then
          begin
            slCiteList.Add('%' + stCite);
          end
          else
          begin
            // In the same footnote, etc. of the previous key
            slCiteList.Add('&' + stCite);
          end;
        end
        else
        begin
          // In another footnote, etc
          slCiteList.Add('|' + stCite);
        end;
        stText := UTF8Copy(stText, i + 1, UTF8Length(stText));
      end;
    end;
    if slCiteList.Count = 0 then
    begin
      MessageDlg('No \cite or similar command was found ' +
        'in the selected LaTex file; ' +
        'anyway, it will be converted as requested.', mtWarning, [mbOK], 0);
      CopyFile(FileName, ExtractFilePath(FileName) +
        ExtractFileNameOnly(FileName) + '-fullcitations.tex', True);
      Exit;
    end
    else
      try
        flNoKeys := False;
        meText.Lines.LoadFromFile(FileName);
        sqGetCitations.FileName := sqItems.FileName;
        sqGetCitations.TableName := 'Items';
        sqGetCitations.PrimaryKey := 'IDItems';
        for i := 0 to slCiteList.Count - 1 do
        begin
          flIbidem := False;
          flAuthorAsIdem := False;
          if fmOptions.cbCitAsIbidem.Checked = True then
          begin
            if UTF8Copy(slCiteList[i], 1, 1) = '%' then
            begin
              flIbidem := True;
            end;
          end;
          if fmOptions.cbAuthorAsIdem.Checked = True then
          begin
            if UTF8Copy(slCiteList[i], 1, 1) = '&' then
            begin
              flAuthorAsIdem := True;
            end;
          end;
          slCiteList[i] := UTF8Copy(slCiteList[i], 2, UTF8Length(slCiteList[i]));
          if slUniqueCiteList.IndexOf(slCiteList[i]) < 0 then
          begin
            flFirstCitation := True;
            slUniqueCiteList.Add(slCiteList[i]);
          end
          else
          begin
            flFirstCitation := False;
          end;
          sqGetCitations.Close;
          sqGetCitations.SQL :=
            'Select * from Items ' + 'where Lower(bibtexkey) = "' +
            UTF8LowerCase(slCiteList[i]) + '"';
          sqGetCitations.Open;
          if sqGetCitations.RecordCount = 0 then
          begin
            flNoKeys := True;
          end
          else
          begin
            stCite := '';
            n := UTF8Pos('{' + slCiteList[i] + '}', meText.Text);
            while UTF8Copy(meText.Text, n, 1) <> '\' do
            begin
              n := n - 1;
              stCite := UTF8Copy(meText.Text, n, 1) + stCite;
              if n < 0 then
                Exit;
            end;
            stCite := stCite + '{' + slCiteList[i] + '}';
            stPreNote := '';
            stPostNote := '';
            if UTF8Pos('][', stCite) > 0 then
            begin
              stPreNote := UTF8Copy(stCite, UTF8Pos('[', stCite) +
                1, UTF8Pos('][', stCite) - UTF8Pos('[', stCite) - 1);
              stPostNote := UTF8Copy(stCite, UTF8Pos('][', stCite) +
                2, UTF8Pos(']{', stCite) - UTF8Pos('][', stCite) - 2);
            end
            else if UTF8Pos('[', stCite) > 0 then
            begin
              stPostNote := UTF8Copy(stCite, UTF8Pos('[', stCite) +
                1, UTF8Pos(']', stCite) - UTF8Pos('[', stCite) - 1);
            end;
            if stPreNote <> '' then
              stPreNote := stPreNote + ' ';
            stTitPages := '';
            if UTF8Pos('|', stPostNote) > 0 then
            begin
              stTitPages := ', ' + UTF8Copy(stPostNote, 1,
                UTF8Pos('|', stPostNote) - 1);
              stPostNote := UTF8Copy(stPostNote,
                UTF8Pos('|', stPostNote) + 1, UTF8Length(stPostNote));
            end;
            if stPostNote <> '' then
            begin
              if UTF8Pos('-', stPostNote) > 0 then
              begin
                if fmOptions.edPostnoteMorePages.Text <> '' then
                  stPostNote := fmOptions.edPostnoteMorePages.Text + ' ' + stPostNote;
              end
              else
              begin
                if fmOptions.edPostnoteOnePage.Text <> '' then
                  stPostNote := fmOptions.edPostnoteOnePage.Text + ' ' + stPostNote;
              end;
              stPostNote := ', ' + stPostNote;
            end;
            if flIbidem = True then
            begin
              meText.Text :=
                StringReplace(meText.Text, stCite, stPreNote +
                fmOptions.edCitAsIbidem.Text + stPostNote, [rfIgnoreCase]);
            end
            else if flFirstCitation = True then
            begin
              if stPostNote <> '' then
                meText.Text :=
                  StringReplace(meText.Text, stCite, stPreNote +
                  CreateCitation(slCiteList[i], True, flAuthorAsIdem,
                  False, True, 0, False, stTitPages) + stPostNote, [rfIgnoreCase])
              else
                meText.Text :=
                  StringReplace(meText.Text, stCite, stPreNote +
                  CreateCitation(slCiteList[i], True, flAuthorAsIdem,
                  False, False, 0, False, stTitPages) + stPostNote, [rfIgnoreCase]);
            end
            else
            begin
              if stPostNote <> '' then
                meText.Text :=
                  StringReplace(meText.Text, stCite, stPreNote +
                  CreateCitation(slCiteList[i], True, flAuthorAsIdem,
                  False, True, 1, False, stTitPages) + stPostNote, [rfIgnoreCase])
              else
                meText.Text :=
                  StringReplace(meText.Text, stCite, stPreNote +
                  CreateCitation(slCiteList[i], True, flAuthorAsIdem,
                  False, False, 1, False, stTitPages) + stPostNote, [rfIgnoreCase]);
            end;
            // Update bibliography
            if flFirstCitation = True then
            begin
              if fmOptions.cbAuthorAsIdem.Checked = True then
              begin
                slBiblioList.Add(CreateCitation(slCiteList[i], False,
                  False, False, False, 2, True, '') + '.' + LineEnding);
              end
              else
              begin
                slBiblioList.Add(CreateCitation(slCiteList[i], False,
                  False, False, False, 2, False, '') + '.' + LineEnding);
              end;
            end;
          end;
        end;
        slBiblioList.Sort;
        if fmOptions.cbAuthorAsIdem.Checked = True then
        begin
          stPrevAuthor := '';
          for y := 0 to slBiblioList.Count - 1 do
          begin
            if UTF8Copy(slBiblioList[y], UTF8Pos(#4, slBiblioList[y]) +
              1, UTF8Pos(#5, slBiblioList[y]) - UTF8Pos(#4, slBiblioList[y]) - 1) =
              stPrevAuthor then
            begin
              slBiblioList[y] :=
                UTF8Copy(slBiblioList[y], 1, UTF8Pos(#4, slBiblioList[y]) - 1) +
                fmOptions.edIdemBiblio.Text +
                UTF8Copy(slBiblioList[y], UTF8Pos(#5, slBiblioList[y]) + 1,
                UTF8Length(slBiblioList[y]));
            end
            else
            begin
              stPrevAuthor :=
                UTF8Copy(slBiblioList[y], UTF8Pos(#4, slBiblioList[y]) + 1,
                UTF8Pos(#5, slBiblioList[y]) - UTF8Pos(#4, slBiblioList[y]) - 1);
              slBiblioList[y] := StringReplace(slBiblioList[y], #4, '', []);
              slBiblioList[y] := StringReplace(slBiblioList[y], #5, '', []);
            end;
          end;
        end;
        meText.Text := StringReplace(meText.Text, '\printbibliography',
          slBiblioList.Text, [rfReplaceAll, rfIgnoreCase]);
        meText.Lines.SaveToFile(ExtractFilePath(FileName) +
          ExtractFileNameOnly(FileName) + '-fullcitations.tex');
        if flNoKeys = True then
        begin
          MessageDlg('Some BibTex keys were not found in the ' +
            'selected Bibfilex file. Look for unresolved ' +
            'citation commands in the converted file.',
            mtWarning, [mbOK], 0);
        end;
      except
        Result := False;
      end;
  finally
    meText.Free;
    slCiteList.Free;
    slUniqueCiteList.Free;
    slBiblioList.Free;
    sqGetCitations.Free;
    Screen.Cursor := crDefault;
  end;
end;

procedure TfmMain.CreateAttachment(FileNames: array of string);
var
  myZipper: TZipper;
  AttDir: string;
  i: integer;
begin
  try
    AttDir := ExtractFileNameWithoutExt(sqItems.FileName);
    if DirectoryExistsUTF8(AttDir) = False then
      CreateDirUTF8(AttDir);
  except
    MessageDlg('It was not possible to create the attachment directory.',
      mtError, [mbOK], 0);
    Abort;
  end;
  for i := 0 to Length(FileNames) - 1 do
  begin
    // The file is a directory
    if FileGetAttrUTF8(FileNames[i]) = 48 then
      Continue
    else if lbAttNames.Items.IndexOf(ExtractFileName(FileNames[i])) > -1 then
    begin
      MessageDlg('A file with the name ' + ExtractFileName(FileNames[i]) +
        ' is already attached.', mtWarning, [mbOK], 0);
      Continue;
    end
    else
      try
        try
          Screen.Cursor := crHourGlass;
          Application.ProcessMessages;
          myZipper := TZipper.Create;
          myZipper.FileName :=
            AttDir + DirectorySeparator + sqItems.FieldByName('IDItems').AsString +
            '-' + ExtractFileName(FileNames[i]) + '.zip';
          myZipper.Entries.AddFileEntry(FileNames[i],
            ExtractFileName(FileNames[i]));
          myZipper.ZipAllFiles;
          lbAttNames.Items.Add(ExtractFileName(FileNames[i]));
          sqItems.Edit;
          sqItems.FieldByName('AttName').AsString := lbAttNames.Items.Text;
          sqItems.Post;
          sqItems.ApplyUpdates;
        except
          MessageDlg('It is not possible to load the file.', mtError, [mbOK], 0);
        end;
      finally
        myZipper.Free;
        Screen.Cursor := crDefault;
      end;
  end;
end;

function TfmMain.CreateCitation(BibTexKey: string; AuthorNameBefore: boolean;
  AuthorAsIdem: boolean; IsHTML: boolean; SkipPages: boolean;
  CitationKind: smallint; MarkAroundAuthor: boolean; TitPages: string): string;
var
  sqCitation: TSqlite3Dataset;
  iType, iField: integer;
  stField, stContent, stType, stCellPattern: string;
  flAddEditorNote: boolean = False;
begin
  // Create citation
  // CitationKind: 0 = normal; 1 = short; 2 = biblio
  Result := '';
  try
    sqCitation := TSqlite3Dataset.Create(Self);
    sqCitation.FileName := sqItems.FileName;
    sqCitation.TableName := 'Items';
    sqCitation.PrimaryKey := 'IDItems';
    sqCitation.SQL := 'Select * from Items ' + 'where Lower(bibtexkey) = "' +
      UTF8LowerCase(BibTexKey) + '" ' + CreateOrderBy;
    sqCitation.Open;
    // Get the type
    stType := LowerCase(sqCitation.FieldByName('entrytype').AsString);
    for iType := 0 to fmOptions.sgCitation.RowCount - 1 do
    begin
      if LowerCase(fmOptions.sgCitation.Cells[0, iType]) = stType + ' normal' then
        Break;
    end;
    // Step to the short or biblio citation pattern.
    iType := iType + CitationKind;
    for iField := 1 to fmOptions.sgCitation.ColCount - 1 do
    begin
      // No field
      if UTF8Pos('|', fmOptions.sgCitation.Cells[iField, iType]) < 1 then
      begin
        Result := Result + fmOptions.sgCitation.Cells[iField, iType];
      end
      // Field
      else
      begin
        stfield := fmOptions.sgCitation.Cells[iField, iType];
        stField := UTF8Copy(stField, UTF8Pos('|', stField) + 1, UTF8Length(stField));
        stField := UTF8Copy(stField, 1, UTF8Pos('|', stField) - 1);
        if sqCitation.FieldDefs.IndexOf(stField) > 0 then
        begin
          stContent := '';
          // Skip pages
          if ((LowerCase(stField) = 'pages') and (SkipPages = True)) then
          begin
            Continue;
          end
          // Entry types in which the editor might take the place of the bookauthor
          else if ((stType = 'inbook') or (stType = 'bookinbook') or
            (stType = 'suppbook') or (stType = 'suppperiodical')) then
          begin
            // Skip editor if used in place of author
            if ((LowerCase(stField) = 'editor') and
              (sqCitation.FieldByName('bookauthor').AsString = '') and
              (fmOptions.cbEditorAsAuthor.Checked = True)) then
            begin
              Continue;
            end
            // No bookauthor, use editor insted; "ed." added below
            else if ((LowerCase(stField) = 'bookauthor') and
              (sqCitation.FieldByName('bookauthor').AsString = '') and
              (sqCitation.FieldByName('editor').AsString <> '') and
              (fmOptions.cbEditorAsAuthor.Checked = True)) then
            begin
              if ((CitationKind = 1) and
                (fmOptions.cbShortenAuthor.Checked = True)) then
              begin
                stContent := FormatAuthorInCitation(
                  sqCitation.FieldByName('editor').AsString, False, True);
                flAddEditorNote := True;
              end
              else
              begin
                stContent := FormatAuthorInCitation(
                  sqCitation.FieldByName('editor').AsString,
                  AuthorNameBefore, False);
                flAddEditorNote := True;
              end;
              CitAutCurrent :=
                FormatAuthorInCitation(sqCitation.FieldByName('editor').AsString,
                True, False);
              if CitAutCurrent = CitAutPrevious then
              begin
                if AuthorAsIdem = True then
                begin
                  stContent := fmOptions.edIdemNotes.Text;
                end;
              end
              else
              begin
                CitAutPrevious := CitAutCurrent;
              end;
            end
            else if ((LowerCase(stField) = 'bookauthor') and
              (sqCitation.FieldByName('bookauthor').AsString = '') and
              (sqCitation.FieldByName('editor').AsString = '') and
              (fmOptions.cbEditorAsAuthor.Checked = True)) then
            begin
              Continue;
            end
            else if ((LowerCase(stField) = 'bookauthor') and
              (sqCitation.FieldByName('bookauthor').AsString = '') and
              (fmOptions.cbEditorAsAuthor.Checked = False)) then
            begin
              Continue;
            end;
          end
          // Entry types in which the editor might take the place of the author
          else
          begin
            // Skip editor if used in place of author
            if ((LowerCase(stField) = 'editor') and
              (sqCitation.FieldByName('author').AsString = '') and
              (fmOptions.cbEditorAsAuthor.Checked = True)) then
            begin
              Continue;
            end
            // No author, use editor insted; "ed." added below
            else if ((LowerCase(stField) = 'author') and
              (sqCitation.FieldByName('author').AsString = '') and
              (sqCitation.FieldByName('editor').AsString <> '') and
              (fmOptions.cbEditorAsAuthor.Checked = True)) then
            begin
              if ((CitationKind = 1) and
                (fmOptions.cbShortenAuthor.Checked = True)) then
              begin
                stContent := FormatAuthorInCitation(
                  sqCitation.FieldByName('editor').AsString, False, True);
                flAddEditorNote := True;
              end
              else
              begin
                stContent := FormatAuthorInCitation(
                  sqCitation.FieldByName('editor').AsString,
                  AuthorNameBefore, False);
                flAddEditorNote := True;
              end;
              CitAutCurrent :=
                FormatAuthorInCitation(sqCitation.FieldByName('editor').AsString,
                True, False);
              if CitAutCurrent = CitAutPrevious then
              begin
                if AuthorAsIdem = True then
                begin
                  stContent := fmOptions.edIdemNotes.Text;
                end;
              end
              else
              begin
                CitAutPrevious := CitAutCurrent;
              end;
            end
            else if ((LowerCase(stField) = 'author') and
              (sqCitation.FieldByName('author').AsString = '') and
              (fmOptions.cbEditorAsAuthor.Checked = False)) then
            begin
              Continue;
            end;
          end;
          // No author or bookuthor: the field must not be empty
          if ((LowerCase(stField) <> 'author') and
            (LowerCase(stField) <> 'bookauthor') and
            (sqCitation.FieldByName(stField).AsString = '')) then
          begin
            Continue;
          end
          else if ((LowerCase(stField) = 'author') and
            (sqCitation.FieldByName('author').AsString <> '')) then
          begin
            if ((CitationKind = 1) and
              (fmOptions.cbShortenAuthor.Checked = True)) then
            begin
              stContent := FormatAuthorInCitation(
                sqCitation.FieldByName(stField).AsString, False, True);
            end
            else
            begin
              stContent := FormatAuthorInCitation(
                sqCitation.FieldByName(stField).AsString,
                AuthorNameBefore, False);
            end;
            CitAutCurrent := FormatAuthorInCitation(
              sqCitation.FieldByName(stField).AsString, True, False);
            if CitAutCurrent = CitAutPrevious then
            begin
              if AuthorAsIdem = True then
              begin
                stContent := fmOptions.edIdemNotes.Text;
              end;
            end
            else
            begin
              CitAutPrevious := CitAutCurrent;
            end;
          end
          else if ((LowerCase(stField) = 'bookauthor') and
            (sqCitation.FieldByName('bookauthor').AsString <> '')) then
          begin
            if ((fmOptions.cbAuthorAsIdem.Checked = True) and
              (sqCitation.FieldByName('bookauthor').AsString =
              sqCitation.FieldByName('author').AsString)) then
            begin
              stContent := fmOptions.edIdemNotes.Text;
            end
            else
            begin
              stContent := FormatAuthorInCitation(
                sqCitation.FieldByName(stField).AsString,
                AuthorNameBefore, False);
            end;
          end
          else if ((LowerCase(stField) = 'editor') and
            (sqCitation.FieldByName('editor').AsString <> '')) then
          begin
            // This editor is not in place of author or bookauthor
            stContent := FormatAuthorInCitation(
              sqCitation.FieldByName(stField).AsString, True, False);
          end
          else if ((LowerCase(stField) = 'title') and
            (sqCitation.FieldByName('title').AsString <> '') and
            (CitationKind = 1) and (fmOptions.cbShortenTitle.Checked =
            True)) then
          begin
            stContent := ShortenTitle(
              sqCitation.FieldByName(stField).AsString);
          end
          else if ((LowerCase(stField) = 'title') and
            (sqCitation.FieldByName('title').AsString <> '')) then
          begin
            stContent := StringReplace(sqCitation.FieldByName(stField).AsString,
              '*', '', [rfReplaceAll]);
          end
          else if ((LowerCase(stField) <> 'author') and
            (LowerCase(stField) <> 'bookauthor') and
            (sqCitation.FieldByName(stField).AsString <> '')) then
          begin
            stContent := sqCitation.FieldByName(stField).AsString;
          end;
          stContent := StringReplace(stContent, '&', '\&', [rfReplaceAll]);
          stContent := StringReplace(stContent, '"', '“', [rfReplaceAll]);
          stCellPattern := fmOptions.sgCitation.Cells[iField, iType];
          // Add ed. mention
          if flAddEditorNote = True then
          begin
            if UTF8Pos('|bookauthor|}', stCellPattern) > 0 then
              stCellPattern :=
                StringReplace(stCellPattern, '|bookauthor|}',
                '|bookauthor|} ' + fmOptions.edEditorMention.Text, [rfIgnoreCase])
            else if UTF8Pos('|bookauthor|', stCellPattern) > 0 then
              stCellPattern :=
                StringReplace(stCellPattern, '|bookauthor|', '|bookauthor| ' +
                fmOptions.edEditorMention.Text, [rfIgnoreCase])
            else if UTF8Pos('|author|}', stCellPattern) > 0 then
              stCellPattern :=
                StringReplace(stCellPattern, '|author|}', '|author|} ' +
                fmOptions.edEditorMention.Text, [rfIgnoreCase])
            else if UTF8Pos('|author|', stCellPattern) > 0 then
              stCellPattern :=
                StringReplace(stCellPattern, '|author|', '|author| ' +
                fmOptions.edEditorMention.Text, [rfIgnoreCase]);
          end;
          // Add the text before pages field
          if UTF8Pos('|pages|', stCellPattern) > 0 then
          begin
            if sqCitation.FieldByName('pages').AsString <> '' then
            begin
              if UTF8Pos('-', sqCitation.FieldByName('pages').AsString) > 0 then
              begin
                if fmOptions.edPostnoteMorePages.Text <> '' then
                  stCellPattern :=
                    StringReplace(stCellPattern, '|pages|',
                    fmOptions.edPostnoteMorePages.Text + ' ' +
                    '|pages|', [rfIgnoreCase]);
              end
              else
              begin
                if fmOptions.edPostnoteOnePage.Text <> '' then
                  stCellPattern :=
                    StringReplace(stCellPattern, '|pages|',
                    fmOptions.edPostnoteOnePage.Text + ' ' +
                    '|pages|', [rfIgnoreCase]);
              end;
            end;
          end;
          if ((MarkAroundAuthor = True) and (UTF8LowerCase(stField) =
            'author')) then
          begin
            Result := Result + StringReplace(stCellPattern, '|' +
              stField + '|', #4 + stContent + #5, [rfReplaceAll, rfIgnoreCase]);
          end
          else
          begin
            Result := Result + StringReplace(stCellPattern, '|' +
              stField + '|', stContent, [rfReplaceAll, rfIgnoreCase]);
          end;
          if ((LowerCase(stField) = 'title') and
            (sqCitation.FieldByName('title').AsString <> '')) then
          begin
            Result := Result + TitPages;
          end;
        end;
      end;
    end;
    // Set the '. ', etc to '\. ', etc to avoid double space in LaTex
    if fmOptions.cbEscDot.Checked = True then
    begin
      Result := StringReplace(Result, '. ', '.\ ', [rfReplaceAll]);
      Result := StringReplace(Result, '! ', '!\ ', [rfReplaceAll]);
      Result := StringReplace(Result, '? ', '?\ ', [rfReplaceAll]);
    end;
    if IsHTML = True then
    begin
      Result := StringReplace(Result, '\textsc{',
        '<span style="font-style:small-caps">', [rfReplaceAll, rfIgnoreCase]);
      Result := StringReplace(Result, '\textit{',
        '<span style="font-style:italic">', [rfReplaceAll, rfIgnoreCase]);
      Result := StringReplace(Result, '\emph{',
        '<span style="font-style:italic">', [rfReplaceAll, rfIgnoreCase]);
      Result := StringReplace(Result, '\textbf{',
        '<span style="font-weight:bold">', [rfReplaceAll, rfIgnoreCase]);
      Result := StringReplace(Result, '}',
        '<span style="font-style:normal;font-weight:normal">',
        [rfReplaceAll, rfIgnoreCase]);
    end;
  finally
    sqCitation.Free;
  end;
end;

function TfmMain.FormatAuthorInCitation(stAuthor: string; NameBefore: boolean;
  OnlyFamilyName: boolean): string;
var
  slAuthorList: TStringList;
  i, n: integer;
  stOrigName, stDefName: string;
  blAndOthers: boolean;
begin
  // Format the author or editor for citation
  if stAuthor = '' then
  begin
    Result := '';
    Exit;
  end;
  try
    blAndOthers := False;
    if UTF8LowerCase(UTF8Copy(stAuthor, UTF8Length(stAuthor) -
      UTF8Length('and others') + 1, UTF8Length('and others'))) = 'and others' then
    begin
      blAndOthers := True;
      stAuthor := UTF8Copy(stAuthor, 1, UTF8Length(stAuthor) -
        UTF8Length('and others') - 1);
    end;
    slAuthorList := TStringList.Create;
    slAuthorList.Text := StringReplace(stAuthor, ' and ', LineEnding,
      [rfReplaceAll, rfIgnoreCase]);
    for i := 0 to slAuthorList.Count - 1 do
    begin
      if UTF8Pos(',', slAuthorList[i]) > 0 then
      begin
        // Only family name
        if OnlyFamilyName = True then
        begin
          slAuthorList[i] :=
            UTF8Copy(slAuthorList[i], 1, UTF8Pos(', ', slAuthorList[i]) - 1);
        end
        // Name with dot
        else if fmOptions.cbDotAftertFirstName.Checked = True then
        begin
          stOrigName := ' ' + UTF8Copy(slAuthorList[i],
            UTF8Pos(', ', slAuthorList[i]) + 2, UTF8Length(slAuthorList[i]));
          stDefName := '';
          for n := 0 to UTF8Length(stOrigName) - 2 do
          begin
            if stOrigName[n] = ' ' then
              stDefName := stDefName + stOrigName[n + 1] + '. ';
          end;
        end
        // Name without dot
        else
        begin
          stDefName := UTF8Copy(slAuthorList[i], UTF8Pos(', ', slAuthorList[i]) +
            2, UTF8Length(slAuthorList[i]));
        end;
        if OnlyFamilyName = False then
        begin
          // Name before
          if NameBefore = True then
          begin
            if fmOptions.cbDotAftertFirstName.Checked = False then
              stDefName := stDefName + ' ';
            slAuthorList[i] :=
              stDefName + UTF8Copy(slAuthorList[i], 1,
              UTF8Pos(', ', slAuthorList[i]) - 1);
          end
          // Family name before
          else
          begin
            if fmOptions.cbDotAftertFirstName.Checked = True then
              stDefName := ', ' + UTF8Copy(stDefName, 1, UTF8Length(stDefName) - 1)
            else
              stDefName := ', ' + stDefName;
            slAuthorList[i] :=
              // Family name
              UTF8Copy(slAuthorList[i], 1, UTF8Pos(', ', slAuthorList[i]) - 1) +
              stDefName;
            // Otherwise the author needs no change
          end;
        end;
      end;
    end;
    Result := '';
    if slAuthorList.Count < 2 then
      Result := slAuthorList[0]
    else
    begin
      for i := 0 to slAuthorList.Count - 1 do
      begin
        Result := Result + slAuthorList[i];
        if i < slAuthorList.Count - 1 then
          Result := Result + ' ' + fmOptions.edAuthSep.Text + ' ';
      end;
    end;
    // Might be necessary to remove other possible CR on Win and OS X
    Result := StringReplace(Result, LineEnding, '', [rfReplaceAll, rfIgnoreCase]);
    if blAndOthers = True then
    begin
      Result := Result + ' ' + fmOptions.edMoreAuth.Text;
    end;
  finally
    slAuthorList.Free;
  end;
end;

function TfmMain.ShortenTitle(stTitle: string): string;
begin
  // Shorten the title
  if UTF8Pos('*', stTitle) > 0 then
    Result := UTF8Copy(stTitle, 1, UTF8Pos('*', stTitle) - 1)
  else if UTF8Pos('.', stTitle) > 0 then
    Result := UTF8Copy(stTitle, 1, UTF8Pos('.', stTitle) - 1)
  else if UTF8Pos('?', stTitle) > 0 then
    Result := UTF8Copy(stTitle, 1, UTF8Pos('?', stTitle) - 1)
  else if UTF8Pos('!', stTitle) > 0 then
    Result := UTF8Copy(stTitle, 1, UTF8Pos('!', stTitle) - 1)
  else if UTF8Pos(':', stTitle) > 0 then
    Result := UTF8Copy(stTitle, 1, UTF8Pos(':', stTitle) - 1)
  else
    Result := stTitle;
end;

procedure TfmMain.CompileSummary;
var
  stTitle: string;
begin
  // Compile the label with author and title
  stTitle := StringReplace(sqItems.FieldByName('title').AsString,
    '\emph{', '', [rfReplaceAll, rfIgnoreCase]);
  stTitle := StringReplace(stTitle, '\textit{', '', [rfReplaceAll, rfIgnoreCase]);
  stTitle := StringReplace(stTitle, '\textsc{', '', [rfReplaceAll, rfIgnoreCase]);
  stTitle := StringReplace(stTitle, '\textbf{', '', [rfReplaceAll, rfIgnoreCase]);
  stTitle := StringReplace(stTitle, '{', '', [rfReplaceAll, rfIgnoreCase]);
  stTitle := StringReplace(stTitle, '}', '', [rfReplaceAll, rfIgnoreCase]);
  stTitle := StringReplace(stTitle, '*', '', [rfReplaceAll, rfIgnoreCase]);
  if sqItems.FieldByName('author').AsString <> '' then
    lbAuthorTitle.Caption := sqItems.FieldByName('author').AsString +
      ' / ' + stTitle
  else
    lbAuthorTitle.Caption := stTitle;
end;

function TfmMain.CleanAbstractReview(stText: string): string;
begin
  // Remove single CR at the end of the lines
  stText := StringReplace(stText, LineEnding + LineEnding, #2, [rfReplaceAll]);
  stText := StringReplace(stText, LineEnding, ' ', [rfReplaceAll]);
  Result := StringReplace(stText, #2, LineEnding + LineEnding, [rfReplaceAll]);
end;

procedure TfmMain.InsertKeyAbstRew(meMemo: TMemo);
var
  i: integer;
begin
  // Insert the bibtex key in Abstract and Review
  SaveItems;
  if sqItems.FieldByName('bibtexkey').AsString <> '' then
  begin
    i := meMemo.SelStart;
    meMemo.SelText := '\cite[]{' + sqItems.FieldByName('bibtexkey').AsString + '}';
    meMemo.SelStart := i + 6;
  end;
end;

function TfmMain.IsDirectoryEmpty(const myDir: string): boolean;
var
  SearchRec: TSearchRec;
begin
  // Check if a directory is empty
  try
    // faSysFile (= normal file) to avoid that also the directory is found
    Result := (FindFirst(myDir + DirectorySeparator + '*', faSysFile,
      SearchRec) <> 0);
  finally
    FindClose(SearchRec);
  end;
end;

procedure TfmMain.SetMarkAbstractReviewTab;
begin
  if sqItems.RecordCount > 0 then
  begin
    if sqItems.FieldByName('abstract').AsString <> '' then
      tsAbstract.Caption := '• Abstract'
    else
      tsAbstract.Caption := 'Abstract';
    if sqItems.FieldByName('review').AsString <> '' then
      tsReview.Caption := '• Review'
    else
      tsReview.Caption := 'Review';
  end
  else
  begin
    tsAbstract.Caption := 'Abstract';
    tsReview.Caption := 'Review';
  end;
end;

procedure TfmMain.MoveFocus(Sender: TObject; var Key: word; Shift: TShiftState);
begin
  // Move focus with Return or Tab
  if ((Key = VK_TAB) and (Shift = [ssShift])) then
  begin
    Key := 0;
    SelectNext(ActiveControl, False, True);
  end
  else
  if ((Key = VK_RETURN) or (Key = VK_TAB)) then
  begin
    Key := 0;
    SelectNext(ActiveControl, True, True);
  end;
end;

end.
(*

******* Notes on the Code ***********

* A unique index on Sqlite3Dataset avoides double data, but no error message is shown.
* The field "file" is present in the database, but the attachments are stored in
"AttName" field, so that in the future the first might be used to store data
according to LaTex rules and the second according to Biblatex needs. Up to now,
the "file" field is useful only for exportation.
* If a file is opened with an old structure and an error is raised, it must
be upgraded with the proper function. Anyway, the sqlite component that had
raised the error raises an error again even opening the new upgraded file,
even if it's fine. It must be opened aain (twice), or quit the software,
run it again and open the file.
* A inputquery or inputbox firse the onactivate form event. So I use a flag to
set option and possibly to open a file only for the first activation.
* To fix a bug with DBGrid, data are saved OnResize.

*)
