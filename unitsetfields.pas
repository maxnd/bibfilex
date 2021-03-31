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

unit UnitSetFields;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils;

procedure SetFieldsInGrid;

implementation

uses Unit1;

procedure SetFieldsInGrid;
begin
  // Set fields in grids
  with fmMain do
  begin

    // ********************************************************
    // ********************** ARTICLE *************************
    // ********************************************************

    if LowerCase(cbItemsKind.Text) = 'article' then
    begin
      // Required fields
      sgRequiredFields.Clean;
      sgRequiredFields.RowCount := 10;
      sgRequiredFields.Cells[0, 0] := 'Author';
      sgRequiredFields.Cells[0, 1] := 'Title';
      sgRequiredFields.Cells[0, 2] := 'Journaltitle';
      sgRequiredFields.Cells[0, 3] := 'Volume';
      sgRequiredFields.Cells[0, 4] := 'Number';
      sgRequiredFields.Cells[0, 5] := 'Year';
      sgRequiredFields.Cells[0, 6] := 'Pages';
      sgRequiredFields.Cells[0, 7] := 'Crossref';
      sgRequiredFields.Cells[0, 8] := 'Keywords';
      sgRequiredFields.Cells[0, 9] := 'BibTexKey';
      // Optional fields 1
      sgOptionalFields1.Clean;
      sgOptionalFields1.RowCount := 12;
      sgOptionalFields1.Cells[0, 0] := 'Translator';
      sgOptionalFields1.Cells[0, 1] := 'Annotator';
      sgOptionalFields1.Cells[0, 2] := 'Commentator';
      sgOptionalFields1.Cells[0, 3] := 'Subtitle';
      sgOptionalFields1.Cells[0, 4] := 'Titleaddon';
      sgOptionalFields1.Cells[0, 5] := 'Editor';
      sgOptionalFields1.Cells[0, 6] := 'Editora';
      sgOptionalFields1.Cells[0, 7] := 'Editorb';
      sgOptionalFields1.Cells[0, 8] := 'Editorc';
      sgOptionalFields1.Cells[0, 9] := 'Journalsubtitle';
      sgOptionalFields1.Cells[0, 10] := 'Issuetitle';
      sgOptionalFields1.Cells[0, 11] := 'Issuesubtitle';
      // Optional fields 2
      sgOptionalFields2.Clean;
      sgOptionalFields2.RowCount := 7;
      sgOptionalFields2.Cells[0, 0] := 'Language';
      sgOptionalFields2.Cells[0, 1] := 'Origlanguage';
      sgOptionalFields2.Cells[0, 2] := 'Series';
      sgOptionalFields2.Cells[0, 3] := 'Eid';
      sgOptionalFields2.Cells[0, 4] := 'Issue';
      sgOptionalFields2.Cells[0, 5] := 'Month';
      sgOptionalFields2.Cells[0, 6] := 'Date';
      // Optional fields 3
      sgOptionalFields3.Clean;
      sgOptionalFields3.RowCount := 11;
      sgOptionalFields3.Cells[0, 0] := 'Version';
      sgOptionalFields3.Cells[0, 1] := 'Note';
      sgOptionalFields3.Cells[0, 2] := 'Issn';
      sgOptionalFields3.Cells[0, 3] := 'Addendum';
      sgOptionalFields3.Cells[0, 4] := 'Pubstate';
      sgOptionalFields3.Cells[0, 5] := 'Doi';
      sgOptionalFields3.Cells[0, 6] := 'Eprint';
      sgOptionalFields3.Cells[0, 7] := 'Eprintclass';
      sgOptionalFields3.Cells[0, 8] := 'Eprinttype';
      sgOptionalFields3.Cells[0, 9] := 'Url';
      sgOptionalFields3.Cells[0, 10] := 'Urldate';
    end

    // ********************************************************
    // ********************** BOOK ****************************
    // ********************************************************

    else if LowerCase(cbItemsKind.Text) = 'book' then
    begin
      // Required fields
      sgRequiredFields.Clean;
      sgRequiredFields.RowCount := 11;
      sgRequiredFields.Cells[0, 0] := 'Author';
      sgRequiredFields.Cells[0, 1] := 'Title';
      sgRequiredFields.Cells[0, 2] := 'Editor';
      sgRequiredFields.Cells[0, 3] := 'Series';
      sgRequiredFields.Cells[0, 4] := 'Number';
      sgRequiredFields.Cells[0, 5] := 'Publisher';
      sgRequiredFields.Cells[0, 6] := 'Location';
      sgRequiredFields.Cells[0, 7] := 'Year';
      sgRequiredFields.Cells[0, 8] := 'Crossref';
      sgRequiredFields.Cells[0, 9] := 'Keywords';
      sgRequiredFields.Cells[0, 10] := 'BibTexKey';
      // Optional fields 1
      sgOptionalFields1.Clean;
      sgOptionalFields1.RowCount := 14;
      sgOptionalFields1.Cells[0, 0] := 'Editora';
      sgOptionalFields1.Cells[0, 1] := 'Editorb';
      sgOptionalFields1.Cells[0, 2] := 'Editorc';
      sgOptionalFields1.Cells[0, 3] := 'Translator';
      sgOptionalFields1.Cells[0, 4] := 'Annotator';
      sgOptionalFields1.Cells[0, 5] := 'Commentator';
      sgOptionalFields1.Cells[0, 6] := 'Introduction';
      sgOptionalFields1.Cells[0, 7] := 'Foreword';
      sgOptionalFields1.Cells[0, 8] := 'Afterword';
      sgOptionalFields1.Cells[0, 9] := 'Subtitle';
      sgOptionalFields1.Cells[0, 10] := 'Titleaddon';
      sgOptionalFields1.Cells[0, 11] := 'Maintitle';
      sgOptionalFields1.Cells[0, 12] := 'Mainsubtitle';
      sgOptionalFields1.Cells[0, 13] := 'Maintitleaddon';
      // Optional fields 2
      sgOptionalFields2.Clean;
      sgOptionalFields2.RowCount := 7;
      sgOptionalFields2.Cells[0, 0] := 'Language';
      sgOptionalFields2.Cells[0, 1] := 'Origlanguage';
      sgOptionalFields2.Cells[0, 2] := 'Volume';
      sgOptionalFields2.Cells[0, 3] := 'Part';
      sgOptionalFields2.Cells[0, 4] := 'Edition';
      sgOptionalFields2.Cells[0, 5] := 'Volumes';
      sgOptionalFields2.Cells[0, 6] := 'Date';
      // Optional fields 3
      sgOptionalFields3.Clean;
      sgOptionalFields3.RowCount := 13;
      sgOptionalFields3.Cells[0, 0] := 'Note';
      sgOptionalFields3.Cells[0, 1] := 'Isbn';
      sgOptionalFields3.Cells[0, 2] := 'Chapter';
      sgOptionalFields3.Cells[0, 3] := 'Pages';
      sgOptionalFields3.Cells[0, 4] := 'Pagetotal';
      sgOptionalFields3.Cells[0, 5] := 'Addendum';
      sgOptionalFields3.Cells[0, 6] := 'Pubstate';
      sgOptionalFields3.Cells[0, 7] := 'Doi';
      sgOptionalFields3.Cells[0, 8] := 'Eprint';
      sgOptionalFields3.Cells[0, 9] := 'Eprintclass';
      sgOptionalFields3.Cells[0, 10] := 'Eprinttype';
      sgOptionalFields3.Cells[0, 11] := 'Url';
      sgOptionalFields3.Cells[0, 12] := 'Urldate';
    end

    // ********************************************************
    // ********************** MVBOOK **************************
    // ********************************************************

    else if LowerCase(cbItemsKind.Text) = 'mvbook' then
    begin
      // Required fields
      sgRequiredFields.Clean;
      sgRequiredFields.RowCount := 10;
      sgRequiredFields.Cells[0, 0] := 'Author';
      sgRequiredFields.Cells[0, 1] := 'Title';
      sgRequiredFields.Cells[0, 2] := 'Editor';
      sgRequiredFields.Cells[0, 3] := 'Series';
      sgRequiredFields.Cells[0, 4] := 'Number';
      sgRequiredFields.Cells[0, 5] := 'Publisher';
      sgRequiredFields.Cells[0, 6] := 'Location';
      sgRequiredFields.Cells[0, 7] := 'Year';
      sgRequiredFields.Cells[0, 8] := 'Keywords';
      sgRequiredFields.Cells[0, 9] := 'BibTexKey';
      // Optional fields 1
      sgOptionalFields1.Clean;
      sgOptionalFields1.RowCount := 11;
      sgOptionalFields1.Cells[0, 0] := 'Editora';
      sgOptionalFields1.Cells[0, 1] := 'Editorb';
      sgOptionalFields1.Cells[0, 2] := 'Editorc';
      sgOptionalFields1.Cells[0, 3] := 'Translator';
      sgOptionalFields1.Cells[0, 4] := 'Annotator';
      sgOptionalFields1.Cells[0, 5] := 'Commentator';
      sgOptionalFields1.Cells[0, 6] := 'Introduction';
      sgOptionalFields1.Cells[0, 7] := 'Foreword';
      sgOptionalFields1.Cells[0, 8] := 'Afterword';
      sgOptionalFields1.Cells[0, 9] := 'Subtitle';
      sgOptionalFields1.Cells[0, 10] := 'Titleaddon';
      // Optional fields 2
      sgOptionalFields2.Clean;
      sgOptionalFields2.RowCount := 5;
      sgOptionalFields2.Cells[0, 0] := 'Language';
      sgOptionalFields2.Cells[0, 1] := 'Origlanguage';
      sgOptionalFields2.Cells[0, 2] := 'Edition';
      sgOptionalFields2.Cells[0, 3] := 'Volumes';
      sgOptionalFields2.Cells[0, 4] := 'Date';
      // Optional fields 3
      sgOptionalFields3.Clean;
      sgOptionalFields3.RowCount := 11;
      sgOptionalFields3.Cells[0, 0] := 'Note';
      sgOptionalFields3.Cells[0, 1] := 'Isbn';
      sgOptionalFields3.Cells[0, 2] := 'Pagetotal';
      sgOptionalFields3.Cells[0, 3] := 'Addendum';
      sgOptionalFields3.Cells[0, 4] := 'Pubstate';
      sgOptionalFields3.Cells[0, 5] := 'Doi';
      sgOptionalFields3.Cells[0, 6] := 'Eprint';
      sgOptionalFields3.Cells[0, 7] := 'Eprintclass';
      sgOptionalFields3.Cells[0, 8] := 'Eprinttype';
      sgOptionalFields3.Cells[0, 9] := 'Url';
      sgOptionalFields3.Cells[0, 10] := 'Urldate';
    end

    // ********************************************************
    // ********************** INBOOK **************************
    // ********************************************************

    else if LowerCase(cbItemsKind.Text) = 'inbook' then
    begin
      // Required fields
      sgRequiredFields.Clean;
      sgRequiredFields.RowCount := 14;
      sgRequiredFields.Cells[0, 0] := 'Author';
      sgRequiredFields.Cells[0, 1] := 'Title';
      sgRequiredFields.Cells[0, 2] := 'Bookauthor';
      sgRequiredFields.Cells[0, 3] := 'Booktitle';
      sgRequiredFields.Cells[0, 4] := 'Editor';
      sgRequiredFields.Cells[0, 5] := 'Series';
      sgRequiredFields.Cells[0, 6] := 'Number';
      sgRequiredFields.Cells[0, 7] := 'Publisher';
      sgRequiredFields.Cells[0, 8] := 'Location';
      sgRequiredFields.Cells[0, 9] := 'Year';
      sgRequiredFields.Cells[0, 10] := 'Pages';
      sgRequiredFields.Cells[0, 11] := 'Crossref';
      sgRequiredFields.Cells[0, 12] := 'Keywords';
      sgRequiredFields.Cells[0, 13] := 'BibTexKey';
      // Optional fields 1
      sgOptionalFields1.Clean;
      sgOptionalFields1.RowCount := 16;
      sgOptionalFields1.Cells[0, 0] := 'Editora';
      sgOptionalFields1.Cells[0, 1] := 'Editorb';
      sgOptionalFields1.Cells[0, 2] := 'Editorc';
      sgOptionalFields1.Cells[0, 3] := 'Translator';
      sgOptionalFields1.Cells[0, 4] := 'Annotator';
      sgOptionalFields1.Cells[0, 5] := 'Commentator';
      sgOptionalFields1.Cells[0, 6] := 'Introduction';
      sgOptionalFields1.Cells[0, 7] := 'Foreword';
      sgOptionalFields1.Cells[0, 8] := 'Afterword';
      sgOptionalFields1.Cells[0, 9] := 'Subtitle';
      sgOptionalFields1.Cells[0, 10] := 'Titleaddon';
      sgOptionalFields1.Cells[0, 11] := 'Maintitle';
      sgOptionalFields1.Cells[0, 12] := 'Mainsubtitle';
      sgOptionalFields1.Cells[0, 13] := 'Maintitleaddon';
      sgOptionalFields1.Cells[0, 14] := 'Booksubtitle';
      sgOptionalFields1.Cells[0, 15] := 'Booktitleaddon';
      // Optional fields 2
      sgOptionalFields2.Clean;
      sgOptionalFields2.RowCount := 7;
      sgOptionalFields2.Cells[0, 0] := 'Language';
      sgOptionalFields2.Cells[0, 1] := 'Origlanguage';
      sgOptionalFields2.Cells[0, 2] := 'Volume';
      sgOptionalFields2.Cells[0, 3] := 'Part';
      sgOptionalFields2.Cells[0, 4] := 'Edition';
      sgOptionalFields2.Cells[0, 5] := 'Volumes';
      sgOptionalFields2.Cells[0, 6] := 'Date';
      // Optional fields 3
      sgOptionalFields3.Clean;
      sgOptionalFields3.RowCount := 11;
      sgOptionalFields3.Cells[0, 0] := 'Note';
      sgOptionalFields3.Cells[0, 1] := 'Isbn';
      sgOptionalFields3.Cells[0, 2] := 'Chapter';
      sgOptionalFields3.Cells[0, 3] := 'Addendum';
      sgOptionalFields3.Cells[0, 4] := 'Pubstate';
      sgOptionalFields3.Cells[0, 5] := 'Doi';
      sgOptionalFields3.Cells[0, 6] := 'Eprint';
      sgOptionalFields3.Cells[0, 7] := 'Eprintclass';
      sgOptionalFields3.Cells[0, 8] := 'Eprinttype';
      sgOptionalFields3.Cells[0, 9] := 'Url';
      sgOptionalFields3.Cells[0, 10] := 'Urldate';
    end

    // ********************************************************
    // ********************** BOOKINBOOK **********************
    // ********************************************************

    else if LowerCase(cbItemsKind.Text) = 'bookinbook' then
    begin
      // Required fields
      sgRequiredFields.Clean;
      sgRequiredFields.RowCount := 14;
      sgRequiredFields.Cells[0, 0] := 'Author';
      sgRequiredFields.Cells[0, 1] := 'Title';
      sgRequiredFields.Cells[0, 2] := 'Bookauthor';
      sgRequiredFields.Cells[0, 3] := 'Booktitle';
      sgRequiredFields.Cells[0, 4] := 'Editor';
      sgRequiredFields.Cells[0, 5] := 'Series';
      sgRequiredFields.Cells[0, 6] := 'Number';
      sgRequiredFields.Cells[0, 7] := 'Publisher';
      sgRequiredFields.Cells[0, 8] := 'Location';
      sgRequiredFields.Cells[0, 9] := 'Year';
      sgRequiredFields.Cells[0, 10] := 'Pages';
      sgRequiredFields.Cells[0, 11] := 'Crossref';
      sgRequiredFields.Cells[0, 12] := 'Keywords';
      sgRequiredFields.Cells[0, 13] := 'BibTexKey';
      // Optional fields 1
      sgOptionalFields1.Clean;
      sgOptionalFields1.RowCount := 16;
      sgOptionalFields1.Cells[0, 0] := 'Editora';
      sgOptionalFields1.Cells[0, 1] := 'Editorb';
      sgOptionalFields1.Cells[0, 2] := 'Editorc';
      sgOptionalFields1.Cells[0, 3] := 'Translator';
      sgOptionalFields1.Cells[0, 4] := 'Annotator';
      sgOptionalFields1.Cells[0, 5] := 'Commentator';
      sgOptionalFields1.Cells[0, 6] := 'Introduction';
      sgOptionalFields1.Cells[0, 7] := 'Foreword';
      sgOptionalFields1.Cells[0, 8] := 'Afterword';
      sgOptionalFields1.Cells[0, 9] := 'Subtitle';
      sgOptionalFields1.Cells[0, 10] := 'Titleaddon';
      sgOptionalFields1.Cells[0, 11] := 'Maintitle';
      sgOptionalFields1.Cells[0, 12] := 'Mainsubtitle';
      sgOptionalFields1.Cells[0, 13] := 'Maintitleaddon';
      sgOptionalFields1.Cells[0, 14] := 'Booksubtitle';
      sgOptionalFields1.Cells[0, 15] := 'Booktitleaddon';
      // Optional fields 2
      sgOptionalFields2.Clean;
      sgOptionalFields2.RowCount := 7;
      sgOptionalFields2.Cells[0, 0] := 'Language';
      sgOptionalFields2.Cells[0, 1] := 'Origlanguage';
      sgOptionalFields2.Cells[0, 2] := 'Volume';
      sgOptionalFields2.Cells[0, 3] := 'Part';
      sgOptionalFields2.Cells[0, 4] := 'Edition';
      sgOptionalFields2.Cells[0, 5] := 'Volumes';
      sgOptionalFields2.Cells[0, 6] := 'Date';
      // Optional fields 3
      sgOptionalFields3.Clean;
      sgOptionalFields3.RowCount := 11;
      sgOptionalFields3.Cells[0, 0] := 'Note';
      sgOptionalFields3.Cells[0, 1] := 'Isbn';
      sgOptionalFields3.Cells[0, 2] := 'Chapter';
      sgOptionalFields3.Cells[0, 3] := 'Addendum';
      sgOptionalFields3.Cells[0, 4] := 'Pubstate';
      sgOptionalFields3.Cells[0, 5] := 'Doi';
      sgOptionalFields3.Cells[0, 6] := 'Eprint';
      sgOptionalFields3.Cells[0, 7] := 'Eprintclass';
      sgOptionalFields3.Cells[0, 8] := 'Eprinttype';
      sgOptionalFields3.Cells[0, 9] := 'Url';
      sgOptionalFields3.Cells[0, 10] := 'Urldate';
    end

    // ********************************************************
    // ********************** SUPPBOOK ************************
    // ********************************************************

    else if LowerCase(cbItemsKind.Text) = 'suppbook' then
    begin
      // Required fields
      sgRequiredFields.Clean;
      sgRequiredFields.RowCount := 14;
      sgRequiredFields.Cells[0, 0] := 'Author';
      sgRequiredFields.Cells[0, 1] := 'Title';
      sgRequiredFields.Cells[0, 2] := 'Bookauthor';
      sgRequiredFields.Cells[0, 3] := 'Booktitle';
      sgRequiredFields.Cells[0, 4] := 'Editor';
      sgRequiredFields.Cells[0, 5] := 'Series';
      sgRequiredFields.Cells[0, 6] := 'Number';
      sgRequiredFields.Cells[0, 7] := 'Publisher';
      sgRequiredFields.Cells[0, 8] := 'Location';
      sgRequiredFields.Cells[0, 9] := 'Year';
      sgRequiredFields.Cells[0, 10] := 'Pages';
      sgRequiredFields.Cells[0, 11] := 'Crossref';
      sgRequiredFields.Cells[0, 12] := 'Keywords';
      sgRequiredFields.Cells[0, 13] := 'BibTexKey';
      // Optional fields 1
      sgOptionalFields1.Clean;
      sgOptionalFields1.RowCount := 16;
      sgOptionalFields1.Cells[0, 0] := 'Editora';
      sgOptionalFields1.Cells[0, 1] := 'Editorb';
      sgOptionalFields1.Cells[0, 2] := 'Editorc';
      sgOptionalFields1.Cells[0, 3] := 'Translator';
      sgOptionalFields1.Cells[0, 4] := 'Annotator';
      sgOptionalFields1.Cells[0, 5] := 'Commentator';
      sgOptionalFields1.Cells[0, 6] := 'Introduction';
      sgOptionalFields1.Cells[0, 7] := 'Foreword';
      sgOptionalFields1.Cells[0, 8] := 'Afterword';
      sgOptionalFields1.Cells[0, 9] := 'Subtitle';
      sgOptionalFields1.Cells[0, 10] := 'Titleaddon';
      sgOptionalFields1.Cells[0, 11] := 'Maintitle';
      sgOptionalFields1.Cells[0, 12] := 'Mainsubtitle';
      sgOptionalFields1.Cells[0, 13] := 'Maintitleaddon';
      sgOptionalFields1.Cells[0, 14] := 'Booksubtitle';
      sgOptionalFields1.Cells[0, 15] := 'Booktitleaddon';
      // Optional fields 2
      sgOptionalFields2.Clean;
      sgOptionalFields2.RowCount := 7;
      sgOptionalFields2.Cells[0, 0] := 'Language';
      sgOptionalFields2.Cells[0, 1] := 'Origlanguage';
      sgOptionalFields2.Cells[0, 2] := 'Volume';
      sgOptionalFields2.Cells[0, 3] := 'Part';
      sgOptionalFields2.Cells[0, 4] := 'Edition';
      sgOptionalFields2.Cells[0, 5] := 'Volumes';
      sgOptionalFields2.Cells[0, 6] := 'Date';
      // Optional fields 3
      sgOptionalFields3.Clean;
      sgOptionalFields3.RowCount := 11;
      sgOptionalFields3.Cells[0, 0] := 'Note';
      sgOptionalFields3.Cells[0, 1] := 'Isbn';
      sgOptionalFields3.Cells[0, 2] := 'Chapter';
      sgOptionalFields3.Cells[0, 3] := 'Addendum';
      sgOptionalFields3.Cells[0, 4] := 'Pubstate';
      sgOptionalFields3.Cells[0, 5] := 'Doi';
      sgOptionalFields3.Cells[0, 6] := 'Eprint';
      sgOptionalFields3.Cells[0, 7] := 'Eprintclass';
      sgOptionalFields3.Cells[0, 8] := 'Eprinttype';
      sgOptionalFields3.Cells[0, 9] := 'Url';
      sgOptionalFields3.Cells[0, 10] := 'Urldate';
    end

    // ********************************************************
    // ********************** BOOKLET *************************
    // ********************************************************

    else if LowerCase(cbItemsKind.Text) = 'booklet' then
    begin
      // Required fields
      sgRequiredFields.Clean;
      sgRequiredFields.RowCount := 8;
      sgRequiredFields.Cells[0, 0] := 'Author';
      sgRequiredFields.Cells[0, 1] := 'Title';
      sgRequiredFields.Cells[0, 2] := 'Editor';
      sgRequiredFields.Cells[0, 3] := 'Howpublished';
      sgRequiredFields.Cells[0, 4] := 'Location';
      sgRequiredFields.Cells[0, 5] := 'Year';
      sgRequiredFields.Cells[0, 6] := 'Keywords';
      sgRequiredFields.Cells[0, 7] := 'BibTexKey';
      // Optional fields 1
      sgOptionalFields1.Clean;
      sgOptionalFields1.RowCount := 2;
      sgOptionalFields1.Cells[0, 0] := 'Subtitle';
      sgOptionalFields1.Cells[0, 1] := 'Titleaddon';
      // Optional fields 2
      sgOptionalFields2.Clean;
      sgOptionalFields2.RowCount := 3;
      sgOptionalFields2.Cells[0, 0] := 'Language';
      sgOptionalFields2.Cells[0, 1] := 'Type';
      sgOptionalFields2.Cells[0, 2] := 'Date';
      // Optional fields 3
      sgOptionalFields3.Clean;
      sgOptionalFields3.RowCount := 12;
      sgOptionalFields3.Cells[0, 0] := 'Note';
      sgOptionalFields3.Cells[0, 1] := 'Chapter';
      sgOptionalFields3.Cells[0, 2] := 'Pages';
      sgOptionalFields3.Cells[0, 3] := 'Pagetotal';
      sgOptionalFields3.Cells[0, 4] := 'Addendum';
      sgOptionalFields3.Cells[0, 5] := 'Pubstate';
      sgOptionalFields3.Cells[0, 6] := 'Doi';
      sgOptionalFields3.Cells[0, 7] := 'Eprint';
      sgOptionalFields3.Cells[0, 8] := 'Eprintclass';
      sgOptionalFields3.Cells[0, 9] := 'Eprinttype';
      sgOptionalFields3.Cells[0, 10] := 'Url';
      sgOptionalFields3.Cells[0, 11] := 'Urldate';
    end

    // ********************************************************
    // ********************** COLLECTION **********************
    // ********************************************************

    else if LowerCase(cbItemsKind.Text) = 'collection' then
    begin
      // Required fields
      sgRequiredFields.Clean;
      sgRequiredFields.RowCount := 10;
      sgRequiredFields.Cells[0, 0] := 'Editor';
      sgRequiredFields.Cells[0, 1] := 'Title';
      sgRequiredFields.Cells[0, 2] := 'Publisher';
      sgRequiredFields.Cells[0, 3] := 'Series';
      sgRequiredFields.Cells[0, 4] := 'Number';
      sgRequiredFields.Cells[0, 5] := 'Location';
      sgRequiredFields.Cells[0, 6] := 'Year';
      sgRequiredFields.Cells[0, 7] := 'Crossref';
      sgRequiredFields.Cells[0, 8] := 'Keywords';
      sgRequiredFields.Cells[0, 9] := 'BibTexKey';
      // Optional fields 1
      sgOptionalFields1.Clean;
      sgOptionalFields1.RowCount := 14;
      sgOptionalFields1.Cells[0, 0] := 'Editora';
      sgOptionalFields1.Cells[0, 1] := 'Editorb';
      sgOptionalFields1.Cells[0, 2] := 'Editorc';
      sgOptionalFields1.Cells[0, 3] := 'Translator';
      sgOptionalFields1.Cells[0, 4] := 'Annotator';
      sgOptionalFields1.Cells[0, 5] := 'Commentator';
      sgOptionalFields1.Cells[0, 6] := 'Introduction';
      sgOptionalFields1.Cells[0, 7] := 'Foreword';
      sgOptionalFields1.Cells[0, 8] := 'Afterword';
      sgOptionalFields1.Cells[0, 9] := 'Subtitle';
      sgOptionalFields1.Cells[0, 10] := 'Titleaddon';
      sgOptionalFields1.Cells[0, 11] := 'Maintitle';
      sgOptionalFields1.Cells[0, 12] := 'Mainsubtitle';
      sgOptionalFields1.Cells[0, 13] := 'Maintitleaddon';
      // Optional fields 2
      sgOptionalFields2.Clean;
      sgOptionalFields2.RowCount := 7;
      sgOptionalFields2.Cells[0, 0] := 'Language';
      sgOptionalFields2.Cells[0, 1] := 'Origlanguage';
      sgOptionalFields2.Cells[0, 2] := 'Volume';
      sgOptionalFields2.Cells[0, 3] := 'Part';
      sgOptionalFields2.Cells[0, 4] := 'Edition';
      sgOptionalFields2.Cells[0, 5] := 'Volumes';
      sgOptionalFields2.Cells[0, 6] := 'Date';
      // Optional fields 3
      sgOptionalFields3.Clean;
      sgOptionalFields3.RowCount := 13;
      sgOptionalFields3.Cells[0, 0] := 'Note';
      sgOptionalFields3.Cells[0, 1] := 'Isbn';
      sgOptionalFields3.Cells[0, 2] := 'Chapter';
      sgOptionalFields3.Cells[0, 3] := 'Pages';
      sgOptionalFields3.Cells[0, 4] := 'Pagetotal';
      sgOptionalFields3.Cells[0, 5] := 'Addendum';
      sgOptionalFields3.Cells[0, 6] := 'Pubstate';
      sgOptionalFields3.Cells[0, 7] := 'Doi';
      sgOptionalFields3.Cells[0, 8] := 'Eprint';
      sgOptionalFields3.Cells[0, 9] := 'Eprintclass';
      sgOptionalFields3.Cells[0, 10] := 'Eprinttype';
      sgOptionalFields3.Cells[0, 11] := 'Url';
      sgOptionalFields3.Cells[0, 12] := 'Urldate';
    end

    // ********************************************************
    // ********************** MVCOLLECTION ********************
    // ********************************************************

    else if LowerCase(cbItemsKind.Text) = 'mvcollection' then
    begin
      // Required fields
      sgRequiredFields.Clean;
      sgRequiredFields.RowCount := 9;
      sgRequiredFields.Cells[0, 0] := 'Editor';
      sgRequiredFields.Cells[0, 1] := 'Title';
      sgRequiredFields.Cells[0, 2] := 'Series';
      sgRequiredFields.Cells[0, 3] := 'Number';
      sgRequiredFields.Cells[0, 4] := 'Publisher';
      sgRequiredFields.Cells[0, 5] := 'Location';
      sgRequiredFields.Cells[0, 6] := 'Year';
      sgRequiredFields.Cells[0, 7] := 'Keywords';
      sgRequiredFields.Cells[0, 8] := 'BibTexKey';
      // Optional fields 1
      sgOptionalFields1.Clean;
      sgOptionalFields1.RowCount := 11;
      sgOptionalFields1.Cells[0, 0] := 'Editora';
      sgOptionalFields1.Cells[0, 1] := 'Editorb';
      sgOptionalFields1.Cells[0, 2] := 'Editorc';
      sgOptionalFields1.Cells[0, 3] := 'Translator';
      sgOptionalFields1.Cells[0, 4] := 'Annotator';
      sgOptionalFields1.Cells[0, 5] := 'Commentator';
      sgOptionalFields1.Cells[0, 6] := 'Introduction';
      sgOptionalFields1.Cells[0, 7] := 'Foreword';
      sgOptionalFields1.Cells[0, 8] := 'Afterword';
      sgOptionalFields1.Cells[0, 9] := 'Subtitle';
      sgOptionalFields1.Cells[0, 10] := 'Titleaddon';
      // Optional fields 2
      sgOptionalFields2.Clean;
      sgOptionalFields2.RowCount := 5;
      sgOptionalFields2.Cells[0, 0] := 'Language';
      sgOptionalFields2.Cells[0, 1] := 'Origlanguage';
      sgOptionalFields2.Cells[0, 2] := 'Edition';
      sgOptionalFields2.Cells[0, 3] := 'Volumes';
      sgOptionalFields2.Cells[0, 4] := 'Date';
      // Optional fields 3
      sgOptionalFields3.Clean;
      sgOptionalFields3.RowCount := 11;
      sgOptionalFields3.Cells[0, 0] := 'Note';
      sgOptionalFields3.Cells[0, 1] := 'Isbn';
      sgOptionalFields3.Cells[0, 2] := 'Pagetotal';
      sgOptionalFields3.Cells[0, 3] := 'Addendum';
      sgOptionalFields3.Cells[0, 4] := 'Pubstate';
      sgOptionalFields3.Cells[0, 5] := 'Doi';
      sgOptionalFields3.Cells[0, 6] := 'Eprint';
      sgOptionalFields3.Cells[0, 7] := 'Eprintclass';
      sgOptionalFields3.Cells[0, 8] := 'Eprinttype';
      sgOptionalFields3.Cells[0, 9] := 'Url';
      sgOptionalFields3.Cells[0, 10] := 'Urldate';
    end

    // ********************************************************
    // ********************** INCOLLECTION ********************
    // ********************************************************

    else if LowerCase(cbItemsKind.Text) = 'incollection' then
    begin
      // Required fields
      sgRequiredFields.Clean;
      sgRequiredFields.RowCount := 13;
      sgRequiredFields.Cells[0, 0] := 'Author';
      sgRequiredFields.Cells[0, 1] := 'Editor';
      sgRequiredFields.Cells[0, 2] := 'Title';
      sgRequiredFields.Cells[0, 3] := 'Booktitle';
      sgRequiredFields.Cells[0, 4] := 'Series';
      sgRequiredFields.Cells[0, 5] := 'Number';
      sgRequiredFields.Cells[0, 6] := 'Publisher';
      sgRequiredFields.Cells[0, 7] := 'Location';
      sgRequiredFields.Cells[0, 8] := 'Year';
      sgRequiredFields.Cells[0, 9] := 'Pages';
      sgRequiredFields.Cells[0, 10] := 'Crossref';
      sgRequiredFields.Cells[0, 11] := 'Keywords';
      sgRequiredFields.Cells[0, 12] := 'BibTexKey';
      // Optional fields 1
      sgOptionalFields1.Clean;
      sgOptionalFields1.RowCount := 16;
      sgOptionalFields1.Cells[0, 0] := 'Editora';
      sgOptionalFields1.Cells[0, 1] := 'Editorb';
      sgOptionalFields1.Cells[0, 2] := 'Editorc';
      sgOptionalFields1.Cells[0, 3] := 'Translator';
      sgOptionalFields1.Cells[0, 4] := 'Annotator';
      sgOptionalFields1.Cells[0, 5] := 'Commentator';
      sgOptionalFields1.Cells[0, 6] := 'Introduction';
      sgOptionalFields1.Cells[0, 7] := 'Foreword';
      sgOptionalFields1.Cells[0, 8] := 'Afterword';
      sgOptionalFields1.Cells[0, 9] := 'Subtitle';
      sgOptionalFields1.Cells[0, 10] := 'Titleaddon';
      sgOptionalFields1.Cells[0, 11] := 'Maintitle';
      sgOptionalFields1.Cells[0, 12] := 'Mainsubtitle';
      sgOptionalFields1.Cells[0, 13] := 'Maintitleaddon';
      sgOptionalFields1.Cells[0, 14] := 'Booksubtitle';
      sgOptionalFields1.Cells[0, 15] := 'Booktitleaddon';
      // Optional fields 2
      sgOptionalFields2.Clean;
      sgOptionalFields2.RowCount := 7;
      sgOptionalFields2.Cells[0, 0] := 'Language';
      sgOptionalFields2.Cells[0, 1] := 'Origlanguage';
      sgOptionalFields2.Cells[0, 2] := 'Volume';
      sgOptionalFields2.Cells[0, 3] := 'Part';
      sgOptionalFields2.Cells[0, 4] := 'Edition';
      sgOptionalFields2.Cells[0, 5] := 'Volumes';
      sgOptionalFields2.Cells[0, 6] := 'Date';
      // Optional fields 3
      sgOptionalFields3.Clean;
      sgOptionalFields3.RowCount := 11;
      sgOptionalFields3.Cells[0, 0] := 'Note';
      sgOptionalFields3.Cells[0, 1] := 'Isbn';
      sgOptionalFields3.Cells[0, 2] := 'Chapter';
      sgOptionalFields3.Cells[0, 3] := 'Addendum';
      sgOptionalFields3.Cells[0, 4] := 'Pubstate';
      sgOptionalFields3.Cells[0, 5] := 'Doi';
      sgOptionalFields3.Cells[0, 6] := 'Eprint';
      sgOptionalFields3.Cells[0, 7] := 'Eprintclass';
      sgOptionalFields3.Cells[0, 8] := 'Eprinttype';
      sgOptionalFields3.Cells[0, 9] := 'Url';
      sgOptionalFields3.Cells[0, 10] := 'Urldate';
    end

    // ********************************************************
    // ********************** SUPPCOLLECTION ******************
    // ********************************************************

    else if LowerCase(cbItemsKind.Text) = 'suppcollection' then
    begin
      // Required fields
      sgRequiredFields.Clean;
      sgRequiredFields.RowCount := 13;
      sgRequiredFields.Cells[0, 0] := 'Author';
      sgRequiredFields.Cells[0, 1] := 'Editor';
      sgRequiredFields.Cells[0, 2] := 'Title';
      sgRequiredFields.Cells[0, 3] := 'Booktitle';
      sgRequiredFields.Cells[0, 4] := 'Series';
      sgRequiredFields.Cells[0, 5] := 'Number';
      sgRequiredFields.Cells[0, 6] := 'Publisher';
      sgRequiredFields.Cells[0, 7] := 'Location';
      sgRequiredFields.Cells[0, 8] := 'Year';
      sgRequiredFields.Cells[0, 9] := 'Pages';
      sgRequiredFields.Cells[0, 10] := 'Crossref';
      sgRequiredFields.Cells[0, 11] := 'Keywords';
      sgRequiredFields.Cells[0, 12] := 'BibTexKey';
      // Optional fields 1
      sgOptionalFields1.Clean;
      sgOptionalFields1.RowCount := 16;
      sgOptionalFields1.Cells[0, 0] := 'Editora';
      sgOptionalFields1.Cells[0, 1] := 'Editorb';
      sgOptionalFields1.Cells[0, 2] := 'Editorc';
      sgOptionalFields1.Cells[0, 3] := 'Translator';
      sgOptionalFields1.Cells[0, 4] := 'Annotator';
      sgOptionalFields1.Cells[0, 5] := 'Commentator';
      sgOptionalFields1.Cells[0, 6] := 'Introduction';
      sgOptionalFields1.Cells[0, 7] := 'Foreword';
      sgOptionalFields1.Cells[0, 8] := 'Afterword';
      sgOptionalFields1.Cells[0, 9] := 'Subtitle';
      sgOptionalFields1.Cells[0, 10] := 'Titleaddon';
      sgOptionalFields1.Cells[0, 11] := 'Maintitle';
      sgOptionalFields1.Cells[0, 12] := 'Mainsubtitle';
      sgOptionalFields1.Cells[0, 13] := 'Maintitleaddon';
      sgOptionalFields1.Cells[0, 14] := 'Booksubtitle';
      sgOptionalFields1.Cells[0, 15] := 'Booktitleaddon';
      // Optional fields 2
      sgOptionalFields2.Clean;
      sgOptionalFields2.RowCount := 7;
      sgOptionalFields2.Cells[0, 0] := 'Language';
      sgOptionalFields2.Cells[0, 1] := 'Origlanguage';
      sgOptionalFields2.Cells[0, 2] := 'Volume';
      sgOptionalFields2.Cells[0, 3] := 'Part';
      sgOptionalFields2.Cells[0, 4] := 'Edition';
      sgOptionalFields2.Cells[0, 5] := 'Volumes';
      sgOptionalFields2.Cells[0, 6] := 'Date';
      // Optional fields 3
      sgOptionalFields3.Clean;
      sgOptionalFields3.RowCount := 11;
      sgOptionalFields3.Cells[0, 0] := 'Note';
      sgOptionalFields3.Cells[0, 1] := 'Isbn';
      sgOptionalFields3.Cells[0, 2] := 'Chapter';
      sgOptionalFields3.Cells[0, 3] := 'Addendum';
      sgOptionalFields3.Cells[0, 4] := 'Pubstate';
      sgOptionalFields3.Cells[0, 5] := 'Doi';
      sgOptionalFields3.Cells[0, 6] := 'Eprint';
      sgOptionalFields3.Cells[0, 7] := 'Eprintclass';
      sgOptionalFields3.Cells[0, 8] := 'Eprinttype';
      sgOptionalFields3.Cells[0, 9] := 'Url';
      sgOptionalFields3.Cells[0, 10] := 'Urldate';
    end

    // ********************************************************
    // ********************** MANUAL **************************
    // ********************************************************

    else if LowerCase(cbItemsKind.Text) = 'manual' then
    begin
      // Required fields
      sgRequiredFields.Clean;
      sgRequiredFields.RowCount := 10;
      sgRequiredFields.Cells[0, 0] := 'Author';
      sgRequiredFields.Cells[0, 1] := 'Title';
      sgRequiredFields.Cells[0, 2] := 'Editor';
      sgRequiredFields.Cells[0, 3] := 'Series';
      sgRequiredFields.Cells[0, 4] := 'Number';
      sgRequiredFields.Cells[0, 5] := 'Publisher';
      sgRequiredFields.Cells[0, 6] := 'Location';
      sgRequiredFields.Cells[0, 7] := 'Year';
      sgRequiredFields.Cells[0, 8] := 'Keywords';
      sgRequiredFields.Cells[0, 9] := 'BibTexKey';
      // Optional fields 1
      sgOptionalFields1.Clean;
      sgOptionalFields1.RowCount := 2;
      sgOptionalFields1.Cells[0, 0] := 'Subtitle';
      sgOptionalFields1.Cells[0, 1] := 'Titleaddon';
      // Optional fields 2
      sgOptionalFields2.Clean;
      sgOptionalFields2.RowCount := 5;
      sgOptionalFields2.Cells[0, 0] := 'Language';
      sgOptionalFields2.Cells[0, 1] := 'Edition';
      sgOptionalFields2.Cells[0, 2] := 'Type';
      sgOptionalFields2.Cells[0, 3] := 'Version';
      sgOptionalFields2.Cells[0, 4] := 'Date';
      // Optional fields 3
      sgOptionalFields3.Clean;
      sgOptionalFields3.RowCount := 14;
      sgOptionalFields3.Cells[0, 0] := 'Note';
      sgOptionalFields3.Cells[0, 1] := 'Organization';
      sgOptionalFields3.Cells[0, 2] := 'Isbn';
      sgOptionalFields3.Cells[0, 3] := 'Chapter';
      sgOptionalFields3.Cells[0, 4] := 'Pages';
      sgOptionalFields3.Cells[0, 5] := 'Pagetotal';
      sgOptionalFields3.Cells[0, 6] := 'Addendum';
      sgOptionalFields3.Cells[0, 7] := 'Pubstate';
      sgOptionalFields3.Cells[0, 8] := 'Doi';
      sgOptionalFields3.Cells[0, 9] := 'Eprint';
      sgOptionalFields3.Cells[0, 10] := 'Eprintclass';
      sgOptionalFields3.Cells[0, 11] := 'Eprinttype';
      sgOptionalFields3.Cells[0, 12] := 'Url';
      sgOptionalFields3.Cells[0, 13] := 'Urldate';
    end

    // ********************************************************
    // ********************** MISC ****************************
    // ********************************************************

    else if LowerCase(cbItemsKind.Text) = 'misc' then
    begin
      // Required fields
      sgRequiredFields.Clean;
      sgRequiredFields.RowCount := 8;
      sgRequiredFields.Cells[0, 0] := 'Author';
      sgRequiredFields.Cells[0, 1] := 'Title';
      sgRequiredFields.Cells[0, 2] := 'Editor';
      sgRequiredFields.Cells[0, 3] := 'Howpublished';
      sgRequiredFields.Cells[0, 4] := 'Location';
      sgRequiredFields.Cells[0, 5] := 'Year';
      sgRequiredFields.Cells[0, 6] := 'Keywords';
      sgRequiredFields.Cells[0, 7] := 'BibTexKey';
      // Optional fields 1
      sgOptionalFields1.Clean;
      sgOptionalFields1.RowCount := 2;
      sgOptionalFields1.Cells[0, 0] := 'Subtitle';
      sgOptionalFields1.Cells[0, 1] := 'Titleaddon';
      // Optional fields 2
      sgOptionalFields2.Clean;
      sgOptionalFields2.RowCount := 3;
      sgOptionalFields2.Cells[0, 0] := 'Language';
      sgOptionalFields2.Cells[0, 1] := 'Type';
      sgOptionalFields2.Cells[0, 2] := 'Version';
      // Optional fields 3
      sgOptionalFields3.Clean;
      sgOptionalFields3.RowCount := 12;
      sgOptionalFields3.Cells[0, 0] := 'Note';
      sgOptionalFields3.Cells[0, 1] := 'Organization';
      sgOptionalFields3.Cells[0, 2] := 'Date';
      sgOptionalFields3.Cells[0, 3] := 'Month';
      sgOptionalFields3.Cells[0, 4] := 'Addendum';
      sgOptionalFields3.Cells[0, 5] := 'Pubstate';
      sgOptionalFields3.Cells[0, 6] := 'Doi';
      sgOptionalFields3.Cells[0, 7] := 'Eprint';
      sgOptionalFields3.Cells[0, 8] := 'Eprintclass';
      sgOptionalFields3.Cells[0, 9] := 'Eprinttype';
      sgOptionalFields3.Cells[0, 10] := 'Url';
      sgOptionalFields3.Cells[0, 11] := 'Urldate';
    end

    // ********************************************************
    // ********************** ONLINE **************************
    // ********************************************************

    else if LowerCase(cbItemsKind.Text) = 'online' then
    begin
      // Required fields
      sgRequiredFields.Clean;
      sgRequiredFields.RowCount := 7;
      sgRequiredFields.Cells[0, 0] := 'Author';
      sgRequiredFields.Cells[0, 1] := 'Title';
      sgRequiredFields.Cells[0, 2] := 'Editor';
      sgRequiredFields.Cells[0, 3] := 'Year';
      sgRequiredFields.Cells[0, 4] := 'Url';
      sgRequiredFields.Cells[0, 5] := 'Keywords';
      sgRequiredFields.Cells[0, 6] := 'BibTexKey';
      // Optional fields 1
      sgOptionalFields1.Clean;
      sgOptionalFields1.RowCount := 2;
      sgOptionalFields1.Cells[0, 0] := 'Subtitle';
      sgOptionalFields1.Cells[0, 1] := 'Titleaddon';
      // Optional fields 2
      sgOptionalFields2.Clean;
      sgOptionalFields2.RowCount := 2;
      sgOptionalFields2.Cells[0, 0] := 'Language';
      sgOptionalFields2.Cells[0, 1] := 'Version';
      // Optional fields 3
      sgOptionalFields3.Clean;
      sgOptionalFields3.RowCount := 7;
      sgOptionalFields3.Cells[0, 0] := 'Note';
      sgOptionalFields3.Cells[0, 1] := 'Organization';
      sgOptionalFields3.Cells[0, 2] := 'Date';
      sgOptionalFields3.Cells[0, 3] := 'Month';
      sgOptionalFields3.Cells[0, 4] := 'Addendum';
      sgOptionalFields3.Cells[0, 5] := 'Pubstate';
      sgOptionalFields3.Cells[0, 6] := 'Urldate';
    end

    // ********************************************************
    // ********************** PATENT **************************
    // ********************************************************

    else if LowerCase(cbItemsKind.Text) = 'patent' then
    begin
      // Required fields
      sgRequiredFields.Clean;
      sgRequiredFields.RowCount := 6;
      sgRequiredFields.Cells[0, 0] := 'Author';
      sgRequiredFields.Cells[0, 1] := 'Title';
      sgRequiredFields.Cells[0, 2] := 'Number';
      sgRequiredFields.Cells[0, 3] := 'Year';
      sgRequiredFields.Cells[0, 4] := 'Keywords';
      sgRequiredFields.Cells[0, 5] := 'BibTexKey';
      // Optional fields 1
      sgOptionalFields1.Clean;
      sgOptionalFields1.RowCount := 3;
      sgOptionalFields1.Cells[0, 0] := 'Holder';
      sgOptionalFields1.Cells[0, 1] := 'Subtitle';
      sgOptionalFields1.Cells[0, 2] := 'Titleaddon';
      // Optional fields 2
      sgOptionalFields2.Clean;
      sgOptionalFields2.RowCount := 3;
      sgOptionalFields2.Cells[0, 0] := 'Type';
      sgOptionalFields2.Cells[0, 1] := 'Version';
      sgOptionalFields2.Cells[0, 2] := 'Location';
      // Optional fields 3
      sgOptionalFields3.Clean;
      sgOptionalFields3.RowCount := 11;
      sgOptionalFields3.Cells[0, 0] := 'Note';
      sgOptionalFields3.Cells[0, 1] := 'Date';
      sgOptionalFields3.Cells[0, 2] := 'Month';
      sgOptionalFields3.Cells[0, 3] := 'Addendum';
      sgOptionalFields3.Cells[0, 4] := 'Pubstate';
      sgOptionalFields3.Cells[0, 5] := 'Doi';
      sgOptionalFields3.Cells[0, 6] := 'Eprint';
      sgOptionalFields3.Cells[0, 7] := 'Eprintclass';
      sgOptionalFields3.Cells[0, 8] := 'Eprinttype';
      sgOptionalFields3.Cells[0, 9] := 'Url';
      sgOptionalFields3.Cells[0, 10] := 'Urldate';
    end

    // ********************************************************
    // ********************** PERIODICAL **********************
    // ********************************************************

    else if LowerCase(cbItemsKind.Text) = 'periodical' then
    begin
      // Required fields
      sgRequiredFields.Clean;
      sgRequiredFields.RowCount := 6;
      sgRequiredFields.Cells[0, 0] := 'Editor';
      sgRequiredFields.Cells[0, 1] := 'Title';
      sgRequiredFields.Cells[0, 2] := 'Issuetitle';
      sgRequiredFields.Cells[0, 3] := 'Year';
      sgRequiredFields.Cells[0, 4] := 'Keywords';
      sgRequiredFields.Cells[0, 5] := 'BibTexKey';
      // Optional fields 1
      sgOptionalFields1.Clean;
      sgOptionalFields1.RowCount := 5;
      sgOptionalFields1.Cells[0, 0] := 'Editora';
      sgOptionalFields1.Cells[0, 1] := 'Editorb';
      sgOptionalFields1.Cells[0, 2] := 'Editorc';
      sgOptionalFields1.Cells[0, 3] := 'Subtitle';
      sgOptionalFields1.Cells[0, 4] := 'Issuesubtitle';
      // Optional fields 2
      sgOptionalFields2.Clean;
      sgOptionalFields2.RowCount := 7;
      sgOptionalFields2.Cells[0, 0] := 'Language';
      sgOptionalFields2.Cells[0, 1] := 'Series';
      sgOptionalFields2.Cells[0, 2] := 'Volume';
      sgOptionalFields2.Cells[0, 3] := 'Number';
      sgOptionalFields2.Cells[0, 4] := 'Issue';
      sgOptionalFields2.Cells[0, 5] := 'Date';
      sgOptionalFields2.Cells[0, 6] := 'Month';
      // Optional fields 3
      sgOptionalFields3.Clean;
      sgOptionalFields3.RowCount := 10;
      sgOptionalFields3.Cells[0, 0] := 'Note';
      sgOptionalFields3.Cells[0, 1] := 'Issn';
      sgOptionalFields3.Cells[0, 2] := 'Addendum';
      sgOptionalFields3.Cells[0, 3] := 'Pubstate';
      sgOptionalFields3.Cells[0, 4] := 'Doi';
      sgOptionalFields3.Cells[0, 5] := 'Eprint';
      sgOptionalFields3.Cells[0, 6] := 'Eprintclass';
      sgOptionalFields3.Cells[0, 7] := 'Eprinttype';
      sgOptionalFields3.Cells[0, 8] := 'Url';
      sgOptionalFields3.Cells[0, 9] := 'Urldate';
    end

    // ********************************************************
    // ********************** SUPPPERIODICAL ******************
    // ********************************************************

    else if LowerCase(cbItemsKind.Text) = 'suppperiodical' then
    begin
      // Required fields
      sgRequiredFields.Clean;
      sgRequiredFields.RowCount := 14;
      sgRequiredFields.Cells[0, 0] := 'Author';
      sgRequiredFields.Cells[0, 1] := 'Title';
      sgRequiredFields.Cells[0, 2] := 'Bookauthor';
      sgRequiredFields.Cells[0, 3] := 'Booktitle';
      sgRequiredFields.Cells[0, 4] := 'Editor';
      sgRequiredFields.Cells[0, 5] := 'Series';
      sgRequiredFields.Cells[0, 6] := 'Number';
      sgRequiredFields.Cells[0, 7] := 'Publisher';
      sgRequiredFields.Cells[0, 8] := 'Location';
      sgRequiredFields.Cells[0, 9] := 'Year';
      sgRequiredFields.Cells[0, 10] := 'Pages';
      sgRequiredFields.Cells[0, 11] := 'Crossref';
      sgRequiredFields.Cells[0, 12] := 'Keywords';
      sgRequiredFields.Cells[0, 13] := 'BibTexKey';
      // Optional fields 1
      sgOptionalFields1.Clean;
      sgOptionalFields1.RowCount := 16;
      sgOptionalFields1.Cells[0, 0] := 'Editora';
      sgOptionalFields1.Cells[0, 1] := 'Editorb';
      sgOptionalFields1.Cells[0, 2] := 'Editorc';
      sgOptionalFields1.Cells[0, 3] := 'Translator';
      sgOptionalFields1.Cells[0, 4] := 'Annotator';
      sgOptionalFields1.Cells[0, 5] := 'Commentator';
      sgOptionalFields1.Cells[0, 6] := 'Introduction';
      sgOptionalFields1.Cells[0, 7] := 'Foreword';
      sgOptionalFields1.Cells[0, 8] := 'Afterword';
      sgOptionalFields1.Cells[0, 9] := 'Subtitle';
      sgOptionalFields1.Cells[0, 10] := 'Titleaddon';
      sgOptionalFields1.Cells[0, 11] := 'Maintitle';
      sgOptionalFields1.Cells[0, 12] := 'Mainsubtitle';
      sgOptionalFields1.Cells[0, 13] := 'Maintitleaddon';
      sgOptionalFields1.Cells[0, 14] := 'Booksubtitle';
      sgOptionalFields1.Cells[0, 15] := 'Booktitleaddon';
      // Optional fields 2
      sgOptionalFields2.Clean;
      sgOptionalFields2.RowCount := 7;
      sgOptionalFields2.Cells[0, 0] := 'Language';
      sgOptionalFields2.Cells[0, 1] := 'Origlanguage';
      sgOptionalFields2.Cells[0, 2] := 'Volume';
      sgOptionalFields2.Cells[0, 3] := 'Part';
      sgOptionalFields2.Cells[0, 4] := 'Edition';
      sgOptionalFields2.Cells[0, 5] := 'Volumes';
      sgOptionalFields2.Cells[0, 6] := 'Date';
      // Optional fields 3
      sgOptionalFields3.Clean;
      sgOptionalFields3.RowCount := 11;
      sgOptionalFields3.Cells[0, 0] := 'Note';
      sgOptionalFields3.Cells[0, 1] := 'Isbn';
      sgOptionalFields3.Cells[0, 2] := 'Chapter';
      sgOptionalFields3.Cells[0, 3] := 'Addendum';
      sgOptionalFields3.Cells[0, 4] := 'Pubstate';
      sgOptionalFields3.Cells[0, 5] := 'Doi';
      sgOptionalFields3.Cells[0, 6] := 'Eprint';
      sgOptionalFields3.Cells[0, 7] := 'Eprintclass';
      sgOptionalFields3.Cells[0, 8] := 'Eprinttype';
      sgOptionalFields3.Cells[0, 9] := 'Url';
      sgOptionalFields3.Cells[0, 10] := 'Urldate';
    end

    // ********************************************************
    // ********************** PROCEEDINGS *********************
    // ********************************************************

    else if LowerCase(cbItemsKind.Text) = 'proceedings' then
    begin
      // Required fields
      sgRequiredFields.Clean;
      sgRequiredFields.RowCount := 10;
      sgRequiredFields.Cells[0, 0] := 'Editor';
      sgRequiredFields.Cells[0, 1] := 'Title';
      sgRequiredFields.Cells[0, 2] := 'Series';
      sgRequiredFields.Cells[0, 3] := 'Number';
      sgRequiredFields.Cells[0, 4] := 'Publisher';
      sgRequiredFields.Cells[0, 5] := 'Location';
      sgRequiredFields.Cells[0, 6] := 'Year';
      sgRequiredFields.Cells[0, 7] := 'Crossref';
      sgRequiredFields.Cells[0, 8] := 'Keywords';
      sgRequiredFields.Cells[0, 9] := 'BibTexKey';
      // Optional fields 1
      sgOptionalFields1.Clean;
      sgOptionalFields1.RowCount := 9;
      sgOptionalFields1.Cells[0, 0] := 'Subtitle';
      sgOptionalFields1.Cells[0, 1] := 'Titleaddon';
      sgOptionalFields1.Cells[0, 2] := 'Maintitle';
      sgOptionalFields1.Cells[0, 3] := 'Mainsubtitle';
      sgOptionalFields1.Cells[0, 4] := 'Maintitleaddon';
      sgOptionalFields1.Cells[0, 5] := 'Eventtitle';
      sgOptionalFields1.Cells[0, 6] := 'Eventtitleaddon';
      sgOptionalFields1.Cells[0, 7] := 'Eventdate';
      sgOptionalFields1.Cells[0, 8] := 'Venue';
      // Optional fields 2
      sgOptionalFields2.Clean;
      sgOptionalFields2.RowCount := 5;
      sgOptionalFields2.Cells[0, 0] := 'Language';
      sgOptionalFields2.Cells[0, 1] := 'Volume';
      sgOptionalFields2.Cells[0, 2] := 'Part';
      sgOptionalFields2.Cells[0, 3] := 'Volumes';
      sgOptionalFields2.Cells[0, 4] := 'Date';
      // Optional fields 3
      sgOptionalFields3.Clean;
      sgOptionalFields3.RowCount := 15;
      sgOptionalFields3.Cells[0, 0] := 'Note';
      sgOptionalFields3.Cells[0, 1] := 'Organization';
      sgOptionalFields3.Cells[0, 2] := 'Month';
      sgOptionalFields3.Cells[0, 3] := 'Isbn';
      sgOptionalFields3.Cells[0, 4] := 'Chapter';
      sgOptionalFields3.Cells[0, 5] := 'Pages';
      sgOptionalFields3.Cells[0, 6] := 'Pagetotal';
      sgOptionalFields3.Cells[0, 7] := 'Addendum';
      sgOptionalFields3.Cells[0, 8] := 'Pubstate';
      sgOptionalFields3.Cells[0, 9] := 'Doi';
      sgOptionalFields3.Cells[0, 10] := 'Eprint';
      sgOptionalFields3.Cells[0, 11] := 'Eprintclass';
      sgOptionalFields3.Cells[0, 12] := 'Eprinttype';
      sgOptionalFields3.Cells[0, 13] := 'Url';
      sgOptionalFields3.Cells[0, 14] := 'Urldate';
    end

    // ********************************************************
    // ********************** MVPROCEEDINGS ******************
    // ********************************************************

    else if LowerCase(cbItemsKind.Text) = 'mvproceedings' then
    begin
      // Required fields
      sgRequiredFields.Clean;
      sgRequiredFields.RowCount := 9;
      sgRequiredFields.Cells[0, 0] := 'Editor';
      sgRequiredFields.Cells[0, 1] := 'Title';
      sgRequiredFields.Cells[0, 2] := 'Series';
      sgRequiredFields.Cells[0, 3] := 'Number';
      sgRequiredFields.Cells[0, 4] := 'Publisher';
      sgRequiredFields.Cells[0, 5] := 'Location';
      sgRequiredFields.Cells[0, 6] := 'Year';
      sgRequiredFields.Cells[0, 7] := 'Keywords';
      sgRequiredFields.Cells[0, 8] := 'BibTexKey';
      // Optional fields 1
      sgOptionalFields1.Clean;
      sgOptionalFields1.RowCount := 6;
      sgOptionalFields1.Cells[0, 0] := 'Subtitle';
      sgOptionalFields1.Cells[0, 1] := 'Titleaddon';
      sgOptionalFields1.Cells[0, 2] := 'Eventtitle';
      sgOptionalFields1.Cells[0, 3] := 'Eventtitleaddon';
      sgOptionalFields1.Cells[0, 4] := 'Eventdate';
      sgOptionalFields1.Cells[0, 5] := 'Venue';
      // Optional fields 2
      sgOptionalFields2.Clean;
      sgOptionalFields2.RowCount := 3;
      sgOptionalFields2.Cells[0, 0] := 'Language';
      sgOptionalFields2.Cells[0, 1] := 'Volumes';
      sgOptionalFields2.Cells[0, 2] := 'Date';
      // Optional fields 3
      sgOptionalFields3.Clean;
      sgOptionalFields3.RowCount := 13;
      sgOptionalFields3.Cells[0, 0] := 'Note';
      sgOptionalFields3.Cells[0, 1] := 'Organization';
      sgOptionalFields3.Cells[0, 2] := 'Month';
      sgOptionalFields3.Cells[0, 3] := 'Isbn';
      sgOptionalFields3.Cells[0, 4] := 'Pagetotal';
      sgOptionalFields3.Cells[0, 5] := 'Addendum';
      sgOptionalFields3.Cells[0, 6] := 'Pubstate';
      sgOptionalFields3.Cells[0, 7] := 'Doi';
      sgOptionalFields3.Cells[0, 8] := 'Eprint';
      sgOptionalFields3.Cells[0, 9] := 'Eprintclass';
      sgOptionalFields3.Cells[0, 10] := 'Eprinttype';
      sgOptionalFields3.Cells[0, 11] := 'Url';
      sgOptionalFields3.Cells[0, 12] := 'Urldate';
    end

    // ********************************************************
    // ********************** INPROCEEDINGS *******************
    // ********************************************************

    else if LowerCase(cbItemsKind.Text) = 'inproceedings' then
    begin
      // Required fields
      sgRequiredFields.Clean;
      sgRequiredFields.RowCount := 13;
      sgRequiredFields.Cells[0, 0] := 'Author';
      sgRequiredFields.Cells[0, 1] := 'Editor';
      sgRequiredFields.Cells[0, 2] := 'Title';
      sgRequiredFields.Cells[0, 3] := 'Booktitle';
      sgRequiredFields.Cells[0, 4] := 'Series';
      sgRequiredFields.Cells[0, 5] := 'Number';
      sgRequiredFields.Cells[0, 6] := 'Publisher';
      sgRequiredFields.Cells[0, 7] := 'Location';
      sgRequiredFields.Cells[0, 8] := 'Year';
      sgRequiredFields.Cells[0, 9] := 'Pages';
      sgRequiredFields.Cells[0, 10] := 'Crossref';
      sgRequiredFields.Cells[0, 11] := 'Keywords';
      sgRequiredFields.Cells[0, 12] := 'BibTexKey';
      // Optional fields 1
      sgOptionalFields1.Clean;
      sgOptionalFields1.RowCount := 11;
      sgOptionalFields1.Cells[0, 0] := 'Subtitle';
      sgOptionalFields1.Cells[0, 1] := 'Titleaddon';
      sgOptionalFields1.Cells[0, 2] := 'Maintitle';
      sgOptionalFields1.Cells[0, 3] := 'Mainsubtitle';
      sgOptionalFields1.Cells[0, 4] := 'Maintitleaddon';
      sgOptionalFields1.Cells[0, 5] := 'Booksubtitle';
      sgOptionalFields1.Cells[0, 6] := 'Booktitleaddon';
      sgOptionalFields1.Cells[0, 7] := 'Eventtitle';
      sgOptionalFields1.Cells[0, 8] := 'Eventtitleaddon';
      sgOptionalFields1.Cells[0, 9] := 'Eventdate';
      sgOptionalFields1.Cells[0, 10] := 'Venue';
      // Optional fields 2
      sgOptionalFields2.Clean;
      sgOptionalFields2.RowCount := 5;
      sgOptionalFields2.Cells[0, 0] := 'Language';
      sgOptionalFields2.Cells[0, 1] := 'Volume';
      sgOptionalFields2.Cells[0, 2] := 'Part';
      sgOptionalFields2.Cells[0, 3] := 'Volumes';
      sgOptionalFields2.Cells[0, 4] := 'Date';
      // Optional fields 3
      sgOptionalFields3.Clean;
      sgOptionalFields3.RowCount := 13;
      sgOptionalFields3.Cells[0, 0] := 'Note';
      sgOptionalFields3.Cells[0, 1] := 'Organization';
      sgOptionalFields3.Cells[0, 2] := 'Month';
      sgOptionalFields3.Cells[0, 3] := 'Isbn';
      sgOptionalFields3.Cells[0, 4] := 'Chapter';
      sgOptionalFields3.Cells[0, 5] := 'Addendum';
      sgOptionalFields3.Cells[0, 6] := 'Pubstate';
      sgOptionalFields3.Cells[0, 7] := 'Doi';
      sgOptionalFields3.Cells[0, 8] := 'Eprint';
      sgOptionalFields3.Cells[0, 9] := 'Eprintclass';
      sgOptionalFields3.Cells[0, 10] := 'Eprinttype';
      sgOptionalFields3.Cells[0, 11] := 'Url';
      sgOptionalFields3.Cells[0, 12] := 'Urldate';
    end

    // ********************************************************
    // ********************** REFERENCE ***********************
    // ********************************************************

    else if LowerCase(cbItemsKind.Text) = 'reference' then
    begin
      // Required fields
      sgRequiredFields.Clean;
      sgRequiredFields.RowCount := 10;
      sgRequiredFields.Cells[0, 0] := 'Editor';
      sgRequiredFields.Cells[0, 1] := 'Title';
      sgRequiredFields.Cells[0, 2] := 'Series';
      sgRequiredFields.Cells[0, 3] := 'Number';
      sgRequiredFields.Cells[0, 4] := 'Publisher';
      sgRequiredFields.Cells[0, 5] := 'Location';
      sgRequiredFields.Cells[0, 6] := 'Year';
      sgRequiredFields.Cells[0, 7] := 'Crossref';
      sgRequiredFields.Cells[0, 8] := 'Keywords';
      sgRequiredFields.Cells[0, 9] := 'BibTexKey';
      // Optional fields 1
      sgOptionalFields1.Clean;
      sgOptionalFields1.RowCount := 14;
      sgOptionalFields1.Cells[0, 0] := 'Editora';
      sgOptionalFields1.Cells[0, 1] := 'Editorb';
      sgOptionalFields1.Cells[0, 2] := 'Editorc';
      sgOptionalFields1.Cells[0, 3] := 'Translator';
      sgOptionalFields1.Cells[0, 4] := 'Annotator';
      sgOptionalFields1.Cells[0, 5] := 'Commentator';
      sgOptionalFields1.Cells[0, 6] := 'Introduction';
      sgOptionalFields1.Cells[0, 7] := 'Foreword';
      sgOptionalFields1.Cells[0, 8] := 'Afterword';
      sgOptionalFields1.Cells[0, 9] := 'Subtitle';
      sgOptionalFields1.Cells[0, 10] := 'Titleaddon';
      sgOptionalFields1.Cells[0, 11] := 'Maintitle';
      sgOptionalFields1.Cells[0, 12] := 'Mainsubtitle';
      sgOptionalFields1.Cells[0, 13] := 'Maintitleaddon';
      // Optional fields 2
      sgOptionalFields2.Clean;
      sgOptionalFields2.RowCount := 7;
      sgOptionalFields2.Cells[0, 0] := 'Language';
      sgOptionalFields2.Cells[0, 1] := 'Origlanguage';
      sgOptionalFields2.Cells[0, 2] := 'Volume';
      sgOptionalFields2.Cells[0, 3] := 'Part';
      sgOptionalFields2.Cells[0, 4] := 'Edition';
      sgOptionalFields2.Cells[0, 5] := 'Volumes';
      sgOptionalFields2.Cells[0, 6] := 'Date';
      // Optional fields 3
      sgOptionalFields3.Clean;
      sgOptionalFields3.RowCount := 13;
      sgOptionalFields3.Cells[0, 0] := 'Note';
      sgOptionalFields3.Cells[0, 1] := 'Isbn';
      sgOptionalFields3.Cells[0, 2] := 'Chapter';
      sgOptionalFields3.Cells[0, 3] := 'Pages';
      sgOptionalFields3.Cells[0, 4] := 'Pagetotal';
      sgOptionalFields3.Cells[0, 5] := 'Addendum';
      sgOptionalFields3.Cells[0, 6] := 'Pubstate';
      sgOptionalFields3.Cells[0, 7] := 'Doi';
      sgOptionalFields3.Cells[0, 8] := 'Eprint';
      sgOptionalFields3.Cells[0, 9] := 'Eprintclass';
      sgOptionalFields3.Cells[0, 10] := 'Eprinttype';
      sgOptionalFields3.Cells[0, 11] := 'Url';
      sgOptionalFields3.Cells[0, 12] := 'Urldate';
    end

    // ********************************************************
    // ********************** MVREFERENCE *********************
    // ********************************************************

    else if LowerCase(cbItemsKind.Text) = 'mvreference' then
    begin
      // Required fields
      sgRequiredFields.Clean;
      sgRequiredFields.RowCount := 9;
      sgRequiredFields.Cells[0, 0] := 'Editor';
      sgRequiredFields.Cells[0, 1] := 'Title';
      sgRequiredFields.Cells[0, 2] := 'Series';
      sgRequiredFields.Cells[0, 3] := 'Number';
      sgRequiredFields.Cells[0, 4] := 'Publisher';
      sgRequiredFields.Cells[0, 5] := 'Location';
      sgRequiredFields.Cells[0, 6] := 'Year';
      sgRequiredFields.Cells[0, 7] := 'Keywords';
      sgRequiredFields.Cells[0, 8] := 'BibTexKey';
      // Optional fields 1
      sgOptionalFields1.Clean;
      sgOptionalFields1.RowCount := 11;
      sgOptionalFields1.Cells[0, 0] := 'Editora';
      sgOptionalFields1.Cells[0, 1] := 'Editorb';
      sgOptionalFields1.Cells[0, 2] := 'Editorc';
      sgOptionalFields1.Cells[0, 3] := 'Translator';
      sgOptionalFields1.Cells[0, 4] := 'Annotator';
      sgOptionalFields1.Cells[0, 5] := 'Commentator';
      sgOptionalFields1.Cells[0, 6] := 'Introduction';
      sgOptionalFields1.Cells[0, 7] := 'Foreword';
      sgOptionalFields1.Cells[0, 8] := 'Afterword';
      sgOptionalFields1.Cells[0, 9] := 'Subtitle';
      sgOptionalFields1.Cells[0, 10] := 'Titleaddon';
      // Optional fields 2
      sgOptionalFields2.Clean;
      sgOptionalFields2.RowCount := 5;
      sgOptionalFields2.Cells[0, 0] := 'Language';
      sgOptionalFields2.Cells[0, 1] := 'Origlanguage';
      sgOptionalFields2.Cells[0, 2] := 'Edition';
      sgOptionalFields2.Cells[0, 3] := 'Volumes';
      sgOptionalFields2.Cells[0, 4] := 'Date';
      // Optional fields 3
      sgOptionalFields3.Clean;
      sgOptionalFields3.RowCount := 11;
      sgOptionalFields3.Cells[0, 0] := 'Note';
      sgOptionalFields3.Cells[0, 1] := 'Isbn';
      sgOptionalFields3.Cells[0, 2] := 'Pagetotal';
      sgOptionalFields3.Cells[0, 3] := 'Addendum';
      sgOptionalFields3.Cells[0, 4] := 'Pubstate';
      sgOptionalFields3.Cells[0, 5] := 'Doi';
      sgOptionalFields3.Cells[0, 6] := 'Eprint';
      sgOptionalFields3.Cells[0, 7] := 'Eprintclass';
      sgOptionalFields3.Cells[0, 8] := 'Eprinttype';
      sgOptionalFields3.Cells[0, 9] := 'Url';
      sgOptionalFields3.Cells[0, 10] := 'Urldate';
    end

    // ********************************************************
    // ********************** INREFERENCE *********************
    // ********************************************************

    else if LowerCase(cbItemsKind.Text) = 'inreference' then
    begin
      // Required fields
      sgRequiredFields.Clean;
      sgRequiredFields.RowCount := 13;
      sgRequiredFields.Cells[0, 0] := 'Author';
      sgRequiredFields.Cells[0, 1] := 'Editor';
      sgRequiredFields.Cells[0, 2] := 'Title';
      sgRequiredFields.Cells[0, 3] := 'Booktitle';
      sgRequiredFields.Cells[0, 4] := 'Series';
      sgRequiredFields.Cells[0, 5] := 'Number';
      sgRequiredFields.Cells[0, 6] := 'Publisher';
      sgRequiredFields.Cells[0, 7] := 'Location';
      sgRequiredFields.Cells[0, 8] := 'Year';
      sgRequiredFields.Cells[0, 9] := 'Pages';
      sgRequiredFields.Cells[0, 10] := 'Crossref';
      sgRequiredFields.Cells[0, 11] := 'Keywords';
      sgRequiredFields.Cells[0, 12] := 'BibTexKey';
      // Optional fields 1
      sgOptionalFields1.Clean;
      sgOptionalFields1.RowCount := 16;
      sgOptionalFields1.Cells[0, 0] := 'Editora';
      sgOptionalFields1.Cells[0, 1] := 'Editorb';
      sgOptionalFields1.Cells[0, 2] := 'Editorc';
      sgOptionalFields1.Cells[0, 3] := 'Translator';
      sgOptionalFields1.Cells[0, 4] := 'Annotator';
      sgOptionalFields1.Cells[0, 5] := 'Commentator';
      sgOptionalFields1.Cells[0, 6] := 'Introduction';
      sgOptionalFields1.Cells[0, 7] := 'Foreword';
      sgOptionalFields1.Cells[0, 8] := 'Afterword';
      sgOptionalFields1.Cells[0, 9] := 'Subtitle';
      sgOptionalFields1.Cells[0, 10] := 'Titleaddon';
      sgOptionalFields1.Cells[0, 11] := 'Maintitle';
      sgOptionalFields1.Cells[0, 12] := 'Mainsubtitle';
      sgOptionalFields1.Cells[0, 13] := 'Maintitleaddon';
      sgOptionalFields1.Cells[0, 14] := 'Booksubtitle';
      sgOptionalFields1.Cells[0, 15] := 'Booktitleaddon';
      // Optional fields 2
      sgOptionalFields2.Clean;
      sgOptionalFields2.RowCount := 7;
      sgOptionalFields2.Cells[0, 0] := 'Language';
      sgOptionalFields2.Cells[0, 1] := 'Origlanguage';
      sgOptionalFields2.Cells[0, 2] := 'Volume';
      sgOptionalFields2.Cells[0, 3] := 'Part';
      sgOptionalFields2.Cells[0, 4] := 'Edition';
      sgOptionalFields2.Cells[0, 5] := 'Volumes';
      sgOptionalFields2.Cells[0, 6] := 'Date';
      // Optional fields 3
      sgOptionalFields3.Clean;
      sgOptionalFields3.RowCount := 11;
      sgOptionalFields3.Cells[0, 0] := 'Note';
      sgOptionalFields3.Cells[0, 1] := 'Isbn';
      sgOptionalFields3.Cells[0, 2] := 'Chapter';
      sgOptionalFields3.Cells[0, 3] := 'Addendum';
      sgOptionalFields3.Cells[0, 4] := 'Pubstate';
      sgOptionalFields3.Cells[0, 5] := 'Doi';
      sgOptionalFields3.Cells[0, 6] := 'Eprint';
      sgOptionalFields3.Cells[0, 7] := 'Eprintclass';
      sgOptionalFields3.Cells[0, 8] := 'Eprinttype';
      sgOptionalFields3.Cells[0, 9] := 'Url';
      sgOptionalFields3.Cells[0, 10] := 'Urldate';
    end

    // ********************************************************
    // ********************** REPORT **************************
    // ********************************************************

    else if LowerCase(cbItemsKind.Text) = 'report' then
    begin
      // Required fields
      sgRequiredFields.Clean;
      sgRequiredFields.RowCount := 8;
      sgRequiredFields.Cells[0, 0] := 'Author';
      sgRequiredFields.Cells[0, 1] := 'Title';
      sgRequiredFields.Cells[0, 2] := 'Type';
      sgRequiredFields.Cells[0, 3] := 'Institution';
      sgRequiredFields.Cells[0, 4] := 'Location';
      sgRequiredFields.Cells[0, 5] := 'Year';
      sgRequiredFields.Cells[0, 6] := 'Keywords';
      sgRequiredFields.Cells[0, 7] := 'BibTexKey';
      // Optional fields 1
      sgOptionalFields1.Clean;
      sgOptionalFields1.RowCount := 2;
      sgOptionalFields1.Cells[0, 0] := 'Subtitle';
      sgOptionalFields1.Cells[0, 1] := 'Titleaddon';
      // Optional fields 2
      sgOptionalFields2.Clean;
      sgOptionalFields2.RowCount := 4;
      sgOptionalFields2.Cells[0, 0] := 'Language';
      sgOptionalFields2.Cells[0, 1] := 'Number';
      sgOptionalFields2.Cells[0, 2] := 'Version';
      sgOptionalFields2.Cells[0, 3] := 'Date';
      // Optional fields 3
      sgOptionalFields3.Clean;
      sgOptionalFields3.RowCount := 14;
      sgOptionalFields3.Cells[0, 0] := 'Note';
      sgOptionalFields3.Cells[0, 1] := 'Month';
      sgOptionalFields3.Cells[0, 2] := 'Isrn';
      sgOptionalFields3.Cells[0, 3] := 'Chapter';
      sgOptionalFields3.Cells[0, 4] := 'Pages';
      sgOptionalFields3.Cells[0, 5] := 'Pagetotal';
      sgOptionalFields3.Cells[0, 6] := 'Addendum';
      sgOptionalFields3.Cells[0, 7] := 'Pubstate';
      sgOptionalFields3.Cells[0, 8] := 'Doi';
      sgOptionalFields3.Cells[0, 9] := 'Eprint';
      sgOptionalFields3.Cells[0, 10] := 'Eprintclass';
      sgOptionalFields3.Cells[0, 11] := 'Eprinttype';
      sgOptionalFields3.Cells[0, 12] := 'Url';
      sgOptionalFields3.Cells[0, 13] := 'Urldate';
    end

    // ********************************************************
    // ********************** THESIS **************************
    // ********************************************************

    else if LowerCase(cbItemsKind.Text) = 'thesis' then
    begin
      // Required fields
      sgRequiredFields.Clean;
      sgRequiredFields.RowCount := 8;
      sgRequiredFields.Cells[0, 0] := 'Author';
      sgRequiredFields.Cells[0, 1] := 'Title';
      sgRequiredFields.Cells[0, 2] := 'Type';
      sgRequiredFields.Cells[0, 3] := 'Institution';
      sgRequiredFields.Cells[0, 4] := 'Location';
      sgRequiredFields.Cells[0, 5] := 'Year';
      sgRequiredFields.Cells[0, 6] := 'Keywords';
      sgRequiredFields.Cells[0, 7] := 'BibTexKey';
      // Optional fields 1
      sgOptionalFields1.Clean;
      sgOptionalFields1.RowCount := 2;
      sgOptionalFields1.Cells[0, 0] := 'Subtitle';
      sgOptionalFields1.Cells[0, 1] := 'Titleaddon';
      // Optional fields 2
      sgOptionalFields2.Clean;
      sgOptionalFields2.RowCount := 2;
      sgOptionalFields2.Cells[0, 0] := 'Language';
      sgOptionalFields2.Cells[0, 1] := 'Date';
      // Optional fields 3
      sgOptionalFields3.Clean;
      sgOptionalFields3.RowCount := 14;
      sgOptionalFields3.Cells[0, 0] := 'Note';
      sgOptionalFields3.Cells[0, 1] := 'Month';
      sgOptionalFields3.Cells[0, 2] := 'Isbn';
      sgOptionalFields3.Cells[0, 3] := 'Chapter';
      sgOptionalFields3.Cells[0, 4] := 'Pages';
      sgOptionalFields3.Cells[0, 5] := 'Pagetotal';
      sgOptionalFields3.Cells[0, 6] := 'Addendum';
      sgOptionalFields3.Cells[0, 7] := 'Pubstate';
      sgOptionalFields3.Cells[0, 8] := 'Doi';
      sgOptionalFields3.Cells[0, 9] := 'Eprint';
      sgOptionalFields3.Cells[0, 10] := 'Eprintclass';
      sgOptionalFields3.Cells[0, 11] := 'Eprinttype';
      sgOptionalFields3.Cells[0, 12] := 'Url';
      sgOptionalFields3.Cells[0, 13] := 'Urldate';
    end

    // ********************************************************
    // ********************** UNPUBLISHED *********************
    // ********************************************************

    else if LowerCase(cbItemsKind.Text) = 'unpublished' then
    begin
      // Required fields
      sgRequiredFields.Clean;
      sgRequiredFields.RowCount := 6;
      sgRequiredFields.Cells[0, 0] := 'Author';
      sgRequiredFields.Cells[0, 1] := 'Title';
      sgRequiredFields.Cells[0, 2] := 'Howpublished';
      sgRequiredFields.Cells[0, 3] := 'Year';
      sgRequiredFields.Cells[0, 4] := 'Keywords';
      sgRequiredFields.Cells[0, 5] := 'BibTexKey';
      // Optional fields 1
      sgOptionalFields1.Clean;
      sgOptionalFields1.RowCount := 2;
      sgOptionalFields1.Cells[0, 0] := 'Subtitle';
      sgOptionalFields1.Cells[0, 1] := 'Titleaddon';
      // Optional fields 2
      sgOptionalFields2.Clean;
      sgOptionalFields2.RowCount := 1;
      sgOptionalFields2.Cells[0, 0] := 'Language';
      // Optional fields 3
      sgOptionalFields3.Clean;
      sgOptionalFields3.RowCount := 8;
      sgOptionalFields3.Cells[0, 0] := 'Note';
      sgOptionalFields3.Cells[0, 1] := 'Isbn';
      sgOptionalFields3.Cells[0, 2] := 'Date';
      sgOptionalFields3.Cells[0, 3] := 'Month';
      sgOptionalFields3.Cells[0, 4] := 'Addendum';
      sgOptionalFields3.Cells[0, 5] := 'Pubstate';
      sgOptionalFields3.Cells[0, 6] := 'Url';
      sgOptionalFields3.Cells[0, 7] := 'Urldate';
    end;
    // To avoid a selection of the bibtex field in the sgRequiredField grid
    // while switching from a type with the last row selected to a type
    // with less rows; this that would allow to change the bibtex fild.
    sgRequiredFields.Row := 0;
    sgOptionalFields1.Row := 0;
    sgOptionalFields2.Row := 0;
    sgOptionalFields3.Row := 0;
  end;
end;

end.
