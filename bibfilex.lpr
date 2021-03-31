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

program bibfilex;

{$mode objfpc}{$H+}

uses {$IFDEF UNIX} {$IFDEF UseCThreads}
  cthreads, {$ENDIF} {$ENDIF}
  Interfaces, // this includes the LCL widgetset
  Forms,
  sqlite3laz,
  Unit1,
  UnitSetFields,
  Copyright,
  Unit2,
  Unit3, Unit4, Unit5, Unit6, Unit7;

{$R *.res}

begin
  Application.Title:='Bibfilex';
  Application.Initialize;
  Application.CreateForm(TfmMain, fmMain);
  Application.CreateForm(TfmCopyright, fmCopyright);
  Application.CreateForm(TfmFilters, fmFilters);
  Application.CreateForm(TfmOptions, fmOptions);
  Application.CreateForm(TfmStopList, fmStopList);
  Application.CreateForm(TfmChar, fmChar);
  Application.CreateForm(TfmKeywords, fmKeywords);
  Application.CreateForm(TfmCheckDoubles, fmCheckDoubles);
  Application.Run;
end.
