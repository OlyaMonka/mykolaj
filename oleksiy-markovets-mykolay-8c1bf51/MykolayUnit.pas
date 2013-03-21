unit MykolayUnit;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, xmldom, XMLIntf, Grids, msxmldom, XMLDoc, StdCtrls, ShellAPI,
   ExtCtrls, FileCtrl, Outline, DirOutln,
  CustomizeDlg, ExtDlgs, ShlObj, XPMan, Menus,  xercesxmldom;

type
  TMykolayMainForm = class(TForm)
    DataBaseXMLDocument: TXMLDocument;
    DataBaseStringGrid: TStringGrid;
    AddButton: TButton;
    DeleteButton: TButton;
    EditButton: TButton;
    SearchNameLabeledEdit: TLabeledEdit;
    SearchAddressLabeledEdit: TLabeledEdit;
    SexLabel: TLabel;
    SearchDistrictLabeledEdit: TLabeledEdit;
    SearchGuardiansAndParensLabeledEdit: TLabeledEdit;
    SearchNotesLabeledEdit: TLabeledEdit;
    SearchSexComboBox: TComboBox;
    SearchNumOfEvelopeLabeledEdit: TLabeledEdit;
    SearchYearOfBirthLabeledEdit: TLabeledEdit;
    SearchAgeLabeledEdit: TLabeledEdit;
    SearchWhiteListComboBox: TComboBox;
    OpenDialog: TOpenDialog;
    BackGroundImage: TImage;
    LogoImage: TImage;
    XPManifest1: TXPManifest;
    MainMenu: TMainMenu;
    N1: TMenuItem;
    N2: TMenuItem;
    N3: TMenuItem;
    N4: TMenuItem;
    N5: TMenuItem;
    N6: TMenuItem;
    N7: TMenuItem;
    N8: TMenuItem;
    N9: TMenuItem;
    N10: TMenuItem;
    N11: TMenuItem;
    N12: TMenuItem;
    N13: TMenuItem;
    N14: TMenuItem;
    N15: TMenuItem;
    ClearSearchButton: TButton;
    SearchStatusComboBox: TComboBox;
    StatusLabel: TLabel;
    CloneButton: TButton;
    N16: TMenuItem;
    N17: TMenuItem;
    N18: TMenuItem;
    N19: TMenuItem;
    N20: TMenuItem;
    procedure FormCreate(Sender: TObject);
    procedure LoadAllBases(dir: string);
    procedure LoadDataBase(WL: boolean);
    procedure FormResize(Sender: TObject);
    procedure DataBaseStringGridDblClick(Sender: TObject);
    procedure AddButtonClick(Sender: TObject);
    procedure DeleteButtonClick(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure FillTable();
    procedure ShowData();
    procedure MakeIDArray();
    procedure Sort;
    procedure SortInterval(start: integer; finish: integer);
    procedure MakeDataBase(WL: boolean);
    procedure DataBaseStringGridMouseDown(Sender: TObject;
      Button: TMouseButton; Shift: TShiftState; X, Y: integer);
    procedure SearchNameLabeledEditChange(Sender: TObject);
    procedure EditButtonClick(Sender: TObject);
    procedure SaveButtonClick(Sender: TObject);
    procedure MakeReserveCopyButtonClick(Sender: TObject);
    procedure SaveAllBases(dir: string);
    procedure LoadReserveCopyButtonClick(Sender: TObject);
    procedure SaveAllButtonClick(Sender: TObject);
    procedure AddFromFileButtonClick(Sender: TObject);
    procedure ColumsViewModeButtonClick(Sender: TObject);
    procedure DataBaseStringGridMouseUp(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: integer);
    procedure ClearSearchButtonClick(Sender: TObject);
    procedure DataBaseStringGridDrawCell(Sender: TObject; ACol, ARow: integer;
      Rect: TRect; State: TGridDrawState);
    procedure CloneButtonClick(Sender: TObject);
    procedure N17Click(Sender: TObject);
    procedure DataBaseStringGridKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure FormCloseQuery(Sender: TObject; var CanClose: Boolean);
    procedure N20Click(Sender: TObject);
    procedure N21Click(Sender: TObject);

  private
    { Private declarations }
  public
    function lessString(s1: string; s2: string): boolean;

  end;

var
  MykolayMainForm: TMykolayMainForm;

type
  TSex = (notAssigned, boy, girl);

  TColumInfo = record
    Num: integer;
    Width: integer;
    Key: string;
    Caption: string;
  end;

  TChildren = record
    Name: String;
    Address: string;
    District: string;
    Sex: TSex;
    YearOfBirth: integer;
    GuardiansParents: string;
    Status: string;
    Notes: string;
    NumOfEvelope: integer;
    NeedsHelp: boolean;
  end;

  TSort = record
    colum: integer;
    sortType: (nosort, up, down);
  end;

const
  notAssignedString = '?';
  girlString = 'дівчинка';
  boyString = 'хлопчик';
  BoysAndGirlsString = 'хлопчики і дівчатка';
  AllString = 'всі';
  NeedsHelpString = 'потребують допомоги';
  dontNeedsHelpString = 'не потребують допомоги';
  GString = 'д';
  BString = 'х';
  YesString = 'так';
  NoString = 'ні';

  defDataBaseFile = 'MykolayLviv.xml';
  defDataBaseBlackListFile = 'MykolayLvivBL.xml';
  keyMykolay = 'MYKOLAY';
  keyChildren = 'CHILDREN';
  keyName = 'NAME';
  keyAddress = 'ADDRESS';
  keyDistrict = 'DISTRICT';
  keySex = 'SEX';
  keyYearOfBirth = 'YEAR_OF_BIRTH';
  keyGuardiansParents = 'GUARDIANS_PARENTS';
  keyStatus = 'STATUS';
  keyNotes = 'NOTES';
  keyNumOfEvelope = 'NUM_OF_EVELOPE';
  keyAge = 'AGE';
  keyNeedsHelp = 'NEEDS_HELP';
  keyFamily = 'FAMILY';
  keyChildrenCount = 'CHILDREN_COUNT';

  cinName = 0;
  cinAddress = 1;
  cinDistrict = 2;
  cinSex = 3;
  cinYearOfBirth = 4;
  cinGuardiansParents = 5;
  cinStatus = 6;
  cinNotes = 7;
  cinNumOfEvelope = 8;
  cinAge = 9;
  cinNeedsHelp = 10;

  ColumCount = 11;

var
  SortedBy: TSort = (colum: - 1; sortType: nosort);
  AppDir: string;

  needRefresh: boolean = true;
  ShowColumCount: integer = 9;
  ColumInfo: array [0 .. ColumCount - 1] of TColumInfo = (
    (
      Num: 0; Width: 140; Key: keyName; Caption: 'Прізвище та Ім''я'),
    (Num: 1; Width: 200; Key: keyAddress; Caption: 'Адреса'), (Num: - 1;
      Width: 70; Key: keyDistrict; Caption: 'Район'), (Num: 2; Width: 60;
      Key: keySex; Caption: 'Стать'), (Num: 3; Width: 96; Key: keyYearOfBirth;
      Caption: 'Рік народження'), (Num: 4; Width: 270;
      Key: keyGuardiansParents; Caption: 'Батьки та опікуни'),
    (Num: 5; Width: 70; Key: keyStatus; Caption: 'Статус'), (Num: 6;
      Width: 100; Key: keyNotes; Caption: 'Нотатки'), (Num: 7; Width: 150;
      Key: keyNumOfEvelope; Caption: 'Номер конверта'), (Num: 8; Width: 50;
      Key: keyAge; Caption: 'Вік'), (Num: - 1; Width: 130; Key: keyNeedsHelp;
      Caption: 'Потребують допомоги'));

  curentYear: integer = 2008;
  IDArray: array of integer;
  IDArrayCount: integer = 0;
  DataBaseCount: integer = 0;
  DataBase: array of TChildren;
  DataBaseChanged: boolean;

function StrToIntEx(s: string): integer;
function IntToStrEx(i: integer): string;
function UpperCaseEx(s: string): string;
function GetFolderDialog(Handle: integer; Caption: string;
  var strFolder: string): boolean;
function EqualAddress(x:integer; y: integer; ignoreAppartment:boolean = false):boolean;
function RemoveAppartmentFromAddres(s:string):string;


implementation

uses EditUnit, ColumsViewModeUnit, StatisticUnit, UnitAbout;
{$R *.dfm}
function RemoveAppartmentFromAddres(s:string):string;
var
  tmp:string;
  x:integer;
begin
    tmp:=s;
    x:=Pos('/',tmp);
    while x<>0 do
    begin
      Delete(tmp,1,x);
      x:=Pos('/',tmp);
    end;
    x:=Pos(tmp,s);
    if x <> 0 then
      Delete(s,x-1,Length(s));
    Result:=s;
end;
function EqualAddress(x:integer; y: integer; ignoreAppartment:boolean):boolean;
  var
    s1,s2 : string;
  begin
    s1:=DataBase[x].Address;
    s2:=DataBase[y].Address;
    while Pos('  ',s1)<>0 do
      Delete(s1,Pos('  ',s1),1);
    while Pos('  ',s2)<>0 do
      Delete(s2,Pos('  ',s2),1);
    if s1[1]=' ' then
      Delete(s1,1,1);
    if s2[1]=' ' then
      Delete(s2,1,1);
    if s1[Length(s1)]=' ' then
      Delete(s1,Length(s1),1);
    if s2[Length(s2)]=' ' then
      Delete(s2,Length(s2),1);
    if ignoreAppartment then
    begin
      s1:=RemoveAppartmentFromAddres(s1);
      s2:=RemoveAppartmentFromAddres(s2);
    end;

    Result:=s1=s2;
  end;

function UpperCaseEx(s: string): string;
var
  i: integer;
  rez: string;
begin
  rez := '';
  for i := 1 to length(s) do
  begin
    if (ord(s[i]) >= 1072) and (ord(s[i]) <= 1103) then
      rez := rez + Chr(ord(s[i]) - (1072 - 1040))
    else
    begin
      case s[i] of
        'є':
          rez := rez + 'Є';
        'і':
          rez := rez + 'І';
        'ї':
          rez := rez + 'Ї';
        'ґ':
          rez := rez + 'Ґ';
      else
        rez := rez + s[i];
      end;
    end;
  end;
  Result := UpperCase(rez);
end;

function StrToIntEx(s: string): integer;
begin
  if (UpperCaseEx(s) = UpperCaseEx(notAssignedString)) or (s = '') then
    Result := -1
  else
    Result := StrToInt(s);
end;

function IntToStrEx(i: integer): string;
begin
  if i < 0 then
    Result := notAssignedString
  else
    Result := IntToStr(i);
end;

function IncludeString(subStr: string; str: string): boolean;
var
  curentS: string;
  start, finish: integer;
  posAp: integer;
begin
  while (subStr <> '') and (subStr[1] = ' ') do
    Delete(subStr, 1, 1);
  if subStr = '' then
    Result := true
  else
  begin
    if subStr[1] <> '=' then
    begin
      start := Pos('"', subStr);
      if start <> 0 then
        subStr[start] := ' ';
      finish := Pos('"', subStr);
      if finish <> 0 then
        subStr[finish] := ' '
      else
        finish := length(subStr) + 1;

      Result := true;
      while (start <> 0) and (finish <> 0) do
      begin
        curentS := Copy(subStr, start + 1, finish - start - 1);
        if (curentS <> '') and (Pos(UpperCaseEx(curentS), UpperCaseEx(str))
            = 0) then
        begin
          Result := false;
          break;
        end;
        Delete(subStr, start, finish - start);
        start := Pos('"', subStr);
        if start <> 0 then
          subStr[start] := ' ';
        finish := Pos('"', subStr);
        if finish <> 0 then
          subStr[finish] := ' '
        else
          finish := length(subStr) + 1;
      end;
      if Result then
      begin
        while Pos('  ', subStr) <> 0 do
        begin
          Delete(subStr, Pos('  ', subStr), 1);
        end;
        if (subStr <> '') and (subStr[length(subStr)] = ' ') then
          Delete(subStr, length(subStr), 1);
        if (subStr <> '') and (subStr[1] = ' ') then
          Delete(subStr, 1, 1);
        posAp := Pos('''''', subStr);
        while posAp <> 0 do
        begin
          Delete(subStr, posAp, 1);
          subStr[posAp] := '"';
          posAp := Pos('''''', subStr);
        end;

        start := 1;
        finish := Pos(' ', subStr);
        while finish <> 0 do
        begin
          curentS := Copy(subStr, start, finish - start);
          Delete(subStr, start, finish - start + 1);
          if (curentS <> '') and (Pos(UpperCaseEx(curentS), UpperCaseEx(str))
              = 0) then
          begin
            Result := false;
            break;
          end;
          finish := Pos(' ', subStr);
        end;
        if (Pos(UpperCaseEx(subStr), UpperCaseEx(str)) = 0) and (subStr <> '')
          then
        begin
          Result := false;
        end;
      end;
    end
    else
    begin
      Delete(subStr, 1, 1);
      while (subStr <> '') and (subStr[length(subStr)] = ' ') do
        Delete(subStr, length(subStr), 1);
      while (subStr <> '') and (subStr[1] = ' ') do
        Delete(subStr, 1, 1);

      if (subStr <> '') and (subStr <> '"') and (subStr[1] = '"') and
        (subStr[length(subStr)] = '"') then
      begin
        Delete(subStr, length(subStr), 1);
        Delete(subStr, 1, 1);
      end;

      Result := UpperCaseEx(subStr) = UpperCaseEx(str);

    end;
  end;
end;

function InInterval(s: string; X: integer): boolean;
var
  curentS, s1, s2: string;
  k, minusCount: integer;
  ok: boolean;
  start, finish: integer;
begin
  curentS := '';
  ok := true;
  Result := false;
  start := 1;
  finish := Pos(',', s);
  while start <> 0 do
  begin
    if finish <> 0 then
    begin
      curentS := Copy(s, start, finish - start);
      s[finish] := ' ';
    end
    else
      curentS := Copy(s, start, length(s) - start + 1);
    while Pos(' ', curentS) <> 0 do
      Delete(curentS, Pos(' ', curentS), 1);
    if (UpperCaseEx(curentS) = UpperCaseEx(notAssignedString)) then
    begin
      if X = -1 then
      begin
        Result := true;
        break;
      end;
    end
    else
    begin
      minusCount := 0;
      while Pos('-', curentS) <> 0 do
      begin
        curentS[Pos('-', curentS)] := ' ';
        inc(minusCount);
      end;
      case minusCount of
        0:
          begin
            ok := true;
            for k := 1 to length(curentS) do
              if not((curentS[k] >= '0') and (curentS[k] <= '9')) then
                ok := false;
            if curentS = '' then
              ok := false;
            if ok and (StrToInt(curentS) = X) then
            begin
              Result := true;
              break;
            end;
          end;
        1:
          begin
            ok := true;
            for k := 1 to length(curentS) do
              if not(((curentS[k] >= '0') and (curentS[k] <= '9')) or
                  (curentS[k] = ' ')) then
                ok := false;
            if ok then
            begin
              s1 := Copy(curentS, 1, Pos(' ', curentS) - 1);
              s2 := Copy(curentS, Pos(' ', curentS) + 1, length(curentS) - Pos
                  (' ', curentS) + 1);
              if (s1 <> '') and (s2 <> '') and (StrToInt(s1) <= X) and
                (StrToInt(s2) >= X) then
              begin
                Result := true;
                break;
              end;
            end;
          end;
      end;
    end;
    start := finish;
    finish := Pos(',', s);
  end;
end;

procedure TMykolayMainForm.SaveButtonClick(Sender: TObject);
begin
  ColumsViewModeForm.ShowInSaveMode;
end;

procedure TMykolayMainForm.SearchNameLabeledEditChange(Sender: TObject);
begin
  if needRefresh then
    FillTable;
end;

function TMykolayMainForm.lessString(s1: string; s2: string): boolean;
const
  exeptCount = 2;
  exeptArray: array [0 .. exeptCount - 1] of string = ('ВУЛ.', 'ПР.');
var
  i, j, k: integer;
  a, b: real;

begin
  i := 1;
  j := 1;
  if (s2 = UpperCaseEx(notAssignedString)) or
    (s1 = UpperCaseEx(notAssignedString)) then
  begin
    Result := (s1 = UpperCaseEx(notAssignedString)) and
      (s2 <> UpperCaseEx(notAssignedString));
  end
  else
  begin
    if SortedBy.colum = ColumInfo[cinAddress].Num then
    begin
      for k := 0 to exeptCount - 1 do
        if Pos(exeptArray[k], s1) = 1 then
        begin
          i := length(exeptArray[k]) + 1;
          break;
        end;
      for k := 0 to exeptCount - 1 do
        if Pos(exeptArray[k], s2) = 1 then
        begin
          j := length(exeptArray[k]) + 1;
          break;
        end;
    end;

    while Pos('  ', s1) <> 0 do
      Delete(s1, Pos('  ', s1), 1);
    while Pos('  ', s2) <> 0 do
      Delete(s2, Pos('  ', s2), 1);

    while (s1[i] = ' ') and (i <= length(s1)) do
      inc(i);
    while (s2[j] = ' ') and (j <= length(s2)) do
      inc(j);
    while (s1[i] = s2[j]) and (i <= length(s1)) and (j <= length(s2)) do
    begin
      inc(i);
      inc(j);
    end;
    if (s1[i] = s2[j]) and ((i = length(s1)) or (j <= length(s2))) then
      Result := length(s1) < length(s2)
    else
    begin
      case s1[i] of
        'Ъ', 'Ы', 'Э':
          a := 2000;
        'Ґ':
          a := 1043.5;
        'Є':
          a := 1045.5;
        'І':
          a := 1048.5;
        'Ї':
          a := 1048.75;
      else
        a := ord(s1[i]);
      end;
      case s2[j] of
        'Ъ', 'Ы', 'Э':
          b := 2000;
        'Ґ':
          b := 1043.5;
        'Є':
          b := 1045.5;
        'І':
          b := 1048.5;
        'Ї':
          b := 1048.75;
      else
        b := ord(s2[j]);
      end;
      Result := a < b;
    end;
  end;
end;

procedure TMykolayMainForm.SortInterval(start: integer; finish: integer);
  function getXString(i: integer): string;
  begin
    if (ColumInfo[cinName].Num = SortedBy.colum) then
    begin
      getXString := UpperCaseEx(DataBase[i].Name);
    end;
    if (ColumInfo[cinAddress].Num = SortedBy.colum) then
    begin
      getXString := UpperCaseEx(DataBase[i].Address);
    end;
    if (ColumInfo[cinDistrict].Num = SortedBy.colum) then
    begin
      getXString := UpperCaseEx(DataBase[i].District);
    end;
    if (ColumInfo[cinGuardiansParents].Num = SortedBy.colum) then
    begin
      getXString := UpperCaseEx(DataBase[i].GuardiansParents);
    end;
    if (ColumInfo[cinStatus].Num = SortedBy.colum) then
    begin
      getXString := UpperCaseEx(DataBase[i].Status);
    end;
    if (ColumInfo[cinNotes].Num = SortedBy.colum) then
    begin
      getXString := UpperCaseEx(DataBase[i].Notes);
    end;
    if (ColumInfo[cinNeedsHelp].Num = SortedBy.colum) then
    begin
      if DataBase[i].NeedsHelp then
        getXString := YesString
      else
        getXString := NoString;
    end;
  end;
  function getXInteger(i: integer): integer;
  begin
    if (ColumInfo[cinYearOfBirth].Num = SortedBy.colum) then
    begin
      getXInteger := DataBase[i].YearOfBirth;
    end;
    if (ColumInfo[cinNumOfEvelope].Num = SortedBy.colum) then
    begin
      getXInteger := DataBase[i].NumOfEvelope;
    end;
    if (ColumInfo[cinAge].Num = SortedBy.colum) then
    begin
      if DataBase[i].YearOfBirth <> -1 then
        getXInteger := curentYear - DataBase[i].YearOfBirth
      else
        getXInteger := -1;
    end;
    if (ColumInfo[cinSex].Num = SortedBy.colum) then
    begin
      case DataBase[i].Sex of
        girl:
          getXInteger := 1;
        boy:
          getXInteger := 2;
        notAssigned:
          getXInteger := -1;
      end;
    end;
  end;
  procedure QSString(L: integer; R: integer);
  var
    X: string;
    i, j, z: integer;
  begin
    X := getXString(IDArray[(L + R) div 2]);
    i := L;
    j := R;
    repeat
      while (lessString(getXString(IDArray[i]), X) and (SortedBy.sortType = up)
          or lessString(X, getXString(IDArray[i])) and
          (SortedBy.sortType = down)) do
        inc(i);
      while (lessString(X, getXString(IDArray[j])) and (SortedBy.sortType = up)
          or lessString(getXString(IDArray[j]), X) and
          (SortedBy.sortType = down)) do
        dec(j);
      if i <= j then
      begin
        z := IDArray[i];
        IDArray[i] := IDArray[j];
        IDArray[j] := z;
        inc(i);
        dec(j);
      end;
    until (i > j);
    if L < j then
      QSString(L, j);
    if i < R then
      QSString(i, R);
  end;
  procedure QSInteger(L: integer; R: integer);
  var
    X: integer;
    i, j, z: integer;
  begin
    X := getXInteger(IDArray[(L + R) div 2]);
    i := L;
    j := R;
    repeat
      while ((getXInteger(IDArray[i]) < X) and (SortedBy.sortType = up) or
          (getXInteger(IDArray[i]) > X) and (SortedBy.sortType = down)) do
        inc(i);
      while ((getXInteger(IDArray[j]) > X) and (SortedBy.sortType = up) or
          (getXInteger(IDArray[j]) < X) and (SortedBy.sortType = down)) do
        dec(j);
      if i <= j then
      begin
        z := IDArray[i];
        IDArray[i] := IDArray[j];
        IDArray[j] := z;
        inc(i);
        dec(j);
      end;
    until (i > j);
    if L < j then
      QSInteger(L, j);
    if i < R then
      QSInteger(i, R);
  end;
  procedure QSNoSort(L: integer; R: integer);
  var
    X: integer;
    i, j, z: integer;
  begin
    X := IDArray[(L + R) div 2];
    i := L;
    j := R;
    repeat
      while IDArray[i] < X do
        inc(i);
      while IDArray[j] > X do
        dec(j);
      if i <= j then
      begin
        z := IDArray[i];
        IDArray[i] := IDArray[j];
        IDArray[j] := z;
        inc(i);
        dec(j);
      end;
    until (i > j);
    if L < j then
      QSNoSort(L, j);
    if i < R then
      QSNoSort(i, R);
  end;

begin
  if ((SortedBy.colum = ColumInfo[cinYearOfBirth].Num) or
      (SortedBy.colum = ColumInfo[cinNumOfEvelope].Num) or
      (SortedBy.colum = ColumInfo[cinAge].Num) or
      (SortedBy.colum = ColumInfo[cinNumOfEvelope].Num) or
      (SortedBy.colum = ColumInfo[cinSex].Num)) and
    (SortedBy.sortType <> nosort) then
  begin
    if IDArrayCount <> 0 then
      QSInteger(start, finish);
  end;
  if ((SortedBy.colum = ColumInfo[cinName].Num) or
      (SortedBy.colum = ColumInfo[cinAddress].Num) or
      (SortedBy.colum = ColumInfo[cinDistrict].Num) or
      (SortedBy.colum = ColumInfo[cinGuardiansParents].Num) or
      (SortedBy.colum = ColumInfo[cinStatus].Num) or
      (SortedBy.colum = ColumInfo[cinNotes].Num) or
      (SortedBy.colum = ColumInfo[cinNeedsHelp].Num)) and
    (SortedBy.sortType <> nosort) then
  begin
    if IDArrayCount <> 0 then
      QSString(start, finish);
  end;
  if SortedBy.sortType = nosort then
    QSNoSort(start, finish);
end;

procedure TMykolayMainForm.Sort;
begin
  SortInterval(0, IDArrayCount - 1);
end;

procedure TMykolayMainForm.MakeIDArray();
var
  i, age: integer;
begin
  IDArrayCount := 0;
  SetLength(IDArray, DataBaseCount);
  for i := 0 to DataBaseCount - 1 do
  begin
    if DataBase[i].YearOfBirth > 0 then
      age := curentYear - DataBase[i].YearOfBirth
    else
      age := -1;
    if ((SearchNameLabeledEdit.Text = '') or IncludeString
        (SearchNameLabeledEdit.Text, DataBase[i].Name)) and
      ((SearchAddressLabeledEdit.Text = '') or IncludeString
        (SearchAddressLabeledEdit.Text, DataBase[i].Address)) and
      ((SearchDistrictLabeledEdit.Text = '') or IncludeString
        (SearchDistrictLabeledEdit.Text, DataBase[i].District)) and
      ((SearchNotesLabeledEdit.Text = '') or IncludeString
        (SearchNotesLabeledEdit.Text, DataBase[i].Notes)) and
      ((SearchStatusComboBox.Text = '') or IncludeString
        (SearchStatusComboBox.Text, DataBase[i].Status)) and
      ((SearchGuardiansAndParensLabeledEdit.Text = '') or IncludeString
        (SearchGuardiansAndParensLabeledEdit.Text, DataBase[i].GuardiansParents)
      ) and ((SearchNumOfEvelopeLabeledEdit.Text = '') or
        (InInterval(SearchNumOfEvelopeLabeledEdit.Text,
          DataBase[i].NumOfEvelope))) and ((SearchAgeLabeledEdit.Text = '') or
        (InInterval(SearchAgeLabeledEdit.Text, age))) and
      ((SearchYearOfBirthLabeledEdit.Text = '') or
        (InInterval(SearchYearOfBirthLabeledEdit.Text, DataBase[i].YearOfBirth))
      ) and ((SearchWhiteListComboBox.ItemIndex = 0) or
        ((SearchWhiteListComboBox.ItemIndex = 1) and
          (DataBase[i].NeedsHelp = true)) or
        ((SearchWhiteListComboBox.ItemIndex = 2) and
          (DataBase[i].NeedsHelp = false))) and
      (((SearchSexComboBox.ItemIndex = 1) and (DataBase[i].Sex = notAssigned))
        or ((SearchSexComboBox.ItemIndex = 2) and (DataBase[i].Sex = boy)) or
        ((SearchSexComboBox.ItemIndex = 3) and (DataBase[i].Sex = girl)) or
        ((SearchSexComboBox.ItemIndex = 4) and (DataBase[i].Sex <> notAssigned)
        ) or (SearchSexComboBox.ItemIndex = 0))

      then
    begin
      inc(IDArrayCount);
      IDArray[IDArrayCount - 1] := i;
    end;
  end;
  SetLength(IDArray, IDArrayCount);

end;

procedure TMykolayMainForm.ShowData();
var
  i: integer;
begin
  DataBaseStringGrid.ColCount := ShowColumCount;
  for i := 0 to ColumCount - 1 do
  begin
    if ColumInfo[i].Num >= 0 then
    begin
      DataBaseStringGrid.ColWidths[ColumInfo[i].Num] := ColumInfo[i].Width;
      if (SortedBy.colum = ColumInfo[i].Num) then
      begin
        case SortedBy.sortType of
          up:
            DataBaseStringGrid.Cells[ColumInfo[i].Num, 0] := '/\ ' + ColumInfo
              [i].Caption;
          down:
            DataBaseStringGrid.Cells[ColumInfo[i].Num, 0] := '\/ ' + ColumInfo
              [i].Caption;
          nosort:
            DataBaseStringGrid.Cells[ColumInfo[i].Num, 0] := ColumInfo[i]
              .Caption;
        end;
      end
      else
        DataBaseStringGrid.Cells[ColumInfo[i].Num, 0] := ColumInfo[i].Caption;
    end;
  end;

  if IDArrayCount <> 0 then
  begin
    DataBaseStringGrid.RowCount := IDArrayCount + 1;
  end
  else
  begin
    DataBaseStringGrid.RowCount := 2;
    DataBaseStringGrid.Rows[1].Clear;
  end;
  for i := 0 to IDArrayCount - 1 do
  begin
    if ColumInfo[cinName].Num >= 0 then
      DataBaseStringGrid.Cells[ColumInfo[cinName].Num, i + 1] := DataBase
        [IDArray[i]].Name;
    if ColumInfo[cinAddress].Num >= 0 then
      DataBaseStringGrid.Cells[ColumInfo[cinAddress].Num, i + 1] := DataBase
        [IDArray[i]].Address;
    if ColumInfo[cinDistrict].Num >= 0 then
      DataBaseStringGrid.Cells[ColumInfo[cinDistrict].Num, i + 1] := DataBase
        [IDArray[i]].District;
    if ColumInfo[cinSex].Num >= 0 then
    begin
      case DataBase[IDArray[i]].Sex of
        girl:
          DataBaseStringGrid.Cells[ColumInfo[cinSex].Num, i + 1] := girlString;
        boy:
          DataBaseStringGrid.Cells[ColumInfo[cinSex].Num, i + 1] := boyString;
        notAssigned:
          DataBaseStringGrid.Cells[ColumInfo[cinSex].Num, i + 1] :=
            notAssignedString;
      end;
    end;
    if ColumInfo[cinYearOfBirth].Num >= 0 then
      DataBaseStringGrid.Cells[ColumInfo[cinYearOfBirth].Num, i + 1] :=
        IntToStrEx(DataBase[IDArray[i]].YearOfBirth);
    if ColumInfo[cinGuardiansParents].Num >= 0 then
      DataBaseStringGrid.Cells[ColumInfo[cinGuardiansParents].Num, i + 1] :=
        DataBase[IDArray[i]].GuardiansParents;
    if ColumInfo[cinStatus].Num >= 0 then
      DataBaseStringGrid.Cells[ColumInfo[cinStatus].Num, i + 1] := DataBase
        [IDArray[i]].Status;
    if ColumInfo[cinNotes].Num >= 0 then
      DataBaseStringGrid.Cells[ColumInfo[cinNotes].Num, i + 1] := DataBase
        [IDArray[i]].Notes;
    if ColumInfo[cinNumOfEvelope].Num >= 0 then
      DataBaseStringGrid.Cells[ColumInfo[cinNumOfEvelope].Num, i + 1] :=
        IntToStrEx(DataBase[IDArray[i]].NumOfEvelope);
    if ColumInfo[cinAge].Num >= 0 then
    begin
      if (DataBase[IDArray[i]].YearOfBirth > 0) then
        DataBaseStringGrid.Cells[ColumInfo[cinAge].Num, i + 1] := IntToStrEx
          (curentYear - DataBase[IDArray[i]].YearOfBirth)
      else
        DataBaseStringGrid.Cells[ColumInfo[cinAge].Num, i + 1] :=
          notAssignedString;
    end;
    if ColumInfo[cinNeedsHelp].Num >= 0 then
    begin
      if DataBase[IDArray[i]].NeedsHelp then
        DataBaseStringGrid.Cells[ColumInfo[cinNeedsHelp].Num, i + 1] :=
          YesString
      else
        DataBaseStringGrid.Cells[ColumInfo[cinNeedsHelp].Num, i + 1] :=
          NoString;
    end;

  end;

end;

procedure TMykolayMainForm.FillTable();
begin
  MakeIDArray;
  if SortedBy.sortType <> nosort then
    Sort;
  ShowData;
end;

procedure TMykolayMainForm.LoadDataBase(WL: boolean);
var
  i: integer;
  j: integer;
  s: string;
  RootNode: IXMLNode;
begin
  i := 0;
  while (DataBaseXMLDocument.ChildNodes[i].NodeName <> keyMykolay) and
    (DataBaseXMLDocument.ChildNodes.Count > i) do
  begin
    inc(i);
  end;
  RootNode := DataBaseXMLDocument.ChildNodes[i];
  SetLength(DataBase, DataBaseCount + RootNode.ChildNodes.Count);
  for i := 0 to RootNode.ChildNodes.Count - 1 do
  begin
    if (RootNode.ChildNodes[i].LocalName = keyChildren) and
      (RootNode.ChildNodes[i].ChildNodes.Count > 0) then
    begin
      inc(DataBaseCount);
      DataBase[DataBaseCount - 1].Name := '';
      DataBase[DataBaseCount - 1].Address := '';
      DataBase[DataBaseCount - 1].District := '';
      DataBase[DataBaseCount - 1].Sex := notAssigned;
      DataBase[DataBaseCount - 1].YearOfBirth := -1;
      DataBase[DataBaseCount - 1].GuardiansParents := '';
      DataBase[DataBaseCount - 1].Status := '';
      DataBase[DataBaseCount - 1].Notes := '';
      DataBase[DataBaseCount - 1].NumOfEvelope := -1;
      DataBase[DataBaseCount - 1].NeedsHelp := WL;

      for j := 0 to RootNode.ChildNodes[i].ChildNodes.Count - 1 do
      begin
        s := RootNode.ChildNodes[i].ChildNodes[j].Text;
        if RootNode.ChildNodes[i].ChildNodes[j].LocalName = keyName then
          DataBase[DataBaseCount - 1].Name := RootNode.ChildNodes[i].ChildNodes
            [j].Text;
        if RootNode.ChildNodes[i].ChildNodes[j].LocalName = keyNeedsHelp then
          DataBase[DataBaseCount - 1].NeedsHelp := UpperCaseEx
            (RootNode.ChildNodes[i].ChildNodes[j].Text) = UpperCaseEx
            (YesString);
        if RootNode.ChildNodes[i].ChildNodes[j].LocalName = keyAddress then
          DataBase[DataBaseCount - 1].Address := RootNode.ChildNodes[i]
            .ChildNodes[j].Text;
        if RootNode.ChildNodes[i].ChildNodes[j].LocalName = keyDistrict then
          DataBase[DataBaseCount - 1].District := RootNode.ChildNodes[i]
            .ChildNodes[j].Text;
        if RootNode.ChildNodes[i].ChildNodes[j].LocalName = keySex then
        begin
          if UpperCaseEx(RootNode.ChildNodes[i].ChildNodes[j].Text)
            = UpperCaseEx(GString) then
            DataBase[DataBaseCount - 1].Sex := girl
          else if UpperCaseEx(RootNode.ChildNodes[i].ChildNodes[j].Text)
            = UpperCaseEx(BString) then
            DataBase[DataBaseCount - 1].Sex := boy
          else
            DataBase[DataBaseCount - 1].Sex := notAssigned;
        end;
        if RootNode.ChildNodes[i].ChildNodes[j].LocalName = keyYearOfBirth then
          DataBase[DataBaseCount - 1].YearOfBirth := StrToIntEx
            (RootNode.ChildNodes[i].ChildNodes[j].Text);
        if RootNode.ChildNodes[i].ChildNodes[j].LocalName =
          keyGuardiansParents then
          DataBase[DataBaseCount - 1].GuardiansParents := RootNode.ChildNodes[i]
            .ChildNodes[j].Text;
        if RootNode.ChildNodes[i].ChildNodes[j].LocalName = keyStatus then
          DataBase[DataBaseCount - 1].Status := RootNode.ChildNodes[i]
            .ChildNodes[j].Text;
        if RootNode.ChildNodes[i].ChildNodes[j].LocalName = keyNotes then
          DataBase[DataBaseCount - 1].Notes := RootNode.ChildNodes[i].ChildNodes
            [j].Text;
        if RootNode.ChildNodes[i].ChildNodes[j].LocalName = keyNumOfEvelope then
          DataBase[DataBaseCount - 1].NumOfEvelope := StrToIntEx
            (RootNode.ChildNodes[i].ChildNodes[j].Text);
      end;
      if DataBase[DataBaseCount - 1].Name = '' then
        DataBase[DataBaseCount - 1].Name := notAssignedString;
      if DataBase[DataBaseCount - 1].Address = '' then
        DataBase[DataBaseCount - 1].Address := notAssignedString;
      if DataBase[DataBaseCount - 1].District = '' then
        DataBase[DataBaseCount - 1].District := notAssignedString;
      if DataBase[DataBaseCount - 1].GuardiansParents = '' then
        DataBase[DataBaseCount - 1].GuardiansParents := notAssignedString;
      if DataBase[DataBaseCount - 1].Status = '' then
        DataBase[DataBaseCount - 1].Status := notAssignedString;
      if DataBase[DataBaseCount - 1].Notes = '' then
        DataBase[DataBaseCount - 1].Notes := notAssignedString;

    end;
  end;
  DataBaseXMLDocument.XML.Clear;
  SetLength(DataBase, DataBaseCount);
  DataBaseChanged := false;
end;

procedure TMykolayMainForm.LoadReserveCopyButtonClick(Sender: TObject);
var
  ReserveDir: string;
  ok: boolean;
begin
  while ok do
  begin
    if GetFolderDialog(Application.Handle, 'Виберіть папку', ReserveDir) then
    begin
      if FileExists(ReserveDir + '\' + defDataBaseFile) and FileExists
        (ReserveDir + '\' + defDataBaseBlackListFile) then
      begin
        LoadAllBases(ReserveDir + '\');

        DataBaseChanged := true;
        FillTable;
        ok := false;
      end
      else
      begin
        ShowMessage('Файли не знайдені');
      end;
    end
    else
    begin
      ok := false;
    end;
  end;

end;

procedure TMykolayMainForm.AddButtonClick(Sender: TObject);
begin
  EditForm.ShowEditWindow(-1, false);
end;

procedure TMykolayMainForm.AddFromFileButtonClick(Sender: TObject);
begin
  if OpenDialog.Execute then
  begin
    DataBaseXMLDocument.LoadFromFile(OpenDialog.FileName);
    LoadDataBase(true);
    DataBaseXMLDocument.XML.Clear;
    DataBaseChanged := true;
    FillTable;
  end;
end;

procedure TMykolayMainForm.ClearSearchButtonClick(Sender: TObject);
begin
  needRefresh := false;
  if SearchNameLabeledEdit.Text <> '' then
    SearchNameLabeledEdit.Clear;
  if SearchAddressLabeledEdit.Text <> '' then
    SearchAddressLabeledEdit.Clear;
  if SearchDistrictLabeledEdit.Text <> '' then
    SearchDistrictLabeledEdit.Clear;
  if SearchGuardiansAndParensLabeledEdit.Text <> '' then
    SearchGuardiansAndParensLabeledEdit.Clear;
  if SearchNotesLabeledEdit.Text <> '' then
    SearchNotesLabeledEdit.Clear;
  if SearchStatusComboBox.Text <> '' then
    SearchStatusComboBox.Text := '';
  if SearchSexComboBox.ItemIndex <> 0 then
    SearchSexComboBox.ItemIndex := 0;
  if SearchNumOfEvelopeLabeledEdit.Text <> '' then
    SearchNumOfEvelopeLabeledEdit.Clear;
  if SearchYearOfBirthLabeledEdit.Text <> '' then
    SearchYearOfBirthLabeledEdit.Clear;
  if SearchAgeLabeledEdit.Text <> '' then
    SearchAgeLabeledEdit.Clear;
  if SearchWhiteListComboBox.ItemHeight <> 0 then
    SearchWhiteListComboBox.ItemIndex := 0;
  needRefresh := true;
  FillTable;
end;

procedure TMykolayMainForm.CloneButtonClick(Sender: TObject);
begin
  if DataBaseStringGrid.Selection.Top > 0 then
    EditForm.ShowEditWindow(IDArray[DataBaseStringGrid.Selection.Top - 1],
      true);
end;

procedure TMykolayMainForm.ColumsViewModeButtonClick(Sender: TObject);
begin
  ColumsViewModeForm.ShowInPreferencesMode;
end;

procedure TMykolayMainForm.DataBaseStringGridDblClick(Sender: TObject);
var
  Col, Row: integer;
  p: TPoint;
begin
  GetCursorPos(p);
  DataBaseStringGrid.MouseToCell(DataBaseStringGrid.ScreenToClient(p).X,
    DataBaseStringGrid.ScreenToClient(p).Y, Col, Row);
  if (Row <> 0) and (Row <> -1) and (Col <> -1) and (IDArrayCount <> 0) then
  begin
    if DataBaseStringGrid.Selection.Top > 0 then
      EditForm.ShowEditWindow(IDArray[DataBaseStringGrid.Selection.Top - 1],
        false);
  end;
end;

procedure TMykolayMainForm.DataBaseStringGridDrawCell
  (Sender: TObject; ACol, ARow: integer; Rect: TRect; State: TGridDrawState);
begin
  if (ARow <> DataBaseStringGrid.Selection.Top) and (ACol = SortedBy.colum) and
    (SortedBy.sortType <> nosort) then
    DataBaseStringGrid.Canvas.Brush.color := TColor($F9CEAD);
  if (ACol >= 0) and (ARow > 0) then
  begin
    DataBaseStringGrid.Canvas.FillRect(Rect);
    DataBaseStringGrid.Canvas.TextOut(Rect.Left, Rect.Top,
      DataBaseStringGrid.Cells[ACol, ARow]);
  end;
end;

procedure TMykolayMainForm.DataBaseStringGridKeyDown
  (Sender: TObject; var Key: Word; Shift: TShiftState);
begin
  if (ord(Key) = VK_DELETE) and (IDArrayCount <> 0) then
    DeleteButton.Click;
  if (ord(Key) = VK_RETURN) then
    EditButton.Click;

end;

procedure TMykolayMainForm.DataBaseStringGridMouseDown
  (Sender: TObject; Button: TMouseButton; Shift: TShiftState; X, Y: integer);
var
  Col, Row: integer;
begin
  DataBaseStringGrid.MouseToCell(X, Y, Col, Row);
  if (Row = 0) and (Col <> -1) then
  begin
    if SortedBy.colum = Col then
    begin
      if SortedBy.sortType < down then
      begin
        inc(SortedBy.sortType);
      end
      else
      begin
        SortedBy.sortType := nosort;
      end;
    end
    else
    begin
      SortedBy.sortType := up;
      SortedBy.colum := Col;
    end;
    Sort;
    ShowData;
  end;
end;

procedure TMykolayMainForm.DataBaseStringGridMouseUp
  (Sender: TObject; Button: TMouseButton; Shift: TShiftState; X, Y: integer);
var
  i: integer;
begin
  for i := 0 to ColumCount - 1 do
  begin
    if ColumInfo[i].Num >= 0 then
    begin
      ColumInfo[i].Width := DataBaseStringGrid.ColWidths[ColumInfo[i].Num];
    end;
  end;
end;

procedure TMykolayMainForm.DeleteButtonClick(Sender: TObject);
var
  i: integer;
begin
  if MessageDlg('Ви впевнені що хочете видалити запис?', mtConfirmation,
    mbYesNo, 0, mbYes) = ID_YES then
  begin
    DataBaseChanged := true;
    for i := IDArray[DataBaseStringGrid.Selection.Top - 1] to DataBaseCount - 2
      do
    begin
      DataBase[i] := DataBase[i + 1];
    end;
    for i := 0 to IDArrayCount - 1 do
    begin
      if IDArray[i] > IDArray[DataBaseStringGrid.Selection.Top - 1] then
        dec(IDArray[i]);
    end;
    for i := DataBaseStringGrid.Selection.Top - 1 to IDArrayCount - 2 do
    begin
      IDArray[i] := IDArray[i + 1];
    end;
    dec(DataBaseCount);
    dec(IDArrayCount);
    SetLength(DataBase, DataBaseCount);
    SetLength(IDArray, IDArrayCount);
    FillTable;
  end;
end;

procedure TMykolayMainForm.EditButtonClick(Sender: TObject);
begin
  if DataBaseStringGrid.Selection.Top > 0 then
    EditForm.ShowEditWindow(IDArray[DataBaseStringGrid.Selection.Top - 1],
      false);
end;

procedure TMykolayMainForm.LoadAllBases(dir: string);
begin
  DataBaseCount := 0;
  DataBaseXMLDocument.LoadFromFile(dir + defDataBaseFile);
  LoadDataBase(true);
  DataBaseXMLDocument.XML.Clear;
  DataBaseXMLDocument.LoadFromFile(dir + defDataBaseBlackListFile);
  LoadDataBase(false);
  DataBaseXMLDocument.XML.Clear;
end;

procedure TMykolayMainForm.SaveAllBases(dir: string);
begin
  DataBaseXMLDocument.Active := true;
  MakeDataBase(true);
  DataBaseXMLDocument.SaveToFile(dir + defDataBaseFile);
  DataBaseXMLDocument.XML.Clear;
  DataBaseXMLDocument.Active := true;
  MakeDataBase(false);
  DataBaseXMLDocument.SaveToFile(dir + defDataBaseBlackListFile);
  DataBaseXMLDocument.XML.Clear;

end;

procedure TMykolayMainForm.SaveAllButtonClick(Sender: TObject);
begin
  SaveAllBases(AppDir);
  DataBaseChanged := false;
end;

procedure TMykolayMainForm.FormCloseQuery(Sender: TObject;
  var CanClose: Boolean);
var
  but : integer;
begin
  if DataBaseChanged then
  begin
    but:=MessageDlg('Зберегти зміни в базі?', mtConfirmation, mbYesNoCancel, 0, mbYes);
    if but= mrYes then
    begin
      SaveAllBases(AppDir);
    end;
    if but= mrCancel then
    begin
      CanClose:=false;
    end;
  end;

end;

procedure TMykolayMainForm.FormCreate(Sender: TObject);
var
  time: _SYSTEMTIME;
begin
  GetSystemTime(time);
  curentYear := time.wYear;
  AppDir := Application.ExeName + '/..\';
  if FileExists(AppDir + '\' + defDataBaseFile) and FileExists
    (AppDir + '\' + defDataBaseBlackListFile) then
    LoadAllBases(AppDir);
  SearchStatusComboBox.Items.Clear;
  SearchStatusComboBox.Items.Add('="-"');
  SearchStatusComboBox.Items.Add('="+"');
  SearchStatusComboBox.Items.Add('="+!"');
  SearchStatusComboBox.Items.Add('="+!!"');
  SearchSexComboBox.Items.Clear;
  SearchSexComboBox.Items.Add(AllString);
  SearchSexComboBox.Items.Add(notAssignedString);
  SearchSexComboBox.Items.Add(boyString);
  SearchSexComboBox.Items.Add(girlString);
  SearchSexComboBox.Items.Add(BoysAndGirlsString);
  SearchSexComboBox.ItemIndex := 0;
  SearchWhiteListComboBox.Clear;
  SearchWhiteListComboBox.Items.Add(AllString);
  SearchWhiteListComboBox.Items.Add(NeedsHelpString);
  SearchWhiteListComboBox.Items.Add(dontNeedsHelpString);
  SearchWhiteListComboBox.ItemIndex := 0;
  FillTable;
end;

procedure TMykolayMainForm.FormDestroy(Sender: TObject);
begin
  SetLength(DataBase, 0);
  SetLength(IDArray, 0);
end;

procedure TMykolayMainForm.FormResize(Sender: TObject);
begin
  DataBaseStringGrid.Width := MykolayMainForm.ClientWidth -
    DataBaseStringGrid.Left * 2;
  DataBaseStringGrid.Height := MykolayMainForm.ClientHeight -
    DataBaseStringGrid.Top - DataBaseStringGrid.Left;
  LogoImage.Width := MykolayMainForm.ClientWidth - 3 * LogoImage.Left -
    SearchGuardiansAndParensLabeledEdit.Width;
  SearchNameLabeledEdit.Left := LogoImage.Width + 2 * LogoImage.Left;
  SearchNumOfEvelopeLabeledEdit.Left := LogoImage.Width + 2 * LogoImage.Left;
  SearchYearOfBirthLabeledEdit.Left := LogoImage.Width + 2 * LogoImage.Left;
  SearchNotesLabeledEdit.Left := LogoImage.Width + 2 * LogoImage.Left;
  SearchAddressLabeledEdit.Left := LogoImage.Width + 2 * LogoImage.Left;
  SearchGuardiansAndParensLabeledEdit.Left :=
    LogoImage.Width + 2 * LogoImage.Left;
  SearchAgeLabeledEdit.Left := LogoImage.Width + 3 * LogoImage.Left +
    SearchYearOfBirthLabeledEdit.Width;
  SearchWhiteListComboBox.Left := MykolayMainForm.ClientWidth - 2 *
    LogoImage.Left - AddButton.Width - SearchWhiteListComboBox.Width;
  AddButton.Left := ClientWidth - AddButton.Width - LogoImage.Left;
  EditButton.Left := ClientWidth - EditButton.Width - LogoImage.Left;
  DeleteButton.Left := ClientWidth - DeleteButton.Width - LogoImage.Left;
  CloneButton.Left := ClientWidth - DeleteButton.Width - LogoImage.Left;
  ClearSearchButton.Left := ClientWidth - ClearSearchButton.Width -
    LogoImage.Left;
  SearchSexComboBox.Left := ClientWidth - SearchSexComboBox.Width -
    LogoImage.Left;
  SearchDistrictLabeledEdit.Left := ClientWidth -
    SearchDistrictLabeledEdit.Width - LogoImage.Left;
  SearchStatusComboBox.Left := ClientWidth - SearchStatusComboBox.Width -
    LogoImage.Left;
  SexLabel.Left := ClientWidth - SearchSexComboBox.Width - LogoImage.Left;
  StatusLabel.Left := ClientWidth - SearchStatusComboBox.Width - LogoImage.Left;

end;

procedure TMykolayMainForm.MakeDataBase(WL: boolean);
var
  RootNode: IXMLNode;
  SubNode: IXMLNode;
  i: integer;
begin
  DataBaseXMLDocument.Options := DataBaseXMLDocument.Options +
    [doNodeAutoIndent];
  RootNode := DataBaseXMLDocument.AddChild(keyMykolay);
  for i := 0 to DataBaseCount - 1 do
  begin
    if DataBase[i].NeedsHelp = WL then
    begin
      SubNode := RootNode.AddChild(keyChildren);
      SubNode.AddChild(keyName);
      SubNode.AddChild(keyAddress);
      SubNode.AddChild(keyDistrict);
      SubNode.AddChild(keySex);
      SubNode.AddChild(keyYearOfBirth);
      SubNode.AddChild(keyGuardiansParents);
      SubNode.AddChild(keyStatus);
      SubNode.AddChild(keyNotes);
      SubNode.AddChild(keyNumOfEvelope);
      SubNode.ChildValues[0 * 3 + 2] := DataBase[i].Name;
      SubNode.ChildValues[1 * 3 + 2] := DataBase[i].Address;
      SubNode.ChildValues[2 * 3 + 2] := DataBase[i].District;
      case DataBase[i].Sex of
        notAssigned:
          SubNode.ChildValues[3 * 3 + 2] := notAssignedString;
        girl:
          SubNode.ChildValues[3 * 3 + 2] := GString;
        boy:
          SubNode.ChildValues[3 * 3 + 2] := BString;
      end;
      SubNode.ChildValues[4 * 3 + 2] := IntToStrEx(DataBase[i].YearOfBirth);
      SubNode.ChildValues[5 * 3 + 2] := DataBase[i].GuardiansParents;
      SubNode.ChildValues[6 * 3 + 2] := DataBase[i].Status;
      SubNode.ChildValues[7 * 3 + 2] := DataBase[i].Notes;
      SubNode.ChildValues[8 * 3 + 2] := IntToStrEx(DataBase[i].NumOfEvelope);
    end;
  end;
  DataBaseXMLDocument.Options := DataBaseXMLDocument.Options -
    [doNodeAutoIndent];
end;

function BrowseCallbackProc(hwnd: hwnd; uMsg: UINT; lParam: lParam;
  lpData: lParam): integer; stdcall;
begin
  if (uMsg = BFFM_INITIALIZED) then
    SendMessage(hwnd, BFFM_SETSELECTION, 1, lpData);
  BrowseCallbackProc := 0;
end;

function GetFolderDialog(Handle: integer; Caption: string;
  var strFolder: string): boolean;

var
  BrowseInfo: TBrowseInfo;
  ItemIDList: PItemIDList;
  JtemIDList: PItemIDList;
  Path: PChar;
begin
  Result := false;
  Path := StrAlloc(MAX_PATH);
  SHGetSpecialFolderLocation(Handle, CSIDL_DRIVES, JtemIDList);
  with BrowseInfo do
  begin
    ulFlags := BIF_NEWDIALOGSTYLE or BIF_RETURNONLYFSDIRS;
    hwndOwner := GetActiveWindow;
    pidlRoot := JtemIDList;
    SHGetSpecialFolderLocation(hwndOwner, CSIDL_DRIVES, JtemIDList);

    { return display name of item selected }
    pszDisplayName := StrAlloc(MAX_PATH);

    { set the title of dialog }
    lpszTitle := PChar(Caption); // 'Select the folder';
    { flags that control the return stuff }
    lpfn := @BrowseCallbackProc;
    { extra info that's passed back in callbacks }
    lParam := LongInt(PChar(strFolder));
  end;

  ItemIDList := SHBrowseForFolder(BrowseInfo);

  if (ItemIDList <> nil) then
    if SHGetPathFromIDList(ItemIDList, Path) then
    begin
      strFolder := Path;
      Result := true
    end;
end;

procedure TMykolayMainForm.MakeReserveCopyButtonClick(Sender: TObject);
var
  ReserveDir, s: string;
  sysTime: _SYSTEMTIME;
  time: TTime;
  tz: TTimeZoneInformation;
begin
  if GetFolderDialog(Application.Handle, 'Виберіть папку', ReserveDir) then
  begin
    GetTimeZoneInformation(tz);
    GetSystemTime(sysTime);
    time := GetTime;
    TimeSeparator := '.';
    s := TimeToStr(time);

    ReserveDir := ReserveDir + '\' + IntToStr(sysTime.wDay) + '-' + IntToStr
      (sysTime.wMonth) + '-' + IntToStr(sysTime.wYear) + ' ' + s + '\';
    CreateDir(ReserveDir);
    SaveAllBases(ReserveDir);
  end;

end;

procedure TMykolayMainForm.N17Click(Sender: TObject);
begin
  StatisticForm.ShowModal;
end;

procedure TMykolayMainForm.N20Click(Sender: TObject);
begin
  WinExec('hh.exe help.chm',SW_SHOW);
end;

procedure TMykolayMainForm.N21Click(Sender: TObject);
begin
FormAbout.Show;
end;

end.
