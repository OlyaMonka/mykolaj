unit StatisticUnit;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, Grids, xmldom, XMLIntf, msxmldom, XMLDoc;

type
  TStatisticForm = class(TForm)
    StatisticStringGrid: TStringGrid;
    CloseButton: TButton;
    StaticticSaveXMLDialog: TSaveDialog;
    StatisticXMLDocument: TXMLDocument;
    SaveXMLButton: TButton;
    SaveHTMLButton: TButton;
    StaticticSaveHTMLDialog: TSaveDialog;
    SaveAddressListButton: TButton;
    procedure FormShow(Sender: TObject);
    procedure CloseButtonClick(Sender: TObject);
    procedure SaveXMLButtonClick(Sender: TObject);
    procedure SaveHTMLButtonClick(Sender: TObject);
    procedure FormCanResize(Sender: TObject; var NewWidth, NewHeight: Integer;
      var Resize: Boolean);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure SaveAddressListButtonClick(Sender: TObject);
  private
  public
    { Public declarations }
  end;

var
  StatisticForm: TStatisticForm;

implementation

uses MykolayUnit;
{$R *.dfm}

type
  TEvelopes = record
    num: Integer;
    AddressCount: Integer;
    ChildrenCount: Integer;
  end;

var
  Evelopes: array of TEvelopes;
  prevIdArray: array of Integer;
  prevSort: TSort;

const
  keyStatictic = 'Statictic';
  keyAddressList = 'AddressList';
  keyNum = 'Num';
  keyAddressCount = 'AddressCount';
  keyChildrenCount = 'ChildrenCount';
  keyEvelope = 'Evelope';

procedure TStatisticForm.CloseButtonClick(Sender: TObject);
begin
  Close;
end;

procedure TStatisticForm.FormCanResize(Sender: TObject; var NewWidth,
  NewHeight: Integer; var Resize: Boolean);
begin
  Resize := false;
end;

procedure TStatisticForm.FormClose(Sender: TObject; var Action: TCloseAction);
var
  i: Integer;
begin
  SortedBy := prevSort;
  for i := 0 to IDArrayCount - 1 do
    IDArray[i] := prevIdArray[i];
  SetLength(prevIdArray, 0);
end;

procedure TStatisticForm.FormShow(Sender: TObject);
var

  i: Integer;
  intervalStart: Integer;
  j: Integer;
begin
  prevSort := SortedBy;
  SetLength(prevIdArray, IDArrayCount);
  for i := 0 to IDArrayCount - 1 do
    prevIdArray[i] := IDArray[i];
  SortedBy.sortType := up;
  SortedBy.colum := columInfo[cinNumOfEvelope].num;
  MykolayMainForm.Sort;

  SortedBy.sortType := up;
  SortedBy.colum := columInfo[cinAddress].num;

  SetLength(Evelopes, 1);
  Evelopes[0].num := DataBase[IDArray[0]].NumOfEvelope;
  Evelopes[0].AddressCount := 0;
  Evelopes[0].ChildrenCount := 0;
  intervalStart := 0;
  for i := 0 to IDArrayCount - 2 do
  begin
    if DataBase[IDArray[i]].NumOfEvelope <> DataBase[IDArray[i + 1]]
      .NumOfEvelope then
    begin
      SetLength(Evelopes, length(Evelopes) + 1);
      Evelopes[length(Evelopes) - 1].num := DataBase[IDArray[i + 1]]
        .NumOfEvelope;
      Evelopes[length(Evelopes) - 1].AddressCount := 0;
      if length(Evelopes) > 1 then
      begin
        Evelopes[length(Evelopes) - 2].ChildrenCount := i - intervalStart + 1;
        if i <> intervalStart then
          MykolayMainForm.SortInterval(intervalStart, i);
        for j := intervalStart to i do
        begin
          if not EqualAddress(IDArray[j], IDArray[j + 1]) then
          //
          // MykolayMainForm.lessString(DataBase[IDArray[j]].Address,
          // DataBase[IDArray[j + 1]].Address) or MykolayMainForm.lessString
          // (DataBase[IDArray[j + 1]].Address, DataBase[IDArray[j]].Address)
          // then
          begin
            inc(Evelopes[length(Evelopes) - 2].AddressCount);
          end;
        end;
        intervalStart := i + 1;
      end;

    end;
  end;
  Evelopes[length(Evelopes) - 1].ChildrenCount :=
    IDArrayCount - 1 - intervalStart + 1;
  if IDArrayCount - 1 <> intervalStart then
    MykolayMainForm.SortInterval(intervalStart, IDArrayCount - 1);
  for j := intervalStart to IDArrayCount - 1 do
  begin
    if MykolayMainForm.lessString(DataBase[IDArray[j - 1]].Address,
      DataBase[IDArray[j - 1 + 1]].Address) or MykolayMainForm.lessString
      (DataBase[IDArray[j - 1 + 1]].Address, DataBase[IDArray[j - 1]].Address)
      then
    begin
      inc(Evelopes[length(Evelopes) - 1].AddressCount);
    end;
  end;

  StatisticStringGrid.ColCount := 3;
  StatisticStringGrid.Cells[0, 0] := '№';
  StatisticStringGrid.Cells[1, 0] := 'К-сть адрес';
  StatisticStringGrid.Cells[2, 0] := 'К-сть дітей';
  StatisticStringGrid.ColWidths[0] := 50;
  StatisticStringGrid.ColWidths[1] := 100;
  StatisticStringGrid.ColWidths[2] := 100;

  StatisticStringGrid.RowCount := length(Evelopes) + 1;
  for i := 0 to length(Evelopes) - 1 do
  begin
    StatisticStringGrid.Cells[0, i + 1] := IntToStrEx(Evelopes[i].num);
    StatisticStringGrid.Cells[1, i + 1] := IntToStr(Evelopes[i].AddressCount);
    StatisticStringGrid.Cells[2, i + 1] := IntToStr(Evelopes[i].ChildrenCount);
  end;

end;

procedure TStatisticForm.SaveAddressListButtonClick(Sender: TObject);
var
  RootNode: IXMLNode;
  SubNode, EvelopeNode: IXMLNode;
  i: Integer;
  sl: TStringList;
  k: Integer;
begin
  if StaticticSaveXMLDialog.Execute then
  begin
    StatisticXMLDocument.Active := true;

    StatisticXMLDocument.Options := StatisticXMLDocument.Options +
      [doNodeAutoIndent];
    RootNode := StatisticXMLDocument.AddChild(keyAddressList);
    EvelopeNode := RootNode.AddChild(keyEvelope);
    EvelopeNode.AddChild(keyNumOfEvelope);
    k := 0;
    EvelopeNode.ChildValues[k * 3 + 2] := DataBase[IDArray[0]].NumOfEvelope;
    EvelopeNode.AddChild(keyAddress);
    inc(k);
    EvelopeNode.ChildValues[k * 3 + 2] := DataBase[IDArray[0]].Address;

    for i := 0 to length(IDArray) - 2 do
    begin
      if DataBase[IDArray[i]].NumOfEvelope <> DataBase[IDArray[i + 1]]
        .NumOfEvelope then
      begin
        EvelopeNode := RootNode.AddChild(keyEvelope);
        EvelopeNode.AddChild(keyNumOfEvelope);
        k := 0;
        EvelopeNode.ChildValues[k * 3 + 2] := DataBase[IDArray[i + 1]]
          .NumOfEvelope;
      end;
      if (DataBase[IDArray[i]].NumOfEvelope <> DataBase[IDArray[i + 1]]
          .NumOfEvelope) or not EqualAddress(IDArray[i], IDArray[i + 1]) then
      begin
        EvelopeNode.AddChild(keyAddress);
        inc(k);
        EvelopeNode.ChildValues[k * 3 + 2] := DataBase[IDArray[i + 1]].Address;
      end;

    end;
    StatisticXMLDocument.Options := StatisticXMLDocument.Options -
      [doNodeAutoIndent];

    sl := TStringList.Create;
    sl.Add('<?xml-stylesheet type="text/css" href="addressList.css"?>');
    sl.AddStrings(StatisticXMLDocument.XML);

    StatisticXMLDocument.Options := StatisticXMLDocument.Options -
      [doNodeAutoIndent];

    sl.SaveToFile(StaticticSaveXMLDialog.FileName, TEncoding.UTF8);
    sl.Clear;
    sl.LoadFromFile(appDir + 'addressList.css');
    sl.SaveToFile(StaticticSaveXMLDialog.FileName + '\../addressList.css',
      TEncoding.UTF8);
    sl.Clear;
    sl.Free;

    StatisticXMLDocument.XML.Clear;
  end;
end;

procedure TStatisticForm.SaveHTMLButtonClick(Sender: TObject);
var
  RootNode: IXMLNode;
  SubNode: IXMLNode;
  i: Integer;
  sl: TStringList;
begin
  if StaticticSaveHTMLDialog.Execute then
  begin
    StatisticXMLDocument.Active := true;

    StatisticXMLDocument.Options := StatisticXMLDocument.Options +
      [doNodeAutoIndent];
    RootNode := StatisticXMLDocument.AddChild('table');
    SubNode := RootNode.AddChild('tr');
    SubNode.AddChild('th');
    SubNode.AddChild('th');
    SubNode.AddChild('th');
    SubNode.ChildValues[0 * 3 + 2] := '№';
    SubNode.ChildValues[1 * 3 + 2] := 'К-сть адрес';
    SubNode.ChildValues[2 * 3 + 2] := 'К-сть дітей';

    for i := 0 to length(Evelopes) - 1 do
    begin
      SubNode := RootNode.AddChild('tr');
      SubNode.AddChild('th');
      SubNode.AddChild('th');
      SubNode.AddChild('th');
      SubNode.ChildValues[0 * 3 + 2] := Evelopes[i].num;
      SubNode.ChildValues[1 * 3 + 2] := Evelopes[i].AddressCount;
      SubNode.ChildValues[2 * 3 + 2] := Evelopes[i].ChildrenCount;

    end;
    StatisticXMLDocument.Options := StatisticXMLDocument.Options -
      [doNodeAutoIndent];

    sl := TStringList.Create;
    // sl.Add('<?xml-stylesheet type="text/css" href="statistic.css"?>');
    sl.AddStrings(StatisticXMLDocument.XML);

    StatisticXMLDocument.Options := StatisticXMLDocument.Options -
      [doNodeAutoIndent];

    sl.SaveToFile(StaticticSaveHTMLDialog.FileName, TEncoding.UTF8);
    sl.Clear;
    // sl.LoadFromFile(appDir+'statistic.css');
    // sl.SaveToFile(StaticticSaveHTMLDialog.FileName+'\../statistic.css',TEncoding.UTF8);
    // sl.Clear;
    sl.Free;

    StatisticXMLDocument.XML.Clear;
  end;
end;

procedure TStatisticForm.SaveXMLButtonClick(Sender: TObject);
var
  RootNode: IXMLNode;
  SubNode: IXMLNode;
  i: Integer;
  sl: TStringList;
begin
  if StaticticSaveXMLDialog.Execute then
  begin
    StatisticXMLDocument.Active := true;

    StatisticXMLDocument.Options := StatisticXMLDocument.Options +
      [doNodeAutoIndent];
    RootNode := StatisticXMLDocument.AddChild(keyStatictic);
    for i := 0 to length(Evelopes) - 1 do
    begin
      SubNode := RootNode.AddChild(keyEvelope);
      SubNode.AddChild(keyNum);
      SubNode.AddChild(keyAddressCount);
      SubNode.AddChild(keyChildrenCount);
      SubNode.ChildValues[0 * 3 + 2] := Evelopes[i].num;
      SubNode.ChildValues[1 * 3 + 2] := Evelopes[i].AddressCount;
      SubNode.ChildValues[2 * 3 + 2] := Evelopes[i].ChildrenCount;

    end;
    StatisticXMLDocument.Options := StatisticXMLDocument.Options -
      [doNodeAutoIndent];

    sl := TStringList.Create;
    sl.Add('<?xml-stylesheet type="text/css" href="statistic.css"?>');
    sl.AddStrings(StatisticXMLDocument.XML);

    StatisticXMLDocument.Options := StatisticXMLDocument.Options -
      [doNodeAutoIndent];

    sl.SaveToFile(StaticticSaveXMLDialog.FileName, TEncoding.UTF8);
    sl.Clear;
    sl.LoadFromFile(appDir + 'statistic.css');
    sl.SaveToFile(StaticticSaveXMLDialog.FileName + '\../statistic.css',
      TEncoding.UTF8);
    sl.Clear;
    sl.Free;

    StatisticXMLDocument.XML.Clear;
  end;
end;

end.
