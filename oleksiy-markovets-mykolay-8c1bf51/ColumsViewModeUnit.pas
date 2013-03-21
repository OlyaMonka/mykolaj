unit ColumsViewModeUnit;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, CheckLst, xmldom, XMLIntf, msxmldom, XMLDoc;

type
  TColumsViewModeForm = class(TForm)
    OkButton: TButton;
    CancelButton: TButton;
    ColumsViewCheckListBox: TCheckListBox;
    SelectAllButton: TButton;
    SelectNoneButton: TButton;
    UpButton: TButton;
    DownButton: TButton;
    SaveDialog: TSaveDialog;
    SaveXMLDocument: TXMLDocument;
    GroupByAddressCheckBox: TCheckBox;
    IgnoreAppartmentCheckBox: TCheckBox;
    procedure CancelButtonClick(Sender: TObject);
    procedure UpButtonClick(Sender: TObject);
    procedure DownButtonClick(Sender: TObject);
    procedure SelectAllButtonClick(Sender: TObject);
    procedure SelectNoneButtonClick(Sender: TObject);
    procedure OkButtonClick(Sender: TObject);
    procedure ShowInSaveMode();
    procedure ShowInPreferencesMode();
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  ColumsViewModeForm: TColumsViewModeForm;

implementation
uses MykolayUnit;
var
  ColNum: array [0..ColumCount-1] of integer;
  LastColNumSaveMode: array [0..ColumCount-1] of integer =
    (
      cinNumOfEvelope,cinAddress,cinName,cinSex,cinAge,cinGuardiansParents,
      cinStatus,cinNotes,cinDistrict,cinYearOfBirth,cinNeedsHelp
    );
  LastSetSaveMode : set of byte =
    [
      cinNumOfEvelope,cinAddress,cinName,cinSex,cinAge,cinGuardiansParents,
      cinStatus,cinNotes
    ];
  mode : (notset,save,view)=notset;
{$R *.dfm}
procedure TColumsViewModeForm.ShowInSaveMode();
var
  i : integer;
begin
  mode:=save;
  GroupByAddressCheckBox.Visible:=true;

  ColumsViewCheckListBox.Items.Clear;
  for i := 0 to ColumCount - 1 do
    ColumsViewCheckListBox.Items.Add('');
  for i := 0 to ColumCount - 1 do
  begin
    ColumsViewCheckListBox.Items[i]:=ColumInfo[LastColNumSaveMode[i]].Caption;
    ColNum[i]:=LastColNumSaveMode[i];
    ColumsViewCheckListBox.Checked[i]:=ColNum[i] in LastSetSaveMode;
  end;
  Show;
end;
procedure TColumsViewModeForm.ShowInPreferencesMode();
var
  i,j : integer;
begin
  mode:=view;
  GroupByAddressCheckBox.Visible:=false;
  ColumsViewCheckListBox.Items.Clear;
  for i := 0 to ColumCount - 1 do
    ColumsViewCheckListBox.Items.Add('');
  j:=0;
  for i := 0 to ColumCount - 1 do
    if ColumInfo[i].Num>=0 then
    begin
      ColumsViewCheckListBox.Items[ColumInfo[i].Num]:=ColumInfo[i].Caption;
      ColumsViewCheckListBox.Checked[ColumInfo[i].Num]:=true;
      ColNum[ColumInfo[i].Num]:=i;
      inc(j);
    end;
  for i := 0 to ColumCount - 1 do
    if ColumInfo[i].Num<0 then
    begin
      ColumsViewCheckListBox.Items[j]:=ColumInfo[i].Caption;
      ColumsViewCheckListBox.Checked[j]:=false;
      ColNum[j]:=i;
      inc(j);
    end;
  Show;
end;
procedure TColumsViewModeForm.CancelButtonClick(Sender: TObject);
begin
  close;
end;

procedure TColumsViewModeForm.DownButtonClick(Sender: TObject);
var
  zS: string;
  zB: boolean;
  zi: integer;
begin
  if ColumsViewCheckListBox.ItemIndex<ColumsViewCheckListBox.Count-1 then
  begin
    zS:=ColumsViewCheckListBox.Items[ColumsViewCheckListBox.ItemIndex];
    zB:=ColumsViewCheckListBox.Checked[ColumsViewCheckListBox.ItemIndex];
    zi:=ColNum[ColumsViewCheckListBox.ItemIndex];
    ColumsViewCheckListBox.Items[ColumsViewCheckListBox.ItemIndex]:=ColumsViewCheckListBox.Items[ColumsViewCheckListBox.ItemIndex+1];
    ColumsViewCheckListBox.Checked[ColumsViewCheckListBox.ItemIndex]:=ColumsViewCheckListBox.Checked[ColumsViewCheckListBox.ItemIndex+1];
    ColNum[ColumsViewCheckListBox.ItemIndex]:=ColNum[ColumsViewCheckListBox.ItemIndex+1];
    ColumsViewCheckListBox.Items[ColumsViewCheckListBox.ItemIndex+1]:=zS;
    ColumsViewCheckListBox.Checked[ColumsViewCheckListBox.ItemIndex+1]:=zB;
    ColNum[ColumsViewCheckListBox.ItemIndex+1]:=zi;
    ColumsViewCheckListBox.ItemIndex:=ColumsViewCheckListBox.ItemIndex+1;
  end;

end;

procedure TColumsViewModeForm.FormClose(Sender: TObject;
  var Action: TCloseAction);
var
  i : integer;
begin
  if mode=view then
  begin
    mode:=notset;
  end;
  if mode=save then
  begin
    LastSetSaveMode:=[];
    for i := 0 to ColumCount - 1 do
      if ColumsViewCheckListBox.Checked[i] then
      begin
        LastColNumSaveMode[i]:=ColNum[i];
        LastSetSaveMode:=LastSetSaveMode+[ColNum[i]];
      end;
      mode:=notset;
  end;
end;
procedure TColumsViewModeForm.OkButtonClick(Sender: TObject);


var
  i,j,k: Integer;
  RootNode : IXMLNode;
  SubNode,FamilyNode : IXMLNode;
  fileName : string;
  oldSort: TSort;
  oldIDArray : array of integer;
  sl : TStringList;
  beginOfInterval : integer;
  ChildrenCount : integer;
  fillerIndex : integer;
begin
  if mode=view then
  begin
    j:=0;
    for i := 0 to ColumCount - 1 do
      if ColumsViewCheckListBox.Checked[i] then
      begin
        ColumInfo[ColNum[i]].Num:=j;
        inc(j);
      end
      else
        ColumInfo[ColNum[i]].Num:=-1;
    if j>0 then
    begin
      close;
      ShowColumCount:=j;
      MykolayMainForm.FillTable;
    end
    else
    begin
      ShowMessage('Необхідно відмітити хоч одне поле');
    end;
  end;
  if mode=save then
  begin
    if SaveDialog.Execute then
    begin
      fileName:=SaveDialog.FileName;
      SaveXMLDocument.Active:=true;
      SaveXMLDocument.Options := SaveXMLDocument.Options + [doNodeAutoIndent];
      RootNode:=SaveXMLDocument.AddChild(keyMykolay);
      if GroupByAddressCheckBox.Checked then
      begin
        SetLength(oldIDArray,IDArrayCount);
        for i := 0 to IDArrayCount - 1 do
          oldIDArray[i]:=IDArray[i];
        oldSort:=SortedBy;
        if not IgnoreAppartmentCheckBox.Checked then
        begin
          SortedBy.colum:=ColumInfo[cinNumOfEvelope].Num;
          SortedBy.sortType:=up;
          MykolayMainForm.Sort;
        end;
        SortedBy.colum:=ColumInfo[cinAddress].Num;
        SortedBy.sortType:=up;
        beginOfInterval:=0;
        if IgnoreAppartmentCheckBox.Checked then
        begin
          MykolayMainForm.Sort;
        end
        else
        begin
          for i := 0 to IDArrayCount - 2 do
            if DataBase[IDArray[i]].NumOfEvelope<>DataBase[IDArray[i+1]].NumOfEvelope then
            begin
                MykolayMainForm.SortInterval(beginOfInterval,i);
                beginOfInterval:=i+1;
            end;
          if beginOfInterval<(IDArrayCount - 1) then
             MykolayMainForm.SortInterval(beginOfInterval,IDArrayCount - 1);
        end;


      end;

      for i := 0 to IDArrayCount - 1 do
      begin
        if GroupByAddressCheckBox.Checked then
        begin
          if (i=0) or (not EqualAddress(IDArray[i-1],IDArray[i],IgnoreAppartmentCheckBox.Checked))
            or ((DataBase[IDArray[i]].NumOfEvelope<>DataBase[IDArray[i-1]].NumOfEvelope) and (not IgnoreAppartmentCheckBox.Checked)) then
          begin
            if i>0 then
            begin
              FamilyNode.ChildValues[fillerIndex*3+2]:=ChildrenCount;
              ChildrenCount:=0;

              inc(fillerIndex);
            end;

            FamilyNode:=RootNode.AddChild(keyFamily);
            for k := 0 to ColumCount - 1 do
            begin
              if (ColumsViewCheckListBox.Checked[k]) and
                ((ColNum[k]=cinAddress) or (ColNum[k]=cinNumOfEvelope) or (ColNum[k]=cinStatus))
                then
                  FamilyNode.AddChild(ColumInfo[ColNum[k]].Key);

            end;
            FamilyNode.AddChild(keyChildrenCount);

            fillerIndex:=0;
            for j := 0 to ColumCount - 1 do
            begin
              if (ColumsViewCheckListBox.Checked[j]) then
              begin
                if (ColNum[j]=cinAddress) then
                begin
                  if IgnoreAppartmentCheckBox.Checked then
                    FamilyNode.ChildValues[fillerIndex*3+2]:=RemoveAppartmentFromAddres(DataBase[IDArray[i]].Address)
                  else
                    FamilyNode.ChildValues[fillerIndex*3+2]:=DataBase[IDArray[i]].Address;
                  inc(fillerIndex);
                end;
                if (ColNum[j]=cinNumOfEvelope) then
                begin
                  FamilyNode.ChildValues[fillerIndex*3+2]:=DataBase[IDArray[i]].NumOfEvelope;
                  inc(fillerIndex);
                end;
                if  (ColNum[j]=cinStatus) then
                begin
                  FamilyNode.ChildValues[fillerIndex*3+2]:=DataBase[IDArray[i]].Status;
                  inc(fillerIndex);
                end;
              end;
            end;
          end;
          subNode:=FamilyNode.AddChild(keyChildren);
          inc(ChildrenCount);
        end
        else
        begin
          SubNode:=RootNode.AddChild(keyChildren);

        end;

        for k := 0 to ColumCount - 1 do
          if ColumsViewCheckListBox.Checked[k] then
          begin
            if ((ColNum[k]<>cinAddress) and (ColNum[k]<>cinStatus) and (ColNum[k]<>cinNumOfEvelope)) or (not GroupByAddressCheckBox.Checked) then
               SubNode.AddChild(ColumInfo[ColNum[k]].Key);
          end;
        k:=0;
        for j := 0 to ColumCount - 1 do
          if (ColumsViewCheckListBox.Checked[j])
            and (
              ((ColNum[j]<>cinAddress) and (ColNum[j]<>cinStatus) and (ColNum[j]<>cinNumOfEvelope))
              or (not GroupByAddressCheckBox.Checked)
              ) then
          begin
            if ColNum[j]=cinName then
              SubNode.ChildValues[k*3+2]:=DataBase[IDArray[i]].Name;
            if (ColNum[j]=cinAddress) and (not GroupByAddressCheckBox.Checked) then
              SubNode.ChildValues[k*3+2]:=DataBase[IDArray[i]].Address;
            if ColNum[j]=cinDistrict then
              SubNode.ChildValues[k*3+2]:=DataBase[IDArray[i]].District;
            if ColNum[j]=cinSex then
              case DataBase[IDArray[i]].Sex of
                notAssigned :  SubNode.ChildValues[k*3+2]:=notAssignedString;
                girl :  SubNode.ChildValues[k*3+2]:=GString;
                boy :  SubNode.ChildValues[k*3+2]:=BString;
              end;
            if ColNum[j]=cinYearOfBirth then
              SubNode.ChildValues[k*3+2]:=IntToStrEx(DataBase[IDArray[i]].YearOfBirth);
            if ColNum[j]=cinGuardiansParents then
              SubNode.ChildValues[k*3+2]:=DataBase[IDArray[i]].GuardiansParents;
            if (ColNum[j]=cinStatus) and (not GroupByAddressCheckBox.Checked) then
              SubNode.ChildValues[k*3+2]:=DataBase[IDArray[i]].Status;
            if ColNum[j]=cinNotes then
              SubNode.ChildValues[k*3+2]:=DataBase[IDArray[i]].Notes;
            if (ColNum[j]=cinNumOfEvelope) and (not GroupByAddressCheckBox.Checked) then
              SubNode.ChildValues[k*3+2]:=IntToStrEx(DataBase[IDArray[i]].NumOfEvelope);
            if ColNum[j]=cinAge then
              if DataBase[IDArray[i]].YearOfBirth>=0 then
                SubNode.ChildValues[k*3+2]:=IntToStrEx(curentYear-DataBase[IDArray[i]].YearOfBirth)
              else
                SubNode.ChildValues[k*3+2]:=notAssignedString;
            if ColNum[j]=cinNeedsHelp then
            begin
              if DataBase[IDArray[i]].NeedsHelp then
                SubNode.ChildValues[k*3+2]:=YesString
              else
                SubNode.ChildValues[k*3+2]:=NoString;
            end;
            inc(k);
          end;
      end;

      if GroupByAddressCheckBox.Checked then
      begin
        FamilyNode.ChildValues[fillerIndex*3+2]:=ChildrenCount;

        SortedBy:=oldSort;
        for i := 0 to IDArrayCount - 1 do
          IDArray[i]:=oldIDArray[i];
        SetLength(oldIDArray,0);
      end;
      sl:=TStringList.Create;

      SaveXMLDocument.XML.Insert(0,'<?xml-stylesheet type="text/css" href="mykolay.css"?>');
      SaveXMLDocument.XML.SaveToFile(fileName,TEncoding.UTF8);
      SaveXMLDocument.XML.Clear;


      sl.Clear;
      sl.LoadFromFile(appDir+'mykolay.css');
      sl.SaveToFile(fileName+'\../mykolay.css',TEncoding.UTF8);
      sl.Clear;
      sl.Free;
      close;
    end;
  end;
end;

procedure TColumsViewModeForm.SelectAllButtonClick(Sender: TObject);
var
  i: Integer;
begin
  for i := 0 to ColumsViewCheckListBox.Count - 1 do
    ColumsViewCheckListBox.Checked[i]:=true;
end;

procedure TColumsViewModeForm.SelectNoneButtonClick(Sender: TObject);
var
  i: Integer;
begin
  for i := 0 to ColumsViewCheckListBox.Count - 1 do
    ColumsViewCheckListBox.Checked[i]:=false;
end;

procedure TColumsViewModeForm.UpButtonClick(Sender: TObject);
var
  zS: string;
  zB: boolean;
  zi: integer;
begin
  if ColumsViewCheckListBox.ItemIndex>0 then
  begin
    zS:=ColumsViewCheckListBox.Items[ColumsViewCheckListBox.ItemIndex];
    zB:=ColumsViewCheckListBox.Checked[ColumsViewCheckListBox.ItemIndex];
    zi:=ColNum[ColumsViewCheckListBox.ItemIndex];
    ColumsViewCheckListBox.Items[ColumsViewCheckListBox.ItemIndex]:=ColumsViewCheckListBox.Items[ColumsViewCheckListBox.ItemIndex-1];
    ColumsViewCheckListBox.Checked[ColumsViewCheckListBox.ItemIndex]:=ColumsViewCheckListBox.Checked[ColumsViewCheckListBox.ItemIndex-1];
    ColNum[ColumsViewCheckListBox.ItemIndex]:=ColNum[ColumsViewCheckListBox.ItemIndex-1];
    ColumsViewCheckListBox.Items[ColumsViewCheckListBox.ItemIndex-1]:=zS;
    ColumsViewCheckListBox.Checked[ColumsViewCheckListBox.ItemIndex-1]:=zB;
    ColNum[ColumsViewCheckListBox.ItemIndex-1]:=zi;
    ColumsViewCheckListBox.ItemIndex:=ColumsViewCheckListBox.ItemIndex-1;
  end;
end;

end.
