unit EditUnit;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, ExtCtrls;

type
  TEditForm = class(TForm)
    NameLabeledEdit: TLabeledEdit;
    AddressLabeledEdit: TLabeledEdit;
    DistrictLabeledEdit: TLabeledEdit;
    GuardiansAndParensLabeledEdit: TLabeledEdit;
    NotesLabeledEdit: TLabeledEdit;
    SexComboBox: TComboBox;
    SexLabel: TLabel;
    YearOfBirthComboBox: TComboBox;
    YearOfBirthLabel: TLabel;
    NumOfEvelopeLabeledEdit: TLabeledEdit;
    OkButton: TButton;
    CancelButton: TButton;
    NeedsHelpCheckBox: TCheckBox;
    StatusComboBox: TComboBox;
    StatusLabel: TLabel;
    Image1: TImage;
    procedure FormCreate(Sender: TObject);
    procedure ShowEditWindow(num : integer; clone : boolean);
    procedure CancelButtonClick(Sender: TObject);
    procedure OkButtonClick(Sender: TObject);
    procedure FillEdits(num : integer);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  EditForm: TEditForm;

implementation
uses MykolayUnit;
{$R *.dfm}
Var
  itemNum : integer;
procedure TEditForm.FillEdits(num : integer);
begin
  NameLabeledEdit.Text:=database[num].Name;
  AddressLabeledEdit.Text:=database[num].Address;
  DistrictLabeledEdit.Text:=database[num].District;
  case database[num].Sex of
    notAssigned :  SexComboBox.ItemIndex:=0;
    boy :  SexComboBox.ItemIndex:=1;
    girl :  SexComboBox.ItemIndex:=2;
  end;
  YearOfBirthComboBox.Text:=IntToStrEx(database[num].YearOfBirth);
  GuardiansAndParensLabeledEdit.Text:=database[num].GuardiansParents;
  StatusComboBox.Text:=database[num].Status;
  NotesLabeledEdit.Text:=database[num].Notes;
  if database[num].NumOfEvelope<>-1 then
    NumOfEvelopeLabeledEdit.Text:=IntToStrEx(database[num].NumOfEvelope);
  NeedsHelpCheckBox.Checked:=DataBase[num].NeedsHelp;

end;
procedure TEditForm.ShowEditWindow(num : integer; clone : boolean);
begin
  if not clone then
    itemNum:=num
  else
    itemNum:=-1;
  NameLabeledEdit.Clear;
  AddressLabeledEdit.Clear;
  DistrictLabeledEdit.Clear;
  SexComboBox.ItemIndex:=0;
  YearOfBirthComboBox.Text:=notAssignedString;
  GuardiansAndParensLabeledEdit.Clear;
  StatusComboBox.Text:='';
  NotesLabeledEdit.Clear;
  NumOfEvelopeLabeledEdit.Clear;
  NeedsHelpCheckBox.Checked:=True;
  if num<>-1 then
  begin
    FillEdits(num);
  end;
  EditForm.ShowModal;

end;

procedure TEditForm.CancelButtonClick(Sender: TObject);
begin
  EditForm.Close;
end;

procedure TEditForm.FormCreate(Sender: TObject);
var
  i: Word;
begin
  YearOfBirthComboBox.Items.Clear;
  YearOfBirthComboBox.Items.Add(notAssignedString);
  for i := curentYear downto 1990 do
    YearOfBirthComboBox.Items.Add(IntToStr(i));
  SexComboBox.Items.Clear;
  SexComboBox.Items.Add(notAssignedString);
  SexComboBox.Items.Add(boyString);
  SexComboBox.Items.Add(girlString);
  SexComboBox.ItemIndex:=0;
  StatusComboBox.Items.Add('-');
  StatusComboBox.Items.Add('+');
  StatusComboBox.Items.Add('+!');
  StatusComboBox.Items.Add('+!!');
  StatusComboBox.ItemIndex:=-1;
end;

procedure TEditForm.OkButtonClick(Sender: TObject);
var
  stage1,stage2: boolean;
  num : integer;
begin
  stage1:=False;
  stage2:=False;
  try
    if (itemNum<>-1) then
    begin
      num:=itemNum;
    end
    else begin
      inc(DataBaseCount);
      SetLength(DataBase,DataBaseCount);
      SetLength(IDArray,DataBaseCount);
      stage1:=true;
      IDArray[DataBaseCount-1]:=DataBaseCount-1;
      num:=DataBaseCount-1;
    end;
    DataBase[Num].Name:=NameLabeledEdit.Text;
    DataBase[Num].Address:=AddressLabeledEdit.Text;
    DataBase[Num].District:=DistrictLabeledEdit.Text;
    case SexComboBox.ItemIndex of
      0: DataBase[Num].Sex:=notAssigned;
      1: DataBase[Num].Sex:=boy;
      2: DataBase[Num].Sex:=girl;
    end;
    DataBase[Num].YearOfBirth:=StrToIntEx(YearOfBirthComboBox.Text);
    DataBase[Num].GuardiansParents:=GuardiansAndParensLabeledEdit.Text;
    DataBase[Num].Status:=StatusComboBox.Text;
    DataBase[Num].Notes:=NotesLabeledEdit.Text;
    DataBase[Num].NumOfEvelope:=StrToIntEx(NumOfEvelopeLabeledEdit.Text);
    DataBase[Num].NeedsHelp:=NeedsHelpCheckBox.Checked;
    stage2:=true;
  finally
    if (not stage2) then
    begin
      if stage1 then
      begin
        dec(DataBaseCount);
        SetLength(DataBase,DataBaseCount);
        SetLength(IDArray,DataBaseCount);
      end;
    end
    else begin
      if DataBase[num].Name='' then
        DataBase[num].Name:=notAssignedString;
      if DataBase[num].Address='' then
        DataBase[num].Address:=notAssignedString;
      if DataBase[num].District='' then
        DataBase[num].District:=notAssignedString;
      if DataBase[num].GuardiansParents='' then
        DataBase[num].GuardiansParents:=notAssignedString;
      if DataBase[num].Status='' then
        DataBase[num].Status:=notAssignedString;
      if DataBase[num].Notes='' then
        DataBase[num].Notes:=notAssignedString;
      EditForm.Close;
      DataBaseChanged:=true;      
      MykolayMainForm.FillTable;
    end;
  end;
end;

end.
