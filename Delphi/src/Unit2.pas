unit Unit2;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, cxGraphics, cxControls, cxLookAndFeels,
  cxLookAndFeelPainters, cxStyles, cxCustomData, cxFilter, cxData,
  cxDataStorage, cxEdit, cxNavigator, Data.DB, cxDBData, Vcl.Menus,
  Vcl.StdCtrls, cxButtons, cxGridCustomTableView, cxGridTableView,
  cxGridDBTableView, dxmdaset, cxGridLevel, cxClasses, cxGridCustomView, cxGrid,

  cxgridExportLink, cxExport, dxSpreadSheet, dxSpreadSheetCore, dxDateRanges,
  dxScrollbarAnnotations, cxContainer, cxProgressBar;

type
  TForm2 = class(TForm, IcxExportBeforeSave, IcxExportProgress)
    cxGrid1: TcxGrid;
    cxGrid1DBTableView1: TcxGridDBTableView;
    cxGrid1Level1: TcxGridLevel;
    DataSource1: TDataSource;
    dxMemData1: TdxMemData;
    dxMemData1CustNo: TFloatField;
    dxMemData1Company: TStringField;
    dxMemData1Addr1: TStringField;
    dxMemData1City: TStringField;
    dxMemData1State: TStringField;
    dxMemData1Zip: TStringField;
    dxMemData1Country: TStringField;
    dxMemData1Phone: TStringField;
    dxMemData1FAX: TStringField;
    dxMemData1Contact: TStringField;
    dxMemData1LastInvoiceDate: TDateTimeField;
    cxGrid1DBTableView1RecId: TcxGridDBColumn;
    cxGrid1DBTableView1CustNo: TcxGridDBColumn;
    cxGrid1DBTableView1Company: TcxGridDBColumn;
    cxGrid1DBTableView1Addr1: TcxGridDBColumn;
    cxGrid1DBTableView1City: TcxGridDBColumn;
    cxGrid1DBTableView1State: TcxGridDBColumn;
    cxGrid1DBTableView1Zip: TcxGridDBColumn;
    cxGrid1DBTableView1Country: TcxGridDBColumn;
    cxGrid1DBTableView1Phone: TcxGridDBColumn;
    cxGrid1DBTableView1FAX: TcxGridDBColumn;
    cxGrid1DBTableView1Contact: TcxGridDBColumn;
    cxGrid1DBTableView1LastInvoiceDate: TcxGridDBColumn;
    cxButton1: TcxButton;
    cxProgressBar1: TcxProgressBar;
    procedure cxButton1Click(Sender: TObject);
    procedure OnBeforeSave(Sender: TdxSpreadSheet);
    procedure OnProgress(Sender: TObject; Percent: Integer);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  Form2: TForm2;

implementation

{$R *.dfm}

procedure TForm2.cxButton1Click(Sender: TObject);
begin
  ExportGridToXLSX('test.xlsx', cxGrid1, True, True, True, 'xlsx', Self)
end;

procedure TForm2.OnBeforeSave(Sender: TdxSpreadSheet);
var
  AView: TdxSpreadSheetTableView;
begin
  AView := Sender.ActiveSheetAsTable;
  AView.InsertRows(0, 1);                           //insert a new row
  AView.CreateCell(0, 0).AsString := 'My text';     // add text to a top left cell
end;

procedure TForm2.OnProgress(Sender: TObject; Percent: Integer);
begin
  cxProgressBar1.Position := Percent; // Updates the progress bar position
  Application.ProcessMessages;
end;



end.
