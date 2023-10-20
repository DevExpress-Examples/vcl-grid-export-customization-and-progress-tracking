//---------------------------------------------------------------------------

#ifndef Unit1H
#define Unit1H
//---------------------------------------------------------------------------
#include <System.Classes.hpp>
#include <Vcl.Controls.hpp>
#include <Vcl.StdCtrls.hpp>
#include <Vcl.Forms.hpp>
#include "cxButtons.hpp"
#include "cxClasses.hpp"
#include "cxContainer.hpp"
#include "cxControls.hpp"
#include "cxCustomData.hpp"
#include "cxData.hpp"
#include "cxDataStorage.hpp"
#include "cxDBData.hpp"
#include "cxEdit.hpp"
#include "cxFilter.hpp"
#include "cxGraphics.hpp"
#include "cxGrid.hpp"
#include "cxGridCustomTableView.hpp"
#include "cxGridCustomView.hpp"
#include "cxGridDBTableView.hpp"
#include "cxGridLevel.hpp"
#include "cxGridTableView.hpp"
#include "cxLookAndFeelPainters.hpp"
#include "cxLookAndFeels.hpp"
#include "cxNavigator.hpp"
#include "cxProgressBar.hpp"
#include "cxStyles.hpp"
#include "dxDateRanges.hpp"
#include "dxmdaset.hpp"
#include "dxScrollbarAnnotations.hpp"
#include <Data.DB.hpp>
#include <Vcl.Menus.hpp>
#include "cxExport.hpp"
#include "cxGridExportLink.hpp"
#include "dxSpreadSheet.hpp"


//---------------------------------------------------------------------------

class TExportHandler : public TdxSpreadSheet, IcxExportBeforeSave, IcxExportProgress
{
  private:
	long m_cRef;
	TcxProgressBar *FProgressBar;
	void __fastcall OnBeforeSave(TdxSpreadSheet* Sender);
	void __fastcall OnProgress(System::TObject* Sender, int Percent);

	HRESULT __stdcall QueryInterface(const GUID& IID, void **Obj);

	ULONG __stdcall AddRef(void);
	ULONG __stdcall Release(void);
  public:
	inline __fastcall virtual TExportHandler(TComponent* AOwner);
	__property int ProgressBar = {read = FProgressBar, write = FProgressBar};

};


class TForm1 : public TForm, IcxExportProgress
{
__published:	// IDE-managed Components
	TcxGrid *cxGrid1;
	TcxGridDBTableView *cxGrid1DBTableView1;
	TcxGridDBColumn *cxGrid1DBTableView1RecId;
	TcxGridDBColumn *cxGrid1DBTableView1CustNo;
	TcxGridDBColumn *cxGrid1DBTableView1Company;
	TcxGridDBColumn *cxGrid1DBTableView1Addr1;
	TcxGridDBColumn *cxGrid1DBTableView1City;
	TcxGridDBColumn *cxGrid1DBTableView1State;
	TcxGridDBColumn *cxGrid1DBTableView1Zip;
	TcxGridDBColumn *cxGrid1DBTableView1Country;
	TcxGridDBColumn *cxGrid1DBTableView1Phone;
	TcxGridDBColumn *cxGrid1DBTableView1FAX;
	TcxGridDBColumn *cxGrid1DBTableView1Contact;
	TcxGridDBColumn *cxGrid1DBTableView1LastInvoiceDate;
	TcxGridLevel *cxGrid1Level1;
	TcxButton *btnExportToXLSX;
	TcxProgressBar *cxProgressBar1;
	TDataSource *DataSource1;
	TdxMemData *dxMemData1;
	TFloatField *dxMemData1CustNo;
	TStringField *dxMemData1Company;
	TStringField *dxMemData1Addr1;
	TStringField *dxMemData1City;
	TStringField *dxMemData1State;
	TStringField *dxMemData1Zip;
	TStringField *dxMemData1Country;
	TStringField *dxMemData1Phone;
	TStringField *dxMemData1FAX;
	TStringField *dxMemData1Contact;
	TDateTimeField *dxMemData1LastInvoiceDate;
	void __fastcall btnExportToXLSXClick(TObject *Sender);
private:	// User declarations
public:		// User declarations
	__fastcall TForm1(TComponent* Owner);
};
//---------------------------------------------------------------------------
extern PACKAGE TForm1 *Form1;
//---------------------------------------------------------------------------
#endif
