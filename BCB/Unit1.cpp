//---------------------------------------------------------------------------

#include <vcl.h>
#pragma hdrstop

#include "Unit1.h"
//---------------------------------------------------------------------------
#pragma package(smart_init)
#pragma link "cxButtons"
#pragma link "cxClasses"
#pragma link "cxContainer"
#pragma link "cxControls"
#pragma link "cxCustomData"
#pragma link "cxData"
#pragma link "cxDataStorage"
#pragma link "cxDBData"
#pragma link "cxEdit"
#pragma link "cxFilter"
#pragma link "cxGraphics"
#pragma link "cxGrid"
#pragma link "cxGridCustomTableView"
#pragma link "cxGridCustomView"
#pragma link "cxGridDBTableView"
#pragma link "cxGridLevel"
#pragma link "cxGridTableView"
#pragma link "cxLookAndFeelPainters"
#pragma link "cxLookAndFeels"
#pragma link "cxNavigator"
#pragma link "cxProgressBar"
#pragma link "cxStyles"
#pragma link "dxDateRanges"
#pragma link "dxmdaset"
#pragma link "dxScrollbarAnnotations"
#pragma link "cxExport"
#pragma link "cxGridExportLink"
#pragma link "dxSpreadSheet"
#pragma resource "*.dfm"

TForm1 *Form1;
//---------------------------------------------------------------------------
__fastcall TForm1::TForm1(TComponent* Owner)
	: TForm(Owner)
{
}

//---------------------------------------------------------------------------
void __fastcall TForm1::btnExportToXLSXClick(TObject *Sender)
{
	TExportHandler* AExportHandler = new TExportHandler(this);
	try
	{
        AExportHandler->Visible = false;
		AExportHandler->Parent = this;
		AExportHandler->ProgressBar = cxProgressBar1;
		ExportGridToXLSX("test.xlsx", cxGrid1, True, True, True,
			"xlsx", AExportHandler);
	}
	__finally
	{
		delete AExportHandler;
	}
}

//---------------------------------------------------------------------------
void __fastcall TExportHandler::OnBeforeSave(TdxSpreadSheet* Sender)
{
  TdxSpreadSheetTableView* AView;
  AView = Sender->ActiveSheetAsTable;
  AView->InsertRows(0, 1);
  AView->CreateCell(0, 0)->AsString = "Countries";
}

void __fastcall TExportHandler::OnProgress(System::TObject* Sender, int Percent)
{
   if(FProgressBar != nullptr)
   {
	   FProgressBar->Position = Percent;
	   Sleep(10);
	   Application->ProcessMessages();
   }
}

HRESULT __stdcall TExportHandler::QueryInterface(const GUID& IID, void **Obj)
{
	return GetInterface(IID, Obj) ? S_OK : E_NOINTERFACE;
}

ULONG __stdcall TExportHandler::AddRef(void)
{
	InterlockedIncrement(&m_cRef);
	return (ULONG)m_cRef;
};

ULONG __stdcall TExportHandler::Release(void)
{
	ULONG ulRefCount = InterlockedDecrement(&m_cRef);
	return ulRefCount;
};

inline __fastcall TExportHandler::TExportHandler(TComponent* AOwner): TdxSpreadSheet(AOwner)
{
}
//---------------------------------------------------------------------------
