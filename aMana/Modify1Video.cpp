// Modify1Video.cpp : implementation file
//

#include "stdafx.h"
#include "aMana.h"
#include "Modify1Video.h"
#include "afxdialogex.h"


// Modify1Video dialog

IMPLEMENT_DYNAMIC(Modify1Video, CDialogEx)

Modify1Video::Modify1Video(CWnd* pParent /*=NULL*/)
	: CDialogEx(Modify1Video::IDD, pParent)
{

}

Modify1Video::~Modify1Video()
{
}

void Modify1Video::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
	DDX_Control(pDX, IDC_CHECK_CHN, check_chn);
	DDX_Control(pDX, IDC_CHECK_NEW, check_new);
	DDX_Control(pDX, IDC_CHECK_DEL, check_del);
	DDX_Control(pDX, IDC_COMBO_SERIE, combo_serie);
	DDX_Control(pDX, IDC_CHECK_CHANGSERIE, check_serie);
}


BEGIN_MESSAGE_MAP(Modify1Video, CDialogEx)
//	ON_BN_CLICKED(IDC_BUTTON1, &Modify1Video::OnBnClickedButton1)
	ON_BN_CLICKED(IDOK, &Modify1Video::OnBnClickedOk)
	ON_BN_CLICKED(IDCANCEL, &Modify1Video::OnBnClickedCancel)
	ON_WM_PAINT()
	ON_BN_CLICKED(IDC_BUTTON_CLEAR, &Modify1Video::OnBnClickedButtonClear)
	ON_BN_CLICKED(IDC_CHECK_CHANGSERIE, &Modify1Video::OnClickedCheckChangserie)
END_MESSAGE_MAP()


BOOL Modify1Video::OnInitDialog()
{
	CDialogEx::OnInitDialog();

	// TODO:  Add extra initialization here
	combo_serie.SetDroppedWidth(250);
	GetDlgItem(IDC_COMBO_SERIE)->EnableWindow(FALSE);
	RecordInit();

	return TRUE;  // return TRUE unless you set the focus to a control
	// EXCEPTION: OCX Property Pages should return FALSE
}

void Modify1Video::OnPaint()
{
	CPaintDC dc(this); // device context for painting
	// TODO: Add your message handler code here
	if( bPicPath == _T("0"))
		SetDlgItemText(IDC_STATIC_CENTER, _T("No Image"));
	else if (!PathFileExists(bPicPath))
		SetDlgItemText(IDC_STATIC_CENTER, _T("!!! Check Path"));
	else
		ShowPic(bPicPath);
	
	// Do not call CDialogEx::OnPaint() for painting messages
}


// Modify1Video message handlers

void Modify1Video::RecordInit()
{
	CaManaDlg* mDlg;
	mDlg = (CaManaDlg*)AfxGetMainWnd();

	CString mySQLtext;
	_variant_t RecordsAffected;

	int SerieCount, k;

	bVID.Format(_T("%d"), mDlg->g_VID);
	SetDlgItemText(IDC_STATIC_VID, _T("ID: ")+bVID);

	mySQLtext.Format(_T("SELECT * FROM qmain WHERE VID =  %d;"), mDlg->g_VID);
	try
	{
		mDlg->iAcc1.m_pRecordset = mDlg->iAcc1.m_pConnection->Execute(_bstr_t(mySQLtext), &RecordsAffected, adCmdText);
		if ((mDlg->iAcc1.m_pRecordset->BOF) && (mDlg->iAcc1.m_pRecordset->adoEOF))	 
			return;

		bChn = mDlg->iAcc1.m_pRecordset->GetCollect("IsChn");
		bDel = mDlg->iAcc1.m_pRecordset->GetCollect("IsDel");
		bNew = mDlg->iAcc1.m_pRecordset->GetCollect("IsNew");
		bFilePath = mDlg->iAcc1.m_pRecordset->GetCollect("VFilePath");
		bFileSize = mDlg->iAcc1.m_pRecordset->GetCollect("VFileSize");
		bName1 = mDlg->iAcc1.m_pRecordset->GetCollect("LName1");
		bName2 = mDlg->iAcc1.m_pRecordset->GetCollect("LName2");
		bPicPath = mDlg->iAcc1.m_pRecordset->GetCollect("PFilePath");
		bSerie = mDlg->iAcc1.m_pRecordset->GetCollect("SerieName");
		bDate = mDlg->iAcc1.m_pRecordset->GetCollect("CDate");

		mySQLtext.Format(_T("SELECT * FROM serie;"));
		mDlg->iAcc1.m_pRecordset = mDlg->iAcc1.m_pConnection->Execute(_bstr_t(mySQLtext), &RecordsAffected, adCmdText);

		SerieCount = 0;

		mDlg->iAcc1.m_pRecordset->MoveFirst();
		while (!mDlg->iAcc1.m_pRecordset->adoEOF)
		{
			sID[SerieCount] = mDlg->iAcc1.m_pRecordset->GetCollect("SerieID");
			sName[SerieCount] = mDlg->iAcc1.m_pRecordset->GetCollect("SerieName");
			SerieCount++;
			mDlg->iAcc1.m_pRecordset->MoveNext();
		}

	}
	catch (_com_error &e)
	{
		CString eorr;
		eorr.Format(L"DB failed! Code:%08X   Info:%s \r\nDescribe:%s\r\n", e.Error(), (wchar_t*)e.ErrorMessage(), (wchar_t*)e.Description());
		mDlg->echoAndReturn(eorr);
		return ;
	}
	if (bSerie == _T("0")) bSerie = _T("");
	
	SetDlgItemText(IDC_EDIT_FH1, bName1);
	SetDlgItemText(IDC_EDIT_FH2, bName2);
	SetDlgItemText(IDC_STATIC_PATH, bFilePath);
	if ((bName1 == _T("0")) && (bName2 == _T("0")))
		GetDlgItem(IDC_CHECK_CHANGSERIE)->EnableWindow(FALSE);
	SetDlgItemText(IDC_STATIC_OTHER, _T("Size: ") + bFileSize + _T("\t\tDate: ") + bDate + _T("\t\tSerie: ")+bSerie);
	check_chn.SetCheck(bChn);
	check_del.SetCheck(bDel);
	check_new.SetCheck(bNew);
	for (k = 1; k <= SerieCount; k++)
		((CComboBox*)GetDlgItem(IDC_COMBO_SERIE))->AddString(sName[k]);

	if (mDlg->isBusyWorking == true)
		GetDlgItem(IDOK)->EnableWindow(FALSE);
}


void Modify1Video::ShowPic(CString pp)
{
	CDC* pDC = this->GetDC();
	CImage img;
	img.Load(pp);
	::SetStretchBltMode(pDC->m_hDC, HALFTONE);
	::SetBrushOrgEx(pDC->m_hDC, 0, 0, NULL);
	img.Draw(pDC->m_hDC, 8, 9, 700, 470);

}



void Modify1Video::OnBnClickedOk()
{
	// TODO: Add your control notification handler code here
	if (MessageBox(_T("Are you sure to change this record?"), _T(""), MB_OKCANCEL) == IDOK)
	{

		
		CaManaDlg* mDlg;
		mDlg = (CaManaDlg*)AfxGetMainWnd();

		CString mySQLtext;
		_variant_t RecordsAffected;

		int mLIB ;
		int i;
		bChn = check_chn.GetCheck();
		bNew = check_new.GetCheck();
		int serieNum = 1;
		GetDlgItemText(IDC_EDIT_FH1, bName1);
		GetDlgItemText(IDC_EDIT_FH2, bName2);

		if (check_serie.GetCheck() == 1) 
		{
			CString tmpName;
			GetDlgItemText(IDC_COMBO_SERIE, tmpName);
			if (tmpName != _T(""))
			{
				for (i = 0; i < combo_serie.GetCount(); i++)
				{
					if (tmpName == sName[i])
					{
						serieNum = sID[i];
						break;
					}
				}
			}
		}



		if (bName1 == _T("") || bName2 == _T(""))
		{
			MessageBox(_T("FanHao should not be blank!"));
			return;
		}

		bName1 = bName1.MakeLower();
		bName2 = bName2.MakeLower();
		try
		{
			mySQLtext.Format(_T("SELECT * FROM library WHERE LName1 = '%s' AND LName2 = '%s';"), bName1,bName2);
			mDlg->iAcc2.m_pRecordset = mDlg->iAcc2.m_pConnection->Execute(_bstr_t(mySQLtext), &RecordsAffected, adCmdText);
			if ((mDlg->iAcc2.m_pRecordset->BOF) && (mDlg->iAcc2.m_pRecordset->adoEOF))
			{
				mySQLtext.Format(_T("INSERT INTO library (LName1,LName2,PID,SerieID) VALUES ('%s','%s',1,1);"), bName1, bName2);
				mDlg->iAcc3.m_pRecordset = mDlg->iAcc3.m_pConnection->Execute(_bstr_t(mySQLtext), &RecordsAffected, adCmdText);
				mySQLtext.Format(_T("SELECT @@Identity AS NewLibID;"));
				mDlg->iAcc3.m_pRecordset = mDlg->iAcc3.m_pConnection->Execute(_bstr_t(mySQLtext), &RecordsAffected, adCmdText);
				mLIB = mDlg->iAcc3.m_pRecordset->GetCollect("NewLibID");
				mDlg->echoAndReturn(_T("\r\n") + bName1 + _T(" ") + bName2 + _T(" \thas been added to library!\r\n"));
			}
			else
			{
				mDlg->iAcc2.m_pRecordset->MoveFirst();
				mLIB = mDlg->iAcc2.m_pRecordset->GetCollect("LibID");

				if (check_serie.GetCheck() == 1)
				{
					mySQLtext.Format(_T("UPDATE library SET SerieID = %d WHERE LibID = %d;"), serieNum, mLIB);
					mDlg->iAcc3.m_pRecordset = mDlg->iAcc3.m_pConnection->Execute(_bstr_t(mySQLtext), &RecordsAffected, adCmdText);
				}
			}

			mySQLtext.Format(_T("UPDATE videos SET IsChn=%d,IsNew=%d , LibID = %d WHERE VID = %d;"), bChn, bNew, mLIB, mDlg->g_VID);
			mDlg->iAcc3.m_pRecordset = mDlg->iAcc3.m_pConnection->Execute(_bstr_t(mySQLtext), &RecordsAffected, adCmdText);
		}
		catch (_com_error &e)
		{
			CString eorr;
			eorr.Format(L"DB failed! Code:%08X   Info:%s \r\nDescribe:%s\r\n", e.Error(), (wchar_t*)e.ErrorMessage(), (wchar_t*)e.Description());
			mDlg->echoAndReturn(eorr);
			return;
		}
		CDialogEx::OnOK();
	}
	else
		return;
}


void Modify1Video::OnBnClickedCancel()
{
	// TODO: Add your control notification handler code here
	CDialogEx::OnCancel();
}






void Modify1Video::OnBnClickedButtonClear()
{
	SetDlgItemText(IDC_EDIT_FH1, _T("0"));
	SetDlgItemText(IDC_EDIT_FH2, _T("0"));
}


void Modify1Video::OnClickedCheckChangserie()
{
	if (check_serie.GetCheck() == 1)
	{
		GetDlgItem(IDC_COMBO_SERIE)->EnableWindow(TRUE);
	}
	else
	{
		GetDlgItem(IDC_COMBO_SERIE)->EnableWindow(FALSE);
		combo_serie.SetCurSel(-1);
	}
}
