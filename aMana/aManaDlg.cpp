
// aManaDlg.cpp : implementation file
//

#include "stdafx.h"
#include "aMana.h"
#include "aManaDlg.h"
#include "afxdialogex.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#endif


// CaManaDlg dialog



CaManaDlg::CaManaDlg(CWnd* pParent /*=NULL*/)
	: CDialogEx(CaManaDlg::IDD, pParent)
{
	m_hIcon = AfxGetApp()->LoadIcon(IDR_MAINFRAME);
}

void CaManaDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
	DDX_Control(pDX, IDC_LIST_RECORDS, m_ListCtrl);
	DDX_Control(pDX, IDC_CHECK_CHN, check_chn);
	DDX_Control(pDX, IDC_CHECK_NEW, check_new);
	DDX_Control(pDX, IDC_CHECK_DEL, check_del);
	DDX_Control(pDX, IDC_CHECK_PIC, check_pic);
}

BEGIN_MESSAGE_MAP(CaManaDlg, CDialogEx)
	ON_WM_PAINT()
	ON_WM_QUERYDRAGICON()
	ON_BN_CLICKED(IDC_BUTTON_TEST, &CaManaDlg::OnBnClickedButtonTest)
	ON_BN_CLICKED(IDC_BUTTON_QUERY, &CaManaDlg::OnBnClickedButtonQuery)
	ON_BN_CLICKED(IDC_BUTTON_CHANGEDIR, &CaManaDlg::OnBnClickedButtonChangedir)
	ON_BN_CLICKED(IDC_BUTTON_CLNLIST, &CaManaDlg::OnBnClickedButtonClnlist)
	ON_BN_CLICKED(IDC_BUTTON_NONEW, &CaManaDlg::OnBnClickedButtonNonew)
	ON_BN_CLICKED(IDC_BUTTON_UPDATEDB, &CaManaDlg::OnBnClickedButtonUpdatedb)
	ON_BN_CLICKED(IDC_BUTTON_AUTOMATCH, &CaManaDlg::OnBnClickedButtonAutomatch)
	ON_NOTIFY(NM_DBLCLK, IDC_LIST_RECORDS, &CaManaDlg::OnDblclkListRecords)
	ON_BN_CLICKED(IDC_BUTTON_LIBQ, &CaManaDlg::OnBnClickedButtonLibq)
	ON_BN_CLICKED(IDC_BUTTON_SET2NEW, &CaManaDlg::OnBnClickedButtonSet2new)
END_MESSAGE_MAP()


// CaManaDlg message handlers

BOOL CaManaDlg::OnInitDialog()
{
	CDialogEx::OnInitDialog();

	// Set the icon for this dialog.  The framework does this automatically
	//  when the application's main window is not a dialog
	SetIcon(m_hIcon, TRUE);			// Set big icon
	SetIcon(m_hIcon, FALSE);		// Set small icon

	// TODO: Add extra initialization here
	InitRecordList();
	if (!InitialDatabase())
	{
		GetDlgItem(IDC_BUTTON_UPDATEDB)->EnableWindow(FALSE);
		GetDlgItem(IDC_BUTTON_AUTOMATCH)->EnableWindow(FALSE);
		GetDlgItem(IDC_BUTTON_NONEW)->EnableWindow(FALSE);
		GetDlgItem(IDC_BUTTON_CHANGEDIR)->EnableWindow(FALSE);
		GetDlgItem(IDC_BUTTON_QUERY)->EnableWindow(FALSE);
		GetDlgItem(IDC_BUTTON_CLNLIST)->EnableWindow(FALSE);
		GetDlgItem(IDC_BUTTON_SET2NEW)->EnableWindow(FALSE);
		return TRUE;
	}

	InitRootPath();

	MyStatistics();

	isBusyWorking = false;
	g_VID = 0;
	return TRUE;  // return TRUE  unless you set the focus to a control
}

// If you add a minimize button to your dialog, you will need the code below
//  to draw the icon.  For MFC applications using the document/view model,
//  this is automatically done for you by the framework.

void CaManaDlg::OnPaint()
{
	if (IsIconic())
	{
		CPaintDC dc(this); // device context for painting

		SendMessage(WM_ICONERASEBKGND, reinterpret_cast<WPARAM>(dc.GetSafeHdc()), 0);

		// Center icon in client rectangle
		int cxIcon = GetSystemMetrics(SM_CXICON);
		int cyIcon = GetSystemMetrics(SM_CYICON);
		CRect rect;
		GetClientRect(&rect);
		int x = (rect.Width() - cxIcon + 1) / 2;
		int y = (rect.Height() - cyIcon + 1) / 2;

		// Draw the icon
		dc.DrawIcon(x, y, m_hIcon);
	}
	else
	{
		CDialogEx::OnPaint();
	}
}

// The system calls this function to obtain the cursor to display while the user drags
//  the minimized window.
HCURSOR CaManaDlg::OnQueryDragIcon()
{
	return static_cast<HCURSOR>(m_hIcon);
}



//--------------------about interface------------------------------------
void CaManaDlg::InterfaceInitial()
{
	DashMessage = "";
}

void CaManaDlg::echo(CString text)
{
	DashMessage += text;
	SetDlgItemText(IDC_EDIT_DASHBOARD, DashMessage);
	CEdit *pedit = (CEdit *)GetDlgItem(IDC_EDIT_DASHBOARD);
	int nline = pedit->GetLineCount();
	pedit->LineScroll(nline - 1);
}

void CaManaDlg::echoAndReturn(CString text)
{
	echo(text);
	echo(_T("\r\n"));
}


void CaManaDlg::InitRecordList()
{
	m_ListCtrl.SetExtendedStyle(LVS_EX_GRIDLINES | LVS_EX_FULLROWSELECT|LVS_EX_CHECKBOXES);
	m_ListCtrl.InsertColumn(0, _T("CurrentPath"), LVCFMT_LEFT, 480);
	m_ListCtrl.InsertColumn(1, _T("FileSize"), LVCFMT_CENTER, 90);
	m_ListCtrl.InsertColumn(2, _T("Fan"), LVCFMT_CENTER, 50);
	m_ListCtrl.InsertColumn(3, _T("Hao"), LVCFMT_CENTER, 50);
	m_ListCtrl.InsertColumn(4, _T("Pic"), LVCFMT_CENTER, 35);
	m_ListCtrl.InsertColumn(5, _T("Serie"), LVCFMT_CENTER, 180);
	m_ListCtrl.InsertColumn(6, _T("Chn"), LVCFMT_CENTER, 35);
	m_ListCtrl.InsertColumn(7, _T("New"), LVCFMT_CENTER, 39);
	m_ListCtrl.InsertColumn(8, _T("Del"), LVCFMT_CENTER, 35);
	m_ListCtrl.InsertColumn(9, _T("Date"), LVCFMT_CENTER, 80);
	m_ListCtrl.InsertColumn(10, _T("VID"), LVCFMT_CENTER, 50);
	m_ListCtrl.InsertColumn(11, _T("LibID"), LVCFMT_CENTER, 50);
}

void CaManaDlg::Show1LineInList(CString pPath,CString pSize, CString pFan,CString pHao, CString pPic,CString pSerie,CString pChn,CString pNew,CString pDel, CString pDate, CString pVID,CString pLibID)
{
	int nRow = m_ListCtrl.InsertItem(0, pPath);
	m_ListCtrl.SetItemText(nRow, 1, pSize);
	m_ListCtrl.SetItemText(nRow, 2, pFan);
	m_ListCtrl.SetItemText(nRow, 3, pHao);
	m_ListCtrl.SetItemText(nRow, 4, pPic);
	m_ListCtrl.SetItemText(nRow, 5, pSerie);
	m_ListCtrl.SetItemText(nRow, 6, pChn);
	m_ListCtrl.SetItemText(nRow, 7, pNew);
	m_ListCtrl.SetItemText(nRow, 8, pDel);
	m_ListCtrl.SetItemText(nRow, 9, pDate);
	m_ListCtrl.SetItemText(nRow, 10, pVID);
	m_ListCtrl.SetItemText(nRow, 11, pLibID);
}

void CaManaDlg::OnBnClickedButtonClnlist()
{
	m_ListCtrl.DeleteAllItems();
}


BOOL CaManaDlg::PreTranslateMessage(MSG* pMsg)		//shortcut keys
{
	if (pMsg->message == WM_KEYDOWN&&pMsg->wParam == VK_ESCAPE)
	{
		SetDlgItemText(IDC_EDIT_QUERY1, _T(""));
		SetDlgItemText(IDC_EDIT_QUERY2, _T(""));
		check_chn.SetCheck(0);
		check_del.SetCheck(0);
		check_new.SetCheck(0);
		return TRUE;
	}
	if (pMsg->message == WM_KEYDOWN&&pMsg->wParam == VK_RETURN)
	{
		OnBnClickedButtonQuery();
		return TRUE;
	}
	if (pMsg->message == WM_KEYDOWN&&pMsg->wParam == VK_F4)
	{
		OnBnClickedButtonClnlist();
		return TRUE;
	}
	return CDialog::PreTranslateMessage(pMsg);
}

void CaManaDlg::OnOK()
{

}

void CaManaDlg::OnCancel()
{

	CDialogEx::OnCancel();
}

//about interface end


//--------------------------about database---------------------------------
BOOL CaManaDlg::InitialDatabase()		//return:  1:success; 0:fail,error info will show on dashboard
{
	//初始化OLE/COM环境, 为连接access准备，清空com
	if (!AfxOleInit())
	{
		echoAndReturn(_T("OLE/COM initialization failed! Close this window! \r\n"));
		return 0;
	}

	//打开mdb数据库语句
	const _bstr_t strSRC = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=aa.mdb;Jet OLEDB:Database Password = ...;Persist Security Info=False;";
	iAcc1.m_pConnection.CreateInstance("ADODB.Connection");
	iAcc2.m_pConnection.CreateInstance("ADODB.Connection");
	iAcc3.m_pConnection.CreateInstance("ADODB.Connection");
	iAcc4.m_pConnection.CreateInstance("ADODB.Connection");
	try{
		iAcc1.m_pConnection->Open(strSRC, "", "", -1);
		iAcc2.m_pConnection->Open(strSRC, "", "", -1);
		iAcc3.m_pConnection->Open(strSRC, "", "", -1);
		iAcc4.m_pConnection->Open(strSRC, "", "", -1);
	}
	catch (_com_error &e)
	{
		CString eorr;
		eorr.Format(L"DB failed! Code:%08X   Info:%s \r\nDescribe:%s\r\n", e.Error(), (wchar_t*)e.ErrorMessage(), (wchar_t*)e.Description());
		echoAndReturn(eorr);
		return 0;
	}
	return 1;
}

//iAcc3
void CaManaDlg::OnBnClickedButtonNonew()
{
	if (MessageBox(_T("Are you sure clear new flags?"), _T(""), MB_OKCANCEL) == IDOK)
	{
		isBusyWorking = true;
		GetDlgItem(IDC_BUTTON_UPDATEDB)->EnableWindow(FALSE);
		GetDlgItem(IDC_BUTTON_AUTOMATCH)->EnableWindow(FALSE);
		GetDlgItem(IDC_BUTTON_NONEW)->EnableWindow(FALSE);
		GetDlgItem(IDC_BUTTON_SET2NEW)->EnableWindow(FALSE);


		CString mySQLtext;
		_variant_t RecordsAffected;
		mySQLtext.Format(_T("UPDATE videos SET IsNew=0;"));
		try
		{
			iAcc3.m_pRecordset = iAcc3.m_pConnection->Execute(_bstr_t(mySQLtext), &RecordsAffected, adCmdText);
		}
		catch (_com_error &e)
		{
			CString eorr;
			eorr.Format(L"DB failed! Code:%08X   Info:%s \r\nDescribe:%s\r\n", e.Error(), (wchar_t*)e.ErrorMessage(), (wchar_t*)e.Description());
			echoAndReturn(eorr);
			return;
		}
		echoAndReturn(_T("\r\n---------------->>>>>>All records are set to No New!<<<<<<---------------\r\n\r\n"));
		isBusyWorking = false;
		GetDlgItem(IDC_BUTTON_UPDATEDB)->EnableWindow(TRUE);
		GetDlgItem(IDC_BUTTON_AUTOMATCH)->EnableWindow(TRUE);
		GetDlgItem(IDC_BUTTON_NONEW)->EnableWindow(TRUE);
		GetDlgItem(IDC_BUTTON_SET2NEW)->EnableWindow(TRUE);

	}
}


//////////////////////////////////////////
//this function is not in use currently
/////////////////////////////////////////
BOOL CaManaDlg::DatabaseExeSQL(int AdoNum, CString SQLtext)		//return:  1:success; 0:fail,error info will show on dashboard
{
	Ado2Access tmpAdo;
	_variant_t RecordsAffected;
	if (AdoNum == 1) tmpAdo = iAcc1;
	else if (AdoNum == 2) tmpAdo = iAcc2;
	else if (AdoNum == 3) tmpAdo = iAcc3;
	else return 0;
	try
	{
		tmpAdo.m_pRecordset = tmpAdo.m_pConnection->Execute(_bstr_t(SQLtext), &RecordsAffected, adCmdText);
	}
	catch (_com_error &e)
	{
		CString eorr;
		eorr.Format(L"DB failed! Code:%08X   Info:%s \r\nDescribe:%s\r\n", e.Error(), (wchar_t*)e.ErrorMessage(), (wchar_t*)e.Description());
		echoAndReturn(eorr);
		return 0;
	}
	return 1;
}
//about database end


//-------------------about search---------------------------
void CaManaDlg::OnBnClickedButtonQuery()
{
	m_ListCtrl.SetRedraw(FALSE);			//solve display slowness
	GetDlgItem(IDC_BUTTON_QUERY)->EnableWindow(FALSE);
	Show1LineInList(NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL);
	Show1LineInList(NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL);
	Show1LineInList(NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL);
	Show1LineInList(_T("<-------------------------Search Complete!-------------------------->"), NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL,NULL,NULL);
	CWinThread* h_thread;
	h_thread = AfxBeginThread(ThreadQuery, (LPVOID) this);
}

UINT CaManaDlg::ThreadQuery(LPVOID pParam)
{
	CaManaDlg *p = (CaManaDlg *)pParam;
	p->FunQuery();
	p->GetDlgItem(IDC_BUTTON_QUERY)->EnableWindow(TRUE);
	p->m_ListCtrl.SetRedraw(TRUE);
	return 0;
	AfxEndThread;
}

//iAcc1
void CaManaDlg::FunQuery()	
{
	int RecordCount=0;
	CString CON1, CON2;
	GetDlgItemText(IDC_EDIT_QUERY1, CON1);
	GetDlgItemText(IDC_EDIT_QUERY2, CON2);
	_variant_t RecordsAffected;
	CString mySQLtext;

	CString vFilePath, vFileSize, vFileName, vFan, vHao, vSerie, vDate,vPic;
	int vDel, vChn, vNew;
	long vVID,vLibID;
	CString sDel, sChn, sNew, sVID,sLibID;

	mySQLtext.Format(_T("SELECT * FROM qmain WHERE InStr(1,LCase(VFileName),LCase('%s'),0) AND InStr(1,LCase(VFileName),LCase('%s'),0) "), CON1, CON2);
	//checkbox条件
	if (check_chn.GetCheck() == 1) mySQLtext += _T(" AND IsChn=1");
	if (check_del.GetCheck() == 1) mySQLtext += _T(" AND IsDel=1");
	if (check_new.GetCheck() == 1) mySQLtext += _T(" AND IsNew=1");
	if (check_pic.GetCheck() == 1) mySQLtext += _T(" AND PFilePath <> '0'");

	mySQLtext += _T(" ORDER BY VFilePath ASC;");

	try{iAcc1.m_pRecordset = iAcc1.m_pConnection->Execute(_bstr_t(mySQLtext), &RecordsAffected, adCmdText);}
	catch (_com_error &e)
	{
		CString eorr;
		eorr.Format(L"DB failed! Code:%08X   Info:%s \r\nDescribe:%s\r\n", e.Error(), (wchar_t*)e.ErrorMessage(), (wchar_t*)e.Description());
		echoAndReturn(eorr);
		return ;
	}

	try
	{
		if ((iAcc1.m_pRecordset->BOF) && (iAcc1.m_pRecordset->adoEOF))		//empty recordset
		{}
		else
		{
			iAcc1.m_pRecordset->MoveFirst();
			while (!iAcc1.m_pRecordset->adoEOF)
			{
				vFilePath = iAcc1.m_pRecordset->GetCollect("VFilePath");
				vFileSize = iAcc1.m_pRecordset->GetCollect("VFileSize");
				vFileName = iAcc1.m_pRecordset->GetCollect("VFileName");
				vFan = iAcc1.m_pRecordset->GetCollect("LName1");
				vHao = iAcc1.m_pRecordset->GetCollect("LName2");
				vSerie = iAcc1.m_pRecordset->GetCollect("SerieName");
				vDate = iAcc1.m_pRecordset->GetCollect("CDate");
				vVID = iAcc1.m_pRecordset->GetCollect("VID");
				vChn = iAcc1.m_pRecordset->GetCollect("IsChn");
				vDel = iAcc1.m_pRecordset->GetCollect("IsDel");
				vNew = iAcc1.m_pRecordset->GetCollect("IsNew");
				vPic = iAcc1.m_pRecordset->GetCollect("PFilePath");

				if (vPic == _T("0")) vPic = ""; else vPic = "PIC";
				if (vChn == 1) sChn = _T("CHN"); else sChn = "";
				if (vNew == 1) sNew = _T("NEW"); else sNew = "";
				if (vDel == 1) sDel = _T("DEL"); else sDel = "";
				if (vSerie == _T("0")) vSerie = "";

				sVID.Format(_T("%d"), vVID);

				Show1LineInList(vFilePath, vFileSize, vFan, vHao,vPic,vSerie, sChn, sNew, sDel, vDate, sVID,NULL);
				RecordCount++;
				iAcc1.m_pRecordset->MoveNext();
			}

		}
	}
	catch (_com_error &e)
	{
		CString eorr;
		eorr.Format(L"DB failed! Code:%08X   Info:%s \r\nDescribe:%s\r\n", e.Error(), (wchar_t*)e.ErrorMessage(), (wchar_t*)e.Description());
		echoAndReturn(eorr);
		return ;
	}

	m_ListCtrl.EnsureVisible(0, FALSE);
	CString aaa;
	aaa.Format(_T("%d"), RecordCount);
//	echoAndReturn(aaa + _T(" Records Found!\r\n"));
	aaa = _T("******* ")+aaa+_T(" Records Found! *******\r\n");
	Show1LineInList(aaa, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL,NULL,NULL);
	return;

}

//about search end


//-------------------about root path---------------------------
//iAcc1
void CaManaDlg::InitRootPath()
{
	CString mySQLtext;
	_variant_t RecordsAffected;
	
	CString vDir,vPath;

	mySQLtext.Format(_T("SELECT * FROM dir;"));
	try{ iAcc1.m_pRecordset = iAcc1.m_pConnection->Execute(_bstr_t(mySQLtext), &RecordsAffected, adCmdText); }
	catch (_com_error &e)
	{
		CString eorr;
		eorr.Format(L"DB failed! Code:%08X   Info:%s \r\nDescribe:%s\r\n", e.Error(), (wchar_t*)e.ErrorMessage(), (wchar_t*)e.Description());
		echoAndReturn(eorr);
		return;
	}
	try
	{
		if ((iAcc1.m_pRecordset->BOF) && (iAcc1.m_pRecordset->adoEOF))		//empty recordset
			echoAndReturn(_T("Directories in DB fail!"));
		else
		{
			iAcc1.m_pRecordset->MoveFirst();
			while (!iAcc1.m_pRecordset->adoEOF)
			{
				vDir = iAcc1.m_pRecordset->GetCollect("DirCata");
				vPath = iAcc1.m_pRecordset->GetCollect("DirPath");
				if (vPath == _T("0")) vPath = "";
				if (vDir == _T("D1")){ Dir[1] = vPath; SetDlgItemText(IDC_EDIT_DIR1, Dir[1]); }
				if (vDir == _T("D2")){ Dir[2] = vPath; SetDlgItemText(IDC_EDIT_DIR2, Dir[2]); }
				if (vDir == _T("D3")){ Dir[3] = vPath; SetDlgItemText(IDC_EDIT_DIR3, Dir[3]); }
				if (vDir == _T("D4")){ Dir[4] = vPath; SetDlgItemText(IDC_EDIT_DIR4, Dir[4]); }
				if (vDir == _T("D5")){ Dir[5] = vPath; SetDlgItemText(IDC_EDIT_DIR5, Dir[5]); }
				if (vDir == _T("D6")){ Dir[6] = vPath; SetDlgItemText(IDC_EDIT_DIR6, Dir[6]); }
				if (vDir == _T("D7")){ Dir[7] = vPath; SetDlgItemText(IDC_EDIT_DIR7, Dir[7]); }
				if (vDir == _T("PIC")){ PicPath = vPath; SetDlgItemText(IDC_EDIT_DIRPIC, PicPath); }
				iAcc1.m_pRecordset->MoveNext();
			}

		}
	}
	catch (_com_error &e)
	{
		CString eorr;
		eorr.Format(L"DB failed! Code:%08X   Info:%s \r\nDescribe:%s\r\n", e.Error(), (wchar_t*)e.ErrorMessage(), (wchar_t*)e.Description());
		echoAndReturn(eorr);
		return;
	}
	ChangeRootPathStatus = 0;
	GetDlgItem(IDC_EDIT_DIR1)->EnableWindow(FALSE);
	GetDlgItem(IDC_EDIT_DIR2)->EnableWindow(FALSE);
	GetDlgItem(IDC_EDIT_DIR3)->EnableWindow(FALSE);
	GetDlgItem(IDC_EDIT_DIR4)->EnableWindow(FALSE);
	GetDlgItem(IDC_EDIT_DIR5)->EnableWindow(FALSE);
	GetDlgItem(IDC_EDIT_DIR6)->EnableWindow(FALSE);
	GetDlgItem(IDC_EDIT_DIR7)->EnableWindow(FALSE);
	GetDlgItem(IDC_EDIT_DIRPIC)->EnableWindow(FALSE);
}

//iAcc3
void CaManaDlg::OnBnClickedButtonChangedir()
{
	if (ChangeRootPathStatus == 0)
	{
		GetDlgItem(IDC_EDIT_DIR1)->EnableWindow(TRUE);
		GetDlgItem(IDC_EDIT_DIR2)->EnableWindow(TRUE);
		GetDlgItem(IDC_EDIT_DIR3)->EnableWindow(TRUE);
		GetDlgItem(IDC_EDIT_DIR4)->EnableWindow(TRUE);
		GetDlgItem(IDC_EDIT_DIR5)->EnableWindow(TRUE);
		GetDlgItem(IDC_EDIT_DIR6)->EnableWindow(TRUE);
		GetDlgItem(IDC_EDIT_DIR7)->EnableWindow(TRUE);
		GetDlgItem(IDC_EDIT_DIRPIC)->EnableWindow(TRUE);
		GetDlgItem(IDC_BUTTON_UPDATEDB)->EnableWindow(FALSE);
		GetDlgItem(IDC_BUTTON_AUTOMATCH)->EnableWindow(FALSE);
		GetDlgItem(IDC_BUTTON_NONEW)->EnableWindow(FALSE);
		GetDlgItem(IDC_BUTTON_SET2NEW)->EnableWindow(FALSE);
		ChangeRootPathStatus = 1;
	}
	else if (ChangeRootPathStatus == 1)
	{
		_variant_t RecordsAffected;
		CString mySQLtext;
		
		CString pDir[8];
		CString pPicPath;
		int i, j;

		GetDlgItemText(IDC_EDIT_DIR1, pDir[1]);
		GetDlgItemText(IDC_EDIT_DIR2, pDir[2]);
		GetDlgItemText(IDC_EDIT_DIR3, pDir[3]);
		GetDlgItemText(IDC_EDIT_DIR4, pDir[4]);
		GetDlgItemText(IDC_EDIT_DIR5, pDir[5]);
		GetDlgItemText(IDC_EDIT_DIR6, pDir[6]);
		GetDlgItemText(IDC_EDIT_DIR7, pDir[7]);
		GetDlgItemText(IDC_EDIT_DIRPIC, pPicPath);

		if (pPicPath == "") { AfxMessageBox(_T("Need pic path!")); return; }		//pic dir do not allowed to be empty
		if (pPicPath.Right(1) != "\\") 	pPicPath += "\\";							//add '\'
		if (!PathFileExists(pPicPath)){ AfxMessageBox(_T("Pic path does not exist!")); return; }	//check path exist

		pPicPath.MakeUpper();

		for (i = 1; i < 8; i++)			//add '\'
		{
			if ((pDir[i].Right(1) != "\\") && (pDir[i] != ""))	pDir[i] += "\\";
		}

		for (i = 1; i < 8; i++)		//check path exist
		{
			if ((!PathFileExists(pDir[i])) && (pDir[i] != ""))
			{
				CString a;
				a = pDir[i] + _T(" does not exist!");
				AfxMessageBox(a);
				return;
			}
		}

		for (i = 1; i < 7; i++)		//cannot be same dir
		{
			for (j = i + 1; j < 8; j++)
			{
				if ((pDir[i].MakeUpper() == pDir[j].MakeUpper()) && (pDir[i] != ""))
				{
					AfxMessageBox(_T("Have the same path!"));
					return;
				}
			}

		}

		//change value to global para
		for (i = 1; i < 8; i++)
			Dir[i] = pDir[i];
		PicPath = pPicPath;

		//output in dashboard
		CString aa;
		aa =_T("PIC: ")+PicPath+ _T("\r\nD1: ") + Dir[1] + _T("\r\nD2: ") + Dir[2] + _T("\r\nD3: ") + Dir[3] + _T("\r\nD4: ") + Dir[4] + _T("\r\nD5: ") + Dir[5] + _T("\r\nD6: ") + Dir[6] + _T("\r\nD7: ") + Dir[7];
		echoAndReturn(aa);
		echoAndReturn(_T(""));
		echoAndReturn(_T(""));

		for (i = 1; i < 8; i++)		//change empty to 0 before write db
		{
			if (pDir[i] == _T("")) pDir[i] = _T("0");
		}
		
		try	//change db
		{
			mySQLtext.Format(_T("UPDATE dir SET DirPath = '%s'  WHERE DirCata = 'D1'"), pDir[1]);
			iAcc3.m_pRecordset = iAcc3.m_pConnection->Execute(_bstr_t(mySQLtext), &RecordsAffected, adCmdText);

			mySQLtext.Format(_T("UPDATE dir SET DirPath = '%s'  WHERE DirCata = 'D2'"), pDir[2]);
			iAcc3.m_pRecordset = iAcc3.m_pConnection->Execute(_bstr_t(mySQLtext), &RecordsAffected, adCmdText);

			mySQLtext.Format(_T("UPDATE dir SET DirPath = '%s'  WHERE DirCata = 'D3'"), pDir[3]);
			iAcc3.m_pRecordset = iAcc3.m_pConnection->Execute(_bstr_t(mySQLtext), &RecordsAffected, adCmdText);

			mySQLtext.Format(_T("UPDATE dir SET DirPath = '%s'  WHERE DirCata = 'D4'"), pDir[4]);
			iAcc3.m_pRecordset = iAcc3.m_pConnection->Execute(_bstr_t(mySQLtext), &RecordsAffected, adCmdText);

			mySQLtext.Format(_T("UPDATE dir SET DirPath = '%s'  WHERE DirCata = 'D5'"), pDir[5]);
			iAcc3.m_pRecordset = iAcc3.m_pConnection->Execute(_bstr_t(mySQLtext), &RecordsAffected, adCmdText);

			mySQLtext.Format(_T("UPDATE dir SET DirPath = '%s'  WHERE DirCata = 'D6'"), pDir[6]);
			iAcc3.m_pRecordset = iAcc3.m_pConnection->Execute(_bstr_t(mySQLtext), &RecordsAffected, adCmdText);

			mySQLtext.Format(_T("UPDATE dir SET DirPath = '%s'  WHERE DirCata = 'D7'"), pDir[7]);
			iAcc3.m_pRecordset = iAcc3.m_pConnection->Execute(_bstr_t(mySQLtext), &RecordsAffected, adCmdText);

			mySQLtext.Format(_T("UPDATE dir SET DirPath = '%s'  WHERE DirCata = 'PIC'"), pPicPath);
			iAcc3.m_pRecordset = iAcc3.m_pConnection->Execute(_bstr_t(mySQLtext), &RecordsAffected, adCmdText);
		}
		catch (_com_error &e)
		{
			CString eorr;
			eorr.Format(L"DB failed! Code:%08X   Info:%s \r\nDescribe:%s\r\n", e.Error(), (wchar_t*)e.ErrorMessage(), (wchar_t*)e.Description());
			echoAndReturn(eorr);
			return;
		}
		
		ChangeRootPathStatus = 0;
		GetDlgItem(IDC_EDIT_DIR1)->EnableWindow(FALSE);
		GetDlgItem(IDC_EDIT_DIR2)->EnableWindow(FALSE);
		GetDlgItem(IDC_EDIT_DIR3)->EnableWindow(FALSE);
		GetDlgItem(IDC_EDIT_DIR4)->EnableWindow(FALSE);
		GetDlgItem(IDC_EDIT_DIR5)->EnableWindow(FALSE);
		GetDlgItem(IDC_EDIT_DIR6)->EnableWindow(FALSE);
		GetDlgItem(IDC_EDIT_DIR7)->EnableWindow(FALSE);
		GetDlgItem(IDC_EDIT_DIRPIC)->EnableWindow(FALSE);
		GetDlgItem(IDC_BUTTON_UPDATEDB)->EnableWindow(TRUE);
		GetDlgItem(IDC_BUTTON_AUTOMATCH)->EnableWindow(TRUE);
		GetDlgItem(IDC_BUTTON_NONEW)->EnableWindow(TRUE);
		GetDlgItem(IDC_BUTTON_SET2NEW)->EnableWindow(TRUE);

	}

}

//about root path end

//-------------------about pic---------------------------
void CaManaDlg::SearchPicFiles(CString MyPath)
{
	if (MyPath.Right(1) != "\\")	MyPath += "\\";		//文件夹加斜杠	
	MyPath += "*.*";								//文件路径下的文件类型

	int FileFoundCount = 0;				//用于找到文件的计数
	char buf[1024];
	strcpy_s(buf, 1024, CT2CA(MyPath));
	CFileFind finder;						//实例化CFileFind类
	BOOL bWorking = finder.FindFile(MyPath);
	while (bWorking)
	{
		bWorking = finder.FindNextFileW();					//找下一个文件
		if (finder.IsDirectory() && !finder.IsDots())		//如果找到目录且不是目录本身
			SearchPicFiles(finder.GetFilePath());				//递归调用本函数
		else if (!finder.IsDirectory() && !finder.IsDots())	//找到一个文件
		{
			if (finder.GetFileName().Left(1) == "$")
				continue;
			
			DWORD length = finder.GetLength();		//文件大小，下面是转换成cstring
			CString a, last3;						//a用来转换文件长度格式，last用来判断文件类型
			a.Format(_T("%u"), length);				//注意这里用%u，正整数
			last3 = finder.GetFileName().Right(4).MakeLower();
			if (last3 == ".jpg")
				UpdatePicDB(finder.GetFileName(), finder.GetFilePath());	
		}

	}

}

//iAcc2,iAcc3
void CaManaDlg::UpdatePicDB(CString MyFileName, CString MyFilePath)
{
	CString mySQLtext;				
	_variant_t RecordsAffected;

	mySQLtext.Format(_T("SELECT * FROM pics WHERE PFileName = '%s' "), MyFileName);
	try
	{
		iAcc2.m_pRecordset = iAcc2.m_pConnection->Execute(_bstr_t(mySQLtext), &RecordsAffected, adCmdText);
		if ((iAcc2.m_pRecordset->BOF) && (iAcc2.m_pRecordset->adoEOF))			//no record
		{
			mySQLtext.Format(_T("INSERT INTO pics ( PFileName, PFilePath) VALUES ('%s','%s');"), MyFileName, MyFilePath);
			iAcc3.m_pRecordset = iAcc3.m_pConnection->Execute(_bstr_t(mySQLtext), &RecordsAffected, adCmdText);
			echoAndReturn(MyFileName + _T("\t\tinsert to DB!"));
		}
		else    //update file path
		{
			mySQLtext.Format(_T("UPDATE pics SET PFilePath = '%s' WHERE PFileName = '%s'; "), MyFilePath, MyFileName);
			iAcc3.m_pRecordset = iAcc3.m_pConnection->Execute(_bstr_t(mySQLtext), &RecordsAffected, adCmdText);
		}
	}
	catch (_com_error &e)
	{
		CString eorr;
		eorr.Format(L"DB failed! Code:%08X   Info:%s \r\nDescribe:%s\r\n", e.Error(), (wchar_t*)e.ErrorMessage(), (wchar_t*)e.Description());
		echoAndReturn(eorr);
		return;
	}
}

//about pic end

//-------------------about video---------------------------
void CaManaDlg::SearchVideoFiles(CString MyPath)
{
	if (MyPath.Right(1) != "\\")	MyPath += "\\";		//文件夹加斜杠	
	MyPath += "*.*";								//文件路径下的文件类型

	char buf[1024];
	strcpy_s(buf, 1024, CT2CA(MyPath));
	CFileFind finder;						//实例化CFileFind类
	BOOL bWorking = finder.FindFile(MyPath);
	while (bWorking)
	{
		bWorking = finder.FindNextFileW();					//找下一个文件
		if (finder.IsDirectory() && !finder.IsDots())		//如果找到目录且不是目录本身
			SearchVideoFiles(finder.GetFilePath());				//递归调用本函数
		else if (!finder.IsDirectory() && !finder.IsDots())	//找到一个文件
		{
			if (finder.GetFileName().Left(1) == "$")
				continue;

			DWORD length = finder.GetLength();		//文件大小，下面是转换成cstring
			CString a,b, c,d,e;						//a用来转换文件长度格式，b为修改时间，cde用来判断文件类型
			a.Format(_T("%u"), length);				//注意这里用%u，正整数
			c = finder.GetFileName().Right(4).MakeLower();
			d = finder.GetFileName().Right(5).MakeLower();
			e = finder.GetFileName().Right(3).MakeLower();
			CTime myTime;
			finder.GetLastWriteTime(myTime);
			b.Format(_T("%d-%d-%d %d:%d:%d"), myTime.GetYear(), myTime.GetMonth(), myTime.GetDay(), myTime.GetHour(), myTime.GetMinute(), myTime.GetSecond());
			if ((c == ".avi") || (c == ".mp4") || (c == ".3gp") || (c == ".mpg") || (c == ".vob") || (d == ".rmvb") || (c == ".wmv") || (c == ".mkv") || (c == ".mov") || (e == ".rm") || (d == ".divx") || (c == ".flv"))
			{
				UpdateVidoeDB(finder.GetFileName(), finder.GetFilePath(), a, b);
			}
				
		}

	}
}

//iAcc2,iAcc3
void CaManaDlg::UpdateVidoeDB(CString MyFileName, CString MyFilePath, CString MyFileSize, CString MyCDate)
{
	CString mySQLtext;
	_variant_t RecordsAffected;


	mySQLtext.Format(_T("SELECT * FROM videos WHERE VFileName = '%s' AND VFileSize = '%s';"), MyFileName, MyFileSize);

	try
	{
		iAcc2.m_pRecordset = iAcc2.m_pConnection->Execute(_bstr_t(mySQLtext), &RecordsAffected, adCmdText);
		if ((iAcc2.m_pRecordset->BOF) && (iAcc2.m_pRecordset->adoEOF))			//no record
		{
			mySQLtext.Format(_T("INSERT INTO videos ( VFileName, VFilePath , VFileSize , IsDel,IsChn,IsNew,LibID,CDate ) VALUES ('%s','%s','%s',0,0,1,1,'%s');"), MyFileName, MyFilePath,MyFileSize,MyCDate);
			iAcc3.m_pRecordset = iAcc3.m_pConnection->Execute(_bstr_t(mySQLtext), &RecordsAffected, adCmdText);
			echoAndReturn(MyFileName + _T("\t\tinsert to DB!"));
		}
		else    //update file path
		{
			mySQLtext.Format(_T("UPDATE videos SET VFilePath = '%s',CDate = '%s' WHERE VFileName = '%s' AND VFileSize = '%s'; "), MyFilePath,MyCDate, MyFileName,MyFileSize);
			iAcc3.m_pRecordset = iAcc3.m_pConnection->Execute(_bstr_t(mySQLtext), &RecordsAffected, adCmdText);
		}
	}
	catch (_com_error &e)
	{
		CString eorr;
		eorr.Format(L"DB failed! Code:%08X   Info:%s \r\nDescribe:%s\r\n", e.Error(), (wchar_t*)e.ErrorMessage(), (wchar_t*)e.Description());
		echoAndReturn(eorr);
		return;
	}
}

//iAcc2,iAcc3
void CaManaDlg::CheckVideoDel()
{
	CString mySQLtext;
	_variant_t RecordsAffected;

	CString MyPath, MyName, MySize;
	int RecordCount;
	int vdel;
	int i;


	mySQLtext.Format(_T("SELECT count(*) AS Expr1000 FROM videos;"));
	try
	{
//		iAcc2.m_pRecordset = iAcc2.m_pConnection->Execute(_bstr_t(mySQLtext), &RecordsAffected, adCmdText);
//		RecordCount = iAcc2.m_pRecordset->GetCollect("Expr1000");

		mySQLtext.Format(_T("SELECT * FROM videos;"));
		iAcc2.m_pRecordset = iAcc2.m_pConnection->Execute(_bstr_t(mySQLtext), &RecordsAffected, adCmdText);
		if ((iAcc2.m_pRecordset->BOF) && (iAcc2.m_pRecordset->adoEOF))
		{
			return;
		}
		iAcc2.m_pRecordset->MoveFirst();
		while (!iAcc2.m_pRecordset->adoEOF)
//		for (i = 0; i < RecordCount; i++)
		{
			MyPath = iAcc2.m_pRecordset->GetCollect("VFilePath");
			MyName = iAcc2.m_pRecordset->GetCollect("VFileName");
			MySize = iAcc2.m_pRecordset->GetCollect("VFileSize");
			vdel = iAcc2.m_pRecordset->GetCollect("IsDel");

			if (!PathFileExists(MyPath))		//如果路径不存在文件，说明已经删除
			{
				mySQLtext.Format(_T("UPDATE videos SET IsDel=1,IsNew=0,Cdate='0' WHERE VFileName = '%s' AND VFileSize = '%s'; "),MyName,MySize);
				iAcc3.m_pRecordset = iAcc3.m_pConnection->Execute(_bstr_t(mySQLtext), &RecordsAffected, adCmdText);
				if (vdel==0)
					echoAndReturn(MyName + _T("\t\thas been deleted!"));
			}
			else if (vdel == 1)
			{
				mySQLtext.Format(_T("UPDATE videos SET IsDel = 0 ,IsNew = 1 WHERE VFilePath = '%s'  ;"), MyPath);
				iAcc3.m_pRecordset = iAcc3.m_pConnection->Execute(_bstr_t(mySQLtext), &RecordsAffected, adCmdText);
				echoAndReturn(MyName + _T("\t\tturns to new again!"));
			}

			iAcc2.m_pRecordset->MoveNext();
		}
	}

	catch (_com_error &e)
	{
		CString eorr;
		eorr.Format(L"DB failed! Code:%08X   Info:%s \r\nDescribe:%s\r\n", e.Error(), (wchar_t*)e.ErrorMessage(), (wchar_t*)e.Description());
		echoAndReturn(eorr);
		return;
	}



}
//about video end

//iAcc2
void CaManaDlg::MyStatistics()
{
	_variant_t RecordsAffected;
	CString mySQLtext;

	CString sChn,sDel,sVideo,sNew,sPic,sSerie,sStar;
	CString sResult;
	try
	{

		mySQLtext.Format(_T("SELECT * FROM chnsum"));
		iAcc2.m_pRecordset = iAcc2.m_pConnection->Execute(_bstr_t(mySQLtext), &RecordsAffected, adCmdText);
		sChn = iAcc2.m_pRecordset->GetCollect("chnsum");

		mySQLtext.Format(_T("SELECT * FROM delsum"));
		iAcc2.m_pRecordset = iAcc2.m_pConnection->Execute(_bstr_t(mySQLtext), &RecordsAffected, adCmdText);
		sDel = iAcc2.m_pRecordset->GetCollect("delsum");

		mySQLtext.Format(_T("SELECT * FROM picsum"));
		iAcc2.m_pRecordset = iAcc2.m_pConnection->Execute(_bstr_t(mySQLtext), &RecordsAffected, adCmdText);
		sPic = iAcc2.m_pRecordset->GetCollect("picsum");

		mySQLtext.Format(_T("SELECT * FROM videosum"));
		iAcc2.m_pRecordset = iAcc2.m_pConnection->Execute(_bstr_t(mySQLtext), &RecordsAffected, adCmdText);
		sVideo = iAcc2.m_pRecordset->GetCollect("videosum");

		mySQLtext.Format(_T("SELECT * FROM newsum"));
		iAcc2.m_pRecordset = iAcc2.m_pConnection->Execute(_bstr_t(mySQLtext), &RecordsAffected, adCmdText);
		sNew = iAcc2.m_pRecordset->GetCollect("newsum");

		mySQLtext.Format(_T("SELECT * FROM seriesum"));
		iAcc2.m_pRecordset = iAcc2.m_pConnection->Execute(_bstr_t(mySQLtext), &RecordsAffected, adCmdText);
		sSerie = iAcc2.m_pRecordset->GetCollect("seriesum");

		mySQLtext.Format(_T("SELECT * FROM starsum"));
		iAcc2.m_pRecordset = iAcc2.m_pConnection->Execute(_bstr_t(mySQLtext), &RecordsAffected, adCmdText);
		sStar = iAcc2.m_pRecordset->GetCollect("starsum");
	}
	catch (_com_error &e)
	{
		CString eorr;
		eorr.Format(L"DB failed! Code:%08X   Info:%s \r\nDescribe:%s\r\n", e.Error(), (wchar_t*)e.ErrorMessage(), (wchar_t*)e.Description());
		echoAndReturn(eorr);
		return;
	}


	sResult = _T("Videos:\t") + sVideo + _T("\r\nChinese:\t") + sChn + _T("\r\nNew:\t\t") + sNew + _T("\r\nDeleted:\t") + sDel + _T("\r\nStars:\t\t") + sStar + _T("\r\nSeries:\t\t") + sSerie + _T("\r\nPictures:\t") + sPic + _T("\r\n");
	SetDlgItemText(IDC_STATIC_ST, sResult);
}

void CaManaDlg::OnBnClickedButtonUpdatedb()
{
	if (MessageBox(_T("Are you sure to update videos&pictures?"), _T(""), MB_OKCANCEL) == IDOK)
	{
		isBusyWorking = true;
		GetDlgItem(IDC_BUTTON_UPDATEDB)->EnableWindow(FALSE);
		GetDlgItem(IDC_BUTTON_AUTOMATCH)->EnableWindow(FALSE);
		GetDlgItem(IDC_BUTTON_NONEW)->EnableWindow(FALSE);
		GetDlgItem(IDC_BUTTON_SET2NEW)->EnableWindow(FALSE);


		CWinThread* h_thread;
		h_thread = AfxBeginThread(ThreadUpdateDB, (LPVOID) this);
	}

}

UINT CaManaDlg::ThreadUpdateDB(LPVOID pParam)
{
	CaManaDlg *p = (CaManaDlg *)pParam;
	
	p->echo(_T("\r\n\r\n---------------------------------------->>Update Starts !!!<<----------------------------------------\r\n"));
	p->echo(_T("**************************************************\r\nNew Pictures:\r\n\r\n"));

	p->SearchPicFiles(p->PicPath);
	p->echo(_T("\r\n**************************************************\r\n\r\nNew Videos:\r\n\r\n"));

	if (p->Dir[1] != "")	p->SearchVideoFiles(p->Dir[1]);
	if (p->Dir[2] != "")	p->SearchVideoFiles(p->Dir[2]);
	if (p->Dir[3] != "")	p->SearchVideoFiles(p->Dir[3]);
	if (p->Dir[4] != "")	p->SearchVideoFiles(p->Dir[4]);
	if (p->Dir[5] != "")	p->SearchVideoFiles(p->Dir[5]);
	if (p->Dir[6] != "")	p->SearchVideoFiles(p->Dir[6]);
	if (p->Dir[7] != "")	p->SearchVideoFiles(p->Dir[7]);

	p->echo(_T("\r\n**************************************************\r\n\r\nNewly Deleted Videos:\r\n\r\n"));
//	Sleep(3000);
	p->CheckVideoDel();
	p->echo(_T("\r\n**************************************************\r\n"));
	p->echo(_T("---------------------------------------->>Update Ends !!!<<----------------------------------------\r\n\r\n"));
//	Sleep(3000);
	p->MyStatistics();
	p->isBusyWorking = false;
	p->GetDlgItem(IDC_BUTTON_UPDATEDB)->EnableWindow(TRUE);
	p->GetDlgItem(IDC_BUTTON_AUTOMATCH)->EnableWindow(TRUE);
	p->GetDlgItem(IDC_BUTTON_NONEW)->EnableWindow(TRUE);
	p->GetDlgItem(IDC_BUTTON_SET2NEW)->EnableWindow(TRUE);

	return 0;
	AfxEndThread;
}

void CaManaDlg::OnBnClickedButtonAutomatch()
{
	if (MessageBox(_T("Are you sure to match videos&pictures to library?"), _T(""), MB_OKCANCEL) == IDOK)
	{
		isBusyWorking = true;
		GetDlgItem(IDC_BUTTON_UPDATEDB)->EnableWindow(FALSE);
		GetDlgItem(IDC_BUTTON_AUTOMATCH)->EnableWindow(FALSE);
		GetDlgItem(IDC_BUTTON_NONEW)->EnableWindow(FALSE);
		GetDlgItem(IDC_BUTTON_SET2NEW)->EnableWindow(FALSE);


		CWinThread* h_thread;
		h_thread = AfxBeginThread(ThreadMatch, (LPVOID) this);
	}
}

UINT CaManaDlg::ThreadMatch(LPVOID pParam)
{
	CaManaDlg *p = (CaManaDlg *)pParam;
	p->echoAndReturn(_T("\r\n\r\n---------------------------------------->>Match Starts !!!<<----------------------------------------\r\n"));
	p->AutoMatch();
	p->echoAndReturn(_T("\r\n---------------------------------------->>Match Ends !!!<<----------------------------------------\r\n\r\n"));
	p->isBusyWorking = false;
	p->GetDlgItem(IDC_BUTTON_UPDATEDB)->EnableWindow(TRUE);
	p->GetDlgItem(IDC_BUTTON_AUTOMATCH)->EnableWindow(TRUE);
	p->GetDlgItem(IDC_BUTTON_NONEW)->EnableWindow(TRUE);
	p->GetDlgItem(IDC_BUTTON_SET2NEW)->EnableWindow(TRUE);


	return 0;
	AfxEndThread;
}

//iAcc2,iAcc3,iAcc4
void CaManaDlg::AutoMatch()
{
	_variant_t RecordsAffected;
	CString mySQLtext;
	
	int RecordCount;
	CString name1, name2;	
	
	int vPID, vVID, vLIB;
	int i,f;

	try
	{
		mySQLtext.Format(_T("SELECT count(*) AS Expr1000 FROM library;"));
		iAcc2.m_pRecordset = iAcc2.m_pConnection->Execute(_bstr_t(mySQLtext), &RecordsAffected, adCmdText);
		
		RecordCount = iAcc2.m_pRecordset->GetCollect("Expr1000");
		
		mySQLtext.Format(_T("SELECT * FROM library ORDER BY LNAME1 ASC, LName2 ASC;"));
		iAcc2.m_pRecordset = iAcc2.m_pConnection->Execute(_bstr_t(mySQLtext), &RecordsAffected, adCmdText);
		iAcc2.m_pRecordset->MoveFirst();
		for (i = 0; i < RecordCount; i++)
		{
			name1 = iAcc2.m_pRecordset->GetCollect("LName1");
			name2 = iAcc2.m_pRecordset->GetCollect("LName2");
			vLIB = iAcc2.m_pRecordset->GetCollect("LibID");
			f = iAcc2.m_pRecordset->GetCollect("PID");


			if ((name1 == _T("0")) && (name2 == _T("0"))) 
			{
				iAcc2.m_pRecordset->MoveNext();
				continue;
			}

			mySQLtext.Format(_T("SELECT * FROM pics WHERE InStr(1,LCase(PFileName),LCase('%s'),0) AND InStr(1,LCase(PFileName),LCase('%s'),0);"), name1, name2);
			iAcc3.m_pRecordset = iAcc3.m_pConnection->Execute(_bstr_t(mySQLtext), &RecordsAffected, adCmdText);
			if ((iAcc3.m_pRecordset->BOF) && (iAcc3.m_pRecordset->adoEOF))
			{}
			else
			{
				if (f == 1) echoAndReturn(name1 + name2 + _T("  newly gets picture!"));
				iAcc3.m_pRecordset->MoveFirst();
				vPID = iAcc3.m_pRecordset->GetCollect("PID");
				mySQLtext.Format(_T("UPDATE library SET PID = %d WHERE LibID = %d;"),vPID,vLIB);
				iAcc3.m_pRecordset = iAcc3.m_pConnection->Execute(_bstr_t(mySQLtext), &RecordsAffected, adCmdText);
			}

			mySQLtext.Format(_T("SELECT * FROM videos WHERE InStr(1,LCase(VFileName),LCase('%s'),0) AND InStr(1,LCase(VFileName),LCase('%s'),0) AND IsNew = 1;"), name1, name2);
			iAcc3.m_pRecordset = iAcc3.m_pConnection->Execute(_bstr_t(mySQLtext), &RecordsAffected, adCmdText);
			if ((iAcc3.m_pRecordset->BOF) && (iAcc3.m_pRecordset->adoEOF))
			{}
			else
			{
				iAcc3.m_pRecordset->MoveFirst();
				while (!iAcc3.m_pRecordset->adoEOF)
				{
					vVID = iAcc3.m_pRecordset->GetCollect("VID");
					mySQLtext.Format(_T("UPDATE videos SET LibID = %d WHERE VID = %d;"), vLIB, vVID);
					iAcc4.m_pRecordset = iAcc4.m_pConnection->Execute(_bstr_t(mySQLtext), &RecordsAffected, adCmdText);
					
					iAcc3.m_pRecordset->MoveNext();
				}
				echoAndReturn(name1 + name2 + _T("  newly gets video(s)!"));
			}
			
			iAcc2.m_pRecordset->MoveNext();
		}
	}
	catch (_com_error &e)
	{
		CString eorr;
		eorr.Format(L"DB failed! Code:%08X   Info:%s \r\nDescribe:%s\r\n", e.Error(), (wchar_t*)e.ErrorMessage(), (wchar_t*)e.Description());
		echoAndReturn(eorr);
		return;
	}

}


void CaManaDlg::OnBnClickedButtonTest()
{

}




















void CaManaDlg::OnDblclkListRecords(NMHDR *pNMHDR, LRESULT *pResult)
{
	LPNMITEMACTIVATE pNMItemActivate = reinterpret_cast<LPNMITEMACTIVATE>(pNMHDR);
	
	CListCtrl* List1 = (CListCtrl*)GetDlgItem(IDC_LIST_RECORDS);

	CString tmpVID,tmpLibID;
	int i = List1->GetNextItem(-1, LVNI_SELECTED);

	tmpVID = List1->GetItemText(i, 10);
	if (tmpVID == _T("")) 
		return;
	else
	{
		g_VID = _ttoi(tmpVID);
		Modify1Video dlg(this);
		dlg.DoModal();
		
	}
	*pResult = 0;
}


void CaManaDlg::OnBnClickedButtonLibq()
{
	
}


void CaManaDlg::OnBnClickedButtonSet2new()
{
	isBusyWorking = true;
	GetDlgItem(IDC_BUTTON_UPDATEDB)->EnableWindow(FALSE);
	GetDlgItem(IDC_BUTTON_AUTOMATCH)->EnableWindow(FALSE);
	GetDlgItem(IDC_BUTTON_NONEW)->EnableWindow(FALSE);
	GetDlgItem(IDC_BUTTON_SET2NEW)->EnableWindow(FALSE);


	CString mySQLtext;
	_variant_t RecordsAffected;
	CString tmpVID, tmpFilePath;
	int NumVID;
	int RecordCount;
	int i;

	echoAndReturn(_T("\r\n\r\n---------------------------------------->>Batch2New Starts !!!<<----------------------------------------\r\n"));

	RecordCount = m_ListCtrl.GetItemCount();
	for (i = 0; i < RecordCount; i++)
	{
		if (m_ListCtrl.GetCheck(i) == 1)
		{
			tmpVID = m_ListCtrl.GetItemText(i, 10);
			if (tmpVID == _T(""))
				continue;
			NumVID = _ttoi(tmpVID);
			tmpFilePath = m_ListCtrl.GetItemText(i, 0);
			mySQLtext.Format(_T("UPDATE videos SET IsNew=1 WHERE VID = %d;"), NumVID);
			try
			{
				iAcc3.m_pRecordset = iAcc3.m_pConnection->Execute(_bstr_t(mySQLtext), &RecordsAffected, adCmdText);
			}
			catch (_com_error &e)
			{
				CString eorr;
				eorr.Format(L"DB failed! Code:%08X   Info:%s \r\nDescribe:%s\r\n", e.Error(), (wchar_t*)e.ErrorMessage(), (wchar_t*)e.Description());
				echoAndReturn(eorr);
				return;
			}

			echoAndReturn(tmpFilePath + _T("\tturns new!"));

		}
	}
	for (i = 0; i < RecordCount; i++)
		m_ListCtrl.SetCheck(i, 0);
	echoAndReturn(_T("\r\n---------------------------------------->>Batch2New Ends !!!<<----------------------------------------\r\n\r\n"));
	isBusyWorking = false;
	GetDlgItem(IDC_BUTTON_UPDATEDB)->EnableWindow(TRUE);
	GetDlgItem(IDC_BUTTON_AUTOMATCH)->EnableWindow(TRUE);
	GetDlgItem(IDC_BUTTON_NONEW)->EnableWindow(TRUE);
	GetDlgItem(IDC_BUTTON_SET2NEW)->EnableWindow(TRUE);

}
