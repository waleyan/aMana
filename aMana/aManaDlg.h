
// aManaDlg.h : header file
//

#pragma once
#include "Ado2Access.h"
#include "Modify1Video.h"
#include "afxcmn.h"
#include "afxwin.h"

// CaManaDlg dialog
class CaManaDlg : public CDialogEx
{
// Construction
public:
	CaManaDlg(CWnd* pParent = NULL);	// standard constructor

// Dialog Data
	enum { IDD = IDD_AMANA_DIALOG };

//--------------------about interface------------------------------------
	CString DashMessage;
	void InterfaceInitial();		
	void echo(CString text);
	void echoAndReturn(CString text);
	
	CListCtrl m_ListCtrl;
	void InitRecordList();
	void Show1LineInList(CString pPath, CString pSize, CString pFan, CString pHao, CString pPic, CString pSerie, CString pChn, CString pNew, CString pDel, CString pDate, CString pVID, CString pLibID);
//about interface end


	
//--------------------------about database---------------------------------
	Ado2Access iAcc1;		//for query
	Ado2Access iAcc2;		//for query
	Ado2Access iAcc3;		//for modify,add,delete
	Ado2Access iAcc4;		//for tmp
	BOOL InitialDatabase();		//return:  1:success; 0:fail,error info will show on dashboard
	BOOL DatabaseExeSQL(int AdoNum, CString SQLtext);		//return:  1:success; 0:fail,error info will show on dashboard
//about database end

	
//-----------------about search---------------------------	
	static afx_msg UINT ThreadQuery(LPVOID pParam);
	void FunQuery();	
//about search end


//-------------------about root path---------------------------
	CString Dir[8];
	CString PicPath;
	bool ChangeRootPathStatus;		//1:can edit;0:cannot edit
	void InitRootPath();

//about root path end

//-------------------about pic---------------------------
	void SearchPicFiles(CString MyPath);
	void UpdatePicDB(CString MyFileName, CString MyFilePath);
//about pic end

//-------------------about video---------------------------
	void SearchVideoFiles(CString MyPath);
	void UpdateVidoeDB(CString MyFileName, CString MyFilePath, CString MyFileSize, CString MyCDate);
	void CheckVideoDel();

//about video end

	void AutoMatch();

	void MyStatistics();

	static afx_msg UINT ThreadUpdateDB(LPVOID pParam);
	static afx_msg UINT ThreadMatch(LPVOID pParam);

	bool isBusyWorking;
	int g_VID, g_LibID;

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);	// DDX/DDV support


// Implementation
protected:
	HICON m_hIcon;

	// Generated message map functions
	virtual BOOL OnInitDialog();
	afx_msg void OnPaint();
	afx_msg HCURSOR OnQueryDragIcon();
	DECLARE_MESSAGE_MAP()

public:
	afx_msg void OnBnClickedButtonTest();
	
	afx_msg void OnBnClickedButtonQuery();
	CButton check_chn;
	CButton check_new;
	CButton check_del;
	CButton check_pic;
	afx_msg void OnBnClickedButtonChangedir();
	afx_msg void OnBnClickedButtonClnlist();
	afx_msg void OnBnClickedButtonNonew();
	virtual BOOL PreTranslateMessage(MSG* pMsg);
	virtual void OnOK();
	virtual void OnCancel();
	afx_msg void OnBnClickedButtonUpdatedb();
	afx_msg void OnBnClickedButtonAutomatch();
	afx_msg void OnDblclkListRecords(NMHDR *pNMHDR, LRESULT *pResult);
	
	afx_msg void OnBnClickedButtonLibq();
	afx_msg void OnBnClickedButtonSet2new();
};
