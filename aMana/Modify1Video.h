#pragma once
#include "aMana.h"
#include "aManaDlg.h"
#include "afxwin.h"
// Modify1Video dialog

class Modify1Video : public CDialogEx
{
	DECLARE_DYNAMIC(Modify1Video)

public:
	Modify1Video(CWnd* pParent = NULL);   // standard constructor
	virtual ~Modify1Video();

	int  bChn, bDel, bNew;
	CString bFilePath, bFileSize, bName1, bName2, bPicPath, bSerie, bDate,bVID;
	void ShowPic(CString pp);
	void RecordInit();

	int sID[1024];
	CString	sName[1024];

// Dialog Data
	enum { IDD = IDD_MODIFY1VIDEO_DIALOG };

protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support

	DECLARE_MESSAGE_MAP()
public:
//	afx_msg void OnBnClickedButton1();
	afx_msg void OnBnClickedOk();
	afx_msg void OnBnClickedCancel();
	virtual BOOL OnInitDialog();
	afx_msg void OnPaint();
	CButton check_chn;
	CButton check_new;
	CButton check_del;
	afx_msg void OnBnClickedButtonClear();
	CComboBox combo_serie;
	CButton check_serie;
	afx_msg void OnClickedCheckChangserie();
};
