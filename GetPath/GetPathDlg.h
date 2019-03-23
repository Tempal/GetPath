
// GetPathDlg.h : 头文件
//

#pragma once
#include "afxcmn.h"
#include "afxwin.h"

class CGetPathDlgAutoProxy;


// CGetPathDlg 对话框
class CGetPathDlg : public CDialogEx
{
	DECLARE_DYNAMIC(CGetPathDlg);
	friend class CGetPathDlgAutoProxy;

// 构造
public:
	CGetPathDlg(CWnd* pParent = NULL);	// 标准构造函数
	virtual ~CGetPathDlg();

// 对话框数据
#ifdef AFX_DESIGN_TIME
	enum { IDD = IDD_GETPATH_DIALOG };
#endif

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);	// DDX/DDV 支持


// 实现
protected:
	CGetPathDlgAutoProxy* m_pAutoProxy;
	HICON m_hIcon;

	BOOL CanExit();

	// 生成的消息映射函数
	virtual BOOL OnInitDialog();
	afx_msg void OnSysCommand(UINT nID, LPARAM lParam);
	afx_msg void OnPaint();
	afx_msg HCURSOR OnQueryDragIcon();
	afx_msg void OnClose();
	virtual void OnOK();
	virtual void OnCancel();
	DECLARE_MESSAGE_MAP()
public:
	// 所有文件列表
	CListCtrl m_List;
	// 文件名
	CEdit m_FileName;
	afx_msg void OnBnClickedOk();
	void TraverseDir(CString& dir, std::vector<CString>& vec);
	afx_msg void OnBnClickedExport();
	CEdit m_FileNameOut;
};
