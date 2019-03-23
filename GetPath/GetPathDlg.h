
// GetPathDlg.h : ͷ�ļ�
//

#pragma once
#include "afxcmn.h"
#include "afxwin.h"

class CGetPathDlgAutoProxy;


// CGetPathDlg �Ի���
class CGetPathDlg : public CDialogEx
{
	DECLARE_DYNAMIC(CGetPathDlg);
	friend class CGetPathDlgAutoProxy;

// ����
public:
	CGetPathDlg(CWnd* pParent = NULL);	// ��׼���캯��
	virtual ~CGetPathDlg();

// �Ի�������
#ifdef AFX_DESIGN_TIME
	enum { IDD = IDD_GETPATH_DIALOG };
#endif

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);	// DDX/DDV ֧��


// ʵ��
protected:
	CGetPathDlgAutoProxy* m_pAutoProxy;
	HICON m_hIcon;

	BOOL CanExit();

	// ���ɵ���Ϣӳ�亯��
	virtual BOOL OnInitDialog();
	afx_msg void OnSysCommand(UINT nID, LPARAM lParam);
	afx_msg void OnPaint();
	afx_msg HCURSOR OnQueryDragIcon();
	afx_msg void OnClose();
	virtual void OnOK();
	virtual void OnCancel();
	DECLARE_MESSAGE_MAP()
public:
	// �����ļ��б�
	CListCtrl m_List;
	// �ļ���
	CEdit m_FileName;
	afx_msg void OnBnClickedOk();
	void TraverseDir(CString& dir, std::vector<CString>& vec);
	afx_msg void OnBnClickedExport();
	CEdit m_FileNameOut;
};
