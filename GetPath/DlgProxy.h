
// DlgProxy.h: ͷ�ļ�
//

#pragma once

class CGetPathDlg;


// CGetPathDlgAutoProxy ����Ŀ��

class CGetPathDlgAutoProxy : public CCmdTarget
{
	DECLARE_DYNCREATE(CGetPathDlgAutoProxy)

	CGetPathDlgAutoProxy();           // ��̬������ʹ�õ��ܱ����Ĺ��캯��

// ����
public:
	CGetPathDlg* m_pDialog;

// ����
public:

// ��д
	public:
	virtual void OnFinalRelease();

// ʵ��
protected:
	virtual ~CGetPathDlgAutoProxy();

	// ���ɵ���Ϣӳ�亯��

	DECLARE_MESSAGE_MAP()
	DECLARE_OLECREATE(CGetPathDlgAutoProxy)

	// ���ɵ� OLE ����ӳ�亯��

	DECLARE_DISPATCH_MAP()
	DECLARE_INTERFACE_MAP()
};

