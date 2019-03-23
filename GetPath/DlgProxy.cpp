
// DlgProxy.cpp : ʵ���ļ�
//

#include "stdafx.h"
#include "GetPath.h"
#include "DlgProxy.h"
#include "GetPathDlg.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#endif


// CGetPathDlgAutoProxy

IMPLEMENT_DYNCREATE(CGetPathDlgAutoProxy, CCmdTarget)

CGetPathDlgAutoProxy::CGetPathDlgAutoProxy()
{
	EnableAutomation();
	
	// ΪʹӦ�ó������Զ��������ڻ״̬ʱһֱ���� 
	//	���У����캯������ AfxOleLockApp��
	AfxOleLockApp();

	// ͨ��Ӧ�ó����������ָ��
	//  �����ʶԻ���  ���ô�����ڲ�ָ��
	//  ָ��Ի��򣬲����öԻ���ĺ���ָ��ָ��
	//  �ô���
	ASSERT_VALID(AfxGetApp()->m_pMainWnd);
	if (AfxGetApp()->m_pMainWnd)
	{
		ASSERT_KINDOF(CGetPathDlg, AfxGetApp()->m_pMainWnd);
		if (AfxGetApp()->m_pMainWnd->IsKindOf(RUNTIME_CLASS(CGetPathDlg)))
		{
			m_pDialog = reinterpret_cast<CGetPathDlg*>(AfxGetApp()->m_pMainWnd);
			m_pDialog->m_pAutoProxy = this;
		}
	}
}

CGetPathDlgAutoProxy::~CGetPathDlgAutoProxy()
{
	// Ϊ������ OLE �Զ����������ж������ֹӦ�ó���
	//	������������ AfxOleUnlockApp��
	//  ���������������⣬�⻹���������Ի���
	if (m_pDialog != NULL)
		m_pDialog->m_pAutoProxy = NULL;
	AfxOleUnlockApp();
}

void CGetPathDlgAutoProxy::OnFinalRelease()
{
	// �ͷ��˶��Զ�����������һ�����ú󣬽�����
	// OnFinalRelease��  ���ཫ�Զ�
	// ɾ���ö���  �ڵ��øû���֮ǰ�����������
	// ��������ĸ���������롣

	CCmdTarget::OnFinalRelease();
}

BEGIN_MESSAGE_MAP(CGetPathDlgAutoProxy, CCmdTarget)
END_MESSAGE_MAP()

BEGIN_DISPATCH_MAP(CGetPathDlgAutoProxy, CCmdTarget)
END_DISPATCH_MAP()

// ע��: ��������˶� IID_IGetPath ��֧��
//  ��֧������ VBA �����Ͱ�ȫ�󶨡�  �� IID ����ͬ���ӵ� .IDL �ļ��е�
//  ���Ƚӿڵ� GUID ƥ�䡣

// {30256D0E-AFEE-4389-9AF8-1AD85B590F6B}
static const IID IID_IGetPath =
{ 0x30256D0E, 0xAFEE, 0x4389, { 0x9A, 0xF8, 0x1A, 0xD8, 0x5B, 0x59, 0xF, 0x6B } };

BEGIN_INTERFACE_MAP(CGetPathDlgAutoProxy, CCmdTarget)
	INTERFACE_PART(CGetPathDlgAutoProxy, IID_IGetPath, Dispatch)
END_INTERFACE_MAP()

// IMPLEMENT_OLECREATE2 ���ڴ���Ŀ�� StdAfx.h �ж���
// {4F4D3B78-E2E9-45D2-987B-8C0FEEEDDDF9}
IMPLEMENT_OLECREATE2(CGetPathDlgAutoProxy, "GetPath.Application", 0x4f4d3b78, 0xe2e9, 0x45d2, 0x98, 0x7b, 0x8c, 0xf, 0xee, 0xed, 0xdd, 0xf9)


// CGetPathDlgAutoProxy ��Ϣ�������
