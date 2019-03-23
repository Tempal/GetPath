
// DlgProxy.cpp : 实现文件
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
	
	// 为使应用程序在自动化对象处于活动状态时一直保持 
	//	运行，构造函数调用 AfxOleLockApp。
	AfxOleLockApp();

	// 通过应用程序的主窗口指针
	//  来访问对话框。  设置代理的内部指针
	//  指向对话框，并设置对话框的后向指针指向
	//  该代理。
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
	// 为了在用 OLE 自动化创建所有对象后终止应用程序，
	//	析构函数调用 AfxOleUnlockApp。
	//  除了做其他事情外，这还将销毁主对话框
	if (m_pDialog != NULL)
		m_pDialog->m_pAutoProxy = NULL;
	AfxOleUnlockApp();
}

void CGetPathDlgAutoProxy::OnFinalRelease()
{
	// 释放了对自动化对象的最后一个引用后，将调用
	// OnFinalRelease。  基类将自动
	// 删除该对象。  在调用该基类之前，请添加您的
	// 对象所需的附加清理代码。

	CCmdTarget::OnFinalRelease();
}

BEGIN_MESSAGE_MAP(CGetPathDlgAutoProxy, CCmdTarget)
END_MESSAGE_MAP()

BEGIN_DISPATCH_MAP(CGetPathDlgAutoProxy, CCmdTarget)
END_DISPATCH_MAP()

// 注意: 我们添加了对 IID_IGetPath 的支持
//  以支持来自 VBA 的类型安全绑定。  此 IID 必须同附加到 .IDL 文件中的
//  调度接口的 GUID 匹配。

// {30256D0E-AFEE-4389-9AF8-1AD85B590F6B}
static const IID IID_IGetPath =
{ 0x30256D0E, 0xAFEE, 0x4389, { 0x9A, 0xF8, 0x1A, 0xD8, 0x5B, 0x59, 0xF, 0x6B } };

BEGIN_INTERFACE_MAP(CGetPathDlgAutoProxy, CCmdTarget)
	INTERFACE_PART(CGetPathDlgAutoProxy, IID_IGetPath, Dispatch)
END_INTERFACE_MAP()

// IMPLEMENT_OLECREATE2 宏在此项目的 StdAfx.h 中定义
// {4F4D3B78-E2E9-45D2-987B-8C0FEEEDDDF9}
IMPLEMENT_OLECREATE2(CGetPathDlgAutoProxy, "GetPath.Application", 0x4f4d3b78, 0xe2e9, 0x45d2, 0x98, 0x7b, 0x8c, 0xf, 0xee, 0xed, 0xdd, 0xf9)


// CGetPathDlgAutoProxy 消息处理程序
