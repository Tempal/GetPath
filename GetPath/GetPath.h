
// GetPath.h : PROJECT_NAME Ӧ�ó������ͷ�ļ�
//

#pragma once

#ifndef __AFXWIN_H__
	#error "�ڰ������ļ�֮ǰ������stdafx.h�������� PCH �ļ�"
#endif

#include "resource.h"		// ������
#include<vector>

// CGetPathApp: 
// �йش����ʵ�֣������ GetPath.cpp
//

class CGetPathApp : public CWinApp
{
public:
	CGetPathApp();

// ��д
public:
	virtual BOOL InitInstance();
	virtual int ExitInstance();

// ʵ��

	DECLARE_MESSAGE_MAP()
};

extern CGetPathApp theApp;