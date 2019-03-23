
// GetPathDlg.cpp : 实现文件
//

#include "stdafx.h"
#include "GetPath.h"
#include "GetPathDlg.h"
#include "DlgProxy.h"
#include "afxdialogex.h"
#include <iostream>
#include <fstream>

#include "CApplication.h"
#include "CRange.h"
#include "CWorkbooks.h"
#include "CWorksheets.h"
#include "CWorkbook.h"
#include "CWorksheet.h"
#ifdef _DEBUG
#define new DEBUG_NEW
#endif


// 用于应用程序“关于”菜单项的 CAboutDlg 对话框

class CAboutDlg : public CDialogEx
{
public:
	CAboutDlg();

// 对话框数据
#ifdef AFX_DESIGN_TIME
	enum { IDD = IDD_ABOUTBOX };
#endif

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV 支持

// 实现
protected:
	DECLARE_MESSAGE_MAP()
};

CAboutDlg::CAboutDlg() : CDialogEx(IDD_ABOUTBOX)
{
}

void CAboutDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
}

BEGIN_MESSAGE_MAP(CAboutDlg, CDialogEx)
END_MESSAGE_MAP()


// CGetPathDlg 对话框


IMPLEMENT_DYNAMIC(CGetPathDlg, CDialogEx);

CGetPathDlg::CGetPathDlg(CWnd* pParent /*=NULL*/)
	: CDialogEx(IDD_GETPATH_DIALOG, pParent)
{
	m_hIcon = AfxGetApp()->LoadIcon(IDR_MAINFRAME);
	m_pAutoProxy = NULL;
}

CGetPathDlg::~CGetPathDlg()
{
	// 如果该对话框有自动化代理，则
	//  将此代理指向该对话框的后向指针设置为 NULL，以便
	//  此代理知道该对话框已被删除。
	if (m_pAutoProxy != NULL)
		m_pAutoProxy->m_pDialog = NULL;
}

void CGetPathDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
	DDX_Control(pDX, IDC_AllFile, m_List);
	DDX_Control(pDX, IDC_FileName, m_FileName);
	DDX_Control(pDX, IDC_FileNameOut, m_FileNameOut);
}

BEGIN_MESSAGE_MAP(CGetPathDlg, CDialogEx)
	ON_WM_SYSCOMMAND()
	ON_WM_CLOSE()
	ON_WM_PAINT()
	ON_WM_QUERYDRAGICON()
	ON_BN_CLICKED(IDOK, &CGetPathDlg::OnBnClickedOk)
	ON_BN_CLICKED(IDC_EXPORT, &CGetPathDlg::OnBnClickedExport)
END_MESSAGE_MAP()


// CGetPathDlg 消息处理程序

BOOL CGetPathDlg::OnInitDialog()
{
	CDialogEx::OnInitDialog();

	// 将“关于...”菜单项添加到系统菜单中。

	// IDM_ABOUTBOX 必须在系统命令范围内。
	ASSERT((IDM_ABOUTBOX & 0xFFF0) == IDM_ABOUTBOX);
	ASSERT(IDM_ABOUTBOX < 0xF000);

	CMenu* pSysMenu = GetSystemMenu(FALSE);
	if (pSysMenu != NULL)
	{
		BOOL bNameValid;
		CString strAboutMenu;
		bNameValid = strAboutMenu.LoadString(IDS_ABOUTBOX);
		ASSERT(bNameValid);
		if (!strAboutMenu.IsEmpty())
		{
			pSysMenu->AppendMenu(MF_SEPARATOR);
			pSysMenu->AppendMenu(MF_STRING, IDM_ABOUTBOX, strAboutMenu);
		}
	}

	// 设置此对话框的图标。  当应用程序主窗口不是对话框时，框架将自动
	//  执行此操作
	SetIcon(m_hIcon, TRUE);			// 设置大图标
	SetIcon(m_hIcon, FALSE);		// 设置小图标

	//ShowWindow(SW_MAXIMIZE);

	// TODO: 在此添加额外的初始化代码
	DWORD dwStyle = m_List.GetExtendedStyle();
	dwStyle |= LVS_EX_FULLROWSELECT;
	dwStyle |= LVS_EX_GRIDLINES;
	m_List.SetExtendedStyle(dwStyle);
	m_List.InsertColumn(0, _T(" "), LVCFMT_LEFT, 25);
	m_List.InsertColumn(1, _T("所有路径"), LVCFMT_LEFT, 60);
	m_List.InsertColumn(2, _T("大小"), LVCFMT_LEFT, 80);
	m_List.InsertColumn(3, _T("类型"), LVCFMT_LEFT, 80);
	m_List.InsertColumn(4, _T("位置（行,列）"), LVCFMT_LEFT, 100);
	if (!AfxOleInit())
	{
		AfxMessageBox(IDP_OLE_INIT_FAILED);
		return FALSE;
	}

	return TRUE;  // 除非将焦点设置到控件，否则返回 TRUE
}

void CGetPathDlg::OnSysCommand(UINT nID, LPARAM lParam)
{
	if ((nID & 0xFFF0) == IDM_ABOUTBOX)
	{
		CAboutDlg dlgAbout;
		dlgAbout.DoModal();
	}
	else
	{
		CDialogEx::OnSysCommand(nID, lParam);
	}
}

// 如果向对话框添加最小化按钮，则需要下面的代码
//  来绘制该图标。  对于使用文档/视图模型的 MFC 应用程序，
//  这将由框架自动完成。

void CGetPathDlg::OnPaint()
{
	if (IsIconic())
	{
		CPaintDC dc(this); // 用于绘制的设备上下文

		SendMessage(WM_ICONERASEBKGND, reinterpret_cast<WPARAM>(dc.GetSafeHdc()), 0);

		// 使图标在工作区矩形中居中
		int cxIcon = GetSystemMetrics(SM_CXICON);
		int cyIcon = GetSystemMetrics(SM_CYICON);
		CRect rect;
		GetClientRect(&rect);
		int x = (rect.Width() - cxIcon + 1) / 2;
		int y = (rect.Height() - cyIcon + 1) / 2;

		// 绘制图标
		dc.DrawIcon(x, y, m_hIcon);
	}
	else
	{
		CDialogEx::OnPaint();
	}
}

//当用户拖动最小化窗口时系统调用此函数取得光标
//显示。
HCURSOR CGetPathDlg::OnQueryDragIcon()
{
	return static_cast<HCURSOR>(m_hIcon);
}

// 当用户关闭 UI 时，如果控制器仍保持着它的某个
//  对象，则自动化服务器不应退出。  这些
//  消息处理程序确保如下情形: 如果代理仍在使用，
//  则将隐藏 UI；但是在关闭对话框时，
//  对话框仍然会保留在那里。

void CGetPathDlg::OnClose()
{
	if (CanExit())
		CDialogEx::OnClose();
}

void CGetPathDlg::OnOK()
{
	if (CanExit())
		CDialogEx::OnOK();
}

void CGetPathDlg::OnCancel()
{
	if (CanExit())
		CDialogEx::OnCancel();
}

BOOL CGetPathDlg::CanExit()
{
	// 如果代理对象仍保留在那里，则自动化
	//  控制器仍会保持此应用程序。
	//  使对话框保留在那里，但将其 UI 隐藏起来。
	if (m_pAutoProxy != NULL)
	{
		ShowWindow(SW_HIDE);
		return FALSE;
	}

	return TRUE;
}

void CGetPathDlg::TraverseDir(CString& dir, std::vector<CString>& vec)
{
	CFileFind ff;
	if (dir.Right(1) != "\\")
		dir += "\\";
	dir += "*.*";

	BOOL ret = ff.FindFile(dir);
	while (ret)
	{
		ret = ff.FindNextFile();
		if (ret != 0)
		{
			if (ff.IsDirectory() && !ff.IsDots())
			{
				CString path = ff.GetFilePath();
				vec.push_back(path);
				TraverseDir(path, vec);
			}
			else if (!ff.IsDirectory() && !ff.IsDots())
			{
				CString name = ff.GetFileName();
				CString path = ff.GetFilePath();
				vec.push_back(name);
			}
		}
	}
}

void CGetPathDlg::OnBnClickedOk()
{
	// TODO: 在此添加控件通知处理程序代码
	m_List.DeleteAllItems();
	std::vector<CString> v;
	CString sFileName;
	m_FileName.GetWindowTextW(sFileName);
	TraverseDir(sFileName, v);

//#ifdef _UNICODE
//	std::wofstream ofs;
//#else
//	ofstream ofs;
//#endif
	//ofs.open("log.txt");
	//if (ofs.is_open())
	//{
		int n=m_List.GetItemCount();
		for (int i = 0; i < v.size(); ++i)
		{
			TCHAR* p = v[i].GetBuffer();
			//ofs << p << std::endl;
			CString str;
			str.Format(_T("%d"), n);
			m_List.InsertItem(n, str);
			m_List.SetItemText(n,1, v[i]);
			n++;
		}
		//ofs.flush();
		//ofs.close();
	//}
	//CDialogEx::OnOK();
}

static void   GetCellName(int nRow, int nCol, CString &strName)
{
	CString strRow;
	char cCell = (char)('A' + nCol - 1);

	strName.Format(_T("%c"), cCell);
	strRow.Format(_T("%d "), nRow);
	strName += strRow;
}

void CGetPathDlg::OnBnClickedExport()
{
	// TODO: 在此添加控件通知处理程序代码
	CFileDialog FileDialog(FALSE, L"xls", NULL, OFN_HIDEREADONLY | OFN_OVERWRITEPROMPT, _T("Microsoft Excel 2007(*.xlsx)|*.xlsx|所有文件(*.*)"), this);
	if (FileDialog.DoModal() != IDOK)
	{
		return;
	}
	CString cStrFile = FileDialog.GetPathName();  //选择保存路径名称  

	if (::PathFileExists(cStrFile))
		DeleteFile(cStrFile);

	//CString cStrFile = _T("E:\\myexcel.xls");  
	COleVariant covTrue((short)TRUE), covFalse((short)FALSE), covOptional((long)DISP_E_PARAMNOTFOUND, VT_ERROR);

	CApplication app; //Excel程序   
	CWorkbooks books; //工作簿集合   
	CWorkbook book;  //工作表   
	CWorksheets sheets;  //工作簿集合   
	CWorksheet sheet; //工作表集合   
	CRange range; //使用区域   

	CoUninitialize();

	book.PrintPreview(_variant_t(false));
	//if (CoInitialize(NULL) == S_FALSE)
	//{
		//MessageBox(L"初始化COM支持库失败！");
		//return;
	//}

	if (!app.CreateDispatch(_T("Excel.Application"))) //创建IDispatch接口对象  
	{
		MessageBox(_T("Error!"));   
		return;
	}


	books = app.get_Workbooks();
	book = books.Add(covOptional);


	sheets = book.get_Worksheets();
	sheet = sheets.get_Item(COleVariant((short)1));  //得到第一个工作表  

	CHeaderCtrl   *pmyHeaderCtrl = m_List.GetHeaderCtrl(); //获取表头  

	int   m_cols = pmyHeaderCtrl->GetItemCount(); //获取列数  
	int   m_rows = m_List.GetItemCount();  //获取行数  


	TCHAR     lpBuffer[256];

	HDITEM   hdi; //This structure contains information about an item in a header control. This structure has been updated to support header item images and order values.  
	hdi.mask = HDI_TEXT;
	hdi.pszText = lpBuffer;
	hdi.cchTextMax = 256;

	CString   colname;
	CString strTemp;

	int   iRow, iCol;
	for (iCol = 0; iCol < m_cols; iCol++)//将列表的标题头写入EXCEL   
	{
		GetCellName(1, iCol + 1, colname); //(colname就是对应表格的A1,B1,C1,D1)  

		range = sheet.get_Range(COleVariant(colname), COleVariant(colname));

		pmyHeaderCtrl->GetItem(iCol, &hdi); //获取表头每列的信息  

		range.put_Value2(COleVariant(hdi.pszText));  //设置每列的内容  

		int   nWidth = m_List.GetColumnWidth(iCol) / 6;

		//得到第iCol+1列      
		range.AttachDispatch(range.get_Item(_variant_t((long)(iCol + 1)), vtMissing).pdispVal, true);

		//设置列宽     
		range.put_ColumnWidth(_variant_t((long)nWidth));

	}

	range = sheet.get_Range(COleVariant(_T("A1 ")), COleVariant(colname));

	range.put_RowHeight(_variant_t((long)50));//设置行的高度   


	range.put_VerticalAlignment(COleVariant((short)-4108));//xlVAlignCenter   =   -4108   

	COleSafeArray   saRet; //COleSafeArray类是用于处理任意类型和维数的数组的类  
	DWORD   numElements[] = { m_rows,m_cols };       //行列写入数组  
	saRet.Create(VT_BSTR, 2, numElements); //创建所需的数组  

	range = sheet.get_Range(COleVariant(_T("A2 ")), covOptional); //从A2开始  
	range = range.get_Resize(COleVariant((short)m_rows), COleVariant((short)m_cols)); //表的区域  

	long   index[2];

	for (iRow = 1; iRow <= m_rows; iRow++)//将列表内容写入EXCEL   
	{
		for (iCol = 1; iCol <= m_cols; iCol++)
		{
			index[0] = iRow - 1;
			index[1] = iCol - 1;

			CString   szTemp;

			szTemp = m_List.GetItemText(iRow - 1, iCol - 1); //取得m_list控件中的内容  

			BSTR   bstr = szTemp.AllocSysString(); //The AllocSysString method allocates a new BSTR string that is Automation compatible  

			saRet.PutElement(index, bstr); //把m_list控件中的内容放入saRet  

			SysFreeString(bstr);
		}
	}



	range.put_Value2(COleVariant(saRet)); //将得到的数据的saRet数组值放入表格  


	book.SaveCopyAs(COleVariant(cStrFile)); //保存到cStrFile文件  
	book.put_Saved(true);



	books.Close();

	//   

	book.ReleaseDispatch();
	books.ReleaseDispatch();

	app.ReleaseDispatch();
	app.Quit();
}
