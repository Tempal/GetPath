
// GetPathDlg.cpp : ʵ���ļ�
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


// ����Ӧ�ó��򡰹��ڡ��˵���� CAboutDlg �Ի���

class CAboutDlg : public CDialogEx
{
public:
	CAboutDlg();

// �Ի�������
#ifdef AFX_DESIGN_TIME
	enum { IDD = IDD_ABOUTBOX };
#endif

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV ֧��

// ʵ��
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


// CGetPathDlg �Ի���


IMPLEMENT_DYNAMIC(CGetPathDlg, CDialogEx);

CGetPathDlg::CGetPathDlg(CWnd* pParent /*=NULL*/)
	: CDialogEx(IDD_GETPATH_DIALOG, pParent)
{
	m_hIcon = AfxGetApp()->LoadIcon(IDR_MAINFRAME);
	m_pAutoProxy = NULL;
}

CGetPathDlg::~CGetPathDlg()
{
	// ����öԻ������Զ���������
	//  ���˴���ָ��öԻ���ĺ���ָ������Ϊ NULL���Ա�
	//  �˴���֪���öԻ����ѱ�ɾ����
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


// CGetPathDlg ��Ϣ�������

BOOL CGetPathDlg::OnInitDialog()
{
	CDialogEx::OnInitDialog();

	// ��������...���˵�����ӵ�ϵͳ�˵��С�

	// IDM_ABOUTBOX ������ϵͳ���Χ�ڡ�
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

	// ���ô˶Ի����ͼ�ꡣ  ��Ӧ�ó��������ڲ��ǶԻ���ʱ����ܽ��Զ�
	//  ִ�д˲���
	SetIcon(m_hIcon, TRUE);			// ���ô�ͼ��
	SetIcon(m_hIcon, FALSE);		// ����Сͼ��

	//ShowWindow(SW_MAXIMIZE);

	// TODO: �ڴ���Ӷ���ĳ�ʼ������
	DWORD dwStyle = m_List.GetExtendedStyle();
	dwStyle |= LVS_EX_FULLROWSELECT;
	dwStyle |= LVS_EX_GRIDLINES;
	m_List.SetExtendedStyle(dwStyle);
	m_List.InsertColumn(0, _T(" "), LVCFMT_LEFT, 25);
	m_List.InsertColumn(1, _T("����·��"), LVCFMT_LEFT, 60);
	m_List.InsertColumn(2, _T("��С"), LVCFMT_LEFT, 80);
	m_List.InsertColumn(3, _T("����"), LVCFMT_LEFT, 80);
	m_List.InsertColumn(4, _T("λ�ã���,�У�"), LVCFMT_LEFT, 100);
	if (!AfxOleInit())
	{
		AfxMessageBox(IDP_OLE_INIT_FAILED);
		return FALSE;
	}

	return TRUE;  // ���ǽ��������õ��ؼ������򷵻� TRUE
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

// �����Ի��������С����ť������Ҫ����Ĵ���
//  �����Ƹ�ͼ�ꡣ  ����ʹ���ĵ�/��ͼģ�͵� MFC Ӧ�ó���
//  �⽫�ɿ���Զ���ɡ�

void CGetPathDlg::OnPaint()
{
	if (IsIconic())
	{
		CPaintDC dc(this); // ���ڻ��Ƶ��豸������

		SendMessage(WM_ICONERASEBKGND, reinterpret_cast<WPARAM>(dc.GetSafeHdc()), 0);

		// ʹͼ���ڹ����������о���
		int cxIcon = GetSystemMetrics(SM_CXICON);
		int cyIcon = GetSystemMetrics(SM_CYICON);
		CRect rect;
		GetClientRect(&rect);
		int x = (rect.Width() - cxIcon + 1) / 2;
		int y = (rect.Height() - cyIcon + 1) / 2;

		// ����ͼ��
		dc.DrawIcon(x, y, m_hIcon);
	}
	else
	{
		CDialogEx::OnPaint();
	}
}

//���û��϶���С������ʱϵͳ���ô˺���ȡ�ù��
//��ʾ��
HCURSOR CGetPathDlg::OnQueryDragIcon()
{
	return static_cast<HCURSOR>(m_hIcon);
}

// ���û��ر� UI ʱ������������Ա���������ĳ��
//  �������Զ�����������Ӧ�˳���  ��Щ
//  ��Ϣ�������ȷ����������: �����������ʹ�ã�
//  ������ UI�������ڹرնԻ���ʱ��
//  �Ի�����Ȼ�ᱣ�������

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
	// �����������Ա�����������Զ���
	//  �������Իᱣ�ִ�Ӧ�ó���
	//  ʹ�Ի���������������� UI ����������
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
	// TODO: �ڴ���ӿؼ�֪ͨ����������
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
	// TODO: �ڴ���ӿؼ�֪ͨ����������
	CFileDialog FileDialog(FALSE, L"xls", NULL, OFN_HIDEREADONLY | OFN_OVERWRITEPROMPT, _T("Microsoft Excel 2007(*.xlsx)|*.xlsx|�����ļ�(*.*)"), this);
	if (FileDialog.DoModal() != IDOK)
	{
		return;
	}
	CString cStrFile = FileDialog.GetPathName();  //ѡ�񱣴�·������  

	if (::PathFileExists(cStrFile))
		DeleteFile(cStrFile);

	//CString cStrFile = _T("E:\\myexcel.xls");  
	COleVariant covTrue((short)TRUE), covFalse((short)FALSE), covOptional((long)DISP_E_PARAMNOTFOUND, VT_ERROR);

	CApplication app; //Excel����   
	CWorkbooks books; //����������   
	CWorkbook book;  //������   
	CWorksheets sheets;  //����������   
	CWorksheet sheet; //��������   
	CRange range; //ʹ������   

	CoUninitialize();

	book.PrintPreview(_variant_t(false));
	//if (CoInitialize(NULL) == S_FALSE)
	//{
		//MessageBox(L"��ʼ��COM֧�ֿ�ʧ�ܣ�");
		//return;
	//}

	if (!app.CreateDispatch(_T("Excel.Application"))) //����IDispatch�ӿڶ���  
	{
		MessageBox(_T("Error!"));   
		return;
	}


	books = app.get_Workbooks();
	book = books.Add(covOptional);


	sheets = book.get_Worksheets();
	sheet = sheets.get_Item(COleVariant((short)1));  //�õ���һ��������  

	CHeaderCtrl   *pmyHeaderCtrl = m_List.GetHeaderCtrl(); //��ȡ��ͷ  

	int   m_cols = pmyHeaderCtrl->GetItemCount(); //��ȡ����  
	int   m_rows = m_List.GetItemCount();  //��ȡ����  


	TCHAR     lpBuffer[256];

	HDITEM   hdi; //This structure contains information about an item in a header control. This structure has been updated to support header item images and order values.  
	hdi.mask = HDI_TEXT;
	hdi.pszText = lpBuffer;
	hdi.cchTextMax = 256;

	CString   colname;
	CString strTemp;

	int   iRow, iCol;
	for (iCol = 0; iCol < m_cols; iCol++)//���б�ı���ͷд��EXCEL   
	{
		GetCellName(1, iCol + 1, colname); //(colname���Ƕ�Ӧ����A1,B1,C1,D1)  

		range = sheet.get_Range(COleVariant(colname), COleVariant(colname));

		pmyHeaderCtrl->GetItem(iCol, &hdi); //��ȡ��ͷÿ�е���Ϣ  

		range.put_Value2(COleVariant(hdi.pszText));  //����ÿ�е�����  

		int   nWidth = m_List.GetColumnWidth(iCol) / 6;

		//�õ���iCol+1��      
		range.AttachDispatch(range.get_Item(_variant_t((long)(iCol + 1)), vtMissing).pdispVal, true);

		//�����п�     
		range.put_ColumnWidth(_variant_t((long)nWidth));

	}

	range = sheet.get_Range(COleVariant(_T("A1 ")), COleVariant(colname));

	range.put_RowHeight(_variant_t((long)50));//�����еĸ߶�   


	range.put_VerticalAlignment(COleVariant((short)-4108));//xlVAlignCenter   =   -4108   

	COleSafeArray   saRet; //COleSafeArray�������ڴ����������ͺ�ά�����������  
	DWORD   numElements[] = { m_rows,m_cols };       //����д������  
	saRet.Create(VT_BSTR, 2, numElements); //�������������  

	range = sheet.get_Range(COleVariant(_T("A2 ")), covOptional); //��A2��ʼ  
	range = range.get_Resize(COleVariant((short)m_rows), COleVariant((short)m_cols)); //�������  

	long   index[2];

	for (iRow = 1; iRow <= m_rows; iRow++)//���б�����д��EXCEL   
	{
		for (iCol = 1; iCol <= m_cols; iCol++)
		{
			index[0] = iRow - 1;
			index[1] = iCol - 1;

			CString   szTemp;

			szTemp = m_List.GetItemText(iRow - 1, iCol - 1); //ȡ��m_list�ؼ��е�����  

			BSTR   bstr = szTemp.AllocSysString(); //The AllocSysString method allocates a new BSTR string that is Automation compatible  

			saRet.PutElement(index, bstr); //��m_list�ؼ��е����ݷ���saRet  

			SysFreeString(bstr);
		}
	}



	range.put_Value2(COleVariant(saRet)); //���õ������ݵ�saRet����ֵ������  


	book.SaveCopyAs(COleVariant(cStrFile)); //���浽cStrFile�ļ�  
	book.put_Saved(true);



	books.Close();

	//   

	book.ReleaseDispatch();
	books.ReleaseDispatch();

	app.ReleaseDispatch();
	app.Quit();
}
