
// AttenDlg.cpp : 实现文件
//

#include "stdafx.h"
#include "Atten.h"
#include "AttenDlg.h"
#include "afxdialogex.h"
#include "Person.h"
#include <vector>
#include <map>
using namespace std;
#define NORMAL  (9*60)
#define SPECIAL (9*60+30)
#define UNUSUAL (10*60)

#define REGULAR  (18*60)
#define OVERTIME (20*60) 
#define SOHARD   (22*60) 

#ifdef _DEBUG
#define new DEBUG_NEW
#endif


// CAttenDlg 对话框
void census(Day &d,int &punctual);


CAttenDlg::CAttenDlg(CWnd* pParent /*=NULL*/)
	: CDialogEx(CAttenDlg::IDD, pParent)
{
	m_hIcon = AfxGetApp()->LoadIcon(IDR_MAINFRAME);
}

void CAttenDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
}

BEGIN_MESSAGE_MAP(CAttenDlg, CDialogEx)
	ON_WM_PAINT()
	ON_WM_QUERYDRAGICON()
	ON_BN_CLICKED(IDC_BUTTON1, &CAttenDlg::OnBnClickedButton1)
END_MESSAGE_MAP()


// CAttenDlg 消息处理程序

BOOL CAttenDlg::OnInitDialog()
{
	CDialogEx::OnInitDialog();

	// 设置此对话框的图标。当应用程序主窗口不是对话框时，框架将自动
	//  执行此操作
	SetIcon(m_hIcon, TRUE);			// 设置大图标
	SetIcon(m_hIcon, FALSE);		// 设置小图标

	// TODO: 在此添加额外的初始化代码

	return TRUE;  // 除非将焦点设置到控件，否则返回 TRUE
}

// 如果向对话框添加最小化按钮，则需要下面的代码
//  来绘制该图标。对于使用文档/视图模型的 MFC 应用程序，
//  这将由框架自动完成。

void CAttenDlg::OnPaint()
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
HCURSOR CAttenDlg::OnQueryDragIcon()
{
	return static_cast<HCURSOR>(m_hIcon);
}



void CAttenDlg::OnBnClickedButton1()
{
	CFileDialog dlg(TRUE, NULL, NULL, OFN_HIDEREADONLY | OFN_OVERWRITEPROMPT|OFN_ALLOWMULTISELECT, 
		_T("excel Files (*.xlsx;*.xls)|*.xlsx;*.xls"), this);
	CString strFilePath; 
	if(dlg.DoModal() == IDOK){
		//xml
		IXMLDOMDocumentPtr xmlFile = NULL;
		IXMLDOMElement* xmlRoot = NULL;
		_variant_t varXml(L"C:\\Users\\anan\\Desktop\\1.xml");
		logFile = CreateFile(L"log.txt", GENERIC_WRITE, 0, NULL, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, NULL);
		CoInitialize(NULL);
		xmlFile.CreateInstance("Msxml2.DOMDocument.4.0");
		VARIANT_BOOL varOut;
		xmlFile->load(varXml, &varOut);
		xmlFile->get_documentElement(&xmlRoot);
		BSTR rootName;
		DWORD dwBytesWritten;
		xmlRoot->get_nodeName(&rootName);
		WriteFile(logFile, rootName, (DWORD)(wcslen(rootName)*sizeof(WCHAR)), &dwBytesWritten, NULL);
		WriteFile(logFile, L"\r\n", (DWORD)(2*sizeof(WCHAR)), &dwBytesWritten, NULL);
		IXMLDOMNodeList* xmlChildNodes = NULL;
		xmlRoot->get_childNodes(&xmlChildNodes);
		PrintChild(xmlChildNodes, 2);
		if(logFile != INVALID_HANDLE_VALUE)CloseHandle(logFile);
		if(xmlChildNodes!=NULL)xmlChildNodes->Release();
		if(xmlRoot!=NULL)xmlRoot->Release();

		map <int, int> m1;
		map <int, int>::iterator m1_Iter;
		m1.insert( pair <int, int>( 11,  4+31*31) );

		CArray<CString, CString> aryFilename; 
		POSITION posFile=dlg.GetStartPosition(); 
		while(posFile!=NULL) 
		{ 
			aryFilename.Add(dlg.GetNextPathName(posFile)); 
		}
		ief.InitExcel();
		vector<Person*> li;
		vector<CString> dates;
		for(int i=0;i<aryFilename.GetSize();i++) 
		{
			ief.OpenExcelFile(aryFilename.GetAt(i));
			for(int shind=0;shind<ief.GetSheetCount();shind++){
				dates.clear();
				ief.LoadSheet(shind+1);
				CString tmp = NULL;
				for(int i=0;i<(ief.GetColumnCount()-1)/2;i++){
					dates.push_back(ief.GetCellString(1,2+i*2));
				}
				int m,n;
				double dm,dn;
				int Predetermined = NORMAL+5;
				int row = ief.GetRowCount();

				for (int i = 3; i <= row; i++){
					tmp = ief.GetCellString(i,1);
					if (tmp.Compare(_T(""))==0||tmp.Find(_T("备注"))>=0){
						USES_CONVERSION;
						TRACE("break in \"%s\"\n",W2A(tmp));
						break;
					}
					Person *p = NULL;
					for (int j=0; j<li.size(); j++){
						if((*li[j]).GetName().Compare(tmp)==0){
							p = li[j];
							break;
						}
					}
					if(!p){
						p = new Person();
						p->SetName(tmp);
						li.push_back(p);
					}
					for (int j = 0; j < (ief.GetColumnCount()-1)/2; j++){
						m = j*2+2;
						n = j*2+3;
						if (ief.IsCellString(i,m)){
							if (!ief.GetCellString(i,m).Compare(_T("未打卡"))){
								dm = -1;
							}else{
								dm = -2;
							}
						}else{
							dm = ief.GetCellDouble(i,m);
						}
						if (ief.IsCellString(i,n)){
							if (!ief.GetCellString(i,n).Compare(_T("未打卡"))){
								dn = -1;
							}else{
								dn = -2;
							}
						}else{
							dn = ief.GetCellDouble(i,n);
						}
						p->AddDay(dates[j],dm,dn);
						Day d = p->LastDay();
						if(d.GetArrival()>0&&d.GetArrival()<=Predetermined){
							ief.SetCellColor(i,m,0,0,0);
							ief.SetCellBold(i,m,false);
						}else if(d.GetArrival()==0){
							ief.SetCellColor(i,m,0,0,0);
							ief.SetCellBold(i,m,false);
						}else{
							ief.SetCellColor(i,m,255,0,0);
							ief.SetCellBold(i,m,true);
						}
						if (d.GetLeave()>=REGULAR){
							Predetermined = NORMAL+5;
							if (d.GetLeave()>=OVERTIME){
								Predetermined = SPECIAL;
							}
							if (d.GetLeave()>=SOHARD){
								Predetermined = UNUSUAL;
							}
							ief.SetCellColor(i,n,0,0,0);
							ief.SetCellBold(i,n,false);
						}else if(d.GetLeave()==0){
							Predetermined = NORMAL+5;
							ief.SetCellColor(i,n,0,0,0);
							ief.SetCellBold(i,n,false);
						}else{
							Predetermined = NORMAL+5;
							ief.SetCellColor(i,n,255,0,0);
							ief.SetCellBold(i,n,true);
						}                        
					}
				}
			}
			ief.Save();
			ief.CloseExcelFile();
		}
		for (int i=0; i<li.size(); i++){
			(*li[i]).SortDay();
		}

		TCHAR buf[100];
		GetCurrentDirectory(sizeof(buf),buf);

		//输出
		if (ief.OpenExcelFile(CString(buf)+_T("\\统计结果.xlsx"))){
			ief.LoadSheet(1,true);
		}else{
			ief.addSheet(NULL);
		}		

		ief.SetCellString(1,1,_T("姓名"));
		ief.SetCellBackground(0xdc,0xe6,0xf1);
		ief.SetCellWidth(9);
		ief.SetCellAlign(TextAlignmentCenter,TextAlignmentCenter);

		ief.SetCellString(1,2,_T("日期"));
		ief.SetCellColor(1,2,255,255,255);
		ief.SetCellBackground(0x2f,0x75,0xb5);
		ief.SetCellWidth(16);
		ief.SetCellAlign(TextAlignmentCenter,TextAlignmentCenter);

		ief.SetCellString(1,3,_T("迟到(分钟)"));
		ief.SetCellBackground(0xdc,0xe6,0xf1);
		ief.SetCellWidth(11);
		ief.SetCellAlign(TextAlignmentCenter,TextAlignmentCenter);

		ief.SetCellString(1,4,_T("早退(分钟)"));
		ief.SetCellColor(1,4,255,255,255);
		ief.SetCellBackground(0x2f,0x75,0xb5);
		ief.SetCellWidth(11);
		ief.SetCellAlign(TextAlignmentCenter,TextAlignmentCenter);

		ief.SetCellString(1,5,_T("请假"));
		ief.SetCellBackground(0xdc,0xe6,0xf1);
		ief.SetCellWidth(8);
		ief.SetCellAlign(TextAlignmentCenter,TextAlignmentCenter);

		ief.SetCellString(1,6,_T("未打卡(次)"));
		ief.SetCellColor(1,6,255,255,255);
		ief.SetCellBackground(0x2f,0x75,0xb5);
		ief.SetCellWidth(11);
		ief.SetCellAlign(TextAlignmentCenter,TextAlignmentCenter);

		ief.SetCellString(1,7,_T("罚金(累计分钟/累计次数)"));
		ief.SetCellBackground(0xdc,0xe6,0xf1);
		ief.SetCellWidth(22);
		ief.SetCellAlign(TextAlignmentCenter,TextAlignmentCenter);

		ief.SelectRange(_T("A1"),_T("G1"));
		ief.SetBoardState(xlContinuous,0);

		int ro = 2,punctual = NORMAL,month,times,sumtime,error,rowtmp,odd[5]={0};
		CString c1,c2;
		for (int i=0; i<li.size(); i++){
			Person *p = li[i];
			punctual = NORMAL;
			ief.SetCellString(ro,1,p->GetName());
			times=sumtime=error=0;
			if (p->GetNumberOfDay()){
				month = (*p)[0].GetMonth();
			}
			rowtmp = ro;
			odd[0]=odd[1]=odd[2]=odd[3]=odd[4]=0;
			for(int j=0;j<p->GetNumberOfDay();j++){
				Day d = (*p)[j];
				if (month!=d.GetMonth()){
					times=sumtime=0;
					month = d.GetMonth();
				}
				census(d,punctual);
				m1_Iter=m1.find(month);
				if (m1_Iter == m1.end()){
					continue;
				}else{
					int s = m1_Iter->second%31;
					int e = m1_Iter->second/31;
					if (s>d.GetDay()||e<d.GetDay()){
						continue;
					}
				}
				if (d.off||d.beLate||d.leaveEarly||d.unPunch){
					c1.Format(_T("A%d"),ro);
					c2.Format(_T("G%d"),ro);
					error = 1;
					ief.SetCellString(ro,2,d.GetDate());
					ief.SetCellAlign(TextAlignmentCenter,TextAlignmentCenter);
					if (d.beLate){
						sumtime += d.beLate;
						times ++;
						ief.SetCellInt(ro,3,d.beLate);
						ief.SetCellAlign(TextAlignmentCenter,TextAlignmentCenter);
						ief.SelectRange(c1,c2);
						if (odd[0]){
							ief.SetCellBackground(0xbf,0xbf,0xbf);
						}else{
							ief.SetCellBackground(0xd9,0xd9,0xd9);
						}
						odd[0]=1-odd[0];
					}else{
						odd[0]=0;
					}
					if (d.leaveEarly){
						if (d.beLate){
							c1.Format(_T("D%d"),ro);
						}
						ief.SetCellInt(ro,4,d.leaveEarly);
						ief.SetCellAlign(TextAlignmentCenter,TextAlignmentCenter);
						ief.SelectRange(c1,c2);
						if (odd[1]){
							ief.SetCellBackground(0xbd,0xb7,0xee);
						}else{
							ief.SetCellBackground(0x9b,0xc2,0xe6);
						}
						odd[1]=1-odd[1];
					}else{
						odd[1]=0;
					}
					if (d.off){
						if (d.beLate||d.leaveEarly){
							c1.Format(_T("E%d"),ro);
						}
						ief.SetCellString(ro,5,d.off==1?_T("上午"):(d.off==2?_T("下午"):_T("全天")));
						ief.SetCellAlign(TextAlignmentCenter,TextAlignmentCenter);
						ief.SelectRange(c1,c2);
						if (odd[2]){
							ief.SetCellBackground(0xbf,0x8f,0x00);
						}else{
							ief.SetCellBackground(0xcc,0x66,0x00);
						}
						odd[2]=1-odd[2];
					}else{
						odd[2]=0;
					}
					if (d.unPunch){
						if (d.beLate||d.leaveEarly||d.off){
							c1.Format(_T("F%d"),ro);
						}
						if (d.unPunch&1){
							sumtime += 10;
							times += 1;
						}
						ief.SetCellString(ro,6,d.unPunch==1?_T("1(上午)"):(d.unPunch==2?_T("1(下午)"):_T("2")));
						ief.SetCellAlign(TextAlignmentCenter,TextAlignmentCenter);
						ief.SelectRange(c1,c2);
						if (odd[3]){
							ief.SetCellBackground(0xf2,0x84,0x40);
						}else{
							ief.SetCellBackground(0xf4,0xb0,0x84);
						}
						odd[3]=1-odd[3];
					}else{
						odd[3]=0;
					}
					if (times>5||sumtime>20){
						if (d.beLate||(d.unPunch&1)){
							CString tmp;
							int coin = 10;
							if (times>=8){
								coin *=2;
								if (times>=15){
									coin *=2;
									if (times>=22){
										coin *=2;
									}
								}
							}
							tmp.Format(_T("%d(%d/%d)"),coin,sumtime,times);
							ief.SetCellString(ro,7,tmp);
							ief.SetCellAlign(TextAlignmentCenter,TextAlignmentCenter);
						}
					}
					ro++;
				}
			}
			if(ro==rowtmp){
				c1.Format(_T("A%d"),ro);
				c2.Format(_T("G%d"),ro);
				ief.SelectRange(c1,c2);
				ief.SetCellBackground(0x92,0xd0,0x50);
				ro++;
			}
			c1.Format(_T("A%d"),rowtmp);
			c2.Format(_T("G%d"),ro-1);
			ief.SelectRange(c1,c2);
			ief.SetBoardState(xlContinuous,0);

		}
		for (int i = 0; i < 7; i++){
			c1.Format(_T("%c1"),'A'+i);
			c2.Format(_T("%c%d"),'A'+i,ro-1);
			ief.SelectRange(c1,c2);
			ief.SetBoardState(xlContinuous,0);
		}
		ief.FreezePanes(_T("B2"));
		ief.Save();
		ief.CloseExcelFile();

		for (int i=0; i<li.size(); i++){
			//USES_CONVERSION;
			//TRACE("姓名：%s,早上时间：%d:%d\n",W2A((*li[i]).GetName()),(*li[i])[0].GetArrival()/60,(*li[i])[0].GetArrival()%60);
			delete li[i];
		}
		ief.ReleaseExcel();
		AfxMessageBox(_T("恭喜恭喜，统计成功"));
	}
}

void census(Day &d,int &punctual){
	d.unPunch = 0;
	d.beLate = 0;
	d.leaveEarly = 0;
	d.off = 0;
	switch (d.GetArrival())
	{
	case -1:
		d.unPunch = 1;
		break;
	case -2:
		d.off = 1;
		break;
	case 0:
		break;
	default:
		int tmp = d.GetArrival();
		if (punctual==NORMAL){
			tmp -= NORMAL+5;
			if (tmp>0){
				d.beLate = tmp+5;
			}
		}else{
			tmp -= punctual;
			if(tmp>0){
				d.beLate = tmp;
			}
		}
		break;
	}
	switch (d.GetLeave())
	{
	case -1:
		d.unPunch = d.unPunch|2;
		punctual = NORMAL;
		break;
	case -2:
		d.off = d.off|2;
		punctual = NORMAL;
		break;
	case 0:
		punctual = NORMAL;
		break;
	default:
		int tmp = d.GetLeave();
		if (tmp<REGULAR){
			d.leaveEarly = REGULAR - tmp;
			punctual = NORMAL;
		}else if(tmp>=SOHARD){
			punctual = UNUSUAL;
		}else if (tmp>=OVERTIME){
			punctual = SPECIAL;
		}else{
			punctual = NORMAL;
		}
		break;
	}
}

void CAttenDlg::PrintChild(IXMLDOMNodeList* nodeList, int level)
{
	if(nodeList == NULL)
		return;
	IXMLDOMNode* currentNode = NULL;
	IXMLDOMNodeList* childNodes = NULL;
	IXMLDOMNamedNodeMap* attributes = NULL;
	IXMLDOMNode* attributeID = NULL;
	while(!FAILED(nodeList->nextNode(&currentNode)) && currentNode != NULL)
	{
		BSTR nodeName;
		TESTHR(currentNode->get_nodeName(&nodeName));
		DWORD dwBytesWritten;
		for(int i=0; i<level*INDENT; i++)
			WriteFile(logFile, L" ", (DWORD)(sizeof(WCHAR)), &dwBytesWritten, NULL);
		WriteFile(logFile, nodeName, (DWORD)(wcslen(nodeName)*sizeof(WCHAR)), &dwBytesWritten, NULL);
		TESTHR(currentNode->get_attributes(&attributes));
		if(attributes!=NULL)
		{
			_bstr_t bstrAttributeName = "id";
			BSTR idVal;
			TESTHR(attributes->getNamedItem(bstrAttributeName, &attributeID));
			if(attributeID != NULL){
				TESTHR(attributeID->get_text(&idVal));
				WriteFile(logFile, L" ", (DWORD)(sizeof(WCHAR)), &dwBytesWritten, NULL);
				WriteFile(logFile, idVal, (DWORD)(wcslen(idVal)*sizeof(WCHAR)), &dwBytesWritten, NULL);
				WriteFile(logFile, L"\r\n", (DWORD)(2*sizeof(WCHAR)), &dwBytesWritten, NULL);
				attributeID->Release();
				attributeID = NULL;
			}else{
				WriteFile(logFile, L"\r\n", (DWORD)(2*sizeof(WCHAR)), &dwBytesWritten, NULL);
			}
			attributes->Release();
			attributes = NULL;
		}
		else
		{
			WriteFile(logFile, L"\r\n", (DWORD)(2*sizeof(WCHAR)), &dwBytesWritten, NULL);
		}
		TESTHR(currentNode->get_childNodes(&childNodes));
		PrintChild(childNodes, level+1);
		currentNode=NULL;
	}
fail:
	if(childNodes!=NULL)
		childNodes->Release();
	if(attributeID!=NULL)
		attributeID->Release();
	if(attributes!=NULL)
		attributes->Release();
	if(currentNode != NULL)
		currentNode->Release();
}
