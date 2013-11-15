
// AttenDlg.h : 头文件
//

#pragma once
#include "IllusionExcelFile.h"

#define INDENT 4
#define TESTHR(hr) {if(FAILED(hr)) goto fail;}
// CAttenDlg 对话框
class CAttenDlg : public CDialogEx
{
// 构造
public:
	CAttenDlg(CWnd* pParent = NULL);	// 标准构造函数

// 对话框数据
	enum { IDD = IDD_ATTEN_DIALOG };

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);	// DDX/DDV 支持


// 实现
protected:
	HICON m_hIcon;
	IllusionExcelFile ief;
	HANDLE logFile;
	// 生成的消息映射函数
	virtual BOOL OnInitDialog();
	afx_msg void OnPaint();
	afx_msg HCURSOR OnQueryDragIcon();
	DECLARE_MESSAGE_MAP()
public:
	afx_msg void OnBnClickedButton1();
	void PrintChild(IXMLDOMNodeList* nodeList, int level);
};
