
// AttenDlg.h : ͷ�ļ�
//

#pragma once
#include "IllusionExcelFile.h"

#define INDENT 4
#define TESTHR(hr) {if(FAILED(hr)) goto fail;}
// CAttenDlg �Ի���
class CAttenDlg : public CDialogEx
{
// ����
public:
	CAttenDlg(CWnd* pParent = NULL);	// ��׼���캯��

// �Ի�������
	enum { IDD = IDD_ATTEN_DIALOG };

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);	// DDX/DDV ֧��


// ʵ��
protected:
	HICON m_hIcon;
	IllusionExcelFile ief;
	HANDLE logFile;
	// ���ɵ���Ϣӳ�亯��
	virtual BOOL OnInitDialog();
	afx_msg void OnPaint();
	afx_msg HCURSOR OnQueryDragIcon();
	DECLARE_MESSAGE_MAP()
public:
	afx_msg void OnBnClickedButton1();
	void PrintChild(IXMLDOMNodeList* nodeList, int level);
};
