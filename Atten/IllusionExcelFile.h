#pragma once
#include "CApplication.h"
#include "CRange.h"
#include "CWorkbook.h"
#include "CWorkbooks.h"
#include "CWorksheet.h"
#include "CWorksheets.h"
#include "CFont0.h"
#include "CWindow0.h"
#include "Cnterior.h"

enum LineStyle {
	xlContinuous    =      0,       //     ʵ�ߡ� 
	xlDash          =      1,       //     ���ߡ� 
	xlDashDot       =      4,       //     �㻮����ߡ� 
	xlDashDotDot    =      5,       //     ���ߺ�������㡣 
	xlDot           =      3,       //     ��ʽ�ߡ� 
	xlDouble        =      9,       //     ˫�ߡ� 
	xlLineStyleNone =      2,       //     �������� 
	xlSlantDashDot  =      13 ,     //     ��б�Ļ��ߡ� 
};

enum TextAlign{
	TextAlignmentLeft = -4131,
	TextAlignmentCenter = -4108,
	TextAlignmentRight = -4152,
	TextAlignmentTop = -4160,
	TextAlignmentBottom = -4107,
};

class IllusionExcelFile
{
public:
	IllusionExcelFile(void);
	~IllusionExcelFile(void);
protected:
	///�򿪵�EXCEL�ļ�����
	CString       open_excel_file_;

	///EXCEL BOOK���ϣ�������ļ�ʱ��
	CWorkbooks    excel_books_; 
	///��ǰʹ�õ�BOOK����ǰ������ļ�
	CWorkbook     excel_work_book_; 
	///EXCEL��sheets����
	CWorksheets   excel_sheets_; 
	///��ǰʹ��sheet
	CWorksheet    excel_work_sheet_; 
	///��ǰ�Ĳ�������
	CRange        excel_current_ranges_;

	CRange        excel_current_range_; 


	///�Ƿ��Ѿ�Ԥ������ĳ��sheet������
	BOOL          already_preload_;
	///Create the SAFEARRAY from the VARIANT ret.
	COleSafeArray ole_safe_array_;

	BOOL          islocal_;

protected:

	///EXCEL�Ľ���ʵ��
	static CApplication excel_application_;

	CString RCString(int irow,int icol);
public:

	///
	void ShowInExcel(BOOL bShow);

	///���һ��CELL�Ƿ����ַ���
	BOOL    IsCellString(long iRow, long iColumn);
	///���һ��CELL�Ƿ�����ֵ
	BOOL    IsCellInt(long iRow, long iColumn);

	///�õ�һ��CELL��String
	CString GetCellString(long iRow, long iColumn);
	///�õ�����
	int     GetCellInt(long iRow, long iColumn);
	///�õ�double������
	double  GetCellDouble(long iRow, long iColumn);

	///ȡ���е�����
	int GetRowCount();
	///ȡ���е�����
	int GetColumnCount();

	///ʹ��ĳ��shet��shit��shit
	BOOL LoadSheet(long table_index,BOOL pre_load = FALSE);
	///ͨ������ʹ��ĳ��sheet��
	BOOL LoadSheet(const TCHAR* sheet,BOOL pre_load = FALSE);
	void addSheet(const TCHAR* sheet);
	///ͨ�����ȡ��ĳ��Sheet������
	CString GetSheetName(long table_index);

	///�õ�Sheet������
	int GetSheetCount();

	void FreezePanes(const TCHAR* cell);

	///���ļ�
	BOOL OpenExcelFile(const TCHAR* file_name);
	///�رմ򿪵�Excel �ļ�����ʱ���EXCEL�ļ���Ҫ
	void CloseExcelFile(BOOL if_save = FALSE);
	//���Ϊһ��EXCEL�ļ�
	void SaveasXSLFile(const CString &xls_file);
	void Save();
	///ȡ�ô��ļ�������
	CString GetOpenFileName();
	///ȡ�ô�sheet������
	CString GetLoadSheetName();

	///д��һ��CELLһ��int
	void SetCellInt(long irow, long icolumn,int new_int);
	///д��һ��CELLһ��string
	void SetCellString(long irow, long icolumn,CString new_string);
	void SetCellColor(long irow, long icolumn,long color);
	void SetCellColor(long irow, long icolumn, int red, int green, int blue);
	void SetCellBold(long irow, long icolumn, bool bold);

	void SetCellBackground(int red,int green,int blue);
	void SetCellWidth(int width);
	void SetCellAlign(TextAlign h,TextAlign v);

	void MergeRange(CString cell1,CString cell2);
	void SelectRange(CString cell1,CString cell2);
	void SetBoardState(LineStyle style,int color);
public:
	///��ʼ��EXCEL OLE
	static BOOL InitExcel();
	///�ͷ�EXCEL�� OLE
	static void ReleaseExcel();
	///ȡ���е����ƣ�����27->AA
	static char *GetColumnName(long iColumn);

protected:

	//Ԥ�ȼ���
	void PreLoadSheet();
};
