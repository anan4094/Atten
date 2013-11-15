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
	xlContinuous    =      0,       //     实线。 
	xlDash          =      1,       //     虚线。 
	xlDashDot       =      4,       //     点划相间线。 
	xlDashDotDot    =      5,       //     划线后跟两个点。 
	xlDot           =      3,       //     点式线。 
	xlDouble        =      9,       //     双线。 
	xlLineStyleNone =      2,       //     无线条。 
	xlSlantDashDot  =      13 ,     //     倾斜的划线。 
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
	///打开的EXCEL文件名称
	CString       open_excel_file_;

	///EXCEL BOOK集合，（多个文件时）
	CWorkbooks    excel_books_; 
	///当前使用的BOOK，当前处理的文件
	CWorkbook     excel_work_book_; 
	///EXCEL的sheets集合
	CWorksheets   excel_sheets_; 
	///当前使用sheet
	CWorksheet    excel_work_sheet_; 
	///当前的操作区域
	CRange        excel_current_ranges_;

	CRange        excel_current_range_; 


	///是否已经预加载了某个sheet的数据
	BOOL          already_preload_;
	///Create the SAFEARRAY from the VARIANT ret.
	COleSafeArray ole_safe_array_;

	BOOL          islocal_;

protected:

	///EXCEL的进程实例
	static CApplication excel_application_;

	CString RCString(int irow,int icol);
public:

	///
	void ShowInExcel(BOOL bShow);

	///检查一个CELL是否是字符串
	BOOL    IsCellString(long iRow, long iColumn);
	///检查一个CELL是否是数值
	BOOL    IsCellInt(long iRow, long iColumn);

	///得到一个CELL的String
	CString GetCellString(long iRow, long iColumn);
	///得到整数
	int     GetCellInt(long iRow, long iColumn);
	///得到double的数据
	double  GetCellDouble(long iRow, long iColumn);

	///取得行的总数
	int GetRowCount();
	///取得列的总数
	int GetColumnCount();

	///使用某个shet，shit，shit
	BOOL LoadSheet(long table_index,BOOL pre_load = FALSE);
	///通过名称使用某个sheet，
	BOOL LoadSheet(const TCHAR* sheet,BOOL pre_load = FALSE);
	void addSheet(const TCHAR* sheet);
	///通过序号取得某个Sheet的名称
	CString GetSheetName(long table_index);

	///得到Sheet的总数
	int GetSheetCount();

	void FreezePanes(const TCHAR* cell);

	///打开文件
	BOOL OpenExcelFile(const TCHAR* file_name);
	///关闭打开的Excel 文件，有时候打开EXCEL文件就要
	void CloseExcelFile(BOOL if_save = FALSE);
	//另存为一个EXCEL文件
	void SaveasXSLFile(const CString &xls_file);
	void Save();
	///取得打开文件的名称
	CString GetOpenFileName();
	///取得打开sheet的名称
	CString GetLoadSheetName();

	///写入一个CELL一个int
	void SetCellInt(long irow, long icolumn,int new_int);
	///写入一个CELL一个string
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
	///初始化EXCEL OLE
	static BOOL InitExcel();
	///释放EXCEL的 OLE
	static void ReleaseExcel();
	///取得列的名称，比如27->AA
	static char *GetColumnName(long iColumn);

protected:

	//预先加载
	void PreLoadSheet();
};
