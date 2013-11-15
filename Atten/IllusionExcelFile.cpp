#include "stdafx.h"
#include "IllusionExcelFile.h"

COleVariant
	covTrue((short)TRUE),
	covFalse((short)FALSE),
	covOptional((long)DISP_E_PARAMNOTFOUND, VT_ERROR);    

//
CApplication IllusionExcelFile::excel_application_;


IllusionExcelFile::IllusionExcelFile():
	already_preload_(FALSE)
{
}

IllusionExcelFile::~IllusionExcelFile()
{
	//
	CloseExcelFile();
}


//��ʼ��EXCEL�ļ���
BOOL IllusionExcelFile::InitExcel()
{

	//����Excel 2000������(����Excel) 
	if (!excel_application_.CreateDispatch(_T("Excel.Application"))) 
	{ 
		TRACE("����Excel����ʧ��,�����û�а�װEXCEL������!\n"); 
		return FALSE;
	}else{
		TRACE("����Excel����ɹ�\n");
	}

	excel_application_.put_DisplayAlerts(FALSE); 
	return TRUE;
}

//
void IllusionExcelFile::ReleaseExcel()
{
	excel_application_.Quit();
	excel_application_.ReleaseDispatch();
	excel_application_=NULL;
}

//��excel�ļ�
BOOL IllusionExcelFile::OpenExcelFile(const TCHAR* file_name)
{
	//�ȹر�
	CloseExcelFile();
	CFileFind filefind;
	if (!filefind.FindFile(CString(file_name)))
	{
		//����ģ���ļ��������ĵ� 
		//excel_books_.AttachDispatch(excel_application_.get_Workbooks(),true); 
		//LPDISPATCH lpDis = NULL;
		//lpDis = excel_books_.Add(COleVariant(file_name)); 
		//if (lpDis)
		//{
		//	excel_work_book_.AttachDispatch(lpDis); 
		//	//�õ�Worksheets 
		//	excel_sheets_.AttachDispatch(excel_work_book_.get_Worksheets(),true); 
		//	//��¼�򿪵��ļ�����
		//	open_excel_file_ = file_name;
		//	return TRUE;
		//}
		//return FALSE;
		COleVariant covOptional((long)DISP_E_PARAMNOTFOUND,VT_ERROR);
		excel_books_ = excel_application_.get_Workbooks();
	    //books.AttachDispatch(app.get_Workbooks());�ɴ�������һ��
	    excel_work_book_ = excel_books_.Add(covOptional);
	    //book.AttachDispatch(books.Add(covOptional),true); �ɴ�������һ��
	    excel_sheets_=excel_work_book_.get_Worksheets();
	    //sheets.AttachDispatch(book.get_Worksheets(),true); �ɴ�������һ��
		islocal_ = false;
		open_excel_file_ = file_name;
		return TRUE;
	}else{
		excel_books_ = excel_application_.get_Workbooks();
		excel_work_book_ = excel_books_.Open(file_name,
                covOptional, covOptional, covOptional, covOptional,
                covOptional, covOptional, covOptional, covOptional,
                covOptional, covOptional, covOptional, covOptional,
				covOptional, covOptional);
		excel_sheets_ = excel_work_book_.get_Worksheets();
		islocal_ = true;
		open_excel_file_ = file_name;
		return FALSE;
	}
}

//�رմ򿪵�Excel �ļ�,Ĭ������������ļ�
void IllusionExcelFile::CloseExcelFile(BOOL if_save)
{
	//����Ѿ��򿪣��ر��ļ�
	if (open_excel_file_.IsEmpty() == FALSE)
	{
		//�������,�����û�����,���û��Լ��棬����Լ�SAVE�������Ī���ĵȴ�
		if (if_save)
		{
			ShowInExcel(TRUE);
		}
		else
		{
			//
			excel_work_book_.Close(COleVariant(short(FALSE)),COleVariant(open_excel_file_),covOptional);
			excel_books_.Close();
		}

		//���ļ����������
		open_excel_file_.Empty();
	}



	excel_sheets_.ReleaseDispatch();
	excel_work_sheet_.ReleaseDispatch();
	excel_current_ranges_.ReleaseDispatch();
	excel_current_range_.ReleaseDispatch();
	excel_work_book_.ReleaseDispatch();
	excel_books_.ReleaseDispatch();
}

void IllusionExcelFile::SaveasXSLFile(const CString &xls_file)
{
	excel_work_book_.SaveAs(COleVariant(xls_file),
		covOptional,
		covOptional,
		covOptional,
		covOptional,
		covOptional,
		0,
		covOptional,
		covOptional,
		covOptional,
		covOptional,
		covOptional);
	return;
}

void IllusionExcelFile::Save(){
	if(islocal_){
		excel_work_book_.Save();
	}else{
		islocal_ = true;
		excel_work_book_.SaveCopyAs(COleVariant(open_excel_file_));
		excel_work_book_.put_Saved(TRUE);
	}
}


int IllusionExcelFile::GetSheetCount(){
	return excel_sheets_.get_Count();
}

void IllusionExcelFile::FreezePanes(const TCHAR* cell){
	CRange range = excel_work_sheet_.get_Range(COleVariant(cell), COleVariant(cell));
	range.Select();
	CWindow0 mwin = excel_application_.get_ActiveWindow();
	mwin.put_FreezePanes(1);
	range.ReleaseDispatch();
}

CString IllusionExcelFile::GetSheetName(long table_index)
{
	CWorksheet sheet;
	sheet.AttachDispatch(excel_sheets_.get_Item(COleVariant((long)table_index)),true);
	CString name = sheet.get_Name();
	sheet.ReleaseDispatch();
	return name;
}

//������ż���Sheet���,������ǰ�������еı���ڲ�����
BOOL IllusionExcelFile::LoadSheet(long table_index,BOOL pre_load)
{
	LPDISPATCH lpDis = NULL;
	excel_current_ranges_.ReleaseDispatch();
	excel_work_sheet_.ReleaseDispatch();
	lpDis = excel_sheets_.get_Item(COleVariant((long)table_index));
	if (lpDis)
	{
		excel_work_sheet_.AttachDispatch(lpDis,true);
		excel_current_ranges_.AttachDispatch(excel_work_sheet_.get_Cells(), true);
	}
	else
	{
		return FALSE;
	}

	already_preload_ = FALSE;
	//�������Ԥ�ȼ���
	if (pre_load)
	{
		PreLoadSheet();
		already_preload_ = TRUE;
	}

	return TRUE;
}

void IllusionExcelFile::addSheet(const TCHAR *sheet){
	LPDISPATCH lpDis = NULL;
	excel_sheets_.Add(covOptional,covOptional,covOptional,covOptional);
	lpDis = excel_sheets_.get_Item(COleVariant((long)1));
	if (lpDis){
		excel_work_sheet_.AttachDispatch(lpDis,true);
	}
}

//�������Ƽ���Sheet���,������ǰ�������еı���ڲ�����
BOOL IllusionExcelFile::LoadSheet(const TCHAR* sheet,BOOL pre_load)
{
	LPDISPATCH lpDis = NULL;
	excel_current_ranges_.ReleaseDispatch();
	excel_work_sheet_.ReleaseDispatch();
	lpDis = excel_sheets_.get_Item(COleVariant(sheet));
	if (lpDis)
	{
		excel_work_sheet_.AttachDispatch(lpDis,true);
		excel_current_ranges_.AttachDispatch(excel_work_sheet_.get_Cells(), true);

	}
	else
	{
		return FALSE;
	}
	//
	already_preload_ = FALSE;
	//�������Ԥ�ȼ���
	if (pre_load)
	{
		already_preload_ = TRUE;
		PreLoadSheet();
	}

	return TRUE;
}

//�õ��е�����
int IllusionExcelFile::GetColumnCount()
{
	CRange range;
	CRange usedRange;
	usedRange.AttachDispatch(excel_work_sheet_.get_UsedRange(), true);
	range.AttachDispatch(usedRange.get_Columns(), true);
	int count = range.get_Count();
	usedRange.ReleaseDispatch();
	range.ReleaseDispatch();
	return count;
}

//�õ��е�����
int IllusionExcelFile::GetRowCount()
{
	CRange range;
	CRange usedRange;
	usedRange.AttachDispatch(excel_work_sheet_.get_UsedRange(), true);
	range.AttachDispatch(usedRange.get_Rows(), true);
	int count = range.get_Count();
	usedRange.ReleaseDispatch();
	range.ReleaseDispatch();
	return count;
}

//���һ��CELL�Ƿ����ַ���
BOOL IllusionExcelFile::IsCellString(long irow, long icolumn)
{
	excel_current_range_.ReleaseDispatch();
	excel_current_range_.AttachDispatch(excel_current_ranges_.get_Item (COleVariant((long)irow),COleVariant((long)icolumn)).pdispVal, true);
	COleVariant vResult =excel_current_range_.get_Value2();
	//VT_BSTR��ʾ�ַ���
	if(vResult.vt == VT_BSTR)       
	{
		return TRUE;
	}
	return FALSE;
}

//���һ��CELL�Ƿ�����ֵ
BOOL IllusionExcelFile::IsCellInt(long irow, long icolumn)
{
	excel_current_range_.ReleaseDispatch();
	excel_current_range_.AttachDispatch(excel_current_ranges_.get_Item (COleVariant((long)irow),COleVariant((long)icolumn)).pdispVal, true);
	COleVariant vResult =excel_current_range_.get_Value2();
	//����һ�㶼��VT_R8
	if(vResult.vt == VT_INT || vResult.vt == VT_R8)       
	{
		return TRUE;
	}
	return FALSE;
}

//
CString IllusionExcelFile::GetCellString(long irow, long icolumn)
{

	COleVariant vResult;
	CString str;
	//�ַ���
	if (already_preload_ == FALSE)
	{
		excel_current_range_.ReleaseDispatch();
		excel_current_range_.AttachDispatch(excel_current_ranges_.get_Item (COleVariant((long)irow),COleVariant((long)icolumn)).pdispVal, true);
		vResult =excel_current_range_.get_Value2();
	}
	//�����������Ԥ�ȼ�����
	else
	{
		long read_address[2];
		VARIANT val;
		read_address[0] = irow;
		read_address[1] = icolumn;
		ole_safe_array_.GetElement(read_address, &val);
		vResult = val;
	}

	if(vResult.vt == VT_BSTR)
	{
		str=vResult.bstrVal;
	}
	//����
	else if (vResult.vt==VT_INT)
	{
		str.Format(_T("%d"),vResult.pintVal);
	}
	//8�ֽڵ����� 
	else if (vResult.vt==VT_R8)     
	{
		str.Format(_T("%0.0f"),vResult.dblVal);
	}
	//ʱ���ʽ
	else if(vResult.vt==VT_DATE)    
	{
		SYSTEMTIME st;
		VariantTimeToSystemTime(vResult.date, &st);
		CTime tm(st); 
		str=tm.Format(_T("%Y-%m-%d"));

	}
	//��Ԫ��յ�
	else if(vResult.vt==VT_EMPTY)   
	{
		str=_T("");
	}  

	return str;
}

double IllusionExcelFile::GetCellDouble(long irow, long icolumn)
{
	double rtn_value = 0;
	COleVariant vresult;
	//�ַ���
	if (already_preload_ == FALSE)
	{
		excel_current_range_.ReleaseDispatch();
		excel_current_range_.AttachDispatch(excel_current_ranges_.get_Item (COleVariant((long)irow),COleVariant((long)icolumn)).pdispVal, true);
		vresult =excel_current_range_.get_Value2();
	}
	//�����������Ԥ�ȼ�����
	else
	{
		long read_address[2];
		VARIANT val;
		read_address[0] = irow;
		read_address[1] = icolumn;
		ole_safe_array_.GetElement(read_address, &val);
		vresult = val;
	}

	if (vresult.vt==VT_R8)     
	{
		rtn_value = vresult.dblVal;
	}

	return rtn_value;
}

//VT_R8
int IllusionExcelFile::GetCellInt(long irow, long icolumn)
{
	int num;
	COleVariant vresult;

	if (already_preload_ == FALSE)
	{
		excel_current_range_.ReleaseDispatch();
		excel_current_range_.AttachDispatch(excel_current_ranges_.get_Item (COleVariant((long)irow),COleVariant((long)icolumn)).pdispVal, true);
		vresult = excel_current_range_.get_Value2();
	}
	else
	{
		long read_address[2];
		VARIANT val;
		read_address[0] = irow;
		read_address[1] = icolumn;
		ole_safe_array_.GetElement(read_address, &val);
		vresult = val;
	}
	//
	num = static_cast<int>(vresult.dblVal);

	return num;
}

void IllusionExcelFile::SetCellString(long irow, long icolumn,CString new_string)
{
	COleVariant new_value(new_string);
	excel_current_range_.ReleaseDispatch();
	excel_current_range_ = excel_work_sheet_.get_Range(COleVariant(RCString(irow,icolumn)),covOptional);
	excel_current_range_.put_Value2(new_value);

}

void IllusionExcelFile::SetCellInt(long irow, long icolumn,int new_int)
{
	COleVariant new_value((long)new_int);
	excel_current_range_.ReleaseDispatch();
	excel_current_range_ = excel_work_sheet_.get_Range(COleVariant(RCString(irow,icolumn)),covOptional);
	excel_current_range_.put_Value2(new_value);
}

void IllusionExcelFile::SetCellColor(long irow, long icolumn,long color){
	COleVariant new_color((long)color);

	excel_current_range_.ReleaseDispatch();
	excel_current_range_ = excel_work_sheet_.get_Range(COleVariant(RCString(irow,icolumn)),covOptional);
	CFont0 font = excel_current_range_.get_Font();
	font.put_Color(new_color);
}

void IllusionExcelFile::SetCellColor(long irow, long icolumn, int red, int green, int blue){
	if(red>255)red=255;else if(red<0)red = 0;
	if(green>255)green=255;else if(green<0)green = 0;
	if(blue>255)blue=255;else if(blue<0)blue = 0;
	COleVariant new_color((long)(blue<<16)|(green<<8)|red);

	excel_current_range_.ReleaseDispatch();
	excel_current_range_ = excel_work_sheet_.get_Range(COleVariant(RCString(irow,icolumn)),covOptional);
	CFont0 font = excel_current_range_.get_Font();
	font.put_Color(new_color);
}

void IllusionExcelFile::SetCellBold(long irow, long icolumn, bool bold){
	COleVariant new_bold((long)(bold?1:0));
	excel_current_range_.ReleaseDispatch();
	excel_current_range_ = excel_work_sheet_.get_Range(COleVariant(RCString(irow,icolumn)),covOptional);
	CFont0 font = excel_current_range_.get_Font();
	font.put_Bold(new_bold);
}

void IllusionExcelFile::SetCellAlign(TextAlign h,TextAlign v){
	//�����뷽ʽΪˮƽ��ֱ����
	//ˮƽ���룺Ĭ�ϣ�1,���У�-4108,��-4131,�ң�-4152
	//��ֱ���룺Ĭ�ϣ�2,���У�-4108,��-4160,�ң�-4107
	excel_current_range_.put_HorizontalAlignment(_variant_t((long)h));
	excel_current_range_.put_VerticalAlignment(_variant_t((long)v));
}

void IllusionExcelFile::MergeRange(CString cell1,CString cell2){
	excel_current_range_ = excel_work_sheet_.get_Range(COleVariant(cell1),COleVariant(cell2));
	excel_current_range_.Merge(covOptional);
}
void IllusionExcelFile::SelectRange(CString cell1,CString cell2){
	excel_current_range_ = excel_work_sheet_.get_Range(COleVariant(cell1),COleVariant(cell2));
}
void IllusionExcelFile::SetBoardState(LineStyle style,int color){
	_variant_t	 vLineStyle;
	long weight = 1;
	vLineStyle.vt=VT_I2;
	vLineStyle.lVal=(long)0;
	switch (style)
	{
	case xlContinuous:
		weight = 2;
		vLineStyle.lVal=(long)1;
		break;
	case xlDash:
		weight = 2;
		vLineStyle.lVal=(long)3;
		break;
	case xlDashDot:
		weight = 1;
		vLineStyle.lVal=(long)1;
		break;
	case xlDashDotDot:
		weight = 2;
		vLineStyle.lVal=(long)5;
		break;
	case xlDot:
		break;
	case xlDouble:
		break;
	case xlLineStyleNone:
		weight = 0;
		vLineStyle.lVal=(long)0;
		break;
	case xlSlantDashDot:
		break;
	default:
		break;
	}
	//xlColorIndexAutomatic(-4105)Ϊ�Զ���ɫ,xlColorIndexNone(-4142).
	excel_current_range_._BorderAround(vLineStyle,weight,(long)0,_variant_t((long)color));
}
void IllusionExcelFile::SetCellBackground(int red,int green,int blue){
	if(red>255)red=255;else if(red<0)red = 0;
	if(green>255)green=255;else if(green<0)green = 0;
	if(blue>255)blue=255;else if(blue<0)blue = 0;
	COleVariant new_color((long)(blue<<16)|(green<<8)|red);
	Cnterior it;
	it.AttachDispatch(excel_current_range_.get_Interior());
	it.put_Color(new_color);
}

void IllusionExcelFile::SetCellWidth(int width){
	excel_current_range_.put_ColumnWidth(_variant_t((long)width));
}
//
void IllusionExcelFile::ShowInExcel(BOOL bShow)
{
	excel_application_.put_Visible(bShow);
	excel_application_.put_UserControl(bShow);
}

//���ش򿪵�EXCEL�ļ�����
CString IllusionExcelFile::GetOpenFileName()
{
	return open_excel_file_;
}

//ȡ�ô�sheet������
CString IllusionExcelFile::GetLoadSheetName()
{
	return excel_work_sheet_.get_Name();
}

//ȡ���е����ƣ�����27->AA
char *IllusionExcelFile::GetColumnName(long icolumn)
{   
	static char column_name[64];
	size_t str_len = 0;

	while(icolumn > 0)
	{
		int num_data = icolumn % 26;
		icolumn /= 26;
		if (num_data == 0)
		{
			num_data = 26;
			icolumn--;
		}
		column_name[str_len] = (char)((num_data-1) + 'A' );
		str_len ++;
	}
	column_name[str_len] = '\0';
	//��ת
	_strrev(column_name);

	return column_name;
}

//Ԥ�ȼ���
void IllusionExcelFile::PreLoadSheet()
{

	CRange used_range;

	used_range = excel_work_sheet_.get_UsedRange();


	VARIANT ret_ary = used_range.get_Value2();
	if (!(ret_ary.vt & VT_ARRAY))
	{
		return;
	}
	//
	ole_safe_array_.Clear();
	ole_safe_array_.Attach(ret_ary); 
};

CString IllusionExcelFile::RCString(int irow,int icol){
	CString tmp;
	tmp.Format(_T("%s%d"),GetColumnName(icol),irow);
	return tmp;
}
