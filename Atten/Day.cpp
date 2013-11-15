#include "stdafx.h"
#include "Day.h"


Day::Day(CString date,double arr,double lea){
	this->date = date;
	int indm = date.Find(_T("ÔÂ")),indd = date.Find(_T("ÈÕ"));
	if (indm>0){
		month = _ttoi(date.Left(indm));
		if (indd>indm){
			day = _ttoi(date.Mid(indm+1,indd-indm));
		}
	}
	
	if(arr<0){
		this->arrival = (int)arr;
	}else{
		this->arrival = (int)(arr*24*60);
	}
	if(lea<0){
		this->leave = (int)lea;
	}else{
		this->leave = (int)(lea*24*60);
	}
}


Day::~Day(void)
{
}

int Day::GetArrival() const{
	return arrival;
}

int Day::GetLeave()const{
	return leave;
}

CString Day::GetDate()const{
	return date;
}

int Day::GetMonth()const{
	return month;
}

int Day::GetDay()const{
	return day;
}