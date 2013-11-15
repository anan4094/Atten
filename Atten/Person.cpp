#include "stdafx.h"
#include "Person.h"
#include <algorithm>
bool less_day(const Day dsrc, const Day ddest);

Person::Person(void)
{
}


Person::~Person(void)
{
}

CString &Person::GetName(){
	return name;
}

void Person::SetName(CString& name){
	this->name = name;
}

void Person::Add(Day day){
	days.push_back(day);
}

void Person::AddDay(CString date,double arr,double lea){
	Add(Day(date,arr,lea));
}

Day &Person::operator[](int i){
	return days[i];
}

Day &Person::LastDay(){
	return days[days.size()-1];
}

void Person::SortDay(){
	sort(days.begin(),days.end(),less_day);
}

int Person::GetNumberOfDay(){
	return days.size();
}

bool less_day(const Day dsrc, const Day ddest)
{
	if (dsrc.GetMonth()<ddest.GetMonth()){
        return true;
	}else if (dsrc.GetMonth()==ddest.GetMonth()&&dsrc.GetDay()<ddest.GetDay()){
		return true;
	}
    return false;
}
