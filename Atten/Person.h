#pragma once
#include "Day.h"
#include <vector>
using namespace std;
class Person
{
private:
	CString name;
	vector<Day> days;
public:
	Person(void);
	~Person(void);

	CString &GetName();
	void SetName(CString &name);
	void Add(Day day);
	void AddDay(CString date,double arr,double lea);
	Day &LastDay();
	Day &operator [](int i);
	void SortDay();
	int GetNumberOfDay();
};

