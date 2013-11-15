#pragma once
class Day
{
private:
	CString date;
	int arrival;
	int leave;
	int month;
	int day;
public:
	int unPunch;
	int beLate;
	int leaveEarly;
	int off;
public:
	Day(CString date,double arr,double lea);
	~Day(void);
	int GetArrival() const;
	int GetLeave() const;
	int GetMonth() const;
	int GetDay() const;
	CString GetDate() const;
};

