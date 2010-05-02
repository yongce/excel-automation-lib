/*!
* @file    ExcelAutomation_example.cpp
* @brief   Example for ExcelAutomation
* @date    2009-12-01
* @author  Tu Yongce <tuyongce@gmail.com>
* @version $Id$
*/


/*!
* @example ExcelAutomation_example.cpp
* The following is an example of how to use the library ExcelAutomation.
*/


#include <iostream>
#include "ExcelAutomationLib.h"

using namespace std;
using namespace ExcelAutomation;


int main()
{
    ExcelApplication app;

    if (!app.Startup())
        return -1;

    ExcelWorkbook file = app.OpenWorkbook(ELtext("D:\\Tyc\\Code\\ExcelAutomation\\Example\\C++0x Features Supported by VC.xls"));
    if (file.IsNull())
        return -2;

    ExcelWorksheet activeWorksheet = file.GetActiveWorksheet();
    if (activeWorksheet.IsNull())
        return -3;

    wcout << L"Active worksheet name: " << activeWorksheet.GetName() << endl;

    ExcelRange range = activeWorksheet.GetRange(ELtext('f'), ELtext('j'), 16, 17);
    if (range.IsNull())
        return -4;

    ELstring encodedData(ELtext("2#5#3#abc2#de4#fghi1#34#52352#234#53530#4#32532#32"));
    bool ret = range.WriteData(encodedData.c_str());
    wcout << L"Range write state: " << boolalpha << ret << endl;

    if (!file.Save())
        return -5;

    ELstring data;
    ret = range.ReadData(data);
    wcout << L"Range read state: " << ret << endl;

    wcout << L"Old: " << encodedData << endl;
    wcout << L"New: " << data << endl;
    assert(encodedData == data);

    ExcelWorksheetSet allWorksheets = file.GetAllWorksheets();
    if (allWorksheets.IsNull())
        return -6;

    ExcelWorksheet thirdWorksheet = allWorksheets.GetWorksheet(3);
    if (thirdWorksheet.IsNull())
        return -7;

    wcout << L"Thrid worksheet name: " << thirdWorksheet.GetName() << endl;

    ExcelRange range2 = thirdWorksheet.GetRange(ELtext('C'), ELtext('g'), 16, 17);
    if (range2.IsNull())
        return -8;

    ret = range2.WriteData(encodedData.c_str());
    wcout << L"Range write state: " << boolalpha << ret << endl;

    ExcelCell cell = activeWorksheet.GetCell(ELtext('A'), 4);
    ELstring valueD4;
    if (!cell.IsNull() && cell.GetValue(valueD4))
    {
        wcout << L"D4: " << valueD4 << endl;
        ExcelCell cell2 = activeWorksheet.GetCell(ELtext('J'), 6);
        if (!cell2.IsNull())
        {
            cell2.SetValue(valueD4);
        }
    }


    if (!file.Save())
        return -9;

    if (!file.Close())
        return -10;

    if (!app.Shutdown())
        return -11;

    wcout << L"Test successfully" << endl;

    return 0;
}



