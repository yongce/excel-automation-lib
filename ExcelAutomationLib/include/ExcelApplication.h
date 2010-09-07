/*!
* @file    ExcelApplication.h
* @brief   Header file for class ExcelApplication
* @date    2009-12-09
* @author  Tu Yongce <tuyongce@gmail.com>
* @version $Id$
*/


#ifndef EXCELAPPLICATION_H_GUID_57C39559_BB99_44E0_A728_B1D6ADBF14CC
#define EXCELAPPLICATION_H_GUID_57C39559_BB99_44E0_A728_B1D6ADBF14CC


#include "LibDef.h"
#include "HandleBody.h"
#include "StringUtil.h"


// <begin> namespace
EXCEL_AUTOMATION_NAMESPACE_START


// Forward declaration
class ExcelWorkbook;


/*!
* @brief Class ExcelApplication represents the concept "Application" in Excel.
* @note ExcelApplication/ExcelApplicationImpl is an implementation of the "Handle/Body" pattern.
*/
class EXCEL_AUTOMATION_DLL_API ExcelApplication : public HandleBase
{
public:
    /*!
    * Default constructor
    */ // Doc is needed by Doxygen
    ExcelApplication();

    bool IsRunning() const;

    bool Startup();

    bool SetVisible(bool visible = true);

    bool Shutdown();

    ExcelWorkbook OpenWorkbook(const ELchar *filename);
    ExcelWorkbook OpenWorkbook(const ELstring &filename);

    ExcelWorkbook CreateWorkbook(const ELchar *filename);
    ExcelWorkbook CreateWorkbook(const ELstring &filename);

private:
    // <begin> Handle/Body pattern implementation
    friend class ExcelApplicationImpl;
    ExcelApplication(ExcelApplicationImpl *impl);
    ExcelApplicationImpl& Body() const;
    // <end> Handle/Body pattern implementation
};


// <end> namespace
EXCEL_AUTOMATION_NAMESPACE_END


#endif //EXCELAPPLICATION_H_GUID_57C39559_BB99_44E0_A728_B1D6ADBF14CC
