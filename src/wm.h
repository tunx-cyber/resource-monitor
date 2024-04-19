#pragma once
#define _WIN32_DCOM
#include <iostream>
using namespace std;
#include <comdef.h>
#include <Wbemidl.h>
#include <string>
#pragma comment(lib, "wbemuuid.lib")

class WinResourceMonitor
{
public:
    WinResourceMonitor()
    {
        init();
    }
    ~WinResourceMonitor()
    {
        clean();
    }
    int init()
    {
        // Initialize COM.
        hres = CoInitializeEx(0, COINIT_MULTITHREADED);
        if (FAILED(hres))
        {
            cout << "Failed to initialize COM library. "
                << "Error code = 0x"
                << hex << hres << endl;
            return 1;              // Program has failed.
        }

        // Initialize 
        hres = CoInitializeSecurity(
            NULL,
            -1,      // COM negotiates service                  
            NULL,    // Authentication services
            NULL,    // Reserved
            RPC_C_AUTHN_LEVEL_DEFAULT,    // authentication
            RPC_C_IMP_LEVEL_IMPERSONATE,  // Impersonation
            NULL,             // Authentication info 
            EOAC_NONE,        // Additional capabilities
            NULL              // Reserved
        );


        if (FAILED(hres))
        {
            cout << "Failed to initialize security. "
                << "Error code = 0x"
                << hex << hres << endl;
            CoUninitialize();
            return 1;          // Program has failed.
        }
        // Obtain the initial locator to Windows Management
        // on a particular host computer.
        pLoc = 0;

        hres = CoCreateInstance(
            CLSID_WbemLocator,
            0,
            CLSCTX_INPROC_SERVER,
            IID_IWbemLocator, (LPVOID*)&pLoc);

        if (FAILED(hres))
        {
            cout << "Failed to create IWbemLocator object. "
                << "Error code = 0x"
                << hex << hres << endl;
            CoUninitialize();
            return 1;       // Program has failed.
        }

        pSvc = 0;

        // Connect to the root\cimv2 namespace with the
        // current user and obtain pointer pSvc
        // to make IWbemServices calls.

        hres = pLoc->ConnectServer(

            _bstr_t(L"ROOT\\CIMV2"), // WMI namespace
            NULL,                    // User name
            NULL,                    // User password
            0,                       // Locale
            NULL,                    // Security flags                 
            0,                       // Authority       
            0,                       // Context object
            &pSvc                    // IWbemServices proxy
        );

        if (FAILED(hres))
        {
            cout << "Could not connect. Error code = 0x"
                << hex << hres << endl;
            pLoc->Release();
            CoUninitialize();
            return 1;                // Program has failed.
        }

        /*cout << "Connected to ROOT\\WMI WMI namespace" << endl;*/

        // Set the IWbemServices proxy so that impersonation
        // of the user (client) occurs.
        hres = CoSetProxyBlanket(

            pSvc,                         // the proxy to set
            RPC_C_AUTHN_WINNT,            // authentication service
            RPC_C_AUTHZ_NONE,             // authorization service
            NULL,                         // Server principal name
            RPC_C_AUTHN_LEVEL_CALL,       // authentication level
            RPC_C_IMP_LEVEL_IMPERSONATE,  // impersonation level
            NULL,                         // client identity 
            EOAC_NONE                     // proxy capabilities     
        );

        if (FAILED(hres))
        {
            cout << "Could not set proxy blanket. Error code = 0x"
                << hex << hres << endl;
            pSvc->Release();
            pLoc->Release();
            CoUninitialize();
            return 1;               // Program has failed.
        }
    }

    void clean()
    {
        // Cleanup
        pSvc->Release();
        pLoc->Release();
        pEnumerator->Release();
        CoUninitialize();
    }
    struct CPU
    {
        const wchar_t* Name;
        unsigned int NumberOfCores;
        unsigned int NumberOfLogicalProcessors;
        unsigned int L1CacheSize;
        unsigned int L2CacheSize;
        unsigned int L3CacheSize;
        unsigned short DataWidth;
        std::wstring to_str()
        {
            std::wstring ret = L"CPU INFO:\n";
            ret += L"CPU Name: " + std::wstring(Name) + L"\n";
            ret += L"Data width: " + std::to_wstring(DataWidth) + L"\n";
            ret += L"Number of cores: " + std::to_wstring(NumberOfCores) + L"\n";
            ret += L"Number of logical processors: " + std::to_wstring(NumberOfLogicalProcessors) + L"\n";
            ret += L"L1 cache size: " + std::to_wstring(L1CacheSize) + L" KB\n";
            ret += L"L2 cache size: " + std::to_wstring(L2CacheSize) + L" KB\n";
            ret += L"L3 cache size: " + std::to_wstring(L3CacheSize) + L" KB\n";
            return ret;
        }
    };
    CPU getCpuInfo()
    {

        pEnumerator = NULL;
        CPU c;

        hres = pSvc->ExecQuery(
            bstr_t("WQL"),
            bstr_t("select * from Win32_Processor"),
            WBEM_FLAG_FORWARD_ONLY | WBEM_FLAG_RETURN_IMMEDIATELY,
            NULL,
            &pEnumerator);

        if (FAILED(hres))
        {
            cout << "Query for processes failed. "
                << "Error code = 0x"
                << hex << hres << endl;
            pSvc->Release();
            pLoc->Release();
            CoUninitialize();
            throw 1;               // Program has failed.
        }
        else
        {
            IWbemClassObject* pclsObj;
            ULONG uReturn = 0;

            while (pEnumerator)
            {
                hres = pEnumerator->Next(WBEM_INFINITE, 1,
                    &pclsObj, &uReturn);

                if (0 == uReturn)
                {
                    break;
                }

                VARIANT vtProp;

                // Get the value of the Name property
                hres = pclsObj->Get(L"Name", 0, &vtProp, 0, 0);
                c.Name = vtProp.bstrVal;
                hres = pclsObj->Get(L"NumberOfCores", 0, &vtProp, 0, 0);
                c.NumberOfCores = vtProp.uintVal;
                hres = pclsObj->Get(L"NumberOfLogicalProcessors", 0, &vtProp, 0, 0);
                c.NumberOfLogicalProcessors = vtProp.uintVal;
                hres = pclsObj->Get(L"L2CacheSize", 0, &vtProp, 0, 0);
                c.L2CacheSize = vtProp.uintVal;
                hres = pclsObj->Get(L"L3CacheSize", 0, &vtProp, 0, 0);
                c.L3CacheSize = vtProp.uintVal;
                hres = pclsObj->Get(L"DataWidth", 0, &vtProp, 0, 0);
                c.DataWidth = vtProp.uintVal;
                VariantClear(&vtProp);
                pclsObj->Release();
                pclsObj = NULL;
            }

        }
        hres = pSvc->ExecQuery(
            bstr_t("WQL"),
            bstr_t("select * from Win32_CacheMemory where DeviceID=\"Cache Memory 0\""),
            WBEM_FLAG_FORWARD_ONLY | WBEM_FLAG_RETURN_IMMEDIATELY,
            NULL,
            &pEnumerator);

        if (FAILED(hres))
        {
            cout << "Query for processes failed. "
                << "Error code = 0x"
                << hex << hres << endl;
            pSvc->Release();
            pLoc->Release();
            CoUninitialize();
            throw 1;               // Program has failed.
        }
        else
        {
            IWbemClassObject* pclsObj;
            ULONG uReturn = 0;

            while (pEnumerator)
            {
                hres = pEnumerator->Next(WBEM_INFINITE, 1,
                    &pclsObj, &uReturn);

                if (0 == uReturn)
                {
                    break;
                }

                VARIANT vtProp;

                // Get the value of the Name property
                hres = pclsObj->Get(L"MaxCacheSize", 0, &vtProp, 0, 0);
                c.L1CacheSize = vtProp.uintVal;
                
                VariantClear(&vtProp);
                pclsObj->Release();
                pclsObj = NULL;
            }

        }
        return c;

    }
    struct Disk
    {
        __int64 Size;
        std::wstring to_str()
        {
            std::wstring ret = L"Disk INFO:\n";
            ret += L"Disk size: " + std::to_wstring(Size) + L" GB\n";
            return ret;
        }
    };
    Disk getDiskInfo()
    {
        pEnumerator = NULL;
        Disk d;

        hres = pSvc->ExecQuery(
            bstr_t("WQL"),
            bstr_t("select * from Win32_DiskDrive"),
            WBEM_FLAG_FORWARD_ONLY | WBEM_FLAG_RETURN_IMMEDIATELY,
            NULL,
            &pEnumerator);

        if (FAILED(hres))
        {
            cout << "Query for processes failed. "
                << "Error code = 0x"
                << hex << hres << endl;
            pSvc->Release();
            pLoc->Release();
            CoUninitialize();
            throw 1;               // Program has failed.
        }
        else
        {
            IWbemClassObject* pclsObj;
            ULONG uReturn = 0;

            while (pEnumerator)
            {
                hres = pEnumerator->Next(WBEM_INFINITE, 1,
                    &pclsObj, &uReturn);

                if (0 == uReturn)
                {
                    break;
                }

                VARIANT vtProp;

                // Get the value of the Name property
                CIMTYPE type = CIM_UINT64;
                hres = pclsObj->Get(L"Size", 0, &vtProp, &type , 0);
                std::cout << vtProp.ullVal << std::endl;
                d.Size = 5012;
                VariantClear(&vtProp);
                pclsObj->Release();
                pclsObj = NULL;
            }

        }
        return d;
    }
    bstr_t getMemoryInfo()
    {

    }
    bstr_t getGpuInfo()
    {

    }


private:
    HRESULT hres;
    IWbemLocator* pLoc;
    IWbemServices* pSvc;
    IEnumWbemClassObject* pEnumerator;
};