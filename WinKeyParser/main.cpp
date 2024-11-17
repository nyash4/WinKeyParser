#include <iostream>
#include <Windows.h>
#include <comdef.h>
#include <Wbemidl.h>

#pragma comment(lib, "wbemuuid.lib")

int main() {
    HRESULT hres;

    // Initialize COM.
    std::cout << "Initializing COM..." << std::endl;
    hres = CoInitializeEx(0, COINIT_MULTITHREADED);
    if (FAILED(hres)) {
        std::cerr << "Failed to initialize COM: " << std::hex << hres << std::endl;
        return 1;
    }

    // Set COM security levels.
    std::cout << "Setting COM security levels..." << std::endl;
    hres = CoInitializeSecurity(
        NULL,
        -1,
        NULL,
        NULL,
        RPC_C_AUTHN_LEVEL_DEFAULT,
        RPC_C_IMP_LEVEL_IMPERSONATE,
        NULL,
        EOAC_NONE,
        NULL
    );

    if (FAILED(hres)) {
        std::cerr << "Failed to set security levels: " << std::hex << hres << std::endl;
        CoUninitialize();
        return 1;
    }

    // Create IWbemLocator instance.
    std::cout << "Creating IWbemLocator instance..." << std::endl;
    IWbemLocator* pLoc = NULL;
    hres = CoCreateInstance(
        CLSID_WbemLocator,
        0,
        CLSCTX_INPROC_SERVER,
        IID_IWbemLocator,
        (LPVOID*)&pLoc
    );

    if (FAILED(hres)) {
        std::cerr << "Failed to create IWbemLocator: " << std::hex << hres << std::endl;
        CoUninitialize();
        return 1;
    }

    // Connect to WMI namespace.
    std::cout << "Connecting to WMI..." << std::endl;
    IWbemServices* pSvc = NULL;
    hres = pLoc->ConnectServer(
        _bstr_t(L"ROOT\\CIMV2"),
        NULL,
        NULL,
        0,
        NULL,
        0,
        0,
        &pSvc
    );

    if (FAILED(hres)) {
        std::cerr << "Failed to connect to WMI: " << std::hex << hres << std::endl;
        pLoc->Release();
        CoUninitialize();
        return 1;
    }
    std::cout << "Successfully connected to WMI." << std::endl;

    // Set security levels on the proxy.
    std::cout << "Setting security levels on the proxy..." << std::endl;
    hres = CoSetProxyBlanket(
        pSvc,
        RPC_C_AUTHN_WINNT,
        RPC_C_AUTHZ_NONE,
        NULL,
        RPC_C_AUTHN_LEVEL_CALL,
        RPC_C_IMP_LEVEL_IMPERSONATE,
        NULL,
        EOAC_NONE
    );

    if (FAILED(hres)) {
        std::cerr << "Failed to set proxy blanket: " << std::hex << hres << std::endl;
        pSvc->Release();
        pLoc->Release();
        CoUninitialize();
        return 1;
    }
    std::cout << "Proxy blanket set successfully." << std::endl;

    // Execute WMI query.
    std::cout << "Executing WMI query..." << std::endl;
    IEnumWbemClassObject* pEnumerator = NULL;
    hres = pSvc->ExecQuery(
        bstr_t("WQL"),
        bstr_t("SELECT OA3xOriginalProductKey FROM SoftwareLicensingService"),
        WBEM_FLAG_FORWARD_ONLY | WBEM_FLAG_RETURN_IMMEDIATELY,
        NULL,
        &pEnumerator
    );

    if (FAILED(hres)) {
        std::cerr << "WMI query failed: " << std::hex << hres << std::endl;
        pSvc->Release();
        pLoc->Release();
        CoUninitialize();
        return 1;
    }
    std::cout << "WMI query executed successfully." << std::endl;

    IWbemClassObject* pclsObj = NULL;
    ULONG uReturn = 0;

    while (pEnumerator) {
        std::cout << "Reading results..." << std::endl;
        HRESULT hr = pEnumerator->Next(WBEM_INFINITE, 1, &pclsObj, &uReturn);
        if (0 == uReturn) {
            std::cout << "No more results." << std::endl;
            break;
        }

        VARIANT vtProp;
        hr = pclsObj->Get(L"OA3xOriginalProductKey", 0, &vtProp, 0, 0);
        if (SUCCEEDED(hr) && vtProp.vt == VT_BSTR) {
            std::wcout << L"Product key found: " << vtProp.bstrVal << std::endl;
        }
        else {
            std::cerr << "Product key not found or error retrieving." << std::endl;
        }
        VariantClear(&vtProp);
        pclsObj->Release();
    }

    // Cleanup.
    pEnumerator->Release();
    pSvc->Release();
    pLoc->Release();
    CoUninitialize();

    std::cout << "Program finished." << std::endl;
    system("pause");

    return 0;
}
