#include <Windows.h>
#include <stdio.h>
#include <atlbase.h>
#include <shobjidl_core.h>

void PrintFunc(CComPtr<ITypeInfo> TypeInfo, FUNCDESC* FunctionDescription)
{
    if (FunctionDescription->invkind != INVOKE_FUNC)
    {
        return;
    }

    BSTR functionName = nullptr;
    HRESULT result = TypeInfo->GetDocumentation(
        FunctionDescription->memid, &functionName, nullptr, nullptr, nullptr);
    if (FAILED(result))
    {
        wprintf(L"GetDocumentation failed with 0x%x\n", result);
        return;
    }

    wprintf(L"  %s\n", functionName);

    SysFreeString(functionName);

    BSTR dllName = nullptr;
    BSTR funcName = nullptr;
    WORD ordinal = 0;
    result = TypeInfo->GetDllEntry(FunctionDescription->memid, INVOKE_FUNC, &dllName, &funcName, &ordinal);
    if (FAILED(result))
    {
        wprintf(L"GetDllEntry failed with 0x%x\n", result);
        return;
    }
    wprintf(L"    DLL: %s, func: %s, ord: %d\n", dllName, funcName, ordinal);

}

void PrintFuncs(wchar_t* ClsidStr)
{
    CLSID clsid = {};
    HRESULT result = CLSIDFromString(ClsidStr, &clsid);
    if (FAILED(result))
    {
        wprintf(L"CLSIDFromString failed with 0x%x\n", result);
        return;
    }

    CComPtr<IDispatch> dispatchableObj;
    result = dispatchableObj.CoCreateInstance(clsid);
    if (FAILED(result))
    {
        wprintf(L"CoCreateInstance failed with 0x%x\n", result);
        return;
    }

    UINT typeInfoCount = 0;
    result = dispatchableObj->GetTypeInfoCount(&typeInfoCount);
    if (FAILED(result))
    {
        wprintf(L"GetTypeInfoCount failed with 0x%x\n", result);
        return;
    }

    wprintf(L"typeInfoCount: 0x%d\n", typeInfoCount);

    CComPtr<ITypeInfo> typeInfo;
    result = dispatchableObj->GetTypeInfo(0, LOCALE_SYSTEM_DEFAULT, &typeInfo);
    if (FAILED(result))
    {
        wprintf(L"GetTypeInfo failed with 0x%x\n", result);
        return;
    }

    TYPEATTR* typeAttributes = nullptr;
    result = typeInfo->GetTypeAttr(&typeAttributes);
    if (FAILED(result))
    {
        wprintf(L"GetTypeAttr failed with 0x%x\n", result);
        return;
    }

    WORD functionCount = typeAttributes->cFuncs;
    typeInfo->ReleaseTypeAttr(typeAttributes);

    wprintf(L"functionCount: %d\n", functionCount);

    for (WORD funcIndex = 0; funcIndex < functionCount; funcIndex++)
    {
        FUNCDESC* funcDescription = nullptr;
        result = typeInfo->GetFuncDesc(funcIndex, &funcDescription);
        if (FAILED(result))
        {
            wprintf(L"GetFuncDesc failed for index %d with 0x%x\n", funcIndex, result);
            return;
        }

        PrintFunc(typeInfo, funcDescription);

        typeInfo->ReleaseFuncDesc(funcDescription);
    }
}

int __cdecl wmain(int argc, wchar_t** argv)
{
    if (argc != 2)
    {
        wprintf(L"Usage: %s <CLSID>\n", argv[0]);
        return 0;
    }

    HRESULT result = CoInitialize(nullptr);
    if (FAILED(result))
    {
        wprintf(L"CoInitialize failed with 0x%x\n", result);
        return 0;
    }

    PrintFuncs(argv[1]);

    CoUninitialize();
    return 0;
}