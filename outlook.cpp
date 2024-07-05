#define NAPI_CPP_EXCEPTIONS 1

#pragma once

#include <Windows.h>
#include <napi.h>
#include "util.h"
#include <iostream>

struct Result {
    HRESULT result;
    VARIANT* value;
};

using NameSpace = IDispatch*;
using Folder = IDispatch*;
using Items = IDispatch*;

class Outlook : public Napi::ObjectWrap<Outlook> {
    public:
        static Napi::Object Init(Napi::Env env, Napi::Object exports);
        Outlook(const Napi::CallbackInfo& info);
        Napi::Value GetFilteredEvents(const Napi::CallbackInfo& info);
    private:
        static Napi::FunctionReference constructor;
        IDispatch* pDispatch;
        DISPID GetDispID(OLECHAR* methodName);
        DISPID GetDispID(OLECHAR* methodName, IDispatch* pDispatch);
        Result Execute(OLECHAR* methodName, DISPPARAMS params);
        Result Execute(OLECHAR* methodName, DISPPARAMS params, IDispatch* pDispatch);
        Result GetProperty(OLECHAR* methodName, DISPPARAMS params);
        Result GetProperty(OLECHAR* methodName, DISPPARAMS params, IDispatch* pDispatch);
        NameSpace GetNameSpace();
        Folder GetDefaultFolder(NameSpace pNamespace, int folderType);
        Items GetFolderItems(Folder pFolder);
        Items RestrictItems(Items pItems, OLECHAR* filter);
        long CountItems(Items pItems);
};

Napi::FunctionReference Outlook::constructor;

Outlook::Outlook(const Napi::CallbackInfo& info) : Napi::ObjectWrap<Outlook>(info) {
    CoInitialize(nullptr);

    CLSID clsid;
    HRESULT result = CLSIDFromProgID(L"Outlook.Application", &clsid);

    if( FAILED(result) ){
        Napi::Error::New(info.Env(), "Failed to get Outlook.Application CLSID").ThrowAsJavaScriptException();
    }

    pDispatch = nullptr;
    result = CoCreateInstance(clsid, nullptr, CLSCTX_LOCAL_SERVER, IID_IDispatch, (void**)&pDispatch);
    if( FAILED(result) ){
        Napi::Error::New(info.Env(), "Failed to create Outlook.Application instance").ThrowAsJavaScriptException();
    }
}

DISPID Outlook::GetDispID(OLECHAR* methodName, IDispatch* pDispatch){
    OLECHAR* oleMethodName = methodName;
    DISPID dispid;
    HRESULT result = pDispatch->GetIDsOfNames(IID_NULL, &oleMethodName, 1, LOCALE_USER_DEFAULT, &dispid);
    if( FAILED(result) ){
        std::cout << "Failed to get DispID for " << WideStringToNarrowString(methodName) << std::endl;
        return -1;
    }

    return dispid;
}

DISPID Outlook::GetDispID(OLECHAR* methodName){
    return GetDispID(methodName, pDispatch);
}

Result Outlook::GetProperty(OLECHAR* methodName, DISPPARAMS params, IDispatch* pDispatch){
    VARIANT* result = nullptr;

    DISPID dispID = GetDispID(methodName, pDispatch);
    if( dispID == -1 ){
        return {DISP_E_MEMBERNOTFOUND, result};
    }

    EXCEPINFO excepInfo;
    UINT argErr;
    HRESULT hResult = pDispatch->Invoke(dispID, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYGET, &params, result, &excepInfo, &argErr);
    return {hResult, result};
}

Result Outlook::GetProperty(OLECHAR* methodName, DISPPARAMS params){
    return Outlook::GetProperty(methodName, params, pDispatch);
}

Result Outlook::Execute(OLECHAR* methodName, DISPPARAMS params, IDispatch* pDispatch){
    VARIANT* result = nullptr;

    DISPID dispID = GetDispID(methodName, pDispatch);
    if( dispID == -1 ){
        return {DISP_E_MEMBERNOTFOUND, result};
    }

    EXCEPINFO excepInfo;
    UINT argErr;
    HRESULT hResult = pDispatch->Invoke(dispID, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, result, &excepInfo, &argErr);
    return {hResult, result};
}

Result Outlook::Execute(OLECHAR* methodName, DISPPARAMS params){
    return Outlook::Execute(methodName, params, pDispatch);
}

NameSpace Outlook::GetNameSpace(){
    // NameSpace Outlook.GetNamespace("MAPI")
    VARIANTARG vtNamespaceArg;
    vtNamespaceArg.vt = VT_BSTR; // String
    vtNamespaceArg.bstrVal = SysAllocString(L"MAPI");
    DISPPARAMS paramsNamespace = {&vtNamespaceArg, nullptr, 1, 0};

    Result namespaceResult = Execute(L"GetNamespace", paramsNamespace);

    if( FAILED(namespaceResult.result) ){
        return nullptr;
    }

    return (*namespaceResult.value).pdispVal;
}

Folder Outlook::GetDefaultFolder(NameSpace pNamespace, int folderType){
    // Folder NameSpace.GetDefaultFolder(OlDefaultFolders.olFolderCalendar)
    VARIANTARG vtFoldersArg;
    vtFoldersArg.vt = VT_I4; // 4-byte integer
    vtFoldersArg.lVal = folderType; // olFolderCalendar
    DISPPARAMS paramsFolders = {&vtFoldersArg, nullptr, 1, 0};

    Result foldersResult = Execute(L"GetDefaultFolder", paramsFolders, pNamespace);

    if( FAILED(foldersResult.result) ){
        return nullptr;
    }

    return (*foldersResult.value).pdispVal;
}

Items Outlook::GetFolderItems(Folder pFolder){
    // Items Folder.Items
    DISPPARAMS noArgs = {nullptr, nullptr, 0, 0};
    Result folderItems = GetProperty(L"Items", noArgs, pFolder);

    if( FAILED(folderItems.result) ){
        return nullptr;
    }

    return (*folderItems.value).pdispVal;
}

Items Outlook::RestrictItems(Items pItems, OLECHAR* filter){
    // Items Items.Restrict("[Start] >= '2021-01-01' AND [End] <= '2021-12-31'")
    VARIANTARG vtRestrictArg;
    vtRestrictArg.vt = VT_BSTR; // String
    vtRestrictArg.bstrVal = SysAllocString(filter);
    DISPPARAMS paramsRestrict = {&vtRestrictArg, nullptr, 1, 0};

    Result restrictResult = Execute(L"Restrict", paramsRestrict, pItems);

    if( FAILED(restrictResult.result) ){
        return nullptr;
    }

    return (*restrictResult.value).pdispVal;

}

long Outlook::CountItems(Items pItems){
    // int Items.Count
    DISPPARAMS noArgs = {nullptr, nullptr, 0, 0};
    Result countResult = GetProperty(L"Count", noArgs, pItems);

    if( FAILED(countResult.result) ){
        return -1;
    }

    return countResult.value->intVal;
}

Napi::Value Outlook::GetFilteredEvents(const Napi::CallbackInfo& info){
    Napi::Env env = info.Env();

    if( info.Length() < 1 ){
        Napi::Error::New(info.Env(), "Invalid number of arguments").ThrowAsJavaScriptException();
        return env.Null();
    }

    std::cout << "Filter: " << info[0].As<Napi::String>().Utf8Value().c_str() << std::endl;

    NameSpace pNamespace = GetNameSpace();
    if( pNamespace == nullptr ){
        Napi::Error::New(info.Env(), "Failed to get NameSpace").ThrowAsJavaScriptException();
        return env.Null();
    }

    std::cout << "got pNamespace" << std::endl;

    Folder pFolder = GetDefaultFolder(pNamespace, 9); // olFolderCalendar
    if( pFolder == nullptr ){
        Napi::Error::New(info.Env(), "Failed to get Calendar Folder").ThrowAsJavaScriptException();
        return env.Null();
    }

    std::cout << "got pFolder" << std::endl;

    Items pFolderItems = GetFolderItems(pFolder);
    if( pFolderItems == nullptr ){
        Napi::Error::New(info.Env(), "Failed to get Calendar Items").ThrowAsJavaScriptException();
        return env.Null();
    }

    std::cout << "got pFolderItems" << std::endl;

    Items pRestrictedItems = RestrictItems(pFolderItems, (OLECHAR*)&NarrowStringToWideString(info[0].As<Napi::String>().Utf8Value().c_str()));
    if( pRestrictedItems == nullptr ){
        Napi::Error::New(info.Env(), "Failed to get Restricted Calendar Items").ThrowAsJavaScriptException();
        return env.Null();
    }

    std::cout << "got pRestrictedItems" << std::endl;

    int count = CountItems(pRestrictedItems);
    if( count == -1 ){
        Napi::Error::New(info.Env(), "Failed to get Calendar Items Count").ThrowAsJavaScriptException();
        return env.Null();
    }

    std::cout << "got count" << std::endl;

    // CLEAN UP the COM objects
    pRestrictedItems->Release();
    pFolderItems->Release();
    pFolder->Release();
    pNamespace->Release();

    return Napi::Number::New(env, count);
}

Napi::Object Outlook::Init(Napi::Env env, Napi::Object exports){
    Napi::HandleScope scope(env);

    Napi::Function func = DefineClass(env, "Outlook", {
        InstanceMethod("getFilteredEvents", &Outlook::GetFilteredEvents)
    });

    constructor = Napi::Persistent(func);
    constructor.SuppressDestruct();

    exports.Set("Outlook", func);
    return exports;
}

Napi::Object Init(Napi::Env env, Napi::Object exports){
    Outlook::Init(env, exports);

    return exports;
}

NODE_API_MODULE(NODE_GYP_MODULE_NAME, Init);