//// THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF
//// ANY KIND, EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO
//// THE IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A
//// PARTICULAR PURPOSE.
////
//// Copyright (c) Microsoft Corporation. All rights reserved

#pragma once

#include "Trace.h"

using namespace ABI::Windows::Foundation;
using namespace ABI::Windows::Foundation::Collections;

HRESULT PropertySetFind(
    ComPtr<IPropertySet> & spPropertySet, 
    LPCWSTR key, 
    ComPtr<IInspectable> & spValue);

template <typename Interface>
HRESULT PropertySetFind(
    ComPtr<IPropertySet> & spPropertySet, 
    LPCWSTR key, 
    ComPtr<Interface> & spValue)
{
    ComPtr<IInspectable> spValue2;
    HRESULT hr = PropertySetFind(spPropertySet, key, spValue2);
    if (SUCCEEDED(hr))
    {
        hr = spValue2.As(&spValue);
    }
    TRACEHR_RET(hr);
}

template <typename Interface>
HRESULT PropertySetFind(
    ComPtr<IPropertySet> & spPropertySet, 
    LPCWSTR key, 
    Interface ** pValue)
{
    ComPtr<Interface> spValue;
    HRESULT hr = PropertySetFind(spPropertySet, key, spValue);
    if (SUCCEEDED(hr))
    {
        *pValue = spValue.Detach();
    }
    TRACEHR_RET(hr);
}

HRESULT PropertySetSet(
    ComPtr<IPropertySet> & spPropertySet, 
    LPCWSTR key, 
    UINT32 value);

HRESULT PropertySetSet(
    ComPtr<IPropertySet> & spPropertySet, 
    LPCWSTR key, 
    LPCWSTR value);

HRESULT PropertySetAddSubMap(
    ComPtr<IPropertySet> & spPropertySet, 
    LPCWSTR key, 
    ComPtr<IPropertySet> & spSubMap);
