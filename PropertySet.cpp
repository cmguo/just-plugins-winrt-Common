// PropertySet.cpp

#include "StdAfx.h"
#include "PropertySet.h"

HRESULT PropertySetFind(
    ComPtr<IPropertySet> & spPropertySet, 
    LPCWSTR key, 
    ComPtr<IInspectable> & spValue)
{
    HStringReference hKey(key);
    ComPtr<IMap<HSTRING, IInspectable*> > spMap;

    HRESULT hr = spPropertySet.As(&spMap);
    if (SUCCEEDED(hr))
    {
        hr = spMap->Lookup(hKey.Get(), &spValue);
    }

    TRACEHR_RET(hr);
}

static ComPtr<IPropertyValueStatics> & GetPropertyFactory()
{

    static ComPtr<IPropertyValueStatics> spFactory;
    if (!spFactory)
    {
        GetActivationFactory(
            HStringReference(RuntimeClass_Windows_Foundation_PropertyValue).Get(),
            &spFactory);
    }
    return spFactory;
}

HRESULT PropertySetSet(
    ComPtr<IPropertySet> & spPropertySet, 
    LPCWSTR key, 
    UINT32 value)
{
    ComPtr<IMap<HSTRING, IInspectable*> > spMap;
    ComPtr<IInspectable> spValue;

    HRESULT hr = spPropertySet.As(&spMap);
    if (SUCCEEDED(hr))
    {
        hr = GetPropertyFactory()->CreateUInt32(value, &spValue);
    }
    if (SUCCEEDED(hr))
    {
        boolean replaced = false;
        hr = spMap->Insert(HStringReference(key).Get(), spValue.Get(), &replaced);
    }

    TRACEHR_RET(hr);
}

HRESULT PropertySetAddSubMap(
    ComPtr<IPropertySet> & spPropertySet, 
    LPCWSTR key, 
    ComPtr<IPropertySet> & spSubMap)
{
    ComPtr<IMap<HSTRING, IInspectable*> > spMap;
    ComPtr<IInspectable> spValue;

    HRESULT hr = spPropertySet.As(&spMap);
    if (SUCCEEDED(hr))
    {
        hr = ::RoActivateInstance(
            HStringReference(RuntimeClass_Windows_Foundation_Collections_PropertySet).Get(),
            &spValue);
    }
    if (SUCCEEDED(hr))
    {
        boolean replaced = false;
        hr = spMap->Insert(HStringReference(key).Get(), spValue.Get(), &replaced);
    }
    if (SUCCEEDED(hr))
    {
        hr = spValue.As(&spSubMap);
    }

    TRACEHR_RET(hr);
}
