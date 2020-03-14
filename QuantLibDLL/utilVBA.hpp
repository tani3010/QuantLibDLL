#ifndef QUANTLIBDLL_UTILVBA_HPP
#define QUANTLIBDLL_UTILVBA_HPP

#include <Windows.h>
#include <string>

LONG getSafeArrayLBound(const LPSAFEARRAY* safeArray) {
    UINT uiDims = SafeArrayGetDim(*safeArray);
    HRESULT hResult;
    LONG lbound = 0;
    for (UINT i = 1; i <= uiDims; ++i) {
        hResult = SafeArrayGetLBound(*safeArray, i, &lbound);
    }
    return lbound;
}

LONG getSafeArrayUBound(const LPSAFEARRAY* safeArray) {
    UINT uiDims = SafeArrayGetDim(*safeArray);
    HRESULT hResult;
    LONG ubound = 0;
    for (UINT i = 1; i <= uiDims; ++i) {
        hResult = SafeArrayGetUBound(*safeArray, i, &ubound);
    }
    return ubound;
}

class SafeArrayUtil {
public:
    SafeArrayUtil(LPSAFEARRAY* safeArray) : safeArray_(safeArray) {
        lbound_ = getSafeArrayLBound();
        ubound_ = getSafeArrayUBound();
        size_ = ubound_ - lbound_ + 1;
    }

    LONG getSafeArrayUBound() {
        UINT uiDims = SafeArrayGetDim(*safeArray_);
        HRESULT hResult;
        LONG lbound = 0;
        for (UINT i = 1; i <= uiDims; ++i) {
            hResult = SafeArrayGetLBound(*safeArray_, i, &lbound);
        }
        return lbound;
    }

    LONG getSafeArrayLBound() {
        UINT uiDims = SafeArrayGetDim(*safeArray_);
        HRESULT hResult;
        LONG lbound = 0;
        for (UINT i = 1; i <= uiDims; ++i) {
            hResult = SafeArrayGetLBound(*safeArray_, i, &lbound);
        }
        return lbound;
    }

    template<typename T>
    void update(const std::vector<T> updatedValue) {
        if (size_ != updateValue.size()) {
            std::cerr << "vector size of updatedValue is not matched safeArray_." << std::endl;
        }

        VARTYPE vt;
        HRESULT hResult = SafeArrayGetVartype(*safeArray_, &vt);

        switch (vt) {
            case VT_I4:

                break;
            case VT_BSTR:

                break;
            case VT_VARIANT:
                break;

            default:
                break;
        }
    }

private:
    LPSAFEARRAY* safeArray_;
    LONG ubound_;
    LONG lbound_;
    LONG size_;
};

/*
void WINAPI SetArrayGE(const LPSAFEARRAY* ppsa) {
    //格納されているデータ型の確認
    VARTYPE vt;
    HRESULT hResult = SafeArrayGetVartype(*ppsa, &vt);

    if (FAILED(hResult)) {
        //VARIANTに格納できない（？）データの配列の場合、この判定ではNG
        return;
    }

    //次元数
    UINT uiDims = SafeArrayGetDim(*ppsa);

    std::wstringstream ss;

    for (UINT i = 1; i <= uiDims; ++i) {
        LONG lLBound, lUBound;
        hResult = SafeArrayGetLBound(*ppsa, i, &lLBound);
        hResult = SafeArrayGetUBound(*ppsa, i, &lUBound);

        ss << i << L"次元\n"
            << L"　LBound：" << lLBound << L"\n"
            << L"　UBound：" << lUBound << L"\n";
    }

    if (vt == VT_I4) {
        ss << L"データ型：Long\n";

        if (uiDims == 2) {
            LONG lIndex[2];

            LONG lLBound[2];
            LONG lUBound[2];

            hResult = SafeArrayGetLBound(*ppsa, 1, &lLBound[0]);
            hResult = SafeArrayGetUBound(*ppsa, 1, &lUBound[0]);

            for (LONG i = lLBound[0]; i <= lUBound[0]; ++i) {
                hResult = SafeArrayGetLBound(*ppsa, 2, &lLBound[1]);
                hResult = SafeArrayGetUBound(*ppsa, 2, &lUBound[1]);

                lIndex[0] = i;

                for (LONG j = lLBound[1]; j <= lUBound[1]; ++j) {
                    lIndex[1] = j;

                    int iValue;
                    hResult = SafeArrayGetElement(*ppsa, lIndex, &iValue);

                    ss << iValue << L"\n";
                }
            }
        }
    } else if (vt == VT_BSTR) {
        ss << L"データ型：String\n";

        if (uiDims == 1) {
            LONG lLBound;
            LONG lUBound;

            hResult = SafeArrayGetLBound(*ppsa, 1, &lLBound);
            hResult = SafeArrayGetUBound(*ppsa, 1, &lUBound);

            for (LONG lIndex = lLBound; lIndex <= lUBound; ++lIndex) {
                BSTR bstr;

                SafeArrayGetElement(*ppsa, &lIndex, &bstr);

                ss << convMbcBstr2Wstr(bstr) << L"\n";

                SysFreeString(bstr);
            }
        }
    } else if (vt == VT_VARIANT) {
        ss << L"データ型：Variant\n";

        if (uiDims == 1) {
            LONG lLBound;
            LONG lUBound;

            hResult = SafeArrayGetLBound(*ppsa, 1, &lLBound);
            hResult = SafeArrayGetUBound(*ppsa, 1, &lUBound);

            for (LONG lIndex = lLBound; lIndex <= lUBound; ++lIndex) {
                VARIANT vValue;
                VariantInit(&vValue);
                hResult = SafeArrayGetElement(*ppsa, &lIndex, &vValue);

                switch (vValue.vt) {
                case VT_I4:
                    ss << std::to_wstring(vValue.intVal) << L"\n";
                    break;
                case VT_R8:
                    ss << std::to_wstring(vValue.dblVal) << L"\n";
                    break;
                case VT_BSTR:
                    ss << vValue.bstrVal << L"\n";
                    break;
                case VT_BSTR | VT_BYREF:
                    ss << vValue.pbstrVal << L"\n";
                    break;
                default:
                    break;
                };

                VariantClear(&vValue);
            }
        }

    } else {
        MessageBox(NULL, L"No Data Type matched.", L"SetArrayGE", MB_OK | MB_ICONINFORMATION);

        return;
    }

    MessageBox(NULL, ss.str().c_str(), L"SetArrayGE", MB_OK | MB_ICONINFORMATION);

    return;
}

ACCESSIBLEFROMVBA_API void WINAPI SetArrayAD(const LPSAFEARRAY* ppsa) {
    //格納されているデータ型の確認
    VARTYPE vt;
    HRESULT hResult = SafeArrayGetVartype(*ppsa, &vt);

    if (FAILED(hResult)) {
        //VARIANTに格納できない（？）データの配列の場合、この判定ではNG
        return;
    }

    //次元数
    UINT uiDims = SafeArrayGetDim(*ppsa);

    std::wstringstream ss;

    for (UINT i = 1; i <= uiDims; ++i) {
        LONG lLBound, lUBound;
        hResult = SafeArrayGetLBound(*ppsa, i, &lLBound);
        hResult = SafeArrayGetUBound(*ppsa, i, &lUBound);

        ss << i << L"次元\n"
            << L"　LBound：" << lLBound << L"\n"
            << L"　UBound：" << lUBound << L"\n";
    }

    if (vt == VT_I4) {
        ss << L"データ型：Long\n";

        //全要素数
        ULONG ulElems(1);
        for (ULONG i = 0; i < uiDims; ++i) {
            ulElems *= (*ppsa)->rgsabound[i].cElements;
        }

        int* piValue(0);

        hResult = SafeArrayAccessData(*ppsa, (void**)&piValue);

        for (ULONG i = 0; i < ulElems; ++i) {
            ss << L"0x" << std::setfill(L'0') << std::right << std::setw(8) << std::hex << *piValue << L"\n";

            ++piValue;
        }

        hResult = SafeArrayUnaccessData(*ppsa);
    } else if (vt == VT_BSTR) {
        ss << L"データ型：String\n";

        BSTR* pbstr(0);

        hResult = SafeArrayAccessData(*ppsa, (void**)&pbstr);

        std::wstring ws = convMbcBstr2Wstr(*pbstr);

        for (ULONG i = 0; i < (*ppsa)->rgsabound[0].cElements; ++i) {
            ws = convMbcBstr2Wstr(*pbstr);
            ss << ws << L"\n";

            ++pbstr;
        }

        hResult = SafeArrayUnaccessData(*ppsa);
    } else if (vt == VT_VARIANT) {
        ss << L"データ型：Variant\n";

        VARIANT* pv(0);

        hResult = SafeArrayAccessData(*ppsa, (void**)&pv);

        for (ULONG i = 0; i < (*ppsa)->rgsabound[0].cElements; ++i) {
            switch (pv->vt) {
            case VT_I4:
                ss << std::to_wstring(pv->intVal) << L"\n";
                break;
            case VT_R8:
                ss << std::to_wstring(pv->dblVal) << L"\n";
                break;
            case VT_BSTR:
                ss << convVstr2Wstr(*pv) << L"\n";
                break;
            default:
                break;
            }

            ++pv;
        }

        hResult = SafeArrayUnaccessData(*ppsa);
    } else {
        MessageBox(NULL, L"No Data Type matched.", L"SetArrayAD", MB_OK | MB_ICONINFORMATION);

        return;
    }

    MessageBox(NULL, ss.str().c_str(), L"SetArrayAD", MB_OK | MB_ICONINFORMATION);

    return;
}

*/

#endif // QUANTLIBDLL_UTILVBA_HPP