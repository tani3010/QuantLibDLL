#ifndef QUANTLIBDLL_UTIL_HPP
#define QUANTLIBDLL_UTIL_HPP

#include <Windows.h>

#include <algorithm>
#include <codecvt>
#include <locale>
#include <vector>

#include <ql/time/daycounter.hpp>
#include <ql/time/daycounters/all.hpp>
#include <ql/time/businessdayconvention.hpp>
#include <boost/algorithm/string/classification.hpp>
#include <boost/algorithm/string/split.hpp>
#include "const.hpp"
#include "utilVBA.hpp"

using namespace QuantLib;

template <typename T>
std::vector<T> array2vector(T* array, const std::size_t size) {
    return std::vector<T>(array, array + size);
}

std::string str2safe(const std::string& str) {
    std::string safeString(str);
    safeString.erase(remove(safeString.begin(), safeString.end(), ' '), safeString.end());
    return safeString;
}

std::wstring vstring2wstring(const VARIANT& vstr) {
    std::wstring ws;
    if (vstr.vt == VT_BSTR) {
        ws = std::wstring(vstr.bstrVal, SysStringLen(vstr.bstrVal));
    } else if (vstr.vt == (VT_BSTR | VT_BYREF)) {
        ws = std::wstring(*vstr.pbstrVal, SysStringLen(*vstr.pbstrVal));
    }
    return ws;
}

void wstring2vstring(const std::wstring& wstr, VARIANT& vstr) {
    if (vstr.vt == VT_BSTR) {
        vstr.vt = VT_BSTR;
        vstr.bstrVal = SysAllocString(wstr.c_str());
    } else if (vstr.vt == VT_BSTR || vstr.vt == VT_BYREF) {
        BSTR bs = SysAllocString(wstr.c_str());
        SysReAllocString(vstr.pbstrVal, bs);
        SysFreeString(bs);
    }
}

std::wstring string2wstring(const std::string& str) {
    using convert_typeX = std::codecvt_utf8<wchar_t>;
    std::wstring_convert<convert_typeX, wchar_t> converterX;
    return converterX.from_bytes(str);
}

std::string wstring2string(const std::wstring& wstr) {
    using convert_typeX = std::codecvt_utf8<wchar_t>;
    std::wstring_convert<convert_typeX, wchar_t> converterX;
    return converterX.to_bytes(wstr);
}


BSTR string2BSTR(const std::string& str) {
    BSTR bstr = SysAllocStringByteLen(str.c_str(), str.length());
    return bstr;
}

std::wstring multiByteChar2wstring(const char* srcChar) {
    if (!srcChar) {
        return std::wstring(L"");
    }

    int bufferSize = MultiByteToWideChar(CP_ACP, 0, srcChar, -1, nullptr, 0);
    wchar_t* pwsz = new wchar_t[bufferSize];
    int iResult = MultiByteToWideChar(CP_ACP, 0, srcChar, -1, pwsz, bufferSize);
    std::wstring ws(pwsz);
    delete[] pwsz;
    return ws;
}

std::wstring multiByteBSTR2wstring(const BSTR& bstr) {
    return multiByteChar2wstring((const char*)bstr);
}

std::string str2upper(const std::string& str) {
    std::string upperString(str);
    std::transform(str.begin(), str.end(), upperString.begin(), ::toupper);
    return upperString;
}

template<typename T>
void sort(std::vector<T>& vec) {
    std::sort(vec.begin()(), vec.end());
}

std::string str2safeTermStr(const std::string& str) {
    std::string safeStr = str2upper(str2safe(str));
    if (safeStr == "TOD" || safeStr == "TODAY")
        return std::string("0D");
    else if (safeStr == QuantLibDLL::TERM_OVERNIGHT || safeStr == QuantLibDLL::TERM_OVERNIGHT_SUB)
        return std::string("1D");
    else if (safeStr == QuantLibDLL::TERM_TOMORROWNEXT || safeStr == QuantLibDLL::TERM_TOMORROWNEXT_SUB)
        return std::string("2D");
    else if (safeStr == QuantLibDLL::TERM_SPOTNEXT || safeStr == QuantLibDLL::TERM_SPOTNEXT_SUB)
        return std::string("3D");
    else if (safeStr == QuantLibDLL::TERM_SPOTWEEK || safeStr == QuantLibDLL::TERM_SPOTWEEK_SUB)
        return std::string("1W2D");
    else if (safeStr == "1.5Y")
        return std::string("18M");
    else if (safeStr == "1.75Y")
        return std::string("21M");
    return safeStr;
}

Period str2period(const std::string& str) {
    std::string safeStr = str2upper(str2safe(str));
    return PeriodParser::parse(str2safeTermStr(safeStr));
}

Period str2period(const char* str) {
    std::string tmpStr(str);
    return str2period(tmpStr);
}

DayCounter str2dayCounter(const std::string& str) {
    std::string safeStr = str2upper(str2safe(str));
    if (safeStr == "ACT/365" || safeStr == "ACT365" || safeStr == "A/365" || safeStr == "A365")
        return Actual365Fixed();
    else if (safeStr == "ACT/365F" || safeStr == "ACT365F" || safeStr == "A/365F" || safeStr == "A365F")
        return Actual365Fixed();
    else if (safeStr == "ACT/360" || safeStr == "ACT360" || safeStr == "A/360" || safeStr == "A360")
        return Actual360();
    else if (safeStr == "30/360" || safeStr == "30360")
        return Thirty360(Thirty360::BondBasis);
    else if (safeStr == "30E/360" || safeStr == "30E360")
        return Thirty360(Thirty360::EurobondBasis);
    else if (safeStr == "ACT/ACT" || safeStr == "ACTACT" || safeStr == "A/A" || safeStr == "ACTUAL/ACTUAL")
        return ActualActual(ActualActual::ISDA);
    else if (safeStr == "BUS/252" || safeStr == "BUS252" || safeStr == "BUSINESS/252" || safeStr == "BUSINESS252")
        return Business252();
    else
        return Actual365Fixed();
}

DayCounter str2dayCounter(const char* str) {
    std::string tmpStr(str);
    return str2dayCounter(tmpStr);
}

Calendar str2calendar(const std::string& str) {
    std::string safeStr = str2upper(str2safe(str));
    if (safeStr == "JAPAN" || safeStr == "JAP" || safeStr == "JPN" || safeStr == "JP" || safeStr == "TOKYO" || safeStr == "TKY" || safeStr == "TK" || safeStr == "JPY" || safeStr == "YEN")
        return Japan();
    else if (safeStr == "UNITEDKINGDOM" || safeStr == "UK" || safeStr == "LDN" || safeStr == "LONDON" || safeStr == "GBP" || safeStr == "STERLIN" || safeStr == "GBR" || safeStr == "GB")
        return UnitedKingdom();
    else if (safeStr == "UNITEDSTATES" || safeStr == "USA" || safeStr == "US" || safeStr == "NEWYORK" || safeStr == "NYK" || safeStr == "NY" || safeStr == "USD" || safeStr == "DOLLER")
        return UnitedStates();
    else if (safeStr == "EUR" || safeStr == "EURO" || safeStr == "EU" || safeStr == "TARGET" || safeStr == "TAR")
        return TARGET();
    else if (safeStr == "AUSTRALIA" || safeStr == "AUD" || safeStr == "AUSSIE" || safeStr == "ADOLLER" || safeStr == "AU" || safeStr == "AUS")
        return Australia();
    else if (safeStr == "SWITZERLAND" || safeStr == "SWISS" || safeStr == "CHF" || safeStr == "SFR" || safeStr == "FR" || safeStr == "CF" || safeStr == "CHE")
        return Switzerland();
    else if (safeStr == "NEWZEALAND" || safeStr == "NZD" || safeStr == "NZ" || safeStr == "NZDOLLER" || safeStr == "KIWI" || safeStr == "KIWIDOLLER")
        return NewZealand();
    else if (safeStr == "CANADA" || safeStr == "CAD" || safeStr == "CDOLLER" || safeStr == "CAN" || safeStr == "CA")
        return Canada();
    else if (safeStr == "CHINA" || safeStr == "CNH" || safeStr == "CNY" || safeStr == "CN" || safeStr == "CHN")
        return China();
    else if (safeStr == "HONGKONG" || safeStr == "HKD" || safeStr == "HK" || safeStr == "HKDOLLER")
        return HongKong();
    else if (safeStr == "SOUTHKOREA" || safeStr == "KRW" || safeStr == "KR" || safeStr == "KOR")
        return SouthKorea();
    else if (safeStr == "SINGAPORE" || safeStr == "SGD" || safeStr == "SDOLLER" || safeStr == "SG" || safeStr == "SGP")
        return Singapore();
    else if (safeStr == "BRAZIL" || safeStr == "BRL" || safeStr == "BR" || safeStr == "BRA")
        return Brazil();
    else if (safeStr == "RUSSIA" || safeStr == "RU" || safeStr == "RUS" || safeStr == "RUB" || safeStr == "RUR")
        return Russia();
    else if (safeStr == "GERMANY" || safeStr == "DEU" || safeStr == "DE")
        return Germany();
    else if (safeStr == "FRANCE" || safeStr == "FRA")
        return France();
    else if (safeStr == "HUNGARY" || safeStr == "HUN" || safeStr == "HU" || safeStr == "HUF" || safeStr == "FT" || safeStr == "FORINT")
        return Hungary();
    else if (safeStr == "INDIA" || safeStr == "IND" || safeStr == "IN" || safeStr == "INR" || safeStr == "INDIARUPEE")
        return India();
    else if (safeStr == "INDONESIA" || safeStr == "IDN" || safeStr == "ID" || safeStr == "IDR" || safeStr == "RUPIAH")
        return Indonesia();
    else if (safeStr == "MEXICO" || safeStr == "MEX" || safeStr == "ME" || safeStr == "MXN")
        return Mexico();
    else if (safeStr == "SOUTHAFRICA" || safeStr == "ZAF" || safeStr == "ZA" || safeStr == "ZAR")
        return SouthAfrica();
    else
        return NullCalendar();
}

Calendar str2calendar(const char* str) {
    std::string tmpStr(str);
    return str2calendar(tmpStr);
}

JointCalendar str2jointCalendar(const std::string& str, const std::string& delim) {
    std::vector<std::string> splittedStr;
    boost::algorithm::split(splittedStr, str, boost::is_any_of(delim));
    Calendar tmpcal = NullCalendar();
    JointCalendar cal(tmpcal, tmpcal);
    for (Size i = 0; i < splittedStr.size(); ++i) {
        cal = JointCalendar(cal, str2calendar(splittedStr[i]));
    }
    return cal;
}

JointCalendar str2jointCalendar(const char* str, const char* delim) {
    std::string tmpStr(str);
    std::string tmpStrDelim(delim);
    return str2jointCalendar(tmpStr, tmpStrDelim);
}

BusinessDayConvention str2bdc(const std::string& str) {
    std::string safeStr = str2upper(str2safe(str));
    if (safeStr == QuantLibDLL::BDC_PRECEDING)
        return QuantLib::BusinessDayConvention::Preceding;
    else if (safeStr == QuantLibDLL::BDC_FOLLOWING)
        return QuantLib::BusinessDayConvention::Following;
    else if (safeStr == QuantLibDLL::BDC_MODIFIEDFOLLOWING)
        return QuantLib::BusinessDayConvention::ModifiedFollowing;
    else if (safeStr == QuantLibDLL::BDC_MODIFIEDPRECEDING)
        return QuantLib::BusinessDayConvention::ModifiedPreceding;
    else if (safeStr == QuantLibDLL::BDC_UNADJUSTED)
        return QuantLib::BusinessDayConvention::Unadjusted;
    else if (safeStr == QuantLibDLL::BDC_HALFMONTHMODIFIEDFOLLOWING)
        return QuantLib::BusinessDayConvention::HalfMonthModifiedFollowing;
    else if (safeStr == QuantLibDLL::BDC_NEAREST)
        return QuantLib::BusinessDayConvention::Nearest;
    else
        return QuantLib::BusinessDayConvention::ModifiedFollowing;
}

BusinessDayConvention str2bdc(const char* str) {
    std::string tmpStr(str);
    return str2bdc(tmpStr);
}

DateGeneration::Rule str2dateGeneration(const std::string& str) {
    std::string safeStr = str2upper(str2safe(str));
    if (safeStr == "BACKWARD")
        return DateGeneration::Backward;
    else if (safeStr == "FORWARD")
        return DateGeneration::Forward;
    else if (safeStr == "ZERO")
        return DateGeneration::Zero;
    else if (safeStr == "THIRDWEDNESDAY")
        return DateGeneration::ThirdWednesday;
    else if (safeStr == "TWENTIETH")
        return DateGeneration::Twentieth;
    else if (safeStr == "TWENTIETHIMM")
        return DateGeneration::TwentiethIMM;
    else if (safeStr == "OLDCDS")
        return DateGeneration::OldCDS;
    else if (safeStr == "CDS")
        return DateGeneration::CDS;
    else if (safeStr == "CDS2015")
        return DateGeneration::CDS2015;
    else
        return DateGeneration::Forward;
}

DateGeneration::Rule str2dateGeneration(const char* str) {
    std::string tmpStr(str);
    return str2dateGeneration(tmpStr);
}

void sortTerms(std::vector<std::string>& targetTerms) {
    std::vector<std::pair<QuantLib::Date, std::string> > sortedGrids(targetTerms.size());
    Calendar cal = NullCalendar();
    QuantLib::Date baseDate = cal.adjust(QuantLib::Date::todaysDate(), Following);
    for (size_t i = 0; i < targetTerms.size(); ++i) {
        QuantLib::Date tmpDate = cal.advance(baseDate, str2period(targetTerms[i]));
        std::pair<QuantLib::Date, std::string> gridPair(tmpDate, targetTerms[i]);
        sortedGrids[i] = gridPair;
    }

    std::sort(sortedGrids.begin(), sortedGrids.end());
    for (size_t i = 0; i < targetTerms.size(); ++i) {
        targetTerms[i] = sortedGrids[i].second;
    }
}

void sortTerms(LPSAFEARRAY* targetTerms) {
    LONG ubound = getSafeArrayUBound(targetTerms);
    LONG lbound = getSafeArrayLBound(targetTerms);

    std::vector<std::string> strVec(ubound - lbound + 1);

    VARTYPE vt;
    HRESULT hResult = SafeArrayGetVartype(*targetTerms, &vt);

    for (LONG i = lbound; i <= ubound; ++i) {
        BSTR bstr;
        SafeArrayGetElement(*targetTerms, &i, &bstr);
        strVec[i] = wstring2string(multiByteBSTR2wstring(bstr));
        SysFreeString(bstr);
    }
    sortTerms(strVec);

    for (LONG i = lbound; i <= ubound; ++i) {
        BSTR bstr = SysAllocStringByteLen(strVec[i].c_str(), strVec[i].length());
        SafeArrayPutElement(*targetTerms, &i, bstr);
        SysFreeString(bstr);
    }

}


#endif  // QUANTLIBDLL_UTIL_HPP