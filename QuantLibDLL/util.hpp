#ifndef QUANTLIBDLL_UTIL_HPP
#define QUANTLIBDLL_UTIL_HPP

#include <Windows.h>

#include <algorithm>
#include <vector>

#include <ql/time/daycounter.hpp>
#include <ql/time/daycounters/all.hpp>
#include <ql/time/businessdayconvention.hpp>
#include <boost/algorithm/string/classification.hpp>
#include <boost/algorithm/string/split.hpp>
#include "const.hpp"

using namespace QuantLib;

template <typename T>
std::vector<T> array2vector(T* array, const std::size_t size) {
    return std::vector<T>(array, array + size);
}

BSTR string2BSTR(const std::string& str) {
    int strlen = ::MultiByteToWideChar(
        CP_ACP, 0, str.data(), str.length(), NULL, 0);

    BSTR ret = ::SysAllocStringLen(NULL, strlen);
    ::MultiByteToWideChar(
        CP_ACP, 0, str.data(), str.length(), ret, strlen);
    return ret;
}

std::string str2upper(const std::string& str) {
    std::string upperString(str);
    std::transform(str.begin(), str.end(), upperString.begin(), ::toupper);
    return upperString;
}

std::string str2safe(const std::string& str) {
    std::string safeString(str);
    safeString.erase(remove(safeString.begin(), safeString.end(), ' '), safeString.end());
    return safeString;
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

#endif  // QUANTLIBDLL_UTIL_HPP