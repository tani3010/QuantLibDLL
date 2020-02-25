#ifndef QUANTLIBDLL_CONST_HPP
#define QUANTLIBDLL_CONST_HPP
#include <string>

namespace QuantLibDLL {

    // return code
    enum returnCode {
        RET_SUCCESS = 1,
        RET_WEARNING = 2,
        RET_ERROR = 8
    };

    // delim
    const std::string DELIM_SLUSH = "/";
    const std::string DELIM_SPACE = " ";
    const std::string DELIM_COLON = ":";
    const std::string DELIM_SEMICOLON = ";";

    // term
    const std::string TERM_OVERNIGHT = "O/N";
    const std::string TERM_TOMORROWNEXT = "T/N";
    const std::string TERM_SPOTNEXT = "S/N";
    const std::string TERM_SPOTWEEK = "S/W";
    const std::string TERM_OVERNIGHT_SUB = "ON";
    const std::string TERM_TOMORROWNEXT_SUB = "TN";
    const std::string TERM_SPOTNEXT_SUB = "SN";
    const std::string TERM_SPOTWEEK_SUB = "SW";

    // sliding rule
    const std::string BDC_FOLLOWING = "FOLLOWING";
    const std::string BDC_MODIFIEDFOLLOWING = "MODIFIEDFOLLOWING";
    const std::string BDC_PRECEDING = "PROCEDING";
    const std::string BDC_MODIFIEDPRECEDING = "MODIFIEDPRECEDING";
    const std::string BDC_UNADJUSTED = "UNADJUSTED";
    const std::string BDC_HALFMONTHMODIFIEDFOLLOWING = "HALFMONTHMODIFIEDFOLLOWING";
    const std::string BDC_NEAREST = "NEAREST";

    // interpolation type


    // daycount
    const std::string DC_ACT360 = "ACT/360";
    const std::string DC_ACT365 = "ACT/365";
    const std::string DC_ACT365F = "ACT/365F";
    const std::string DC_30360 = "30/360";
    const std::string DC_30E360 = "30E/360";
    const std::string DC_ACTACT = "ACT/ACT";
}

#endif  // QUANTLIB_CONST_HPP