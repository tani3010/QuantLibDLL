#include <ql/quantlib.hpp>
#include "interpolator.hpp"
#include "util.hpp"

using namespace QuantLib;

namespace QuantLibDLL {
    void __stdcall QLDLL_sortTerms(LPSAFEARRAY* terms) {
        sortTerms(terms);
    }

    long __stdcall QLDLL_term2date(const long baseDate, const char* term, const char* cal, const char* delim, const char* convention) {
        try {
            const Calendar cal_ = str2jointCalendar(cal, delim);
            QuantLib::Date baseDate_ = cal_.adjust(QuantLib::Date(baseDate), str2bdc(convention));
            return cal_.advance(baseDate_, str2period(term)).serialNumber();
        } catch (std::exception e) {
            std::cerr << e.what() << std::endl;
            return RET_ERROR;
        } catch (...) {
            std::cerr << "Unknown error was occurred." << std::endl;
            return RET_ERROR;
        }
    }

    long __stdcall QLDLL_isHoliday(const long baseDate, const char* cal, const char* delim) {
        try {
            const Calendar cal_ = str2jointCalendar(cal, delim);
            return static_cast<long>(cal_.isHoliday(QuantLib::Date(baseDate)));
        } catch (std::exception e) {
            std::cerr << e.what() << std::endl;
            return RET_ERROR;
        } catch (...) {
            std::cerr << "Unknown error was occurred." << std::endl;
            return RET_ERROR;
        }
    }

    double __stdcall QLDLL_getDayCount(const long d1, const long d2, char* dayCountConvention, const BOOL isYearFraction) {
        try {
            QuantLib::Date d1_ = QuantLib::Date(d1);
            QuantLib::Date d2_ = QuantLib::Date(d2);
            DayCounter dc = str2dayCounter(dayCountConvention);
            return isYearFraction ? dc.yearFraction(d1_, d2_) : dc.dayCount(d1_, d2_);
        } catch (std::exception e) {
            std::cerr << e.what() << std::endl;
            return RET_ERROR;
        } catch (...) {
            std::cerr << "Unknown error was occurred." << std::endl;
            return RET_ERROR;
        }
    }

    long __stdcall QLDLL_getNextIMMdate(const long baseDate, const BOOL isMainCycle) {
        try {
            std::unique_ptr<IMM> imm;
            return imm->nextDate(static_cast<QuantLib::Date>(baseDate), isMainCycle).serialNumber();
        } catch (std::exception e) {
            std::cerr << e.what() << std::endl;
            return RET_ERROR;
        } catch (...) {
            std::cerr << "Unknown error was occurred." << std::endl;
            return RET_ERROR;
        }
    }

    BSTR __stdcall QLDLL_getNextIMMcode(const long baseDate, const BOOL isMainCycle) {
        try {
            std::unique_ptr<IMM> imm;
            std::string tmpStr(imm->nextCode(static_cast<QuantLib::Date>(baseDate), static_cast<bool>(isMainCycle)));
            return string2BSTR(tmpStr);
        } catch (std::exception e) {
            std::cerr << e.what() << std::endl;
            return 0;
        } catch (...) {
            std::cerr << "Unknown error was occurred." << std::endl;
            return 0;
        }
    }

    double __stdcall QLDLL_interpolate(double* x, double* y, double target, const char* interpolationType, const long size) {
        try {
            std::unique_ptr<Interpolator> interpolator = std::make_unique<Interpolator>(x, y, interpolationType, size);
            return interpolator->interpolate(target);
        } catch (std::exception e) {
            std::cerr << e.what() << std::endl;
            return RET_ERROR;
        } catch (...) {
            std::cerr << "Unknown error was occurred." << std::endl;
            return RET_ERROR;
        }
    }

    void __stdcall QLDLL_interpolate1D(
        double* x, double* y, double* target, double* output, const char* interpolationType, const long size, const long targetSize) {
        try {
            std::unique_ptr<Interpolator> interpolator = std::make_unique<Interpolator>(x, y, interpolationType, size);
            interpolator->interpolate1D(target, output, targetSize);
        } catch (std::exception e) {
            std::cerr << e.what() << std::endl;
        } catch (...) {
            std::cerr << "Unknown error was occurred." << std::endl;
        }
    }

    void __stdcall QLDLL_getHolidayList(
        const long beginDate, const long endDate, const BOOL isIncludeWeekend,
        const char* cal, const char* delim,
        long* holidays, const long sizeHolidays) {

        try {
            Calendar cal_ = str2jointCalendar(cal, delim);
            QuantLib::Date beginDate_ = QuantLib::Date(beginDate);
            QuantLib::Date endDate_ = QuantLib::Date(endDate);

            long i = 0;
            while (beginDate_ <= endDate_ && i < sizeHolidays) {
                if (cal_.isHoliday(beginDate_)) {
                    if (isIncludeWeekend) {
                        holidays[i] = beginDate_.serialNumber();
                        i++;
                    } else {
                        switch (beginDate_.weekday()) {
                            case Saturday:
                            case Sunday:
                                break;
                            default:
                                holidays[i] = beginDate_.serialNumber();
                                i++;
                                break;
                        }
                    }
                }
                beginDate_++;
            }
        } catch (std::exception e) {
            std::cerr << e.what() << std::endl;
        } catch (...) {
            std::cerr << "Unknown error was occurred." << std::endl;
        }

    }

    void __stdcall QLDLL_createScheduleByTerminationDate(
        long* outputDates, const long outputDatesSize,
        const long effectiveDate, const long terminationDate,
        const char* tenor, const char* calendar, const char* delim,
        const char* convention, const char* terminationDateConvention,
        const char* dateGeneration, const BOOL endOfMonth,
        const long firstDate, const long nextToLastDate) {
        try {
            std::unique_ptr<Schedule> sche = std::make_unique<Schedule>(
                QuantLib::Date(effectiveDate), QuantLib::Date(terminationDate),
                PeriodParser::parse(str2safeTermStr(tenor)),
                str2jointCalendar(calendar, delim), 
                str2bdc(convention),
                str2bdc(terminationDateConvention),
                str2dateGeneration(dateGeneration),
                endOfMonth,
                firstDate == 0 ? Date() : QuantLib::Date(firstDate),
                nextToLastDate == 0 ? Date() : QuantLib::Date(nextToLastDate)
            );

            for (size_t i = 0; i < sche->dates().size(); ++i) {
                outputDates[i] = sche->dates()[i].serialNumber();
            }
        } catch (std::exception e) {
            std::cerr << e.what() << std::endl;
        } catch (...) {
            std::cerr << "Unknown error was occurred." << std::endl;
        }
    }

    void __stdcall QLDLL_createScheduleByTerminationPeriod(
        long* outputDates, const long outputDatesSize,
        const long effectiveDate, const char* terminationPeriod,
        const char* tenor, const char* calendar, const char* delim,
        const char* convention, const char* terminationDateConvention,
        const char* dateGeneration, const BOOL endOfMonth,
        const long firstDate, const long nextToLastDate) {
        try {
            const long terminationDate = QuantLib::Date(
                QLDLL_term2date(effectiveDate, terminationPeriod, calendar, delim, terminationDateConvention)).serialNumber();

            QLDLL_createScheduleByTerminationDate(
                outputDates, outputDatesSize,
                effectiveDate, terminationDate,
                tenor, calendar, delim,
                convention, terminationDateConvention,
                dateGeneration, endOfMonth,
                firstDate, nextToLastDate);
        } catch (std::exception e) {
            std::cerr << e.what() << std::endl;
        } catch (...) {
            std::cerr << "Unknown error was occurred." << std::endl;
        }
    }

}