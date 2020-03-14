#ifndef QUANTLIBDLL_INTERPOLATOR_HPP
#define QUANTLIBDLL_INTERPOLATOR_HPP

#include <iostream>

#include <ql/math/interpolation.hpp>
#include <ql/math/interpolations/all.hpp>

#include "const.hpp"
#include "util.hpp"

class Interpolator {
public:
    Interpolator(double* x, double* y, const char* interpolationType, const long size) :
        x_(array2vector<double>(x, size)), y_(array2vector<double>(y, size)),
        interpolationType_(str2upper(str2safe(interpolationType))) {
        setInterpolator();
    }
    double interpolate(double target) {
        if (interpolator_.get() == nullptr)
            return QuantLibDLL::RET_ERROR;

        if (interpolationType_ == "TSPLINE") {
            return interpolator_->operator()(target, true) / target;
        } else {
            return interpolator_->operator()(target, true);
        }
    }

    void interpolate1D(double* target, double* output, const long size) {
        for (long i = 0; i < size; ++i)
            output[i] = interpolate(target[i]);
    }
private:
    const std::vector<double> x_;
    const std::vector<double> y_;
    std::vector<double> xy_;
    const std::string interpolationType_;
    std::unique_ptr<Interpolation> interpolator_;
    void setInterpolator();
};

void Interpolator::setInterpolator() {
    try {
        if (interpolationType_ == "FLAT" || interpolationType_ == "FLATFORWARD" || interpolationType_ == "FORWARDFLAT") {
            interpolator_ = std::make_unique<ForwardFlatInterpolation>(x_.begin(), x_.end(), y_.begin());
        } else if (interpolationType_ == "FLATBACKWARD" || interpolationType_ == "BACKWARDFLAT") {
            interpolator_ = std::make_unique<BackwardFlatInterpolation>(x_.begin(), x_.end(), y_.begin());
        } else if (interpolationType_ == "LINEAR" || interpolationType_ == "LIN" || interpolationType_ == "LINE") {
            interpolator_ = std::make_unique<LinearInterpolation>(x_.begin(), x_.end(), y_.begin());
        } else if (interpolationType_ == "SPLINE" || interpolationType_ == "SPLIN" || interpolationType_ == "CUBICSPLINE" || interpolationType_ == "CSPLINE") {
            interpolator_ = std::make_unique<CubicNaturalSpline>(x_.begin(), x_.end(), y_.begin());
        } else if (interpolationType_ == "MONOTONESPLINE" || interpolationType_ == "MONOSPLIN" || interpolationType_ == "MONOSPLINE" || interpolationType_ == "MONOTONECUBICSPLINE" || interpolationType_ == "MCSPLINE") {
            interpolator_ = std::make_unique<MonotonicCubicNaturalSpline>(x_.begin(), x_.end(), y_.begin());
        } else if (interpolationType_ == "OVERSHOOTINGMINIMIZATION1") {
            interpolator_ = std::make_unique<CubicSplineOvershootingMinimization1>(x_.begin(), x_.end(), y_.begin());
        } else if (interpolationType_ == "OVERSHOOTINGMINIMIZATION2") {
            interpolator_ = std::make_unique<CubicSplineOvershootingMinimization2>(x_.begin(), x_.end(), y_.begin());
        } else if (interpolationType_ == "TSPLINE") {
            xy_.resize(x_.size());
            for (Size i = 0; i < x_.size(); ++i)
                xy_[i] = x_[i] * y_[i];
            interpolator_ = std::make_unique<CubicNaturalSpline>(x_.begin(), x_.end(), xy_.begin());
        } else if (interpolationType_ == "LAGRANGE") {
            interpolator_ = std::make_unique<LagrangeInterpolation>(x_.begin(), x_.end(), y_.begin());
        } else if (interpolationType_ == "LOGLINEAR" || interpolationType_ == "LLINEAR" || interpolationType_ == "LOGLINE" || interpolationType_ == "LLINE") {
            interpolator_ = std::make_unique<LogLinearInterpolation>(x_.begin(), x_.end(), y_.begin());
        } else if (interpolationType_ == "LOGSPLINE" || interpolationType_ == "LSPLINE") {
            interpolator_ = std::make_unique<LogCubicNaturalSpline>(x_.begin(), x_.end(), y_.begin());
        } else if (interpolationType_ == "LOGMONOTONESPLINE" || interpolationType_ == "LMONOSPLINE") {
            interpolator_ = std::make_unique<MonotonicLogCubicNaturalSpline>(x_.begin(), x_.end(), y_.begin());
        } else {
            interpolator_ = nullptr;
        }
    } catch (std::exception e) {
        std::cerr << e.what() << std::endl;
        interpolator_ = nullptr;
    } catch (...) {
        std::cerr << "Unknown error was occurred." << std::endl;
        interpolator_ = nullptr;
    }
}

#endif  // QUANTLIBDLL_INTERPOLATOR_HPP