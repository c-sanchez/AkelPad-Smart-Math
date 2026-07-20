#include "MathParser.hpp"
#include "MathParserInternal.hpp"

using smartmath_internal::gcdUInt64;
using smartmath_internal::isWithinExactIntFromDoubleRange;
using smartmath_internal::tryExtractExactInt64FromDoubleStrict;
using smartmath_internal::tryMulUInt64Checked;
using smartmath_internal::tryPowUInt64Checked;

#include <algorithm>
#include <cassert>
#include <cmath>
#include <cctype>
#include <cstring>
#include <cstdlib>
#include <limits>
#include <numeric>
#include <unordered_set>
#include <utility>

constexpr MathParser::BuiltinFlags operator|(
    MathParser::BuiltinFlags a,
    MathParser::BuiltinFlags b) noexcept {
  return static_cast<MathParser::BuiltinFlags>(
      static_cast<unsigned>(a) | static_cast<unsigned>(b));
}

#define BF MathParser::BuiltinFlags
#define BHK MathParser::BuiltinHintKind
#define BC MathParser::BuiltinCategory

const MathParser::BuiltinMetaRow MathParser::kBuiltinMeta[] = {
    { BF::NonCalculating, 0, 0, BHK::EmptyPar, BC::Rand },  // Rand
    { BF::FiniteRequired, 2, 2, BHK::MinMax, BC::ScalarBinary },  // Random
    { BF::Format | BF::NonCalculating | BF::TrailingFormatter, 1, MathParser::kBuiltinArityUnbounded, BHK::DotDotDot, BC::BaseFormat },  // Bin
    { BF::Format | BF::NonCalculating | BF::TrailingFormatter, 1, MathParser::kBuiltinArityUnbounded, BHK::DotDotDot, BC::BaseFormat },  // Hex
    { BF::Format | BF::NonCalculating | BF::TrailingFormatter, 1, MathParser::kBuiltinArityUnbounded, BHK::DotDotDot, BC::BaseFormat },  // Oct
    { BF::None, 2, 2, BHK::ValuePower, BC::Pow },  // Pow
    { BF::None, 2, 2, BHK::YX, BC::ScalarBinary },  // Atan2
    { BF::Unary, 1, 1, BHK::Angle, BC::UnaryMath },  // Sin
    { BF::Unary, 1, 1, BHK::Angle, BC::UnaryMath },  // Cos
    { BF::Unary, 1, 1, BHK::Angle, BC::UnaryMath },  // Tan
    { BF::Unary, 1, 1, BHK::Value, BC::UnaryMath },  // Asin
    { BF::Unary, 1, 1, BHK::Value, BC::UnaryMath },  // Acos
    { BF::Unary, 1, 1, BHK::Value, BC::UnaryMath },  // Atan
    { BF::Unary, 1, 1, BHK::Value, BC::UnaryMath },  // Sinh
    { BF::Unary, 1, 1, BHK::Value, BC::UnaryMath },  // Cosh
    { BF::Unary, 1, 1, BHK::Value, BC::UnaryMath },  // Tanh
    { BF::Unary, 1, 1, BHK::Value, BC::UnaryMath },  // Acosh
    { BF::Unary, 1, 1, BHK::Value, BC::UnaryMath },  // Asinh
    { BF::Unary, 1, 1, BHK::Value, BC::UnaryMath },  // Atanh
    { BF::Unary, 1, 1, BHK::Value, BC::UnaryMath },  // Exp
    { BF::None, 2, 2, BHK::ValueBase, BC::Log },  // Log
    { BF::Unary, 1, 1, BHK::Value, BC::UnaryMath },  // Ln
    { BF::Unary, 1, 1, BHK::Value, BC::UnaryMath },  // Log10
    { BF::Unary, 1, 1, BHK::Value, BC::UnaryMath },  // Sqrt
    { BF::Unary, 1, 1, BHK::Value, BC::UnaryMath },  // Sqr
    { BF::Unary, 1, 1, BHK::Value, BC::UnaryMath },  // Int
    { BF::Unary, 1, 1, BHK::Value, BC::UnaryMath },  // Frac
    { BF::Unary, 1, 1, BHK::Value, BC::UnaryMath },  // Abs
    { BF::Unary, 1, 1, BHK::Value, BC::UnaryMath },  // Floor
    { BF::Unary, 1, 1, BHK::Value, BC::UnaryMath },  // Ceil
    { BF::Unary, 1, 1, BHK::Value, BC::UnaryMath },  // Trunc
    { BF::Unary, 1, 1, BHK::Value, BC::UnaryMath },  // Round
    { BF::Unary, 1, 1, BHK::Value, BC::UnaryMath },  // Sign
    { BF::TrailingFormatter, 1, MathParser::kBuiltinArityUnbounded, BHK::DotDotDot, BC::DegRad },  // Deg
    { BF::TrailingFormatter, 1, MathParser::kBuiltinArityUnbounded, BHK::DotDotDot, BC::DegRad },  // Rad
    { BF::None, 1, MathParser::kBuiltinArityUnbounded, BHK::DotDotDot, BC::Aggregate },  // Sum
    { BF::None, 1, MathParser::kBuiltinArityUnbounded, BHK::DotDotDot, BC::Aggregate },  // Median
    { BF::None, 1, MathParser::kBuiltinArityUnbounded, BHK::DotDotDot, BC::Aggregate },  // Variance
    { BF::None, 1, MathParser::kBuiltinArityUnbounded, BHK::DotDotDot, BC::Aggregate },  // Stddev
    { BF::NonCalculating, 1, MathParser::kBuiltinArityUnbounded, BHK::DotDotDot, BC::ArrayTransform },  // Sort
    { BF::NonCalculating, 2, 2, BHK::ArrayFunc, BC::Sortby },  // Sortby
    { BF::Unary, 1, 1, BHK::Value, BC::Ratio },  // Ratio
    { BF::NonCalculating, 2, 3, BHK::StartStopStep, BC::Range },  // Range
    { BF::NonCalculating, 1, MathParser::kBuiltinArityUnbounded, BHK::DotDotDot, BC::ArrayTransform },  // Reverse
    { BF::NonCalculating, 1, MathParser::kBuiltinArityUnbounded, BHK::DotDotDot, BC::ArrayTransform },  // Unique
    { BF::NonCalculating, 1, MathParser::kBuiltinArityUnbounded, BHK::DotDotDot, BC::ArrayTransform },  // Unpack
    { BF::NonCalculating, 2, 2, BHK::XN, BC::Repeat },  // Repeat
    { BF::NonCalculating, 0, MathParser::kBuiltinArityUnbounded, BHK::X, BC::ArrayTransform },  // Len
    { BF::Unary | BF::IntegerOnly, 1, 1, BHK::N, BC::Factorial },  // Fact
    { BF::Unary | BF::IntegerOnly, 1, 1, BHK::N, BC::Factorint },  // Factorint
    { BF::None, 1, MathParser::kBuiltinArityUnbounded, BHK::DotDotDot, BC::Aggregate },  // Avg
    { BF::None, 1, MathParser::kBuiltinArityUnbounded, BHK::DotDotDot, BC::Aggregate },  // Mean
    { BF::IntegerOnly, 2, 2, BHK::ValueDivisor, BC::Mod },  // Mod
    { BF::None, 3, 3, BHK::ValueMinMax, BC::ScalarBinary },  // Clamp
    { BF::None, 2, 2, BHK::XY, BC::ScalarBinary },  // Hypot
    { BF::IntegerOnly, 2, 2, BHK::AB, BC::ScalarBinary },  // Gcd
    { BF::IntegerOnly, 2, 2, BHK::AB, BC::ScalarBinary },  // Lcm
    { BF::IntegerOnly, 2, 2, BHK::AB, BC::ScalarBinary },  // Ncr
    { BF::IntegerOnly, 2, 2, BHK::AB, BC::ScalarBinary },  // Npr
    { BF::None, 1, MathParser::kBuiltinArityUnbounded, BHK::DotDotDot, BC::Aggregate },  // Product
    { BF::None, 1, MathParser::kBuiltinArityUnbounded, BHK::DotDotDot, BC::Aggregate },  // Min
    { BF::None, 1, MathParser::kBuiltinArityUnbounded, BHK::DotDotDot, BC::Aggregate },  // Max
    { BF::Format | BF::NonCalculating | BF::TrailingFormatter, 1, MathParser::kBuiltinArityUnbounded, BHK::DotDotDot, BC::BaseFormat },  // Uhex
    { BF::Format | BF::NonCalculating | BF::TrailingFormatter, 1, MathParser::kBuiltinArityUnbounded, BHK::DotDotDot, BC::BaseFormat },  // Uoct
    { BF::Format | BF::NonCalculating | BF::TrailingFormatter, 1, MathParser::kBuiltinArityUnbounded, BHK::DotDotDot, BC::BaseFormat },  // Ubin
    { BF::None, 1, 1, BHK::Value, BC::TimeUnit },  // Milliseconds
    { BF::None, 1, 1, BHK::Value, BC::TimeUnit },  // Seconds
    { BF::None, 1, 1, BHK::Value, BC::TimeUnit },  // Minutes
    { BF::None, 1, 1, BHK::Value, BC::TimeUnit },  // Hours
    { BF::None, 1, 1, BHK::Value, BC::TimeUnit },  // Days
    { BF::Unary, 1, 1, BHK::Value, BC::UnaryMath },  // Real
    { BF::Unary, 1, 1, BHK::Value, BC::UnaryMath },  // Imag
    { BF::Unary, 1, 1, BHK::Value, BC::UnaryMath },  // Phase
    { BF::Unary, 1, 1, BHK::Value, BC::PolarCart },  // Polar
    { BF::None, 1, 2, BHK::Value, BC::PolarCart },  // Cart
    { BF::Unary, 1, 1, BHK::Value, BC::UnaryMath },  // Conj
};
#undef BC
#undef BHK
#undef BF

constexpr std::size_t builtinMetaRowCountForAssert() {
  return sizeof(MathParser::kBuiltinMeta) / sizeof(MathParser::kBuiltinMeta[0]);
}
static_assert(builtinMetaRowCountForAssert() == 76, "kBuiltinMeta size mismatch");


namespace {
constexpr double kPi = 3.1415926535897932384626433832795;
constexpr double kInvLn10 = 0.43429448190325182046017448968723710;  // 1 / ln(10)
constexpr double kTrigMaxAbsRadians = 9007199254740992.0;  // 2^53
constexpr int kMaxEvalDepth = 128;

bool isTrigRadiansInRange(double radians) {
  if (!std::isfinite(radians)) {
    return false;
  }
  return std::fabs(radians) < kTrigMaxAbsRadians;
}

const char* skipAsciiSpacesPtr(const char* p) {
  while (*p && std::isspace(static_cast<unsigned char>(*p))) {
    ++p;
  }
  return p;
}

bool localIsIdentStart(char c) {
  return std::isalpha(static_cast<unsigned char>(c)) || c == '_';
}

bool localIsIdentChar(char c) {
  return std::isalnum(static_cast<unsigned char>(c)) || c == '_';
}

const char* endOfIdentTokenPtr(const char* p) {
  if (!localIsIdentStart(*p)) {
    return p;
  }
  ++p;
  while (localIsIdentChar(*p)) {
    ++p;
  }
  return p;
}

bool peekIdentFollowedByChar(const char* p, char ch) {
  p = skipAsciiSpacesPtr(endOfIdentTokenPtr(p));
  return *p == ch;
}

bool peekIdentFollowedByAssignEquals(const char* p) {
  p = skipAsciiSpacesPtr(endOfIdentTokenPtr(p));
  return *p == '=' && p[1] != '=';
}

#if SMARTMATH_LAMBDA_FUNCTIONS
bool peekUnwrappedLambdaParamsThenColon(const char* p) {
  p = skipAsciiSpacesPtr(p);
  if (!localIsIdentStart(*p)) {
    return false;
  }
  p = skipAsciiSpacesPtr(endOfIdentTokenPtr(p));
  while (*p == ',') {
    p = skipAsciiSpacesPtr(p + 1);
    if (!localIsIdentStart(*p)) {
      return false;
    }
    p = skipAsciiSpacesPtr(endOfIdentTokenPtr(p));
  }
  return *p == ':';
}

bool peekParenParamListThenColon(const char* p) {
  p = skipAsciiSpacesPtr(p);
  if (*p != '(') {
    return false;
  }
  p = skipAsciiSpacesPtr(p + 1);
  if (*p == ')') {
    p = skipAsciiSpacesPtr(p + 1);
    return *p == ':';
  }
  if (!localIsIdentStart(*p)) {
    return false;
  }
  for (;;) {
    p = skipAsciiSpacesPtr(endOfIdentTokenPtr(p));
    if (*p == ',') {
      p = skipAsciiSpacesPtr(p + 1);
      if (!localIsIdentStart(*p)) {
        return false;
      }
      continue;
    }
    if (*p == ')') {
      p = skipAsciiSpacesPtr(p + 1);
      return *p == ':';
    }
    return false;
  }
}

bool peekRhsMayBeLambdaSyntaxAt(const char* p) {
  p = skipAsciiSpacesPtr(p);
  if (*p == '(') {
    // (x,y):body, ():body, or one extra grouping paren around those forms
    if (peekParenParamListThenColon(p)) {
      return true;
    }
    return peekRhsMayBeLambdaSyntaxAt(p + 1);
  }
  // x:body, x,y:body, x,y,z:body
  return peekUnwrappedLambdaParamsThenColon(p);
}
#endif

bool isMultipleOf(double x, double x_mult)
{
  if (x_mult == 0.0) {
    return false;
  }
  const double q = x / x_mult;
  if (!std::isfinite(q)) {
    return false;
  }
  const double n = std::round(q);
  return std::fabs(q - n) <= 1e-9;
}

bool tryTrigHalfPiQuotient(double x, long long& outK) {
  constexpr double kHalfPi = kPi / 2.0;
  if (!isMultipleOf(x, kHalfPi)) {
    return false;
  }
  const double q = std::round(x / kHalfPi);
  if (!std::isfinite(q) || std::fabs(q) >= kTrigMaxAbsRadians) {
    return false;
  }
  outK = static_cast<long long>(q);
  return true;
}

bool tryTrigQuarterPiQuotient(double x, long long& outK) {
  constexpr double kQuarterPi = kPi / 4.0;
  if (!isMultipleOf(x, kQuarterPi)) {
    return false;
  }
  const double q = std::round(x / kQuarterPi);
  if (!std::isfinite(q) || std::fabs(q) >= kTrigMaxAbsRadians) {
    return false;
  }
  outK = static_cast<long long>(q);
  return true;
}

double calcSin(double x)
{
  if (x == 0.0) {
    return 0.0;
  }
  if (!std::isfinite(x)) {
    return std::sin(x);
  }
  if (isMultipleOf(x, kPi)) {
    return 0.0;
  }
  long long k = 0;
  if (tryTrigHalfPiQuotient(x, k)) {
    long long r = k % 4;
    if (r < 0) {
      r += 4;
    }
    if (r == 1) {
      return 1.0;
    }
    if (r == 3) {
      return -1.0;
    }
    return 0.0;
  }
  return std::sin(x);
}

double calcCos(double x)
{
  if (!std::isfinite(x)) {
    return std::cos(x);
  }
  if (!isMultipleOf(x, kPi)) {
    long long k = 0;
    if (tryTrigHalfPiQuotient(x, k) && (k % 2) != 0) {
      return 0.0;
    }
  }
  return std::cos(x);
}

double calcTan(double x)
{
  if (x == 0.0) {
    return 0.0;
  }
  if (!std::isfinite(x)) {
    return std::tan(x);
  }
  if (isMultipleOf(x, kPi)) {
    return 0.0;
  }
  long long k = 0;
  if (tryTrigHalfPiQuotient(x, k) && (k % 2) != 0) {
    return (k > 0) ? std::numeric_limits<double>::infinity()
                   : -std::numeric_limits<double>::infinity();
  }
  if (tryTrigQuarterPiQuotient(x, k) && (k % 2) != 0) {
    long long r = k % 4;
    if (r < 0) {
      r += 4;
    }
    if (r == 1) {
      return 1.0;
    }
    if (r == 3) {
      return -1.0;
    }
  }
  return std::tan(x);
}

double calcAtan2Basic(double y, double x) {
  if (x > 0.0) {
    return std::atan(y / x);
  }
  if (x < 0.0) {
    if (y >= 0.0) {
      return std::atan(y / x) + kPi;
    }
    return std::atan(y / x) - kPi;
  }
  if (y > 0.0) {
    return kPi / 2.0;
  }
  if (y < 0.0) {
    return -kPi / 2.0;
  }
  return 0.0;
}

double calcHypotBasic(double x, double y) {
  return std::sqrt((x * x) + (y * y));
}

bool fractionalPowerIsOddUnitRoot(double p, long long& outN);
bool fractionalPowerResolveRootDegree(double p, long long& outN);

#if SMARTMATH_COMPLEX_NUMBERS
void snapComplexNearZeroAxis(double& zr, double& zi) {
  if (!std::isfinite(zr) || !std::isfinite(zi)) {
    return;
  }
  constexpr double kRel = 1e-13;
  if (std::fabs(zr) <= kRel * std::fmax(1.0, std::fabs(zi))) {
    zr = 0.0;
  }
  if (std::fabs(zi) <= kRel * std::fmax(1.0, std::fabs(zr))) {
    zi = 0.0;
  }
}

void complexCartesianPrincipalNthRoot(double ar, double ai, double invN, double& outR, double& outI) {
  const double nanv = std::numeric_limits<double>::quiet_NaN();
  if (!std::isfinite(ar) || !std::isfinite(ai)) {
    outR = nanv;
    outI = nanv;
    return;
  }
  const double mag = std::hypot(ar, ai);
  if (mag == 0.0) {
    outR = 0.0;
    outI = 0.0;
    return;
  }
  const double angN = calcAtan2Basic(ai, ar) * invN;
  double rmN = 0.0;
  if (std::fabs(invN - 0.5) < 1e-12) {
    rmN = std::sqrt(mag);
  } else {
    rmN = std::exp(std::log(mag) * invN);
  }
  outR = rmN * calcCos(angN);
  outI = rmN * calcSin(angN);
  snapComplexNearZeroAxis(outR, outI);
}

void scalarPrincipalLnCartesian(double ar, double ai, double& outRe, double& outIm) {
  const double nanv = std::numeric_limits<double>::quiet_NaN();
  if (!std::isfinite(ar) || !std::isfinite(ai)) {
    outRe = nanv;
    outIm = nanv;
    return;
  }
  const double mag = std::hypot(ar, ai);
  if (mag == 0.0) {
    outRe = -std::numeric_limits<double>::infinity();
    outIm = 0.0;
    return;
  }
  constexpr double kAxisEps = 1e-15;
  if (std::fabs(ai) <= kAxisEps * std::fmax(1.0, std::fabs(ar)) && ar < 0.0) {
    outRe = std::log(-ar);
    outIm = kPi;
    snapComplexNearZeroAxis(outRe, outIm);
    return;
  }
  outRe = std::log(mag);
  outIm = calcAtan2Basic(ai, ar);
  snapComplexNearZeroAxis(outRe, outIm);
}

void complexPowPrincipal(double ar, double ai, double br, double bi, double& outRe, double& outIm) {
  const double nanv = std::numeric_limits<double>::quiet_NaN();
  if (!std::isfinite(ar) || !std::isfinite(ai) || !std::isfinite(br) || !std::isfinite(bi)) {
    outRe = nanv;
    outIm = nanv;
    return;
  }
  const double mag = std::hypot(ar, ai);
  if (mag == 0.0) {
    if (br == 0.0 && bi == 0.0) {
      outRe = 1.0;
      outIm = 0.0;
    } else if (br > 0.0) {
      outRe = 0.0;
      outIm = 0.0;
    } else if (br == 0.0 && bi != 0.0) {
      outRe = 0.0;
      outIm = 0.0;
    } else {
      outRe = nanv;
      outIm = nanv;
    }
    return;
  }
  constexpr double kEps = 1e-14;
  if (std::fabs(bi) < kEps) {
    long long nRootFrac = 0;
    if (fractionalPowerResolveRootDegree(br, nRootFrac)) {
      complexCartesianPrincipalNthRoot(ar, ai, 1.0 / static_cast<double>(nRootFrac), outRe, outIm);
      return;
    }
    if (std::fabs(br - std::trunc(br)) < 1e-12) {
      const long long n = static_cast<long long>(std::trunc(br));
      if (n >= -256 && n <= 256) {
        if (n >= 0) {
          double cr = 1.0;
          double ci = 0.0;
          for (long long k = 0; k < n; ++k) {
            const double nr = cr * ar - ci * ai;
            const double ni = cr * ai + ci * ar;
            cr = nr;
            ci = ni;
            if (!std::isfinite(cr) && !std::isfinite(ci)) {
              break;
            }
          }
          outRe = cr;
          outIm = ci;
        } else {
          double pr = 0.0;
          double pi = 0.0;
          double cr = 1.0;
          double ci = 0.0;
          for (long long k = 0; k < -n; ++k) {
            const double nr = cr * ar - ci * ai;
            const double ni = cr * ai + ci * ar;
            cr = nr;
            ci = ni;
            if (!std::isfinite(cr) && !std::isfinite(ci)) {
              break;
            }
          }
          pr = cr;
          pi = ci;
          const double den = pr * pr + pi * pi;
          if (den == 0.0) {
            outRe = nanv;
            outIm = nanv;
          } else {
            outRe = pr / den;
            outIm = (-pi) / den;
          }
        }
        if (std::isfinite(outRe) && std::isfinite(outIm)) {
          snapComplexNearZeroAxis(outRe, outIm);
        }
        return;
      }
    }
  }
  const double loR = std::log(mag);
  const double loI = calcAtan2Basic(ai, ar);
  const double powRe = br * loR - bi * loI;
  const double powIm = br * loI + bi * loR;
  if (!isTrigRadiansInRange(powIm)) {
    outRe = std::numeric_limits<double>::quiet_NaN();
    outIm = std::numeric_limits<double>::quiet_NaN();
    return;
  }
  outRe = std::exp(powRe) * calcCos(powIm);
  outIm = std::exp(powRe) * calcSin(powIm);
  snapComplexNearZeroAxis(outRe, outIm);
}

void complexDivide(double numR, double numI, double denR, double denI, double& outR, double& outI) {
  const double den = denR * denR + denI * denI;
  if (den == 0.0) {
    outR = std::numeric_limits<double>::quiet_NaN();
    outI = std::numeric_limits<double>::quiet_NaN();
  } else {
    outR = (numR * denR + numI * denI) / den;
    outI = (numI * denR - numR * denI) / den;
  }
}

void complexMultiply(double ar, double ai, double br, double bi, double& outR, double& outI) {
  if (ai == 0.0 && br == 0.0) {
    outR = 0.0;
  } else if (bi == 0.0) {
    outR = ar * br;
    if (ar == 0.0 && std::isnan(outR)) {
      outR = 0.0;
    }
  } else if (ai == 0.0) {
    outR = ar * br;
  } else if (br == 0.0) {
    outR = -ai * bi;
  } else {
    outR = ar * br - ai * bi;
  }
  if (ai == 0.0) {
    outI = ar * bi;
  } else if (bi == 0.0) {
    outI = ai * br;
  } else {
    outI = ar * bi + ai * br;
  }
}

bool complexExpCartesian(double ar, double ai, double& outR, double& outI) {
  if (!isTrigRadiansInRange(ai)) {
    return false;
  }
  const double ea = std::exp(ar);
  outR = ea * calcCos(ai);
  outI = ea * calcSin(ai);
  return true;
}

void complexPrincipalSqrt(double ar, double ai, double& outR, double& outI) {
  complexCartesianPrincipalNthRoot(ar, ai, 0.5, outR, outI);
}

void complexGamma(double zr, double zi, double& outR, double& outI) {
  const double nanv = std::numeric_limits<double>::quiet_NaN();
  if (!std::isfinite(zr) || !std::isfinite(zi)) {
    outR = nanv;
    outI = nanv;
    return;
  }
  static constexpr double kLanczosC[] = {
      0.99999999999980993,
      676.5203681218851,
      -1259.1392167224028,
      771.32342877765313,
      -176.61502916214059,
      12.507343278686905,
      -0.13857109526572012,
      9.9843695780195716e-6,
      1.5056327351493116e-7};
  constexpr double kG = 7.0;
  const double zmr = zr - 1.0;
  const double zmi = zi;
  double xRe = kLanczosC[0];
  double xIm = 0.0;
  for (int i = 1; i <= 8; ++i) {
    double quotR = 0.0;
    double quotI = 0.0;
    complexDivide(kLanczosC[i], 0.0, zmr + static_cast<double>(i), zmi, quotR, quotI);
    xRe += quotR;
    xIm += quotI;
  }
  const double tRe = zmr + kG + 0.5;
  const double tIm = zmi;
  double lnRe = 0.0;
  double lnIm = 0.0;
  scalarPrincipalLnCartesian(tRe, tIm, lnRe, lnIm);
  const double pwRe = zmr + 0.5;
  const double pwIm = zmi;
  double lnPowRe = 0.0;
  double lnPowIm = 0.0;
  complexMultiply(pwRe, pwIm, lnRe, lnIm, lnPowRe, lnPowIm);
  double powRe = 0.0;
  double powIm = 0.0;
  if (!complexExpCartesian(lnPowRe, lnPowIm, powRe, powIm)) {
    outR = nanv;
    outI = nanv;
    return;
  }
  double expNegPrR = 0.0;
  double expNegPrI = 0.0;
  if (!complexExpCartesian(-tRe, -tIm, expNegPrR, expNegPrI)) {
    outR = nanv;
    outI = nanv;
    return;
  }
  const double scale = std::sqrt(2.0 * kPi);
  double prodRe = 0.0;
  double prodIm = 0.0;
  complexMultiply(scale * powRe, scale * powIm, expNegPrR, expNegPrI, prodRe, prodIm);
  complexMultiply(prodRe, prodIm, xRe, xIm, outR, outI);
  snapComplexNearZeroAxis(outR, outI);
}
#else
void complexCartesianPrincipalNthRoot(double, double, double, double& outR, double& outI) {
  outR = std::numeric_limits<double>::quiet_NaN();
  outI = std::numeric_limits<double>::quiet_NaN();
}
void scalarPrincipalLnCartesian(double, double, double& outRe, double& outIm) {
  outRe = std::numeric_limits<double>::quiet_NaN();
  outIm = std::numeric_limits<double>::quiet_NaN();
}
void complexPowPrincipal(double, double, double, double, double& outRe, double& outIm) {
  outRe = std::numeric_limits<double>::quiet_NaN();
  outIm = std::numeric_limits<double>::quiet_NaN();
}
void complexDivide(double, double, double, double, double& outR, double& outI) {
  outR = std::numeric_limits<double>::quiet_NaN();
  outI = std::numeric_limits<double>::quiet_NaN();
}
void complexMultiply(double, double, double, double, double& outR, double& outI) {
  outR = std::numeric_limits<double>::quiet_NaN();
  outI = std::numeric_limits<double>::quiet_NaN();
}
bool complexExpCartesian(double, double, double& outR, double& outI) {
  outR = std::numeric_limits<double>::quiet_NaN();
  outI = std::numeric_limits<double>::quiet_NaN();
  return false;
}
void complexPrincipalSqrt(double, double, double& outR, double& outI) {
  outR = std::numeric_limits<double>::quiet_NaN();
  outI = std::numeric_limits<double>::quiet_NaN();
}
void complexGamma(double, double, double& outR, double& outI) {
  outR = std::numeric_limits<double>::quiet_NaN();
  outI = std::numeric_limits<double>::quiet_NaN();
}
#endif

constexpr const char* STR_NAN = "nan";
constexpr const char* STR_NEG_INF = "-inf";
constexpr const char* STR_INF = "inf";
constexpr const char* STR_NEG_ZERO = "-0";
constexpr const char* STR_HEX_DIGITS_LOWER = "0123456789abcdef";
constexpr const char* STR_HEX_DIGITS_UPPER = "0123456789ABCDEF";
constexpr const char* STR_COMMA = ", ";
constexpr const char* STR_PI = "pi";
constexpr const char* STR_E = "e";
constexpr const char* STR_I = "i";
constexpr const char* STR_ANS = "ans";
constexpr const char* STR_FORMAL_VALIDATION_PROBE = "_";
constexpr const char* STR_RAND = "rand";
constexpr const char* STR_RANDOM = "random";
constexpr const char* STR_BIN = "bin";
constexpr const char* STR_HEX = "hex";
constexpr const char* STR_OCT = "oct";
constexpr const char* STR_POW = "pow";
constexpr const char* STR_ATAN2 = "atan2";
constexpr const char* STR_SIN = "sin";
constexpr const char* STR_COS = "cos";
constexpr const char* STR_TAN = "tan";
constexpr const char* STR_ASIN = "asin";
constexpr const char* STR_ARCSIN = "arcsin";
constexpr const char* STR_ACOS = "acos";
constexpr const char* STR_ARCCOS = "arccos";
constexpr const char* STR_ATAN = "atan";
constexpr const char* STR_ARCTAN = "arctan";
constexpr const char* STR_SINH = "sinh";
constexpr const char* STR_COSH = "cosh";
constexpr const char* STR_TANH = "tanh";
constexpr const char* STR_ACOSH = "acosh";
constexpr const char* STR_ASINH = "asinh";
constexpr const char* STR_ATANH = "atanh";
constexpr const char* STR_EXP = "exp";
constexpr const char* STR_LOG = "log";
constexpr const char* STR_LN = "ln";
constexpr const char* STR_LOG10 = "log10";
constexpr const char* STR_SQRT = "sqrt";
constexpr const char* STR_SQR = "sqr";
constexpr const char* STR_INT = "int";
constexpr const char* STR_FRAC = "frac";
constexpr const char* STR_FRACT = "fract";
constexpr const char* STR_ABS = "abs";
constexpr const char* STR_FLOOR = "floor";
constexpr const char* STR_CEIL = "ceil";
constexpr const char* STR_TRUNC = "trunc";
constexpr const char* STR_ROUND = "round";
constexpr const char* STR_SIGN = "sign";
constexpr const char* STR_DEG = "deg";
constexpr const char* STR_RAD = "rad";
constexpr const char* STR_SUM = "sum";
constexpr const char* STR_MEDIAN = "median";
constexpr const char* STR_VARIANCE = "variance";
constexpr const char* STR_STDDEV = "stddev";
constexpr const char* STR_SORT = "sort";
constexpr const char* STR_SORTBY = "sortby";
constexpr const char* STR_RATIO = "ratio";
constexpr const char* STR_RANGE = "range";
constexpr const char* STR_SORTED = "sorted";
constexpr const char* STR_SORTBY_EXPECTS_ONE_FUNCTION = "sortby expects exactly one function";
constexpr const char* STR_SORTBY_EXPECTS_UNARY_FUNCTION = "sortby expects a function that takes 1 parameter";
constexpr const char* STR_SORTBY_KEY_MUST_BE_SCALAR_OR_ARRAY = "sortby key function must return a scalar or an array";
constexpr const char* STR_PAR_ARRAY_COMMA_FUNC = "(array, func)";
constexpr double RATIO_APPROX_EPS = 1e-14;
constexpr long long RATIO_MAX_DENOMINATOR = 10000000LL;
constexpr int RATIO_MAX_POWER10_EXP = 18;
constexpr int RATIO_SEMICONV_LINEAR_THRESH = 64;
constexpr const char* STR_REVERSE = "reverse";
constexpr const char* STR_REVERSED = "reversed";
constexpr const char* STR_UNIQUE = "unique";
constexpr const char* STR_UNPACK = "unpack";
constexpr const char* STR_REPEAT = "repeat";
constexpr const char* STR_LEN = "len";
constexpr const char* STR_LENGTH = "length";
constexpr const char* STR_FACT = "fact";
constexpr const char* STR_FACTORINT = "factorint";
constexpr const char* STR_FACTORIAL = "factorial";
constexpr const char* STR_AVG = "avg";
constexpr const char* STR_MEAN = "mean";
constexpr const char* STR_MOD = "mod";
constexpr const char* STR_CLAMP = "clamp";
constexpr const char* STR_HYPOT = "hypot";
constexpr const char* STR_GCD = "gcd";
constexpr const char* STR_LCM = "lcm";
constexpr const char* STR_NCR = "ncr";
constexpr const char* STR_NPR = "npr";
constexpr const char* STR_PRODUCT = "product";
constexpr const char* STR_PROD = "prod";
constexpr const char* STR_MIN = "min";
constexpr const char* STR_MAX = "max";
constexpr const char* STR_UHEX = "uhex";
constexpr const char* STR_UOCT = "uoct";
constexpr const char* STR_UBIN = "ubin";
constexpr const char* STR_MILLISECONDS = "milliseconds";
constexpr const char* STR_SECONDS = "seconds";
constexpr const char* STR_MINUTES = "minutes";
constexpr const char* STR_HOURS = "hours";
constexpr const char* STR_DAYS = "days";
constexpr const char* STR_REAL = "real";
constexpr const char* STR_IMAG = "imag";
constexpr const char* STR_PHASE = "phase";
constexpr const char* STR_POLAR = "polar";
constexpr const char* STR_CART = "cart";
constexpr const char* STR_CONJ = "conj";
constexpr const char* STR_MILLISECOND = "millisecond";
constexpr const char* STR_SECOND = "second";
constexpr const char* STR_MINUTE = "minute";
constexpr const char* STR_HOUR = "hour";
constexpr const char* STR_DAY = "day";
constexpr const char* STR_NOT = "not";
constexpr const char* STR_AND = "and";
constexpr const char* STR_OR = "or";
constexpr const char* STR_HINT_PREFIX = "function: ";
constexpr const char* STR_PAR_EMPTY = "()";
constexpr const char* STR_PAR_MIN_COMMA_MAX = "(min, max)";
constexpr const char* STR_PAR_VALUE_COMMA_POWER = "(value, power)";
constexpr const char* STR_PAR_Y_COMMA_X = "(y, x)";
constexpr const char* STR_PAR_ANGLE = "(angle)";
constexpr const char* STR_PAR_VALUE = "(value)";
constexpr const char* STR_PAR_VALUE_COMMA_BASE = "(value, base)";
constexpr const char* STR_PAR_N = "(n)";
constexpr const char* STR_PAR_VALUE_COMMA_DIVISOR = "(value, divisor)";
constexpr const char* STR_PAR_VALUE_COMMA_MIN_COMMA_MAX = "(value, min, max)";
constexpr const char* STR_PAR_X_COMMA_Y = "(x, y)";
constexpr const char* STR_PAR_A_COMMA_B = "(a, b)";
constexpr const char* STR_PAR_N_COMMA_R = "(n, r)";
constexpr const char* STR_PAR_X_COMMA_N = "(x, n)";
constexpr const char* STR_PAR_X = "(x)";
constexpr const char* STR_PAR_START_COMMA_STOP_COMMA_STEP = "(start, stop [, step])";
constexpr const char* STR_PAR_STEP_MUST_NOT_BE_ZERO = "() step must not be zero";
constexpr const char* STR_PAR_PRODUCES_NO_VALUES = "() produces no values";
constexpr const char* STR_PAR_PRODUCES_TOO_MANY_VALUES = "() produces too many values";
constexpr int BUILTIN_MAX_PRODUCED_ITEMS = 1000;
constexpr const char* STR_PAR_DOTDOTDOT = "(...)";
constexpr const char* STR_UNEXPECTED_TOKEN = "unexpected token";
constexpr const char* STR_INCOMPATIBLE_OPERANDS = "incompatible operands";
constexpr const char* STR_POS_POW63_DEC_TEXT = "9.223372036854778e+018";
constexpr const char* STR_NEG_POW63_DEC_TEXT = "-9.223372036854778e+018";
constexpr double K_MAX_EXACT_INT_FROM_DOUBLE = smartmath_internal::kMaxExactIntFromDouble;
constexpr const char* STR_PREFIX_HEX = "0x";
constexpr const char* STR_PREFIX_OCT = "0o";
constexpr const char* STR_PREFIX_BIN = "0b";
constexpr const char* STR_EMPTY_STATEMENT = "empty statement";
constexpr const char* STR_UNEXPECTED_COMMA = "unexpected comma";
constexpr const char* STR_MISSING_CLOSING_PARENTHESIS = "missing closing parenthesis";
constexpr const char* STR_INVALID_HEX_LITERAL = "invalid hex literal";
constexpr const char* STR_INVALID_BINARY_LITERAL = "invalid binary literal";
constexpr const char* STR_INVALID_OCTAL_LITERAL = "invalid octal literal";
constexpr const char* STR_INVALID_NUMERIC_LITERAL = "invalid numeric literal";
constexpr const char* STR_MISSING_INDEX = "missing index";
constexpr const char* STR_MISSING_CLOSING_BRACKET = "missing closing bracket";
constexpr const char* STR_MISMATCHED_CLOSING_PARENTHESIS = "mismatched closing parenthesis";
constexpr const char* STR_MISMATCHED_CLOSING_BRACKET = "mismatched closing bracket";
constexpr const char* STR_MISMATCHED_CLOSING_BRACE = "mismatched closing brace";
constexpr const char* STR_LT_EQ = "<=";
constexpr const char* STR_GT_EQ = ">=";
constexpr const char* STR_EQ_EQ = "==";
constexpr const char* STR_NOT_EQ = "!=";
constexpr const char* STR_LT_GT = "<>";
constexpr const char* STR_DUPLICATE_PARAMETER_NAME = "duplicate parameter name";
constexpr const char* STR_RESERVED_CONSTANT_NAME = "reserved constant name";
constexpr const char* STR_RESERVED_FUNCTION_NAME = "reserved function name";
constexpr const char* STR_RESERVED_BUILTIN_VARIABLE_NAME = "reserved built-in variable name";
constexpr const char* STR_RECURSIVE_USER_FUNCTION_CALL_COLON = "recursive function call: ";
constexpr const char* STR_UNEXPECTED_TOKEN_AFTER_EXPRESSION = "unexpected token after expression";
constexpr const char* STR_SCALAR_ONLY_EXPRESSION_ENCOUNTERED_NON = "scalar-only expression encountered non-scalar value";
constexpr const char* STR_BITWISE_OPERANDS_MUST_BE_INTEGER_VALUES = "bitwise operands must be integer values";
constexpr const char* STR_INTERNAL_UNARY_OP = "internal unary op";
constexpr const char* STR_MODULO_OPERANDS_MUST_BE_INTEGER_VALUES = "modulo operands must be integer values";
constexpr const char* STR_NUMERIC_ERROR_IN_POWER_OPERATION = "numeric error in power operation";
constexpr const char* STR_NUMERIC_ERROR_IN_EXPRESSION = "numeric error in expression";
constexpr const char* STR_INTERNAL_BINARY_OP = "internal binary op";
constexpr const char* STR_INTERNAL_EVAL_ERROR = "internal eval error";
constexpr const char* STR_PERCENTAGE_REQUIRES_SCALAR_VALUE = "percentage requires scalar value";
constexpr const char* STR_FAILED_TO_BUILD_ARRAY_LITERAL = "failed to build array literal";
constexpr const char* STR_TIME_LITERAL_EMPTY_SEGMENT = "time literal: empty segment between colons";
constexpr const char* STR_TIME_LITERAL_INVALID_SEGMENT = "time literal: invalid segment";
constexpr const char* STR_TIME_COMPACT_EXPECTED_UNIT = "compact time literal: expected unit suffix";
constexpr const char* STR_TIME_COMPACT_UNIT_ORDER = "compact time literal: unit order or duplicate unit";
constexpr const char* STR_TIME_COMPACT_INVALID_SUFFIX = "compact time literal: invalid suffix";
constexpr const char* STR_TIME_LITERAL_NEGATIVE_SEGMENT = "time literal: negative segment";
constexpr const char* STR_TIME_NON_FINITE = "time value: non-finite operand";
constexpr const char* STR_TIME_ARRAY_MIXED = "time values cannot be mixed with non-time values";
constexpr const char* STR_TIME_EXPECTS_TIME_ARG = "() expects a time value";
constexpr const char* STR_INDEXING_REQUIRES_AN_ARRAY_VALUE = "indexing requires an array value";
constexpr const char* STR_ARRAY_INDEX_MUST_BE_A_SCALAR = "array index must be a scalar integer";
constexpr const char* STR_ARRAY_INDEX_MUST_BE_AN_INTEGER = "array index must be an integer";
constexpr const char* STR_ARRAY_INDEX_IS_OUT_OF_RANGE = "array index is out of range";
constexpr const char* STR_SLICE_STEP_MUST_NOT_BE_ZERO = "slice step must not be zero";
constexpr const char* STR_PAR_EXPECTS_AT_LEAST = "() expects at least ";
constexpr const char* STR_ARGUMENT = " argument";
constexpr const char* STR_ARGUMENTS = " arguments";
constexpr const char* STR_INTERNAL_ERROR_IN_AGGREGATE_BUILTIN = "internal error in aggregate builtin";
constexpr const char* STR_UNEXPECTED_DOUBLE_SIZE = "unexpected double size";
constexpr const char* STR_PAR_EXPECTS_INTEGER_VALUES = "() expects integer values";
constexpr const char* STR_NUMERIC_ERROR_IN = "numeric error in ";
constexpr const char* STR_PAR = "()";
constexpr const char* STR_PAR_EXPECTS_SCALAR_MIN_SLASH = "() expects scalar min/max";
constexpr const char* STR_PAR_EXPECTS_SCALAR_VALUES = "() expects scalar values";
constexpr const char* STR_INTERNAL_ERROR_IN_SCALAR_BINARY_BUILTIN = "internal error in scalar binary builtin";
constexpr const char* STR_PAR_EXPECTS_A_NON_DASH = "() expects a non-negative integer";
constexpr const char* STR_PAR_EXPECTS_A_POSITIVE_INTEGER = "() expects a positive integer";
constexpr const char* STR_INTERNAL_ERROR_IN_UNARY_MATH_BUILTIN = "internal error in unary math builtin";
constexpr const char* STR_PAR_EXPECTS = "() expects ";
constexpr const char* STR_ARGUMENT_PAR_S = " argument(s)";
constexpr const char* STR_MAX_EVALUATION_DEPTH_REACHED = "max evaluation depth reached";
constexpr const char* STR_USER_FUNCTION_CALL_STACK_OVERFLOW = "function call stack overflow";
constexpr const char* STR_FAILED_TO_PARSE_USER_FUNCTION_BODY = "failed to parse user function body";
constexpr const char* STR_FUNCTION_BODY_IS_EMPTY = "function body is empty";
constexpr const char* STR_UNEXPECTED_INPUT = "unexpected characters";
constexpr const char* STR_PARSE_FAILED = "parse failed";
constexpr const char* STR_NOTHING_COMPILED_SEMICOLON_CALL_COMPILE_PAR = "nothing compiled; call compile() first";
constexpr const char* STR_UNKNOWN_VARIABLE_COLON = "unknown variable: ";
constexpr const char* STR_USER_DEFINED_FUNCTION_COLON = "user-defined function: ";
constexpr const char* STR_SEMICOLON_UNKNOWN_FUNCTION_COLON = "; unknown function: ";
constexpr const char* STR_UNKNOWN_FUNCTION_COLON = "unknown function: ";
constexpr const char* STR_INVALID_USER_FUNCTION_EXPRESSION = "invalid user function expression";
constexpr const char* STR_UNEXPECTED_CONTENT_AFTER_FUNCTION_DEFINITION = "unexpected content after function definition";

static constexpr std::uint64_t pow10_u64[20] = {
  1ull,
  10ull,
  100ull,
  1000ull,
  10000ull,
  100000ull,
  1000000ull,
  10000000ull,
  100000000ull,
  1000000000ull,
  10000000000ull,
  100000000000ull,
  1000000000000ull,
  10000000000000ull,
  100000000000000ull,
  1000000000000000ull,
  10000000000000000ull,
  100000000000000000ull,
  1000000000000000000ull,
  10000000000000000000ull
};

inline std::uint64_t quickMult10(std::uint64_t x) {
  // x*10 = x*(8+2) = x*8 + x*2 = (x<<3) + (x<<1)
  return ((x << 3) + (x << 1));
}

bool identAsciiEqualsLower(const std::string& body, std::size_t i0, std::size_t len, const std::string& bLower) {
  if (len != bLower.size()) {
    return false;
  }
  for (std::size_t k = 0; k < bLower.size(); ++k) {
    if (std::tolower(static_cast<unsigned char>(body[i0 + k])) != static_cast<unsigned char>(bLower[k])) {
      return false;
    }
  }
  return true;
}

bool udfBodyCallsDefinedFunction(const std::string& body, const std::string& fnNameLower) {
  if (fnNameLower.empty()) {
    return false;
  }
  const std::size_t n = body.size();
  std::size_t i = 0;
  while (i < n) {
    unsigned char ch = static_cast<unsigned char>(body[i]);
    const bool identStart = std::isalpha(ch) != 0 || ch == '_';
    if (!identStart) {
      ++i;
      continue;
    }
    const std::size_t i0 = i;
    ++i;
    while (i < n) {
      ch = static_cast<unsigned char>(body[i]);
      if (std::isalnum(ch) != 0 || ch == '_') {
        ++i;
      } else {
        break;
      }
    }
    if (identAsciiEqualsLower(body, i0, i - i0, fnNameLower)) {
      std::size_t j = i;
      while (j < n && std::isspace(static_cast<unsigned char>(body[j])) != 0) {
        ++j;
      }
      if (j < n && body[j] == '(') {
        return true;
      }
    }
  }
  return false;
}

double randomUnitScalar() {
  return static_cast<double>(std::rand()) / static_cast<double>(RAND_MAX);
}

int parseDigitForRadix(char c, unsigned int radix) {
  int d = -1;
  if (c >= '0' && c <= '9') d = c - '0';
  else if (c >= 'a' && c <= 'f') d = 10 + (c - 'a');
  else if (c >= 'A' && c <= 'F') d = 10 + (c - 'A');
  if (d < 0 || static_cast<unsigned int>(d) >= radix) {
    return -1;
  }
  return d;
}

bool tryParsePrefixedUIntLiteral(const char* p, char prefixLower, unsigned int radix, const char*& outEnd, std::uint64_t& outValue) {
  if (p[0] != '0') {
    return false;
  }
  const char px = static_cast<char>(std::tolower(static_cast<unsigned char>(p[1])));
  if (px != prefixLower) {
    return false;
  }

  const char* cur = p + 2;
  std::uint64_t u = 0;
  int digits = 0;
  while (*cur) {
    const int d = parseDigitForRadix(*cur, radix);
    if (d < 0) {
      break;
    }
    const std::uint64_t ud = static_cast<std::uint64_t>(d);
    const std::uint64_t base = static_cast<std::uint64_t>(radix);
    if (u > (((std::numeric_limits<std::uint64_t>::max)() - ud) / base)) {
      return false;
    }
    u = (u * base) + ud;
    ++cur;
    ++digits;
  }
  if (digits == 0) {
    return false;
  }
  outEnd = cur;
  outValue = u;
  return true;
}

// Helper: Convert a digit array to uint64_t with overflow checking
inline bool digitsToUint64(const unsigned char* digits, int digitCount, std::uint64_t& outValue) {
  std::uint64_t value = 0;
  for (int i = 0; i < digitCount; ++i) {
    const unsigned char digit = digits[i];
    if (value > ((std::numeric_limits<std::uint64_t>::max)() - digit) / 10u) {
      return false;
    }
    value = value * 10 + digit;
  }
  outValue = value;
  return true;
}

bool tryParseInputNumberAsInteger(const char* numStart, const char* numEnd, std::uint64_t& outValue) {

  if (!numStart || !numEnd || numStart >= numEnd) {
    return false;
  }

  const char* p = numStart;
  if (*p == '+') {
    ++p;
  }
  if (p >= numEnd || *p == '-') {
    return false;
  }

  constexpr int MaxDigits = 128;
  unsigned char digits[MaxDigits];
  int storedDigitCount = 0;
  int intDigitCount = 0;
  int fracDigits = 0;
  bool intPartStarted = false;
  bool hasDigit = false;

  while (p < numEnd && *p >= '0' && *p <= '9') {
    const int digit = *p - '0';
    hasDigit = true;
    if (digit != 0) {
      intPartStarted = true;
    }
    if (intPartStarted) {
      if (storedDigitCount < MaxDigits) {
        digits[storedDigitCount++] = static_cast<unsigned char>(digit);
      } else {
        return false;
      }
      ++intDigitCount;
    }
    ++p;
  }

  if (p < numEnd && *p == '.') {
    ++p;
    intPartStarted = true;
    while (p < numEnd && *p >= '0' && *p <= '9') {
      const int digit = *p - '0';
      hasDigit = true;
      if (storedDigitCount < MaxDigits) {
        digits[storedDigitCount++] = static_cast<unsigned char>(digit);
      } else if (digit != 0) {
        return false;
      }
      ++fracDigits;
      ++p;
    }
  }

  if (!hasDigit) {
    return false;
  }

  if (storedDigitCount == 0) {
    outValue = 0;
    return true;
  }

  int exponent = 0;
  bool exponentNegative = false;
  if (p < numEnd && (*p == 'e' || *p == 'E')) {
    ++p;
    if (p < numEnd && (*p == '+' || *p == '-')) {
      exponentNegative = (*p == '-');
      ++p;
    }
    if (p >= numEnd || *p < '0' || *p > '9') {
      return false;
    }

    int exponentDigits = 0;
    while (p < numEnd && *p >= '0' && *p <= '9') {
      const int digit = *p - '0';
      if (digit != 0 || exponentDigits != 0) {
        if (++exponentDigits > 2) {
          return false; // exponent is too large
        }
        exponent = exponent * 10 + digit;
      }
      ++p;
    }
    if (exponentNegative) {
      exponent = -exponent;
    }
  }

  if (p != numEnd) {
    return false;
  }

  const int significantDigits = intDigitCount + fracDigits;
  const int adjust = exponent - fracDigits;

  if (adjust >= 0) {
    if (significantDigits + adjust > 20) {
      return false;
    }

    std::uint64_t value = 0;
    if (!digitsToUint64(digits, storedDigitCount, value)) {
      return false;
    }

    if (adjust > 0) {
      if (value > (std::numeric_limits<std::uint64_t>::max)() / pow10_u64[adjust]) {
        return false;
      }
      value *= pow10_u64[adjust];
    }

    outValue = value;
    return true;
  }

  const int divisor = -adjust;
  if (divisor > storedDigitCount) {
    return false;
  }
  for (int i = storedDigitCount - divisor; i < storedDigitCount; ++i) {
    if (digits[i] != 0) {
      return false;
    }
  }

  const int resultDigitCount = storedDigitCount - divisor;
  if (resultDigitCount == 0) {
    outValue = 0;
    return true;
  }

  std::uint64_t value = 0;
  if (!digitsToUint64(digits, resultDigitCount, value)) {
    return false;
  }

  outValue = value;
  return true;
}

long long gcdInt64(long long a, long long b) {
  return static_cast<long long>(gcdUInt64(
      static_cast<std::uint64_t>(std::llabs(a)), static_cast<std::uint64_t>(std::llabs(b))));
}

bool tryLcmUInt64(std::uint64_t a, std::uint64_t b, std::uint64_t& out) {
  if (a == 0u || b == 0u) {
    out = 0u;
    return true;
  }
  const std::uint64_t g = gcdUInt64(a, b);
  const std::uint64_t a1 = a / g;
  if (a1 > ((std::numeric_limits<std::uint64_t>::max)() / b)) {
    return false;
  }
  out = a1 * b;
  return true;
}

bool tryAddUInt64Checked(std::uint64_t a, std::uint64_t b, std::uint64_t& out) {
  const std::uint64_t s = a + b;
  if (s < a) {
    return false;
  }
  out = s;
  return true;
}

bool tryAddInt64Checked(long long a, long long b, long long& out) {
  out = a + b;
  if (b > 0 && out < a) {
    return false;
  }
  if (b < 0 && out > a) {
    return false;
  }
  return true;
}

bool tryMulInt64Checked(long long a, long long b, long long& out) {
  if (a == 0 || b == 0) {
    out = 0;
    return true;
  }
  if (a == -1) {
    if (b == (std::numeric_limits<long long>::min)()) return false;
    out = -b;
    return true;
  }
  if (b == -1) {
    if (a == (std::numeric_limits<long long>::min)()) return false;
    out = -a;
    return true;
  }
  if (a > 0) {
    if (b > 0) {
      if (a > (std::numeric_limits<long long>::max)() / b) return false;
    } else {
      if (b < (std::numeric_limits<long long>::min)() / a) return false;
    }
  } else {
    if (b > 0) {
      if (a < (std::numeric_limits<long long>::min)() / b) return false;
    } else {
      if (a != 0 && b < (std::numeric_limits<long long>::max)() / a) return false;
    }
  }
  out = a * b;
  return true;
}

bool tryPowInt64Checked(long long base, long long exp, long long& out) {
  if (exp < 0) {
    return false;
  }
  long long r = 1;
  long long b = base;
  long long e = exp;
  while (e > 0) {
    if ((e & 1) != 0 && !tryMulInt64Checked(r, b, r)) {
      return false;
    }
    e >>= 1;
    if (e > 0 && !tryMulInt64Checked(b, b, b)) {
      return false;
    }
  }
  out = r;
  return true;
}

bool tryComputeNprInt64(long long n, long long r, long long& out) {
  if (n < 0 || r < 0 || r > n) return false;
  long long acc = 1;
  for (long long i = 0; i < r; ++i) {
    long long term = n - i;
    if (!tryMulInt64Checked(acc, term, acc)) return false;
  }
  out = acc;
  return true;
}

bool tryComputeNcrInt64(long long n, long long r, long long& out) {
  if (n < 0 || r < 0 || r > n) return false;
  if (r == 0) {
    out = 1;
    return true;
  }
  long long rEff = r;
  if ((n - rEff) < rEff) rEff = (n - rEff);
  long long acc = 1;
  for (long long i = 1; i <= rEff; ++i) {
    long long num = (n - rEff) + i;
    long long den = i;
    long long g1 = gcdInt64(num, den);
    if (g1 > 1) {
      num /= g1;
      den /= g1;
    }
    long long g2 = gcdInt64(acc, den);
    if (g2 > 1) {
      acc /= g2;
      den /= g2;
    }
    if (den != 1) return false;
    if (!tryMulInt64Checked(acc, num, acc)) return false;
  }
  out = acc;
  return true;
}

constexpr long long kFactorialTable21[21] = {
    1LL,
    1LL,
    2LL,
    6LL,
    24LL,
    120LL,
    720LL,
    5040LL,
    40320LL,
    362880LL,
    3628800LL,
    39916800LL,
    479001600LL,
    6227020800LL,
    87178291200LL,
    1307674368000LL,
    20922789888000LL,
    355687428096000LL,
    6402373705728000LL,
    121645100408832000LL,
    2432902008176640000LL};

double factorialScalarFloatFromInt(long long n) {
  double d = static_cast<double>(kFactorialTable21[20]);
  for (long long i = 21; i <= n; ++i) {
    d *= static_cast<double>(i);
    if (!std::isfinite(d)) {
      break;
    }
  }
  return d;
}

std::string formatDoubleFast(double v) {
  // Custom "general" formatter: 16 significant digits in fixed mode, 16 in scientific (smoke parity).
  // Uses only direct character-buffer operations.
  if (std::isnan(v)) return STR_NAN;
  if (std::isinf(v)) return (v < 0.0) ? STR_NEG_INF : STR_INF;
  if (v == 0.0) return "0";
  if (std::fabs(v) > 0.0 && std::fabs(v) < 1e-10) {
    char buf[64];
    std::snprintf(buf, sizeof(buf), "%.15e", v);
    return buf;
  }

  constexpr int kGenDigits = 17;
  constexpr int kSciOutDigits = 16;
  constexpr int kFixedOutDigits = 16;
  long double x = static_cast<long double>(v);
  bool neg = (x < 0.0L);
  if (neg) x = -x;

  int exp10 = static_cast<int>(std::floor(std::log10(x)));
  long double p10 = std::pow(10.0L, exp10);
  if (p10 == 0.0L) {
    return neg ? STR_NEG_ZERO : "0";
  }

  long double norm = x / p10;
  if (norm >= 10.0L) {
    norm /= 10.0L;
    ++exp10;
  } else if (norm < 1.0L) {
    norm *= 10.0L;
    --exp10;
  }

  constexpr int kExtraDigits = 4;
  int digits[kGenDigits + kExtraDigits] = {0};
  for (int i = 0; i < kGenDigits + kExtraDigits; ++i) {
    int d = static_cast<int>(norm);
    if (d < 0) d = 0;
    if (d > 9) d = 9;
    digits[i] = d;
    norm = (norm - static_cast<long double>(d)) * 10.0L;
  }

  // Round last kept digit using the extra digit.
  bool roundUp = false;
  if (digits[kGenDigits] > 5) {
    roundUp = true;
  } else if (digits[kGenDigits] == 5) {
    for (int i = kGenDigits + 1; i < kGenDigits + kExtraDigits; ++i) {
      if (digits[i] != 0) {
        roundUp = true;
        break;
      }
    }
    if (!roundUp && (digits[kGenDigits - 1] % 2 != 0)) {
      roundUp = true;
    }
  }
  if (roundUp) {
    int i = kGenDigits - 1;
    for (; i >= 0; --i) {
      if (digits[i] < 9) {
        ++digits[i];
        break;
      }
      digits[i] = 0;
    }
    if (i < 0) {
      // 9.999... rounded -> 1.000... and bump exponent.
      digits[0] = 1;
      for (int j = 1; j < kGenDigits; ++j) {
        digits[j] = 0;
      }
      ++exp10;
    }
  }

  // Match "general": fixed for [-4, 15), scientific otherwise.
  const bool useScientific = (exp10 < -4 || exp10 >= kSciOutDigits);
  int sciDigits = kGenDigits;
  if (useScientific) {
    bool sciRound = false;
    if (digits[kSciOutDigits] > 5) {
      sciRound = true;
    } else if (digits[kSciOutDigits] == 5) {
      // There are no fractional digits beyond digits[kSciOutDigits] in the
      // scientific output window (kGenDigits == kSciOutDigits + 1).
      if (!sciRound && (digits[kSciOutDigits - 1] % 2 != 0)) {
        sciRound = true;
      }
    }
    if (sciRound) {
      int i = kSciOutDigits - 1;
      for (; i >= 0; --i) {
        if (digits[i] < 9) {
          ++digits[i];
          break;
        }
        digits[i] = 0;
      }
      if (i < 0) {
        digits[0] = 1;
        for (int j = 1; j < kSciOutDigits; ++j) {
          digits[j] = 0;
        }
        ++exp10;
      }
    }
    sciDigits = kSciOutDigits;
  }
  char out[128] = {0};
  int pos = 0;

  if (neg) out[pos++] = '-';

  if (useScientific) {
    out[pos++] = static_cast<char>('0' + digits[0]);
    // Fractional part
    int fracEnd = sciDigits - 1;
    while (fracEnd > 0 && digits[fracEnd] == 0) {
      --fracEnd;
    }
    if (fracEnd > 0) {
      out[pos++] = '.';
      for (int i = 1; i <= fracEnd; ++i) {
        out[pos++] = static_cast<char>('0' + digits[i]);
      }
    }
    out[pos++] = 'e';
    out[pos++] = (exp10 >= 0) ? '+' : '-';
    int e = (exp10 >= 0) ? exp10 : -exp10;
    char eb[16];
    int ep = 0;
    do {
      eb[ep++] = static_cast<char>('0' + (e % 10));
      e /= 10;
    } while (e > 0);
    while (ep < 3) {
      eb[ep++] = '0';
    }
    for (int i = ep - 1; i >= 0; --i) {
      out[pos++] = eb[i];
    }
  } else {
    // Fixed notation: place decimal point after exp10+1 digits.
    int intDigits = exp10 + 1;
    if (intDigits <= 0) {
      out[pos++] = '0';
      out[pos++] = '.';
      for (int i = 0; i < -intDigits; ++i) {
        out[pos++] = '0';
      }
      for (int i = 0; i < kFixedOutDigits; ++i) {
        out[pos++] = static_cast<char>('0' + digits[i]);
      }
    } else {
      int i = 0;
      for (; i < intDigits; ++i) {
        out[pos++] = (i < kGenDigits) ? static_cast<char>('0' + digits[i]) : '0';
      }
      if (intDigits < kFixedOutDigits) {
        out[pos++] = '.';
        for (; i < kFixedOutDigits; ++i) {
          out[pos++] = static_cast<char>('0' + digits[i]);
        }
      }
    }
    // Trim trailing zeros/dot in fixed mode.
    while (pos > 0 && out[pos - 1] == '0') {
      --pos;
    }
    if (pos > 0 && out[pos - 1] == '.') {
      --pos;
    }
  }

  if (pos <= 0) return "0";
  return std::string(out, static_cast<std::size_t>(pos));
}

std::string formatUnsignedBase(std::uint64_t u, unsigned base, const char* prefix, bool uppercase) {
  static constexpr const char* kDigitsLo = STR_HEX_DIGITS_LOWER;
  static constexpr const char* kDigitsHi = STR_HEX_DIGITS_UPPER;
  const char* digits = uppercase ? kDigitsHi : kDigitsLo;

  char tmp[65];
  int pos = 0;
  if (u == 0) {
    tmp[pos++] = '0';
  } else {
    while (u > 0) {
      tmp[pos++] = digits[u % base];
      u /= base;
    }
  }

  std::string out(prefix);
  out.reserve(out.size() + static_cast<std::size_t>(pos));
  const std::size_t digitsCount = static_cast<std::size_t>(pos);
  for (std::size_t i = 0; i < digitsCount; ++i) {
    out.push_back(tmp[digitsCount - 1 - i]);
  }
  return out;
}

unsigned renderBaseRadix(int baseCode) {
  if (baseCode == 16) return 16U;
  if (baseCode == 8) return 8U;
  return 2U;
}

const char* renderBasePrefix(int baseCode) {
  if (baseCode == 16) return STR_PREFIX_HEX;
  if (baseCode == 8) return STR_PREFIX_OCT;
  return STR_PREFIX_BIN;
}

bool renderBaseUppercase(int baseCode) {
  return baseCode == 16;
}

std::string formatUnsignedForRenderBase(std::uint64_t u, int baseCode) {
  return formatUnsignedBase(u, renderBaseRadix(baseCode), renderBasePrefix(baseCode), renderBaseUppercase(baseCode));
}

std::string formatSignedMagnitudeForRenderBase(
    long long iv,
    int baseCode,
    bool asUnsigned) {
  if (asUnsigned) {
    return formatUnsignedForRenderBase(static_cast<std::uint64_t>(iv), baseCode);
  }
  std::uint64_t mag = 0;
  if (iv < 0) {
    if (iv == (std::numeric_limits<long long>::min)()) {
      mag = (1ull << 63);
    } else {
      mag = static_cast<std::uint64_t>(-iv);
    }
    return "-" + formatUnsignedForRenderBase(mag, baseCode);
  }
  return formatUnsignedForRenderBase(static_cast<std::uint64_t>(iv), baseCode);
}

std::string stripLineComments(const std::string& s) {
  for (std::size_t i = 0; i < s.size(); ++i) {
    if (s[i] == '#') {
      return s.substr(0, i);
    }
    if (s[i] == '/' && (i + 1) < s.size() && s[i + 1] == '/') {
      return s.substr(0, i);
    }
  }
  return s;
}

bool parseSourceHasLineComment(const char* text, std::size_t len) {
  for (std::size_t i = 0; i < len; ++i) {
    if (text[i] == '#') {
      return true;
    }
    if (text[i] == '/' && (i + 1) < len && text[i + 1] == '/') {
      return true;
    }
  }
  return false;
}

bool parseSourceNeedsTrailingSemicolonStrip(const char* text, std::size_t len) {
  if (len == 0) {
    return false;
  }
  std::size_t end = len;
  while (end > 0 && std::isspace(static_cast<unsigned char>(text[end - 1]))) {
    --end;
  }
  return end > 0 && text[end - 1] == ';';
}

bool identHasDigitOrUnderscore(const char* begin, const char* end) {
  for (const char* q = begin; q < end; ++q) {
    const unsigned char c = static_cast<unsigned char>(*q);
    if (c == '_' || (c >= '0' && c <= '9')) {
      return true;
    }
  }
  return false;
}

void assignLowerIdentFromRange(std::string& out, const char* begin, const char* end) {
  out.clear();
  out.reserve(static_cast<std::size_t>(end - begin));
  for (const char* q = begin; q < end; ++q) {
    const unsigned char c = static_cast<unsigned char>(*q);
    if (c >= 'A' && c <= 'Z') {
      out.push_back(static_cast<char>(std::tolower(c)));
    } else {
      out.push_back(*q);
    }
  }
}

bool checkedAddLL(long long a, long long b, long long& out) {
  out = a + b;
  if (b > 0 && out < a) {
    return false;
  }
  if (b < 0 && out > a) {
    return false;
  }
  return true;
}

bool checkedSubLL(long long a, long long b, long long& out) {
  out = a - b;
  if (b > 0 && out > a) {
    return false;
  }
  if (b < 0 && out < a) {
    return false;
  }
  return true;
}

bool checkedMulLL(long long a, long long b, long long& out) {
  if (a == 0 || b == 0) {
    out = 0;
    return true;
  }
  if (a > 0) {
    if (b > 0) {
      if (a > (std::numeric_limits<long long>::max)() / b) return false;
    } else {
      if (b < (std::numeric_limits<long long>::min)() / a) return false;
    }
  } else {
    if (b > 0) {
      if (a < (std::numeric_limits<long long>::min)() / b) return false;
    } else {
      if (a != 0 && b < (std::numeric_limits<long long>::max)() / a) return false;
    }
  }
  out = a * b;
  return true;
}

long long bitwiseShiftLeftDefined(long long value, unsigned int shiftCount) {
  const std::uint64_t u = static_cast<std::uint64_t>(value);
  const std::uint64_t shifted = u << shiftCount;
  if (shifted <= static_cast<std::uint64_t>((std::numeric_limits<long long>::max)())) {
    return static_cast<long long>(shifted);
  }
  const std::uint64_t magnitude = (~shifted) + 1u;
  if (magnitude == (1ull << 63)) {
    return (std::numeric_limits<long long>::min)();
  }
  return -static_cast<long long>(magnitude);
}

long long bitwiseShiftRightDefined(long long value, unsigned int shiftCount) {
  const std::uint64_t u = static_cast<std::uint64_t>(value);
  std::uint64_t shifted = 0;
  if (value >= 0) {
    shifted = u >> shiftCount;
  } else {
    shifted = ~((~u) >> shiftCount);
  }
  if (shifted <= static_cast<std::uint64_t>((std::numeric_limits<long long>::max)())) {
    return static_cast<long long>(shifted);
  }
  const std::uint64_t magnitude = (~shifted) + 1u;
  if (magnitude == (1ull << 63)) {
    return (std::numeric_limits<long long>::min)();
  }
  return -static_cast<long long>(magnitude);
}

bool parseUInt64FromDecimalLiteral(const char* begin, const char* end, std::uint64_t& out) {
  if (!begin || !end || begin >= end) {
    return false;
  }
  std::uint64_t v = 0;
  for (const char* p = begin; p < end; ++p) {
    const char c = *p;
    if (c < '0' || c > '9') {
      return false;
    }
    const std::uint64_t digit = static_cast<std::uint64_t>(c - '0');
    if (v > ((std::numeric_limits<std::uint64_t>::max)() - digit) / 10u) {
      return false;
    }
    v = quickMult10(v) + digit;
  }
  out = v;
  return true;
}

void appendUniqueName(std::string& listText, const std::string& n) {
  const std::string token = std::string(",") + listText + ",";
  const std::string needle = std::string(",") + n + ",";
  if (token.find(needle) != std::string::npos) {
    return;
  }
  if (listText.empty()) {
    listText = n;
  } else {
    listText += STR_COMMA;
    listText += n;
  }
}

void mergeUnknownNameList(std::string& dest, const std::string& src) {
  if (src.empty()) {
    return;
  }
  std::size_t i = 0;
  while (i < src.size()) {
    while (i < src.size() && (src[i] == ' ' || src[i] == ',')) {
      ++i;
    }
    if (i >= src.size()) {
      break;
    }
    std::size_t j = i;
    while (j < src.size() && src[j] != ',') {
      ++j;
    }
    std::string token = src.substr(i, j - i);
    while (!token.empty() && std::isspace(static_cast<unsigned char>(token.front()))) {
      token.erase(token.begin());
    }
    while (!token.empty() && std::isspace(static_cast<unsigned char>(token.back()))) {
      token.pop_back();
    }
    if (!token.empty()) {
      appendUniqueName(dest, token);
    }
    i = j;
  }
}

std::string pad2DigitsLl(long long n) {
  if (n < 0) {
    n = 0;
  }
  if (n >= 100) {
    return std::to_string(n);
  }
  if (n >= 10) {
    return std::string(1, static_cast<char>('0' + (n / 10))) + std::string(1, static_cast<char>('0' + (n % 10)));
  }
  return std::string("0") + std::string(1, static_cast<char>('0' + (n % 10)));
}

#if SMARTMATH_TIME_VALUES
long long roundHalfUpDoubleToLongLong(double x) {
  if (!std::isfinite(x)) {
    return 0;
  }
  if (x >= 0.0) {
    return static_cast<long long>(std::floor(x + 0.5));
  }
  return -static_cast<long long>(std::floor(-x + 0.5));
}

long long secondsFieldToMsRounded(long long wholeSec, const std::string& fracDigits) {
  double d = static_cast<double>(wholeSec);
  if (!fracDigits.empty()) {
    const std::string s = std::string("0.") + fracDigits;
    char* end = nullptr;
    const double fd = std::strtod(s.c_str(), &end);
    if (end == s.c_str() || !std::isfinite(fd)) {
      return 0;
    }
    d += fd;
  }
  return roundHalfUpDoubleToLongLong(d * 1000.0);
}

bool tryAddTimeMsChecked(long long a, long long b, long long& outMs) {
  const double t = static_cast<double>(a) + static_cast<double>(b);
  const double kMin = static_cast<double>((std::numeric_limits<long long>::min)());
  const double kMax = static_cast<double>((std::numeric_limits<long long>::max)());
  if (t < kMin || t > kMax) {
    return false;
  }
  outMs = roundHalfUpDoubleToLongLong(t);
  return true;
}

bool trySubTimeMsChecked(long long a, long long b, long long& outMs) {
  const double t = static_cast<double>(a) - static_cast<double>(b);
  const double kMin = static_cast<double>((std::numeric_limits<long long>::min)());
  const double kMax = static_cast<double>((std::numeric_limits<long long>::max)());
  if (t < kMin || t > kMax) {
    return false;
  }
  outMs = roundHalfUpDoubleToLongLong(t);
  return true;
}

bool tryProductCoeffUnitMs(long long coeff, long long factor, long long& outDelta) {
  if (coeff == 0 || factor == 0) {
    outDelta = 0;
    return true;
  }
  if (coeff > 0 && factor > 0) {
    if (coeff > (std::numeric_limits<long long>::max)() / factor) {
      return false;
    }
  } else {
    return false;
  }
  outDelta = coeff * factor;
  return true;
}

/** Parse `[litStart,litEnd)` as MM:SS, HH:MM:SS, or DD:HH:MM:SS (optional fractional last field). */
bool parseTimeLiteralStringToMs(const char* litStart, const char* litEnd, long long& outMs, const char*& err) {
  err = nullptr;
  if (!litStart || !litEnd || litStart >= litEnd) {
    err = STR_TIME_LITERAL_INVALID_SEGMENT;
    return false;
  }
  const int n = static_cast<int>(litEnd - litStart);
  if (n <= 0) {
    err = STR_TIME_LITERAL_INVALID_SEGMENT;
    return false;
  }
  std::vector<int> colonPos;
  colonPos.reserve(8);
  for (int i = 0; i < n; ++i) {
    if (litStart[i] == ':') {
      if (colonPos.size() > 6U) {
        err = STR_TIME_LITERAL_INVALID_SEGMENT;
        return false;
      }
      colonPos.push_back(i);
    }
  }
  const int segCount = static_cast<int>(colonPos.size()) + 1;
  if (segCount < 2 || segCount > 4) {
    err = STR_TIME_LITERAL_INVALID_SEGMENT;
    return false;
  }

  long long d = 0, h = 0, m = 0;
  long long lastWhole = 0;
  std::string fracPart;
  int getSegStart = 0;
  for (int si = 0; si < segCount; ++si) {
    const int segEnd =
        (si < static_cast<int>(colonPos.size())) ? (colonPos[static_cast<std::size_t>(si)] - 1) : (n - 1);
    if (segEnd < getSegStart) {
      err = STR_TIME_LITERAL_EMPTY_SEGMENT;
      return false;
    }
    const char* segPtr = litStart + getSegStart;
    const char* segEndPtr = litStart + segEnd + 1;
    if (segPtr >= segEndPtr) {
      err = STR_TIME_LITERAL_EMPTY_SEGMENT;
      return false;
    }
    const char* wholeEndPtr = segEndPtr;
    fracPart.clear();
    if (si == segCount - 1) {
      for (const char* t = segPtr; t < segEndPtr; ++t) {
        if (*t == '.') {
          wholeEndPtr = t;
          fracPart.assign(t + 1, segEndPtr);
          break;
        }
      }
    }
    if (wholeEndPtr == segPtr) {
      err = STR_TIME_LITERAL_EMPTY_SEGMENT;
      return false;
    }
    for (const char* t = segPtr; t < wholeEndPtr; ++t) {
      if (*t < '0' || *t > '9') {
        err = STR_TIME_LITERAL_INVALID_SEGMENT;
        return false;
      }
    }
    for (char ch : fracPart) {
      if (ch < '0' || ch > '9') {
        err = STR_TIME_LITERAL_INVALID_SEGMENT;
        return false;
      }
    }
    std::uint64_t uv = 0;
    for (const char* t = segPtr; t < wholeEndPtr; ++t) {
      const int dig = *t - '0';
      if (uv > ((std::numeric_limits<std::uint64_t>::max)() - static_cast<std::uint64_t>(dig)) / 10u) {
        err = STR_TIME_LITERAL_INVALID_SEGMENT;
        return false;
      }
      uv = uv * 10u + static_cast<std::uint64_t>(dig);
    }
    if (uv > static_cast<std::uint64_t>((std::numeric_limits<long long>::max)())) {
      err = STR_TIME_LITERAL_INVALID_SEGMENT;
      return false;
    }
    lastWhole = static_cast<long long>(uv);
    if (si < segCount - 1) {
      switch (segCount) {
        case 2:
          if (si == 0) m = lastWhole;
          break;
        case 3:
          if (si == 0) h = lastWhole;
          if (si == 1) m = lastWhole;
          break;
        case 4:
          if (si == 0) d = lastWhole;
          if (si == 1) h = lastWhole;
          if (si == 2) m = lastWhole;
          break;
        default: break;
      }
    }
    getSegStart = segEnd + 2;
  }

  const long long secMs = secondsFieldToMsRounded(lastWhole, fracPart);
  double t = 0.0;
  switch (segCount) {
    case 2:
      t = static_cast<double>(m) * 60000.0 + static_cast<double>(secMs);
      break;
    case 3:
      t = (static_cast<double>(h) * 3600.0 + static_cast<double>(m) * 60.0) * 1000.0 + static_cast<double>(secMs);
      break;
    case 4:
      t = ((static_cast<double>(d) * 24.0 + static_cast<double>(h)) * 3600.0 + static_cast<double>(m) * 60.0) * 1000.0 +
          static_cast<double>(secMs);
      break;
    default:
      err = STR_TIME_LITERAL_INVALID_SEGMENT;
      return false;
  }
  if (!std::isfinite(t)) {
    err = STR_TIME_LITERAL_INVALID_SEGMENT;
    return false;
  }
  const double kMin = static_cast<double>((std::numeric_limits<long long>::min)());
  const double kMax = static_cast<double>((std::numeric_limits<long long>::max)());
  if (t < kMin || t > kMax) {
    err = STR_TIME_LITERAL_INVALID_SEGMENT;
    return false;
  }
  outMs = roundHalfUpDoubleToLongLong(t);
  return true;
}

bool parseTimeLiteralStringToMs(const std::string& lit, long long& outMs, const char*& err) {
  return parseTimeLiteralStringToMs(lit.data(), lit.data() + lit.size(), outMs, err);
}

std::string formatTimeCanonicalFromMs(long long totalMs) {
  const bool neg = (totalMs < 0);
  std::uint64_t rU = 0;
  if (neg) {
    if (totalMs == (std::numeric_limits<long long>::min)()) {
      rU = static_cast<std::uint64_t>(1ull << 63);
    } else {
      rU = static_cast<std::uint64_t>(-totalMs);
    }
  } else {
    rU = static_cast<std::uint64_t>(totalMs);
  }
  long long msPart = static_cast<long long>(rU % 1000ull);
  rU /= 1000ull;
  long long sPart = static_cast<long long>(rU % 60ull);
  rU /= 60ull;
  long long mPart = static_cast<long long>(rU % 60ull);
  rU /= 60ull;
  long long hPart = static_cast<long long>(rU % 24ull);
  rU /= 24ull;
  const std::uint64_t dPart = rU;
  std::string body;
  if (dPart > 0ull) {
    body = std::to_string(dPart) + ":" + pad2DigitsLl(hPart) + ":" + pad2DigitsLl(mPart) + ":" + pad2DigitsLl(sPart);
  } else if (hPart > 0) {
    body = pad2DigitsLl(hPart) + ":" + pad2DigitsLl(mPart) + ":" + pad2DigitsLl(sPart);
  } else {
    body = pad2DigitsLl(mPart) + ":" + pad2DigitsLl(sPart);
  }
  if (msPart != 0) {
    body.push_back('.');
    body.push_back(static_cast<char>('0' + (msPart / 100)));
    body.push_back(static_cast<char>('0' + ((msPart / 10) % 10)));
    body.push_back(static_cast<char>('0' + (msPart % 10)));
  }
  if (neg) {
    return std::string("-") + body;
  }
  return body;
}

std::string constructTimeArrayMixedError(const std::string& name, bool isFn = true) {
  std::string err(name);
  if (isFn) {
    err += STR_PAR;
  }
  err += ": ";
  err += STR_TIME_ARRAY_MIXED;
  return err;
}
#endif

}  // namespace

#if SMARTMATH_COMPLEX_NUMBERS
bool MathParser::isComplexUnaryTrigBuiltin(BuiltinFunctionId id) {
  switch (id) {
    case BuiltinFunctionId::Sin:
    case BuiltinFunctionId::Cos:
    case BuiltinFunctionId::Tan:
    case BuiltinFunctionId::Asin:
    case BuiltinFunctionId::Acos:
    case BuiltinFunctionId::Atan:
    case BuiltinFunctionId::Sinh:
    case BuiltinFunctionId::Cosh:
    case BuiltinFunctionId::Tanh:
    case BuiltinFunctionId::Acosh:
    case BuiltinFunctionId::Asinh:
    case BuiltinFunctionId::Atanh:
      return true;
    default:
      return false;
  }
}

bool MathParser::complexUnaryTrigCartesian(BuiltinFunctionId id, double ar, double ai, double& outR, double& outI) {
  const double nanv = std::numeric_limits<double>::quiet_NaN();
  if (id == BuiltinFunctionId::Sin || id == BuiltinFunctionId::Cos || id == BuiltinFunctionId::Tan ||
      id == BuiltinFunctionId::Sinh || id == BuiltinFunctionId::Cosh || id == BuiltinFunctionId::Tanh) {
    if (!isTrigRadiansInRange(ar) || !isTrigRadiansInRange(ai)) {
      return false;
    }
  } else if (!std::isfinite(ar) || !std::isfinite(ai)) {
    outR = nanv;
    outI = nanv;
    return true;
  }
  switch (id) {
    case BuiltinFunctionId::Sin:
      outR = calcSin(ar) * std::cosh(ai);
      outI = calcCos(ar) * std::sinh(ai);
      break;
    case BuiltinFunctionId::Cos:
      outR = calcCos(ar) * std::cosh(ai);
      outI = -calcSin(ar) * std::sinh(ai);
      break;
    case BuiltinFunctionId::Sinh:
      outR = std::sinh(ar) * calcCos(ai);
      outI = std::cosh(ar) * calcSin(ai);
      break;
    case BuiltinFunctionId::Cosh:
      outR = std::cosh(ar) * calcCos(ai);
      outI = std::sinh(ar) * calcSin(ai);
      break;
    case BuiltinFunctionId::Tan:
    case BuiltinFunctionId::Tanh: {
      double numR = 0.0;
      double numI = 0.0;
      double denR = 0.0;
      double denI = 0.0;
      if (id == BuiltinFunctionId::Tan) {
        numR = calcSin(ar) * std::cosh(ai);
        numI = calcCos(ar) * std::sinh(ai);
        denR = calcCos(ar) * std::cosh(ai);
        denI = -calcSin(ar) * std::sinh(ai);
      } else {
        numR = std::sinh(ar) * calcCos(ai);
        numI = std::cosh(ar) * calcSin(ai);
        denR = std::cosh(ar) * calcCos(ai);
        denI = std::sinh(ar) * calcSin(ai);
      }
      complexDivide(numR, numI, denR, denI, outR, outI);
      break;
    }
    case BuiltinFunctionId::Asinh: {
      const double z2r = ar * ar - ai * ai;
      const double z2i = 2.0 * ar * ai;
      double sqrR = 0.0;
      double sqrI = 0.0;
      complexPrincipalSqrt(z2r + 1.0, z2i, sqrR, sqrI);
      scalarPrincipalLnCartesian(ar + sqrR, ai + sqrI, outR, outI);
      break;
    }
    case BuiltinFunctionId::Acosh: {
      const double z2r = ar * ar - ai * ai;
      const double z2i = 2.0 * ar * ai;
      double sqrR = 0.0;
      double sqrI = 0.0;
      complexPrincipalSqrt(z2r - 1.0, z2i, sqrR, sqrI);
      scalarPrincipalLnCartesian(ar + sqrR, ai + sqrI, outR, outI);
      break;
    }
    case BuiltinFunctionId::Atanh: {
      double quotR = 0.0;
      double quotI = 0.0;
      complexDivide(1.0 + ar, ai, 1.0 - ar, -ai, quotR, quotI);
      double lnR = 0.0;
      double lnI = 0.0;
      scalarPrincipalLnCartesian(quotR, quotI, lnR, lnI);
      outR = lnR * 0.5;
      outI = lnI * 0.5;
      break;
    }
    case BuiltinFunctionId::Asin: {
      const double izR = -ai;
      const double izI = ar;
      const double z2r = ar * ar - ai * ai;
      const double z2i = 2.0 * ar * ai;
      double sqrR = 0.0;
      double sqrI = 0.0;
      complexPrincipalSqrt(1.0 - z2r, -z2i, sqrR, sqrI);
      double lnR = 0.0;
      double lnI = 0.0;
      scalarPrincipalLnCartesian(izR + sqrR, izI + sqrI, lnR, lnI);
      outR = lnI;
      outI = -lnR;
      break;
    }
    case BuiltinFunctionId::Acos: {
      double asR = 0.0;
      double asI = 0.0;
      if (!complexUnaryTrigCartesian(BuiltinFunctionId::Asin, ar, ai, asR, asI)) {
        return false;
      }
      outR = kPi / 2.0 - asR;
      outI = -asI;
      break;
    }
    case BuiltinFunctionId::Atan: {
      const double izR = -ai;
      const double izI = ar;
      double ln1mR = 0.0;
      double ln1mI = 0.0;
      double ln1pR = 0.0;
      double ln1pI = 0.0;
      scalarPrincipalLnCartesian(1.0 - izR, -izI, ln1mR, ln1mI);
      scalarPrincipalLnCartesian(1.0 + izR, izI, ln1pR, ln1pI);
      const double dR = ln1mR - ln1pR;
      const double dI = ln1mI - ln1pI;
      outR = -dI * 0.5;
      outI = dR * 0.5;
      break;
    }
    default:
      return false;
  }
  snapComplexNearZeroAxis(outR, outI);
  return true;
}
#endif

MathParser::MathParser() {
  assert(functionNames().size() == static_cast<std::size_t>(BuiltinFunctionId::Count));
  assert(operatorNames().size() == static_cast<std::size_t>(OperatorNameId::Count));
  addConst(STR_PI, kPi);
  addConst(STR_E, std::exp(1.0));
  addConst(STR_INF, std::numeric_limits<double>::infinity());
  addConst(STR_NAN, std::numeric_limits<double>::quiet_NaN());
#if SMARTMATH_TIME_VALUES
  setVariable(STR_MILLISECOND, makeScalarTimeMs(1LL));
  setVariable(STR_SECOND, makeScalarTimeMs(1000LL));
  setVariable(STR_MINUTE, makeScalarTimeMs(60000LL));
  setVariable(STR_HOUR, makeScalarTimeMs(3600000LL));
  setVariable(STR_DAY, makeScalarTimeMs(86400000LL));
#endif
  setVariable(STR_ANS, makeScalarInt(0));
  setVariable(STR_FORMAL_VALIDATION_PROBE, makeScalarInt(1));
}

std::string MathParser::toLower(std::string s) {
  std::transform(s.begin(), s.end(), s.begin(), [](unsigned char c) { return static_cast<char>(std::tolower(c)); });
  return s;
}

const std::vector<std::string>& MathParser::functionNames() {
  static const std::vector<std::string> kNames = {
      STR_RAND,   STR_RANDOM, STR_BIN,   STR_HEX,    STR_OCT,      STR_POW,  STR_ATAN2, STR_SIN,      STR_COS,  STR_TAN,
      STR_ASIN,   STR_ACOS, STR_ATAN, STR_SINH, STR_COSH,     STR_TANH, STR_ACOSH, STR_ASINH, STR_ATANH, STR_EXP,
      STR_LOG,    STR_LN,     STR_LOG10, STR_SQRT,   STR_SQR,      STR_INT,  STR_FRAC,    STR_ABS,  STR_FLOOR,
      STR_CEIL,   STR_TRUNC,  STR_ROUND, STR_SIGN,   STR_DEG,      STR_RAD,  STR_SUM,   STR_MEDIAN,   STR_VARIANCE, STR_STDDEV,
      STR_SORT,   STR_SORTBY, STR_RATIO, STR_RANGE, STR_REVERSE, STR_UNIQUE, STR_UNPACK, STR_REPEAT, STR_LEN, STR_FACT, STR_FACTORINT, STR_AVG, STR_MEAN,
      STR_MOD,    STR_CLAMP,  STR_HYPOT, STR_GCD,    STR_LCM,      STR_NCR, STR_NPR, STR_PRODUCT, STR_MIN, STR_MAX,
      STR_UHEX,   STR_UOCT,   STR_UBIN, STR_MILLISECONDS, STR_SECONDS, STR_MINUTES, STR_HOURS, STR_DAYS,
      STR_REAL,   STR_IMAG,   STR_PHASE, STR_POLAR, STR_CART, STR_CONJ};
  return kNames;
}

const std::unordered_map<std::string, MathParser::BuiltinFunctionId>& MathParser::functionNameToId() {
  static const std::unordered_map<std::string, BuiltinFunctionId> kByName = [] {
    std::unordered_map<std::string, BuiltinFunctionId> m;
    m.reserve(functionNames().size() + 8u);
    for (std::size_t i = 0; i < functionNames().size(); ++i) {
      m.emplace(functionNames()[i], static_cast<BuiltinFunctionId>(i));
    }
    m.emplace(STR_ARCSIN, BuiltinFunctionId::Asin);
    m.emplace(STR_ARCCOS, BuiltinFunctionId::Acos);
    m.emplace(STR_ARCTAN, BuiltinFunctionId::Atan);
    m.emplace(STR_FRACT, BuiltinFunctionId::Frac);
    m.emplace(STR_SORTED, BuiltinFunctionId::Sort);
    m.emplace(STR_REVERSED, BuiltinFunctionId::Reverse);
    m.emplace(STR_FACTORIAL, BuiltinFunctionId::Fact);
    m.emplace(STR_PROD, BuiltinFunctionId::Product);
    m.emplace(STR_LENGTH, BuiltinFunctionId::Len);
    return m;
  }();
  return kByName;
}

const std::vector<std::string>& MathParser::operatorNames() {
  static const std::vector<std::string> kNames = {STR_NOT, STR_AND, STR_OR};
  return kNames;
}

const std::string& MathParser::getFunctionName(BuiltinFunctionId id) {
  return functionNames().at(static_cast<std::size_t>(id));
}

const std::string& MathParser::opName(OperatorNameId id) {
  return operatorNames().at(static_cast<std::size_t>(id));
}

bool MathParser::tryGetBuiltinFunctionId(const std::string& nameText, BuiltinFunctionId& outId) {
  const auto& byName = functionNameToId();
  auto it = byName.find(nameText);
  if (it == byName.end()) {
    return false;
  }
  outId = it->second;
  return true;
}

MathParser::BuiltinHintKind MathParser::getBuiltinHintKind(BuiltinFunctionId id) {
  if (id >= BuiltinFunctionId::Count) {
    return BuiltinHintKind::None;
  }
  return kBuiltinMeta[static_cast<std::size_t>(id)].hintKind;
}

std::string MathParser::getBuiltinFunctionMissingCallHint(BuiltinFunctionId id) {
  const MathParser::BuiltinHintKind kind = getBuiltinHintKind(id);
  if (kind == BuiltinHintKind::None) return "";
  const std::string& fnName = getFunctionName(id);
  switch (kind) {
    case BuiltinHintKind::EmptyPar: return fnName + STR_PAR_EMPTY;
    case BuiltinHintKind::MinMax: return fnName + STR_PAR_MIN_COMMA_MAX;
    case BuiltinHintKind::DotDotDot: return fnName + STR_PAR_DOTDOTDOT;
    case BuiltinHintKind::ValuePower: return fnName + STR_PAR_VALUE_COMMA_POWER;
    case BuiltinHintKind::YX: return fnName + STR_PAR_Y_COMMA_X;
    case BuiltinHintKind::Angle: return fnName + STR_PAR_ANGLE;
    case BuiltinHintKind::Value: return fnName + STR_PAR_VALUE;
    case BuiltinHintKind::ValueBase: return fnName + STR_PAR_VALUE_COMMA_BASE;
    case BuiltinHintKind::N: return fnName + STR_PAR_N;
    case BuiltinHintKind::ValueDivisor: return fnName + STR_PAR_VALUE_COMMA_DIVISOR;
    case BuiltinHintKind::ValueMinMax: return fnName + STR_PAR_VALUE_COMMA_MIN_COMMA_MAX;
    case BuiltinHintKind::XY: return fnName + STR_PAR_X_COMMA_Y;
    case BuiltinHintKind::XN: return fnName + STR_PAR_X_COMMA_N;
    case BuiltinHintKind::X: return fnName + STR_PAR_X;
    case BuiltinHintKind::StartStopStep: return fnName + STR_PAR_START_COMMA_STOP_COMMA_STEP;
    case BuiltinHintKind::AB:
      if (id == BuiltinFunctionId::Ncr || id == BuiltinFunctionId::Npr) return fnName + STR_PAR_N_COMMA_R;
      return fnName + STR_PAR_A_COMMA_B;
    case BuiltinHintKind::ArrayFunc:
      return fnName + STR_PAR_ARRAY_COMMA_FUNC;
    default: return "";
  }
}

bool MathParser::isOpKeyword(const std::string& nameText, OperatorNameId id) {
  return nameText == opName(id);
}

bool MathParser::isLogicalBinaryOperatorKeyword(const std::string& nameText) {
  return isOpKeyword(nameText, OperatorNameId::And) || isOpKeyword(nameText, OperatorNameId::Or);
}

bool MathParser::isReservedFunctionName(const std::string& nameText) {
  BuiltinFunctionId id = BuiltinFunctionId::Count;
  return tryGetBuiltinFunctionId(nameText, id)
      || isOpKeyword(nameText, OperatorNameId::Not)
      || isLogicalBinaryOperatorKeyword(nameText);
}

const char* MathParser::getReservedIdentifierError(const std::string& ident) const {
  if (isReservedFunctionName(ident)) {
    return STR_RESERVED_FUNCTION_NAME;
  }
#if SMARTMATH_COMPLEX_NUMBERS
  if (getSupportComplexNumbers()) {
    std::string low = ident;
    std::transform(low.begin(), low.end(), low.begin(), [](unsigned char c) { return static_cast<char>(std::tolower(c)); });
    if (low == STR_I) {
      return STR_RESERVED_CONSTANT_NAME;
    }
  }
#endif
  if (ident == STR_PI || ident == STR_E || ident == STR_INF || ident == STR_NAN) {
    return STR_RESERVED_CONSTANT_NAME;
  }
#if SMARTMATH_TIME_VALUES
  if (getSupportTimeValues() &&
      (ident == STR_MILLISECOND || ident == STR_SECOND || ident == STR_MINUTE || ident == STR_HOUR ||
       ident == STR_DAY)) {
    return STR_RESERVED_CONSTANT_NAME;
  }
#endif
  return nullptr;
}

bool MathParser::isTrailingFormatterFunctionName(const std::string& nameText) {
  BuiltinFunctionId id = BuiltinFunctionId::Count;
  if (!tryGetBuiltinFunctionId(nameText, id)) {
    return false;
  }
  return hasBuiltinFlag(id, BuiltinFlags::TrailingFormatter);
}

bool MathParser::isBareFunctionNameAtExpressionTail(EvalContext& ctx) const {
  if (!ctx.p) {
    return false;
  }
  skipSpaces(ctx);
  return *ctx.p == '\0';
}

bool MathParser::trySetBareFunctionImmediateCloserError(
    EvalContext& ctx,
    const char* identStart) const {
  if (!ctx.p || !identStart || !ctx.start || identStart != ctx.start) {
    return false;
  }
  switch (*ctx.p) {
    case ')':
      setParseError(ctx, ParseErrorId::MismatchedClosingParenthesis);
      return true;
    case ']':
      setParseError(ctx, ParseErrorId::MismatchedClosingBracket);
      return true;
    case '}':
      setParseError(ctx, ParseErrorId::MismatchedClosingBrace);
      return true;
    default:
      return false;
  }
}

bool MathParser::identIsBareFunctionOrUdfName(const std::string& ident, const EvalContext& ctx) const {
  BuiltinFunctionId bid = BuiltinFunctionId::Count;
  if (tryGetBuiltinFunctionId(ident, bid)) {
    return true;
  }
  if (findUserFunction(ident)) {
    return true;
  }
  if (ctx.compilingUserFunctionParams) {
    return ctx.compilingUserFunctionParams->find(ident) != ctx.compilingUserFunctionParams->end();
  }
  return false;
}

bool MathParser::trimmedStmtIsBareFunctionOrUdfName(const std::string& stmt) const {
  return trimmedStmtIsBareFunctionOrUdfName(stmt.data(), stmt.data() + stmt.size());
}

bool MathParser::trimmedStmtIsBareFunctionOrUdfName(const char* begin, const char* end) const {
  while (begin < end && std::isspace(static_cast<unsigned char>(*begin))) {
    ++begin;
  }
  while (end > begin && std::isspace(static_cast<unsigned char>(end[-1]))) {
    --end;
  }
  if (begin >= end || !isIdentStart(*begin)) {
    return false;
  }
  const char* identEnd = begin + 1;
  while (identEnd < end && isIdentChar(*identEnd)) {
    ++identEnd;
  }
  if (identEnd != end) {
    return false;
  }
  std::string ident;
  assignLowerIdentFromRange(ident, begin, identEnd);
  EvalContext ctx;
  return identIsBareFunctionOrUdfName(ident, ctx);
}

bool MathParser::stripTrailingSemicolonsForTopLevelInput(std::string& s) const {
  bool strippedAny = false;
  for (;;) {
    while (!s.empty() && std::isspace(static_cast<unsigned char>(s.back()))) {
      s.pop_back();
    }
    if (!s.empty() && s.back() == ';') {
      const std::string beforeSemi = trim(std::string(s.data(), s.size() - 1));
      if (trimmedStmtIsBareFunctionOrUdfName(beforeSemi)) {
        return strippedAny;
      }
      // Keep trailing ';' on multi-statement input so the last segment stays semicolon-terminated.
      if (beforeSemi.find(';') != std::string::npos) {
        return strippedAny;
      }
      s.pop_back();
      strippedAny = true;  // Path C: stripped trailing ';' hides semicolon from the parse buffer.
      continue;
    }
    break;
  }
  return strippedAny;
}

// Builtins are lowercase a-z except names with digits (atan2, log10). Skip hash lookup for other idents with digit/_ (e.g. myconst1).
bool MathParser::identMayBeBareBuiltinName(const char* begin, const char* end) {
  if (!identHasDigitOrUnderscore(begin, end)) {
    return true;
  }
  std::string low;
  assignLowerIdentFromRange(low, begin, end);
  BuiltinFunctionId bid = BuiltinFunctionId::Count;
  return tryGetBuiltinFunctionId(low, bid);
}

bool MathParser::identMayBeBareBuiltinName(const std::string& ident) {
  return identMayBeBareBuiltinName(ident.data(), ident.data() + ident.size());
}

bool MathParser::tryHandleBareFunctionNameCommonTail(
    EvalContext& ctx,
    const std::string& ident,
    const char* identStartForCloser) const {
  skipSpaces(ctx);
  if (*ctx.p == ';') {
    setParseError(ctx, ParseErrorId::UnexpectedToken);
    return true;
  }
  if (!isBareFunctionNameAtExpressionTail(ctx)) {
    if (identStartForCloser != nullptr
        && trySetBareFunctionImmediateCloserError(ctx, identStartForCloser)) {
      return true;
    }
    setError(ctx, buildUnknownVariableErrorText(ident));
    return true;
  }
  return false;
}

bool MathParser::trySetMissingFunctionCallError(
    EvalContext& ctx,
    const std::string& ident,
    const char* identStartForCloser) const {
  BuiltinFunctionId bid = BuiltinFunctionId::Count;
  if (!tryGetBuiltinFunctionId(ident, bid)) {
    return false;
  }
  if (tryHandleBareFunctionNameCommonTail(ctx, ident, identStartForCloser)) {
    return true;
  }
  if (ctx.suppressBareFunctionTailHint) {
    setError(ctx, buildUnknownVariableErrorText(ident));
    return true;
  }
  std::string hint = getBuiltinFunctionMissingCallHint(bid);
  if (!hint.empty()) setFunctionHintError(ctx, hint);
  else setFunctionHintError(ctx, ident + STR_PAR_DOTDOTDOT);
  return true;
}

bool MathParser::trySetIncompleteOpenedFunctionCallHint(
    EvalContext& ctx,
    const std::string& ident) const {
  if (ctx.suppressBareFunctionTailHint) {
    return false;
  }
  if (!isBareFunctionNameAtExpressionTail(ctx)) {
    return false;
  }
  if (MathParser::identMayBeBareBuiltinName(ident) && trySetMissingFunctionCallError(ctx, ident)) {
    return true;
  }
  if ((!userFunctionIndex_.empty() || ctx.compilingUserFunctionParams != nullptr)
      && trySetBareUserFunctionNameError(ctx, ident, nullptr)) {
    return true;
  }
  return false;
}

bool MathParser::trySetBareUserFunctionNameError(
    EvalContext& ctx,
    const std::string& ident,
    const char* identStart) const {
  const UserFunction* uf = findUserFunction(ident);
  const std::vector<std::string>* compilingParams = nullptr;
  if (!uf && ctx.compilingUserFunctionParams) {
    const auto it = ctx.compilingUserFunctionParams->find(ident);
    if (it != ctx.compilingUserFunctionParams->end()) {
      compilingParams = &it->second;
    }
  }
  if (!uf && !compilingParams) {
    return false;
  }
  if (tryHandleBareFunctionNameCommonTail(ctx, ident, identStart)) {
    return true;
  }
  if (ctx.suppressBareFunctionTailHint
      || (ctx.validatingUserFunctionName != nullptr && ident == ctx.validatingUserFunctionName)
      || (ctx.userFunctionCallStack != nullptr && !ctx.userFunctionCallStack->empty()
          && ctx.userFunctionCallStack->back() == ident)) {
    setError(ctx, buildUnknownVariableErrorText(ident));
    return true;
  }
  if (uf) {
    setError(ctx, std::string(STR_USER_DEFINED_FUNCTION_COLON) + formatUserFunctionSignature(*uf));
  } else {
    UserFunction stub;
    stub.name = ident;
    stub.params = *compilingParams;
    setError(ctx, std::string(STR_USER_DEFINED_FUNCTION_COLON) + formatUserFunctionSignature(stub));
  }
  return true;
}

std::string MathParser::formatUserFunctionSignature(const UserFunction& uf) {
  std::string sig = uf.name + "(";
  for (std::size_t i = 0; i < uf.params.size(); ++i) {
    if (i > 0) {
      sig += ",";
    }
    sig += uf.params[i];
  }
  sig += ")";
  return sig;
}

bool MathParser::finishSubEvalFromCompiledProgram(EvalContext& parent, EvalContext& sub) const {
  if (!sub.unknownVarsText.empty()) {
    setError(parent, buildUnknownVariableErrorText(sub.unknownVarsText));
    return false;
  }
  if (sub.parseError) {
    setError(parent, sub.errorText);
    return false;
  }
  mergeUnknownNameList(parent.unknownVarsText, sub.unknownVarsText);
  mergeUnknownNameList(parent.unknownFuncsText, sub.unknownFuncsText);
  return true;
}

bool MathParser::tryCompileExprToProgram(
    const std::string& expr,
    EvalContext& parentCtx,
    std::vector<std::string>* callStack,
    const char* validatingUserFunctionName,
    bool suppressBareFunctionTailHint,
    std::vector<AstStatement>& outProgram) {
  EvalContext parseCtx;
  parseCtx.p = expr.c_str();
  parseCtx.start = parseCtx.p;
  parseCtx.sourceExpr = expr;
  parseCtx.userFunctionCallStack = callStack;
  parseCtx.validatingUserFunctionName = validatingUserFunctionName;
  parseCtx.compilingUserFunctionParams = parentCtx.compilingUserFunctionParams;
  parseCtx.suppressBareFunctionTailHint = suppressBareFunctionTailHint;
  if (!parseProgram(parseCtx, outProgram)) {
    setError(
        parentCtx,
        parseCtx.errorText.empty() ? STR_FAILED_TO_PARSE_USER_FUNCTION_BODY : parseCtx.errorText);
    return false;
  }
  skipSpaces(parseCtx);
  if (!parseCtx.parseError && *parseCtx.p != '\0') {
    setParseError(parentCtx, ParseErrorId::UnexpectedInput);
    return false;
  }
  return true;
}

bool MathParser::handleUnknownIdentifier(EvalContext& ctx, const std::string& ident, std::string& unknownList) const {
  if (isLogicalBinaryOperatorKeyword(ident)) {
    setParseError(ctx, ParseErrorId::UnexpectedToken);
    return true;
  }
  if (!ident.empty()) {
    if (MathParser::identMayBeBareBuiltinName(ident) && trySetMissingFunctionCallError(ctx, ident)) {
      return true;
    }
    if ((!userFunctionIndex_.empty() || ctx.compilingUserFunctionParams != nullptr)
        && trySetBareUserFunctionNameError(ctx, ident, nullptr)) {
      return true;
    }
  }
  appendUniqueName(unknownList, ident);
  return false;
}

bool MathParser::tryResolveVariableValue(
    const Expr& e,
    const std::unordered_map<std::string, EvalValue>* scopedVars,
    EvalValue& out) const {
  if (scopedVars) {
    auto it = scopedVars->find(e.name);
    if (it != scopedVars->end()) {
      if (it->second.kind == ValueKind::UdfFormalValidationDummy) {
        auto pit = variables_.find(STR_FORMAL_VALIDATION_PROBE);
        if (pit != variables_.end() && pit->second.kind == ValueKind::Scalar) {
          out = pit->second;
        } else {
          out = makeScalarInt(1);
        }
      } else {
        out = it->second;
      }
      return true;
    }
  }
  if (e.boundVariable != nullptr) {
    out = *e.boundVariable;
    return true;
  }
  auto it = variables_.find(e.name);
  if (it != variables_.end()) {
    out = it->second;
    return true;
  }
  return false;
}

MathParser::EvalValue MathParser::makeUdfFormalValidationDummy() {
  EvalValue v;
  v.kind = ValueKind::UdfFormalValidationDummy;
  return v;
}

static bool isReservedBuiltinConstantName(const std::string& nameText, bool supportTimeValues) {
  if (nameText == STR_PI || nameText == STR_E || nameText == STR_INF || nameText == STR_NAN) {
    return true;
  }
  return supportTimeValues &&
         (nameText == STR_MILLISECOND || nameText == STR_SECOND || nameText == STR_MINUTE ||
          nameText == STR_HOUR || nameText == STR_DAY);
}

static bool isReservedBuiltinVariableNameForUserFunctionDefinition(const std::string& fnName) {
  std::string lowered = fnName;
  std::transform(lowered.begin(), lowered.end(), lowered.begin(), [](unsigned char c) {
    return static_cast<char>(std::tolower(c));
  });
  if (lowered == STR_ANS) {
    return true;
  }
  if (fnName == STR_FORMAL_VALIDATION_PROBE) {
    return true;
  }
  return false;
}

const char* MathParser::validateUserFunctionDefinitionNames(
    const std::string& fnName,
    const std::vector<std::string>& fnParams) const {
  if (isReservedBuiltinVariableNameForUserFunctionDefinition(fnName)) {
    return STR_RESERVED_BUILTIN_VARIABLE_NAME;
  }
  if (const char* reservedErr = getReservedIdentifierError(fnName)) {
    return reservedErr;
  }
  std::unordered_map<std::string, bool> seen;
  for (const auto& p : fnParams) {
    if (seen.find(p) != seen.end()) {
      return STR_DUPLICATE_PARAMETER_NAME;
    }
    if (isReservedBuiltinConstantName(p, getSupportTimeValues())) {
      return STR_RESERVED_CONSTANT_NAME;
    }
#if SMARTMATH_COMPLEX_NUMBERS
    {
      std::string low = p;
      std::transform(low.begin(), low.end(), low.begin(), [](unsigned char c) { return static_cast<char>(std::tolower(c)); });
      if (getSupportComplexNumbers() && low == STR_I) {
        return STR_RESERVED_CONSTANT_NAME;
      }
    }
#endif
    seen[p] = true;
  }
  return nullptr;
}

std::string MathParser::getUserFunctionDefinitionErrorText(
    const std::string& fnName,
    const std::vector<std::string>& fnParams,
    const std::string& fnExpr,
    const bool evaluateBody) {
  if (const char* udfNameErr = validateUserFunctionDefinitionNames(fnName, fnParams)) {
    return udfNameErr;
  }
  if (trim(fnExpr).empty() || udfBodyIsEmptyTupleLiteral(fnExpr)) {
    return STR_FUNCTION_BODY_IS_EMPTY;
  }
  if (udfBodyCallsDefinedFunction(fnExpr, fnName)) {
    return STR_RECURSIVE_USER_FUNCTION_CALL_COLON + fnName;
  }
  EvalContext bodyCtx;
  bodyCtx.sourceExpr = fnExpr;
  bodyCtx.p = fnExpr.c_str();
  bodyCtx.start = bodyCtx.p;
  std::unordered_map<std::string, std::vector<std::string>> validatingUserFunctionParams;
  validatingUserFunctionParams[fnName] = fnParams;
  bodyCtx.compilingUserFunctionParams = &validatingUserFunctionParams;
  bodyCtx.validatingUserFunctionName = fnName.c_str();
  std::unique_ptr<Expr> ex = parseExpression(bodyCtx);
  if (bodyCtx.parseError || !ex) {
    if (!bodyCtx.errorText.empty()) {
      return bodyCtx.errorText;
    }
    return STR_FAILED_TO_PARSE_USER_FUNCTION_BODY;
  }
  skipSpaces(bodyCtx);
  if (*bodyCtx.p != '\0') {
    return STR_UNEXPECTED_INPUT;
  }

  if (!evaluateBody) {
    return "";
  }

  std::unordered_map<std::string, EvalValue> formalScoped;
  formalScoped.reserve(fnParams.size());
  for (const auto& p : fnParams) {
    formalScoped.emplace(p, makeUdfFormalValidationDummy());
  }

  bodyCtx.parseError = false;
  bodyCtx.errorText.clear();
  bodyCtx.unknownVarsText.clear();
  bodyCtx.unknownFuncsText.clear();
  (void)evalExpr(*ex, bodyCtx, &formalScoped);

  if (bodyCtx.parseError) {
    if (!bodyCtx.errorText.empty()) {
      return bodyCtx.errorText;
    }
    return STR_FAILED_TO_PARSE_USER_FUNCTION_BODY;
  }
  // Like Basic TryValidateUserFunctionBodyExpression: do not treat accumulated unknown
  // identifiers as definition errors (late-bound UDFs may appear only after later statements).
  return "";
}

bool MathParser::trySetUserFunctionDefinitionError(
    EvalContext& ctx,
    const std::string& fnName,
    const std::vector<std::string>& fnParams,
    const std::string& fnExpr) {
  // Parse-only here: formal-probe evaluation needs variables from earlier statements in the
  // same program, which are not applied until runCompiledProgram (Basic validates the body now).
  const std::string err = getUserFunctionDefinitionErrorText(fnName, fnParams, fnExpr, false);
  if (err.empty()) {
    return false;
  }
  setValidationError(ctx, err.c_str());
  return true;
}

const char* MathParser::validateAssignmentTargetName(const std::string& ident) const {
  return getReservedIdentifierError(ident);
}

std::string MathParser::buildUnknownVariableErrorText(const std::string& unknownVarsText) {
  return STR_UNKNOWN_VARIABLE_COLON + unknownVarsText;
}

std::string MathParser::buildUnknownFunctionErrorText(const std::string& unknownFuncsText) {
  return STR_UNKNOWN_FUNCTION_COLON + unknownFuncsText;
}

void MathParser::appendUnknownFunctionErrorText(std::string& errorText, const std::string& unknownFuncsText) {
  errorText += STR_SEMICOLON_UNKNOWN_FUNCTION_COLON + unknownFuncsText;
}

void MathParser::setValidationError(EvalContext& ctx, const char* errorText) const {
  setError(ctx, errorText);
}

bool MathParser::trySetUnknownNameError(const EvalContext& ctx) {
  lastError_ = buildUnknownNameErrorText(ctx.unknownVarsText, ctx.unknownFuncsText);
  return !lastError_.empty();
}

std::string MathParser::buildUnknownNameErrorText(
    const std::string& unknownVarsText,
    const std::string& unknownFuncsText) {
  if (!unknownVarsText.empty()) {
    std::string errorText = buildUnknownVariableErrorText(unknownVarsText);
    if (!unknownFuncsText.empty()) {
      appendUnknownFunctionErrorText(errorText, unknownFuncsText);
    }
    return errorText;
  }
  if (!unknownFuncsText.empty()) {
    return buildUnknownFunctionErrorText(unknownFuncsText);
  }
  return "";
}

bool MathParser::tryAppendParsedExpressionStatement(
    EvalContext& ctx,
    std::vector<AstStatement>& out) {
  auto ex = parseExpression(ctx);
  if (ctx.parseError || !ex) {
    return false;
  }
  AstStatement st;
  st.kind = AstStatement::Kind::Expr;
  st.expr = std::move(ex);
  out.emplace_back(std::move(st));
  return true;
}

bool MathParser::tryAppendFunctionDefinitionStatement(EvalContext& ctx, std::vector<AstStatement>& out) {
  std::string fnName;
  std::vector<std::string> fnParams;
  std::string fnExpr;
  if (!parseFunctionDefinition(ctx, fnName, fnParams, fnExpr)) {
    return false;
  }
  if (trySetUserFunctionDefinitionError(ctx, fnName, fnParams, fnExpr)) {
    return false;
  }
  AstStatement st;
  st.kind = AstStatement::Kind::FunDef;
  st.fun = UserFunction{fnName, fnParams, fnExpr};
  out.emplace_back(std::move(st));
  return true;
}

bool MathParser::tryAppendAssignOrExpressionStatement(EvalContext& ctx, std::vector<AstStatement>& out) {
  const char* assignSave = ctx.p;
  if (peekIdentFollowedByAssignEquals(assignSave)) {
    std::string ident = consumeLowerIdentToken(ctx);
    skipSpaces(ctx);
    if (*ctx.p == '=' && ctx.p[1] != '=') {
      if (const char* assignNameErr = validateAssignmentTargetName(ident)) {
        setValidationError(ctx, assignNameErr);
        return false;
      }
      ++ctx.p;
      const char* afterEq = ctx.p;
#if SMARTMATH_LAMBDA_FUNCTIONS
      if (getSupportLambdaFunctions()) {
        std::vector<std::string> lamParams;
        std::string lamBody;
        if (peekRhsMayBeLambdaSyntaxAt(afterEq) &&
            tryParseLambdaRhsAfterEquals(ctx, lamParams, lamBody)) {
          if (trySetUserFunctionDefinitionError(ctx, ident, lamParams, lamBody)) {
            return false;
          }
          AstStatement st;
          st.kind = AstStatement::Kind::FunDef;
          st.fun = UserFunction{ident, lamParams, lamBody};
          out.emplace_back(std::move(st));
          return true;
        }
      } else if (peekRhsMayBeLambdaSyntaxAt(afterEq)) {
        setParseError(ctx, ParseErrorId::UnexpectedToken);
        return false;
      }
#endif
      ctx.p = afterEq;
      auto ex = parseExpression(ctx);
      if (ctx.parseError || !ex) {
        return false;
      }
      AstStatement st;
      st.kind = AstStatement::Kind::Assign;
      st.assignName = std::move(ident);
      st.expr = std::move(ex);
      out.emplace_back(std::move(st));
      return true;
    }
    ctx.p = assignSave;
  }
  return tryAppendParsedExpressionStatement(ctx, out);
}

bool MathParser::consumeProgramStatementSeparator(EvalContext& ctx) {
  skipSpaces(ctx);
  if (*ctx.p == ';') {
    ++ctx.p;
    return true;
  }
  if (*ctx.p == '\0') {
    return false;
  }
  setParseError(ctx, ParseErrorId::UnexpectedTokenAfterExpression);
  return false;
}

bool MathParser::hasExprParseFailure(const EvalContext& ctx, const std::unique_ptr<Expr>& node) {
  return ctx.parseError || !node;
}

void MathParser::setParseError(EvalContext& ctx, ParseErrorId id) const {
  static const char* const texts[static_cast<std::size_t>(ParseErrorId::Count)] = {
      STR_UNEXPECTED_COMMA,
      STR_INDEXING_REQUIRES_AN_ARRAY_VALUE,
      STR_MISSING_INDEX,
      STR_ARRAY_INDEX_MUST_BE_A_SCALAR,
      STR_ARRAY_INDEX_MUST_BE_AN_INTEGER,
      STR_ARRAY_INDEX_IS_OUT_OF_RANGE,
      STR_MISSING_CLOSING_BRACKET,
      STR_MISSING_CLOSING_PARENTHESIS,
      STR_MISMATCHED_CLOSING_PARENTHESIS,
      STR_MISMATCHED_CLOSING_BRACKET,
      STR_MISMATCHED_CLOSING_BRACE,
      STR_BITWISE_OPERANDS_MUST_BE_INTEGER_VALUES,
      STR_MODULO_OPERANDS_MUST_BE_INTEGER_VALUES,
      STR_INCOMPATIBLE_OPERANDS,
      STR_INVALID_HEX_LITERAL,
      STR_INVALID_BINARY_LITERAL,
      STR_INVALID_OCTAL_LITERAL,
      STR_INTERNAL_UNARY_OP,
      STR_INTERNAL_BINARY_OP,
      STR_INTERNAL_EVAL_ERROR,
      STR_NUMERIC_ERROR_IN_POWER_OPERATION,
      STR_NUMERIC_ERROR_IN_EXPRESSION,
      STR_USER_FUNCTION_CALL_STACK_OVERFLOW,
      STR_MAX_EVALUATION_DEPTH_REACHED,
      STR_INVALID_NUMERIC_LITERAL,
      STR_PERCENTAGE_REQUIRES_SCALAR_VALUE,
      STR_FAILED_TO_BUILD_ARRAY_LITERAL,
      STR_INTERNAL_ERROR_IN_AGGREGATE_BUILTIN,
      STR_INTERNAL_ERROR_IN_SCALAR_BINARY_BUILTIN,
      STR_INTERNAL_ERROR_IN_UNARY_MATH_BUILTIN,
      STR_UNEXPECTED_TOKEN_AFTER_EXPRESSION,
      STR_SCALAR_ONLY_EXPRESSION_ENCOUNTERED_NON,
      STR_PARSE_FAILED,
      STR_UNEXPECTED_TOKEN,
      STR_UNEXPECTED_INPUT,
  };
  setError(ctx, texts[static_cast<std::size_t>(id)]);
}

void MathParser::setNumericErrorInFunction(EvalContext& ctx, const std::string& fnName) const {
  setError(ctx, STR_NUMERIC_ERROR_IN + fnName + STR_PAR);
}

void MathParser::setAtLeastNArgsError(
    EvalContext& ctx,
    const std::string& fnName,
    uint8_t minArgs) const {
  setError(
      ctx,
      fnName + STR_PAR_EXPECTS_AT_LEAST + std::to_string(minArgs) +
          (minArgs == 1 ? STR_ARGUMENT : STR_ARGUMENTS));
}

void MathParser::setAtLeastOneArgError(EvalContext& ctx, const std::string& fnName) const {
  setAtLeastNArgsError(ctx, fnName, 1);
}

void MathParser::setExactArgCountError(
    EvalContext& ctx,
    const std::string& fnName,
    size_t expectedCount) const {
  setError(ctx, fnName + STR_PAR_EXPECTS + std::to_string(expectedCount) + STR_ARGUMENT_PAR_S);
}

#if SMARTMATH_TIME_VALUES
bool MathParser::rejectBinaryBuiltinTimeOperands(
    EvalContext& ctx,
    const EvalValue& left,
    const EvalValue& right) const {
  if (evalValueInvolvesTime(left) || evalValueInvolvesTime(right)) {
    setParseError(ctx, ParseErrorId::IncompatibleOperands);
    return true;
  }
  return false;
}
#endif

bool MathParser::rejectInt64BinaryOperands(
    EvalContext& ctx,
    const EvalValue& left,
    const EvalValue& right,
    const bool isModulo) const {
#if SMARTMATH_TIME_VALUES
  if (getSupportTimeValues() && (evalValueInvolvesTime(left) || evalValueInvolvesTime(right))) {
    if (isModulo) {
      setParseError(ctx, ParseErrorId::ModuloIntegerOperands);
    } else {
      setParseError(ctx, ParseErrorId::IncompatibleOperands);
    }
    return true;
  }
#endif
#if SMARTMATH_COMPLEX_NUMBERS
  if (getSupportComplexNumbers() &&
      (MathParser::evalValueHasNonzeroImaginary(left) || MathParser::evalValueHasNonzeroImaginary(right))) {
    setParseError(ctx, ParseErrorId::IncompatibleOperands);
    return true;
  }
#endif
  return false;
}

#if SMARTMATH_TIME_VALUES
bool MathParser::rejectNumericBinaryPowWithTime(
    EvalContext& ctx,
    const EvalValue& left,
    const EvalValue& right,
    const char op) const {
  if (op == '^' && getSupportTimeValues() && (evalValueInvolvesTime(left) || evalValueInvolvesTime(right))) {
    setParseError(ctx, ParseErrorId::IncompatibleOperands);
    return true;
  }
  return false;
}
#endif

MathParser::EvalValue MathParser::builtinMapBinaryTwoArg(
    EvalContext& ctx,
    const std::string& fnName,
    BuiltinFunctionId id,
    const std::vector<EvalValue>& args,
    const bool rejectComplexOperands,
    const bool numericErrorOnMapFailure) const {
#if SMARTMATH_TIME_VALUES
  if (rejectBinaryBuiltinTimeOperands(ctx, args[0], args[1])) {
    return makeScalar(0);
  }
#endif
#if SMARTMATH_COMPLEX_NUMBERS
  if (rejectComplexOperands && getSupportComplexNumbers() &&
      (evalValueHasNonzeroImaginary(args[0]) || evalValueHasNonzeroImaginary(args[1]))) {
    setParseError(ctx, ParseErrorId::IncompatibleOperands);
    return makeScalar(0);
  }
#endif
  bool ok = false;
  EvalValue out = mapBinaryBuiltinMathFunction(args[0], args[1], id, ok);
  if (!ok) {
    if (!ctx.parseError) {
      setBinaryBuiltinBroadcastFailure(ctx, fnName, args[0], args[1], numericErrorOnMapFailure ? 2 : 3);
    }
    return makeScalar(0);
  }
  return out;
}

void MathParser::setScalarValuesError(EvalContext& ctx, const std::string& fnName) const {
  setError(ctx, fnName + STR_PAR_EXPECTS_SCALAR_VALUES);
}

void MathParser::setIntegerValuesError(EvalContext& ctx, const std::string& fnName) const {
  setError(ctx, fnName + STR_PAR_EXPECTS_INTEGER_VALUES);
}

void MathParser::setScalarMinMaxError(EvalContext& ctx, const std::string& fnName) const {
  setError(ctx, fnName + STR_PAR_EXPECTS_SCALAR_MIN_SLASH);
}

void MathParser::setNonNegativeIntegerError(EvalContext& ctx, const std::string& fnName) const {
  setError(ctx, fnName + STR_PAR_EXPECTS_A_NON_DASH);
}

void MathParser::setPositiveIntegerError(EvalContext& ctx, const std::string& fnName) const {
  setError(ctx, fnName + STR_PAR_EXPECTS_A_POSITIVE_INTEGER);
}

void MathParser::setRangeStepZeroError(EvalContext& ctx, const std::string& fnName) const {
  setError(ctx, fnName + STR_PAR_STEP_MUST_NOT_BE_ZERO);
}

void MathParser::setRangeEmptyError(EvalContext& ctx, const std::string& fnName) const {
  setError(ctx, fnName + STR_PAR_PRODUCES_NO_VALUES);
}

void MathParser::setTooManyValuesError(EvalContext& ctx, const std::string& fnName) const {
  setError(ctx, fnName + STR_PAR_PRODUCES_TOO_MANY_VALUES);
}

namespace {

// CPython-style length of range(start, stop, step) using unsigned wraparound arithmetic.
std::uint64_t estimateRangeItemCount(long long startV, long long stopV, long long stepV) {
  std::uint64_t uLo = 0;
  std::uint64_t uHi = 0;
  std::uint64_t uStep = 0;
  if (stepV > 0) {
    if (startV >= stopV) {
      return 0;
    }
    uLo = static_cast<std::uint64_t>(startV);
    uHi = static_cast<std::uint64_t>(stopV);
    uStep = static_cast<std::uint64_t>(stepV);
  } else {
    if (startV <= stopV) {
      return 0;
    }
    uLo = static_cast<std::uint64_t>(stopV);
    uHi = static_cast<std::uint64_t>(startV);
    uStep = 0ull - static_cast<std::uint64_t>(stepV);
  }
  return (uHi - uLo - 1ull) / uStep + 1ull;
}

bool checkedMulU64(std::uint64_t a, std::uint64_t b, std::uint64_t& out) {
  if (b != 0ull && a > (std::numeric_limits<std::uint64_t>::max)() / b) {
    return false;
  }
  out = a * b;
  return true;
}

}  // namespace

bool MathParser::acceptProducedItemCount(EvalContext& ctx, const std::string& fnName, std::uint64_t count) const {
  if (count > static_cast<std::uint64_t>(BUILTIN_MAX_PRODUCED_ITEMS)) {
    setTooManyValuesError(ctx, fnName);
    return false;
  }
  return true;
}

bool MathParser::evalValuesHaveMismatchedArrayLengths(const EvalValue& left, const EvalValue& right) const {
  return left.kind == ValueKind::Array && right.kind == ValueKind::Array && left.arr.size() != right.arr.size();
}

void MathParser::setBinaryBuiltinBroadcastFailure(
    EvalContext& ctx,
    const std::string& fnName,
    const EvalValue& left,
    const EvalValue& right,
    int pairStatus) const {
  if (pairStatus == 1) {
    setIntegerValuesError(ctx, fnName);
    return;
  }
  if (pairStatus == 3 || evalValuesHaveMismatchedArrayLengths(left, right)) {
    setParseError(ctx, ParseErrorId::IncompatibleOperands);
    return;
  }
  if (!ctx.parseError && ctx.errorText.empty()) {
    setNumericErrorInFunction(ctx, fnName);
  }
}

void MathParser::setFunctionHintError(EvalContext& ctx, const std::string& hintText) const {
  setError(ctx, STR_HINT_PREFIX + hintText);
}

void MathParser::setInvalidPrefixedLiteralError(EvalContext& ctx, char prefixChar) const {
  switch (prefixChar) {
    case 'x':
    case 'X':
      setParseError(ctx, ParseErrorId::InvalidHexLiteral);
      return;
    case 'b':
    case 'B':
      setParseError(ctx, ParseErrorId::InvalidBinaryLiteral);
      return;
    case 'o':
    case 'O':
      setParseError(ctx, ParseErrorId::InvalidOctalLiteral);
      return;
    default:
      setParseError(ctx, ParseErrorId::InvalidNumericLiteral);
      return;
  }
}

void MathParser::setRecursiveUserFunctionCallError(EvalContext& ctx, const std::string& fnName) const {
  setError(ctx, STR_RECURSIVE_USER_FUNCTION_CALL_COLON + fnName);
}

bool MathParser::isIdentStart(char c) {
  return localIsIdentStart(c);
}

bool MathParser::isIdentChar(char c) {
  return localIsIdentChar(c);
}

bool MathParser::isNumericLiteralStart(char c) {
  return std::isdigit(static_cast<unsigned char>(c)) || c == '.';
}

std::string MathParser::consumeLowerIdentToken(EvalContext& ctx) {
  const char* start = ctx.p++;
  while (isIdentChar(*ctx.p)) {
    ++ctx.p;
  }
  const char* const end = ctx.p;
  std::string out;
  assignLowerIdentFromRange(out, start, end);
  return out;
}

bool MathParser::tryConsumeCommaArgSeparator(EvalContext& ctx, bool& hasComma) const {
  hasComma = false;
  if (*ctx.p != ',') {
    return true;
  }
  hasComma = true;
  ++ctx.p;
  skipSpaces(ctx);
  if (*ctx.p == ')' || *ctx.p == ',') {
    setParseError(ctx, ParseErrorId::UnexpectedComma);
    return false;
  }
  return true;
}

void MathParser::skipSpaces(EvalContext& ctx) {
  while (*ctx.p && std::isspace(static_cast<unsigned char>(*ctx.p))) {
    ++ctx.p;
  }
}

void MathParser::setMissingClosingParenLikeError(EvalContext& ctx) const {
  if (isIdentStart(*ctx.p) || isNumericLiteralStart(*ctx.p)) {
    setParseError(ctx, ParseErrorId::UnexpectedToken);
  } else {
    setParseError(ctx, ParseErrorId::MissingClosingParenthesis);
  }
}

bool MathParser::consumeKeyword(EvalContext& ctx, const char* kw) {
  skipSpaces(ctx);
  std::size_t n = 0;
  while (kw[n] != '\0') {
    const unsigned char a = static_cast<unsigned char>(ctx.p[n]);
    const unsigned char b = static_cast<unsigned char>(kw[n]);
    if (a == '\0' || std::tolower(a) != std::tolower(b)) return false;
    ++n;
  }
  if (isIdentChar(ctx.p[n])) return false;
  ctx.p += n;
  return true;
}

static bool isPercentPostfixContext(const char* p) {
  if (*p != '%') {
    return false;
  }
  const char* q = p + 1;
  while (*q && std::isspace(static_cast<unsigned char>(*q))) {
    ++q;
  }
  if (*q == '\0') return true;
  if (*q == ')' || *q == ',' || *q == ';') return true;
  if (*q == '+' || *q == '-' || *q == '*' || *q == '/' || *q == '^') return true;
  if (*q == '<' || *q == '>' || *q == '=' || *q == '!') return true;
  if (*q == '&' || *q == '|') return true;
  return false;
}

bool MathParser::tryConsumeLogicalBinaryOperator(EvalContext& ctx, OperatorNameId keywordId, char symbol) const {
  skipSpaces(ctx);
  if (ctx.p[0] == symbol && ctx.p[1] == symbol) {
    ctx.p += 2;
    return true;
  }
  return consumeKeyword(ctx, opName(keywordId).c_str());
}

bool MathParser::parseParenthesizedExprList(
    EvalContext& ctx,
    std::vector<std::unique_ptr<Expr>>& outValues) {
  skipSpaces(ctx);
  if (*ctx.p == ')') {
    return true;
  }
  while (true) {
    outValues.emplace_back(parseExpression(ctx));
    if (ctx.parseError) {
      return false;
    }
    skipSpaces(ctx);
    bool hasComma = false;
    if (!tryConsumeCommaArgSeparator(ctx, hasComma)) {
      return false;
    }
    if (!hasComma) {
      return true;
    }
  }
}

std::unique_ptr<MathParser::Expr> MathParser::makeBinaryExpr(
    std::unique_ptr<Expr> left,
    std::unique_ptr<Expr> right,
    Expr::BinaryOp op,
    bool setPercentFlag) {
  auto out = std::make_unique<Expr>();
  out->tag = Expr::Tag::Binary;
  out->binaryOp = op;
  out->left = std::move(left);
  out->right = std::move(right);
  if (setPercentFlag) {
    out->rhsIsDirectPostfixPercent = exprIsDirectPostfixPercent(*out->right);
  }
  return out;
}

std::unique_ptr<MathParser::Expr> MathParser::parseLeftAssocBinary(
    EvalContext& ctx,
    std::unique_ptr<Expr> (MathParser::*parseOperand)(EvalContext&),
    bool (MathParser::*tryConsumeOp)(EvalContext&, Expr::BinaryOp&),
    bool setPercentFlag) {
  auto left = (this->*parseOperand)(ctx);
  if (ctx.parseError || !left) {
    return nullptr;
  }
  while (true) {
    Expr::BinaryOp op = Expr::BinaryOp::None;
    if (!(this->*tryConsumeOp)(ctx, op)) {
      break;
    }
    auto right = (this->*parseOperand)(ctx);
    if (ctx.parseError || !right) {
      return nullptr;
    }
    left = makeBinaryExpr(std::move(left), std::move(right), op, setPercentFlag);
  }
  return left;
}

bool MathParser::tryConsumeBitAndOp(EvalContext& ctx, Expr::BinaryOp& outOp) {
  skipSpaces(ctx);
  if (*ctx.p != '&' || ctx.p[1] == '&') {
    return false;
  }
  ++ctx.p;
  outOp = Expr::BinaryOp::BitAnd;
  return true;
}

bool MathParser::tryConsumeBitXorOp(EvalContext& ctx, Expr::BinaryOp& outOp) {
  skipSpaces(ctx);
  if (*ctx.p != '^') {
    return false;
  }
  ++ctx.p;
  outOp = Expr::BinaryOp::BitXor;
  return true;
}

bool MathParser::tryConsumeBitOrOp(EvalContext& ctx, Expr::BinaryOp& outOp) {
  skipSpaces(ctx);
  if (*ctx.p != '|' || ctx.p[1] == '|') {
    return false;
  }
  ++ctx.p;
  outOp = Expr::BinaryOp::BitOr;
  return true;
}

bool MathParser::tryConsumeShiftOp(EvalContext& ctx, Expr::BinaryOp& outOp) {
  skipSpaces(ctx);
  if (ctx.p[0] == '<' && ctx.p[1] == '<') {
    ctx.p += 2;
    outOp = Expr::BinaryOp::ShiftLeft;
    return true;
  }
  if (ctx.p[0] == '>' && ctx.p[1] == '>') {
    ctx.p += 2;
    outOp = Expr::BinaryOp::ShiftRight;
    return true;
  }
  return false;
}

bool MathParser::tryConsumeAddSubOp(EvalContext& ctx, Expr::BinaryOp& outOp) {
  skipSpaces(ctx);
  if (*ctx.p == '+') {
    ++ctx.p;
    outOp = Expr::BinaryOp::Add;
    return true;
  }
  if (*ctx.p == '-') {
    ++ctx.p;
    outOp = Expr::BinaryOp::Sub;
    return true;
  }
  return false;
}

bool MathParser::tryConsumeMulDivModOp(EvalContext& ctx, Expr::BinaryOp& outOp) {
  skipSpaces(ctx);
  if (*ctx.p == '*') {
    ++ctx.p;
    outOp = Expr::BinaryOp::Mul;
    return true;
  }
  if (*ctx.p == '/') {
    ++ctx.p;
    outOp = Expr::BinaryOp::Div;
    return true;
  }
  if (*ctx.p == '%') {
    ++ctx.p;
    outOp = Expr::BinaryOp::Modulo;
    return true;
  }
  if (*ctx.p == '(') {
    outOp = Expr::BinaryOp::Mul;
    return true;
  }
  return false;
}

bool MathParser::tryConsumeCompareOp(EvalContext& ctx, Expr::BinaryOp& outOp) {
  skipSpaces(ctx);
  if (std::strncmp(ctx.p, STR_LT_EQ, 2) == 0) {
    ctx.p += 2;
    outOp = Expr::BinaryOp::CmpLe;
    return true;
  }
  if (std::strncmp(ctx.p, STR_GT_EQ, 2) == 0) {
    ctx.p += 2;
    outOp = Expr::BinaryOp::CmpGe;
    return true;
  }
  if (std::strncmp(ctx.p, STR_EQ_EQ, 2) == 0) {
    ctx.p += 2;
    outOp = Expr::BinaryOp::CmpEq;
    return true;
  }
  if (std::strncmp(ctx.p, STR_NOT_EQ, 2) == 0 || std::strncmp(ctx.p, STR_LT_GT, 2) == 0) {
    ctx.p += 2;
    outOp = Expr::BinaryOp::CmpNe;
    return true;
  }
  if (*ctx.p == '=') {
    ++ctx.p;
    outOp = Expr::BinaryOp::CmpEq;
    return true;
  }
  if (*ctx.p == '<') {
    ++ctx.p;
    outOp = Expr::BinaryOp::CmpLt;
    return true;
  }
  if (*ctx.p == '>') {
    ++ctx.p;
    outOp = Expr::BinaryOp::CmpGt;
    return true;
  }
  return false;
}

bool MathParser::tryConsumeLogicalAndOp(EvalContext& ctx, Expr::BinaryOp& outOp) {
  if (!tryConsumeLogicalBinaryOperator(ctx, OperatorNameId::And, '&')) {
    return false;
  }
  outOp = Expr::BinaryOp::LogicalAnd;
  return true;
}

bool MathParser::tryConsumeLogicalOrOp(EvalContext& ctx, Expr::BinaryOp& outOp) {
  if (!tryConsumeLogicalBinaryOperator(ctx, OperatorNameId::Or, '|')) {
    return false;
  }
  outOp = Expr::BinaryOp::LogicalOr;
  return true;
}

bool MathParser::isTruthy(const EvalValue& v) {
  if (v.kind == ValueKind::Array) {
    return !v.arr.empty();
  }
  const EvalValue::ScalarValue& sv = v.scalarValue;
  double re = 0.0;
  double im = 0.0;
  scalarLoadCartesian(sv, re, im);
  if (std::isnan(re) || std::isnan(im)) {
    return false;
  }
  return std::fabs(re) > 0.0 || std::fabs(im) > 0.0;
}

std::string MathParser::trim(const std::string& s) {
  std::size_t l = 0;
  while (l < s.size() && std::isspace(static_cast<unsigned char>(s[l]))) {
    ++l;
  }
  std::size_t r = s.size();
  while (r > l && std::isspace(static_cast<unsigned char>(s[r - 1]))) {
    --r;
  }
  return s.substr(l, r - l);
}

bool MathParser::udfBodyIsEmptyTupleLiteral(const std::string& bodyExpr) {
  const std::string t = trim(bodyExpr);
  if (t.size() < 2 || t.front() != '(' || t.back() != ')') {
    return false;
  }
  for (std::size_t i = 1; i + 1 < t.size(); ++i) {
    if (!std::isspace(static_cast<unsigned char>(t[i]))) {
      return false;
    }
  }
  return true;
}

bool MathParser::nearlyInt(double v, long long& out) {
  if (!std::isfinite(v)) {
    return false;
  }
  if (std::fabs(v) > 9007199254740992.0) {
    return false;
  }
  const long double lv = static_cast<long double>(v);
  const long double r = std::round(lv);
  if (std::fabs(lv - r) > 1e-12L) {
    return false;
  }
  if (r < static_cast<long double>((std::numeric_limits<long long>::min)()) ||
      r > static_cast<long double>((std::numeric_limits<long long>::max)())) {
    return false;
  }
  out = static_cast<long long>(r);
  return true;
}

bool MathParser::tryGetExactSignedInt64FromScalarPolicy(
    const EvalValue::ScalarValue& sIn,
    long long& outI,
    ExactSignedInt64Policy policy) {
  if (sIn.scalarKind == ScalarKind::Time || scalarHasNonzeroImaginaryPart(sIn)) {
    return false;
  }
  EvalValue::ScalarValue s = sIn;
  if (policy == ExactSignedInt64Policy::NoUIntWrapWithRepairAndStorageKind) {
    scalarRepairExactMetadata(s);
  }
  if (s.hasExactInt64()) {
    outI = s.exactInt64;
    return true;
  }
  if (s.hasExactUInt64()) {
    if (policy == ExactSignedInt64Policy::AllowUIntBitReinterpret) {
      static_assert(sizeof(long long) == sizeof(std::uint64_t), "int64/uint64 size mismatch");
      std::uint64_t bits = s.exactUInt64;
      std::memcpy(&outI, &bits, sizeof(bits));
      return true;
    }
    if (s.exactUInt64 <= static_cast<std::uint64_t>((std::numeric_limits<long long>::max)())) {
      outI = static_cast<long long>(s.exactUInt64);
      return true;
    }
    return false;
  }
  if (policy == ExactSignedInt64Policy::NoUIntWrapWithRepairAndStorageKind) {
    if (s.scalarKind == ScalarKind::Int64) {
      outI = s.exactInt64;
      return true;
    }
    if (s.scalarKind == ScalarKind::UInt64 &&
        s.exactUInt64 <= static_cast<std::uint64_t>((std::numeric_limits<long long>::max)())) {
      outI = static_cast<long long>(s.exactUInt64);
      return true;
    }
  }
  return false;
}

bool MathParser::tryGetExactImagInt64Strict(const EvalValue::ScalarValue& s, long long& out) {
  if (s.hasImagExactInt64()) {
    out = s.imagExactInt64;
    return true;
  }
  return tryExtractExactInt64FromDoubleStrict(s.imag, out);
}

void MathParser::scalarRepairExactMetadata(EvalValue::ScalarValue& sv) {
  switch (sv.scalarKind) {
    case ScalarKind::Int64:
      sv.setExactInt64Valid(true);
      if (sv.exactInt64 >= 0) {
        sv.setExactUInt64Valid(true);
        sv.exactUInt64 = static_cast<std::uint64_t>(sv.exactInt64);
      } else if (sv.exactUInt64 != 0u) {
        sv.setExactUInt64Valid(true);
      } else {
        sv.setExactUInt64Valid(false);
        sv.exactUInt64 = 0;
      }
      break;
    case ScalarKind::UInt64:
      sv.setExactUInt64Valid(true);
      if (sv.exactUInt64 <= static_cast<std::uint64_t>((std::numeric_limits<long long>::max)())) {
        sv.setExactInt64Valid(true);
        sv.exactInt64 = static_cast<long long>(sv.exactUInt64);
      } else {
        sv.setExactInt64Valid(false);
        sv.exactInt64 = 0;
      }
      break;
    default:
      if (sv.hasRenderRational()) {
        sv.setExactInt64Valid(true);
        sv.setExactUInt64Valid(true);
      } else {
        sv.setExactInt64Valid(false);
        sv.setExactUInt64Valid(false);
      }
      break;
  }
#if SMARTMATH_FACTORINT
  if (!sv.hasRenderIntPower()) {
#endif
    if (!sv.hasImagExactInt64() && sv.imagExactInt64 != 0) {
      long long tIm = 0;
      if (tryExtractExactInt64FromDoubleStrict(sv.imag, tIm) && tIm == sv.imagExactInt64) {
        sv.setImagExactInt64Valid(true);
        if (sv.imagExactInt64 >= 0) {
          sv.setImagExactUInt64Valid(true);
          sv.imagExactUInt64 = static_cast<std::uint64_t>(sv.imagExactInt64);
        }
      }
    } else if (!sv.hasImagExactUInt64() && sv.imagExactUInt64 != 0u) {
      long long tIm = 0;
      if (tryExtractExactInt64FromDoubleStrict(sv.imag, tIm) &&
          static_cast<std::uint64_t>(tIm) == sv.imagExactUInt64) {
        sv.setImagExactUInt64Valid(true);
        if (sv.imagExactUInt64 <= static_cast<std::uint64_t>((std::numeric_limits<long long>::max)())) {
          sv.setImagExactInt64Valid(true);
          sv.imagExactInt64 = static_cast<long long>(sv.imagExactUInt64);
        }
      }
    }
#if SMARTMATH_FACTORINT
  }
#endif
}

bool MathParser::scalarHasExactIntegerPayload(const EvalValue::ScalarValue& sv) {
  if (sv.scalarKind == ScalarKind::Time) {
    return false;
  }
  return sv.hasExactInt64() || sv.hasExactUInt64() || sv.scalarKind == ScalarKind::Int64 ||
         sv.scalarKind == ScalarKind::UInt64;
}

bool MathParser::tryMulExactInt64Square(long long i, long long& outSq) {
  std::uint64_t u = 0;
  if (i >= 0) {
    u = static_cast<std::uint64_t>(i);
  } else if (i == (std::numeric_limits<long long>::min)()) {
    return false;
  } else {
    u = static_cast<std::uint64_t>(-i);
  }
  std::uint64_t sqU = 0;
  if (!tryMulUInt64Checked(u, u, sqU) ||
      sqU > static_cast<std::uint64_t>((std::numeric_limits<long long>::max)())) {
    return false;
  }
  outSq = static_cast<long long>(sqU);
  return true;
}

bool MathParser::tryGetExactSignedInt64NoUIntWrapScalarStrict(const EvalValue::ScalarValue& s, long long& outI) {
  return tryGetExactSignedInt64FromScalarPolicy(
      s, outI, ExactSignedInt64Policy::NoUIntWrapWithRepairAndStorageKind);
}

void MathParser::applySqrtScalarValue(const EvalValue::ScalarValue& sv, EvalValue& outV) {
  EvalValue::ScalarValue s = sv;
  scalarRepairExactMetadata(s);
  double xIn = 0.0;
  double ai = 0.0;
  scalarLoadCartesian(s, xIn, ai);
  const double r = std::sqrt(xIn);
  if (!scalarHasExactIntegerPayload(s)) {
    outV = makeScalar(r);
    return;
  }
  if (!std::isfinite(r)) {
    outV = makeScalar(r);
    return;
  }
  std::uint64_t inpU = 0;
  if (!tryGetExactNonNegativeUInt64FromScalar(s, inpU)) {
    outV = makeScalar(r);
    return;
  }
  const std::uint64_t n = static_cast<std::uint64_t>(std::round(r));
  std::uint64_t sq = 0;
  if (tryMulUInt64Checked(n, n, sq) && sq == inpU) {
    if (n <= static_cast<std::uint64_t>((std::numeric_limits<long long>::max)())) {
      outV = makeScalarInt(static_cast<long long>(n));
    } else {
      outV = makeScalarUInt(n);
    }
  } else {
    outV = makeScalar(r);
  }
}

bool MathParser::tryApplySqrExactScalar(const EvalValue::ScalarValue& sv, EvalValue& outV) {
  EvalValue::ScalarValue s = sv;
  scalarRepairExactMetadata(s);
  if (!scalarHasExactIntegerPayload(s)) {
    return false;
  }
  long long i = 0;
  if (tryGetExactSignedInt64NoUIntWrapScalarStrict(s, i)) {
    long long sq = 0;
    if (tryMulExactInt64Square(i, sq)) {
      outV = makeScalarInt(sq);
      return true;
    }
  } else {
    std::uint64_t u = 0;
    if (tryGetExactNonNegativeUInt64FromScalar(s, u)) {
      std::uint64_t sqU = 0;
      if (tryMulUInt64Checked(u, u, sqU)) {
        if (sqU <= static_cast<std::uint64_t>((std::numeric_limits<long long>::max)())) {
          outV = makeScalarInt(static_cast<long long>(sqU));
        } else {
          outV = makeScalarUInt(sqU);
        }
        return true;
      }
    }
  }
  return false;
}

bool MathParser::tryApplyHypotExactScalars(
    const EvalValue::ScalarValue& leftS,
    const EvalValue::ScalarValue& rightS,
    EvalValue& outV) {
  EvalValue::ScalarValue l = leftS;
  EvalValue::ScalarValue r = rightS;
  scalarRepairExactMetadata(l);
  scalarRepairExactMetadata(r);
  if (!scalarHasExactIntegerPayload(l) || !scalarHasExactIntegerPayload(r)) {
    return false;
  }
  long long la = 0;
  long long lb = 0;
  if (tryGetExactSignedInt64NoUIntWrapScalarStrict(l, la) &&
      tryGetExactSignedInt64NoUIntWrapScalarStrict(r, lb)) {
    long long aa = 0;
    long long bb = 0;
    long long sumSq = 0;
    if (tryMulExactInt64Square(la, aa) && tryMulExactInt64Square(lb, bb) && checkedAddLL(aa, bb, sumSq)) {
      outV = makeScalar(std::sqrt(static_cast<double>(sumSq)));
      return true;
    }
    return false;
  }
  std::uint64_t ua = 0;
  std::uint64_t ub = 0;
  if (tryGetExactNonNegativeUInt64FromScalar(l, ua) && tryGetExactNonNegativeUInt64FromScalar(r, ub)) {
    std::uint64_t aaU = 0;
    std::uint64_t bbU = 0;
    std::uint64_t sumSqU = 0;
    if (tryMulUInt64Checked(ua, ua, aaU) && tryMulUInt64Checked(ub, ub, bbU) &&
        tryAddUInt64Checked(aaU, bbU, sumSqU)) {
      outV = makeScalar(std::sqrt(static_cast<double>(sumSqU)));
      return true;
    }
  }
  return false;
}

namespace {

bool fractionalPowerIsOddUnitRoot(double p, long long& outN) {
  if (p <= 0.0 || p >= 1.0) {
    return false;
  }
  const double inv = 1.0 / p;
  if (inv < 2.0 || inv > 63.0) {
    return false;
  }
  outN = static_cast<long long>(std::llround(inv));
  if (outN < 2 || outN > 63) {
    return false;
  }
  if (std::fabs(inv - static_cast<double>(outN)) > 1e-6) {
    return false;
  }
  return (outN % 2) != 0;
}

bool fractionalPowerResolveRootDegree(double p, long long& outN) {
  if (fractionalPowerIsOddUnitRoot(p, outN)) {
    return true;
  }
  if (std::fabs(p - 0.5) < 1e-12) {
    outN = 2;
    return true;
  }
  return false;
}

bool tryPowFloatGuess(long long valueInt, double baseScalar, double p, double& outR) {
  if (valueInt >= 0 || p <= 0.0 || p >= 1.0) {
    outR = std::pow(baseScalar, p);
    return std::isfinite(outR);
  }
  long long nRoot = 0;
  if (!fractionalPowerIsOddUnitRoot(p, nRoot)) {
    outR = std::pow(baseScalar, p);
    return std::isfinite(outR);
  }
  const double magD = (valueInt == (std::numeric_limits<long long>::min)())
    ? static_cast<double>(1ull << 63)
    : static_cast<double>(-valueInt);
  const double rootMag = std::pow(magD, p);
  if (!std::isfinite(rootMag)) {
    return false;
  }
  outR = -rootMag;
  return true;
}

}  // namespace

bool MathParser::tryApplyPowExactScalarsByKind(
    ScalarExactKind kind,
    long long valueSigned,
    std::uint64_t valueUnsigned,
    const EvalValue::ScalarValue& leftS,
    const EvalValue::ScalarValue& rightS,
    EvalValue& outV) {
  const bool isSigned = (kind == ScalarExactKind::Signed);
  if (isSigned) {
    if (valueSigned == 1) {
      outV = makeScalarInt(1);
      return true;
    }
  } else if (valueUnsigned == 1u) {
    outV = makeScalarInt(1);
    return true;
  }
  const double p = rightS.scalar;
  if (!std::isfinite(p)) {
    return false;
  }
  if (p == 0.0) {
    outV = makeScalarInt(1);
    return true;
  }
  double r = 0.0;
  if (isSigned) {
    if (!tryPowFloatGuess(valueSigned, leftS.scalar, p, r)) {
      return false;
    }
  } else {
    r = std::pow(leftS.scalar, p);
    if (!std::isfinite(r)) {
      return false;
    }
  }
  if (p > 0.0 && p < 1.0) {
    long long nRoot = 0;
    if (fractionalPowerIsOddUnitRoot(p, nRoot)) {
      if (isSigned) {
        const long long rInt = static_cast<long long>(std::round(r));
        if (rInt >= 0 && valueSigned >= 0) {
          std::uint64_t reconU = 0;
          if (tryPowUInt64Checked(static_cast<std::uint64_t>(rInt), static_cast<std::uint64_t>(nRoot), reconU) &&
              static_cast<long long>(reconU) == valueSigned) {
            outV = makeScalarInt(rInt);
            return true;
          }
        } else {
          long long reconI = 0;
          if (tryPowInt64Checked(rInt, nRoot, reconI) && reconI == valueSigned) {
            outV = makeScalarInt(rInt);
            return true;
          }
        }
      } else {
        const std::uint64_t n = static_cast<std::uint64_t>(std::round(r));
        std::uint64_t sq = 0;
        if (tryPowUInt64Checked(n, static_cast<std::uint64_t>(nRoot), sq) && sq == valueUnsigned) {
          if (n <= static_cast<std::uint64_t>((std::numeric_limits<long long>::max)())) {
            outV = makeScalarInt(static_cast<long long>(n));
          } else {
            outV = makeScalarUInt(n);
          }
          return true;
        }
      }
    }
    return false;
  }
  if (p >= 1.0) {
    const long long nExp = static_cast<long long>(std::round(p));
    if (nExp >= 0 && nExp <= 63) {
      if (nExp == 0) {
        outV = makeScalarInt(1);
        return true;
      }
      if (isSigned) {
        const long long rInt = static_cast<long long>(std::round(r));
        long long recon = 0;
        if (tryPowInt64Checked(valueSigned, nExp, recon) && recon == rInt) {
          outV = makeScalarInt(rInt);
          return true;
        }
      } else {
        const std::uint64_t powResult = static_cast<std::uint64_t>(std::round(r));
        std::uint64_t recon = 0;
        if (tryPowUInt64Checked(valueUnsigned, static_cast<std::uint64_t>(nExp), recon) && recon == powResult) {
          if (powResult <= static_cast<std::uint64_t>((std::numeric_limits<long long>::max)())) {
            outV = makeScalarInt(static_cast<long long>(powResult));
          } else {
            outV = makeScalarUInt(powResult);
          }
          return true;
        }
      }
    }
  }
  return false;
}

bool MathParser::tryApplyPowExactScalars(
    const EvalValue::ScalarValue& leftS,
    const EvalValue::ScalarValue& rightS,
    EvalValue& outV) {
  if (scalarHasNonzeroImaginaryPart(leftS) || scalarHasNonzeroImaginaryPart(rightS)) {
    return false;
  }
  if (leftS.scalarKind == ScalarKind::Time) {
    return false;
  }
  long long valueInt = 0;
  if (tryGetExactSignedInt64NoUIntWrapScalarStrict(leftS, valueInt)) {
    if (valueInt == (std::numeric_limits<long long>::min)()) {
      return false;
    }
    const std::uint64_t mag = (valueInt >= 0) ? static_cast<std::uint64_t>(valueInt)
                                              : static_cast<std::uint64_t>(-valueInt);
    if (mag >= static_cast<std::uint64_t>(K_MAX_EXACT_INT_FROM_DOUBLE)) {
      return false;
    }
    return MathParser::tryApplyPowExactScalarsByKind(
        ScalarExactKind::Signed, valueInt, 0u, leftS, rightS, outV);
  }
  std::uint64_t inpU = 0;
  if (tryGetExactNonNegativeUInt64FromScalar(leftS, inpU)) {
    return MathParser::tryApplyPowExactScalarsByKind(
        ScalarExactKind::Unsigned, 0LL, inpU, leftS, rightS, outV);
  }
  return false;
}

bool MathParser::tryApplyRealScalarPowNegFractional(
    const EvalValue::ScalarValue& leftS,
    double p,
    EvalValue& outV) {
  long long nRoot = 0;
  if (!fractionalPowerIsOddUnitRoot(p, nRoot)) {
    return false;
  }
  double ar = 0.0;
  double ai = 0.0;
  scalarLoadCartesian(leftS, ar, ai);
  if (ai != 0.0 || ar >= 0.0) {
    return false;
  }
  long long valueInt = 0;
  if (tryGetExactSignedInt64NoUIntWrapScalarStrict(leftS, valueInt)) {
    if (valueInt != (std::numeric_limits<long long>::min)()) {
      double realR = 0.0;
      if (tryPowFloatGuess(valueInt, leftS.scalar, p, realR)) {
        outV = makeScalar(realR);
        return true;
      }
    }
  }
  const double rootMag = std::pow(-ar, p);
  if (!std::isfinite(rootMag)) {
    return false;
  }
  outV = makeScalar(-rootMag);
  return true;
}

bool MathParser::tryApplyScalarPowSpecialPaths(
    const EvalValue::ScalarValue& leftS,
    const EvalValue::ScalarValue& rightS,
    EvalValue& outV) {
  if (tryApplyPowExactScalars(leftS, rightS, outV)) {
    return true;
  }
  return tryApplyRealScalarPowNegFractional(leftS, rightS.scalar, outV);
}

static bool tryPowVerifyRootExactValue(long long valueInt, long long rootCand, long long rootDeg) {
  if (rootDeg < 2 || rootDeg > 63) {
    return false;
  }
  if (valueInt < 0 && (rootDeg % 2) == 0) {
    return false;
  }
  if (rootCand >= 0 && valueInt >= 0) {
    std::uint64_t reconU = 0;
    if (!tryPowUInt64Checked(static_cast<std::uint64_t>(rootCand), static_cast<std::uint64_t>(rootDeg), reconU)) {
      return false;
    }
    return static_cast<long long>(reconU) == valueInt;
  }
  long long reconI = 0;
  if (!tryPowInt64Checked(rootCand, rootDeg, reconI)) {
    return false;
  }
  return reconI == valueInt;
}

#if SMARTMATH_COMPLEX_NUMBERS
bool MathParser::tryRefinePowPrincipalToExactScalarResult(
    const EvalValue::ScalarValue& leftS,
    const EvalValue::ScalarValue& rightS,
    double powR,
    double powI,
    EvalValue& outV) {
  constexpr double kEps = 1e-12;
  double ar = 0.0;
  double ai = 0.0;
  double br = 0.0;
  double bi = 0.0;
  scalarLoadCartesian(leftS, ar, ai);
  scalarLoadCartesian(rightS, br, bi);
  if (std::fabs(ai) > kEps || std::fabs(bi) > kEps) {
    return false;
  }
  const double p = br;
  if (!std::isfinite(p) || !std::isfinite(powR) || !std::isfinite(powI)) {
    return false;
  }
  long long valueInt = 0;
  const bool hasSigned = tryGetExactSignedInt64NoUIntWrapScalarStrict(leftS, valueInt);
  std::uint64_t inpU = 0;
  const bool hasUInt = tryGetExactNonNegativeUInt64FromScalar(leftS, inpU);

  if (p > 0.0 && p < 1.0) {
    long long nRoot = 0;
    if (!fractionalPowerResolveRootDegree(p, nRoot)) {
      return false;
    }
    if (hasSigned && valueInt < 0) {
      if (std::fabs(powR) > kEps) {
        return false;
      }
      const long long rootCand = static_cast<long long>(std::round(powI));
      long long verifyBase = valueInt;
      if (nRoot == 2 && valueInt != (std::numeric_limits<long long>::min)()) {
        verifyBase = -valueInt;
      }
      if (!tryPowVerifyRootExactValue(verifyBase, rootCand, nRoot)) {
        if (nRoot != 2 || !tryPowVerifyRootExactValue(valueInt, rootCand, nRoot)) {
          return false;
        }
      }
      setPureImaginaryFromMagnitudeScalar(outV, makeScalarInt(rootCand).scalarValue);
      return true;
    }
    if (hasSigned && valueInt >= 0) {
      if (std::fabs(powI) > kEps) {
        return false;
      }
      const long long rootCand = static_cast<long long>(std::round(powR));
      if (!tryPowVerifyRootExactValue(valueInt, rootCand, nRoot)) {
        return false;
      }
      outV = makeScalarInt(rootCand);
      return true;
    }
    if (hasUInt) {
      if (std::fabs(powI) > kEps) {
        return false;
      }
      const std::uint64_t rootCand = static_cast<std::uint64_t>(std::round(powR));
      std::uint64_t recon = 0;
      if (!tryPowUInt64Checked(rootCand, static_cast<std::uint64_t>(nRoot), recon) || recon != inpU) {
        return false;
      }
      if (rootCand <= static_cast<std::uint64_t>((std::numeric_limits<long long>::max)())) {
        outV = makeScalarInt(static_cast<long long>(rootCand));
      } else {
        outV = makeScalarUInt(rootCand);
      }
      return true;
    }
    return false;
  }
  if (p >= 1.0) {
    const long long nExp = static_cast<long long>(std::round(p));
    if (nExp < 0 || nExp > 63 || std::fabs(p - static_cast<double>(nExp)) > 1e-6) {
      return false;
    }
    if (std::fabs(powI) > kEps) {
      return false;
    }
    const long long powResult = static_cast<long long>(std::round(powR));
    if (hasSigned) {
      long long recon = 0;
      if (!tryPowInt64Checked(valueInt, nExp, recon) || recon != powResult) {
        return false;
      }
      outV = makeScalarInt(powResult);
      return true;
    }
    if (hasUInt) {
      std::uint64_t recon = 0;
      if (!tryPowUInt64Checked(inpU, static_cast<std::uint64_t>(nExp), recon) ||
          recon != static_cast<std::uint64_t>(powResult)) {
        return false;
      }
      if (static_cast<std::uint64_t>(powResult) <=
          static_cast<std::uint64_t>((std::numeric_limits<long long>::max)())) {
        outV = makeScalarInt(powResult);
      } else {
        outV = makeScalarUInt(static_cast<std::uint64_t>(powResult));
      }
      return true;
    }
  }
  return false;
}

bool MathParser::tryVerifyComplexCartesianSquareExact(
    long long rootR, long long rootI, long long expR, long long expI) {
  long long rr = 0;
  long long ii = 0;
  long long ri = 0;
  long long sqR = 0;
  long long sqI = 0;
  if (!tryMulExactInt64Square(rootR, rr) || !tryMulExactInt64Square(rootI, ii)) {
    return false;
  }
  if (!checkedSubLL(rr, ii, sqR)) {
    return false;
  }
  if (!tryMulInt64Checked(rootR, rootI, ri)) {
    return false;
  }
  if (!tryAddInt64Checked(ri, ri, sqI)) {
    return false;
  }
  return sqR == expR && sqI == expI;
}

bool MathParser::tryRefineSqrtPrincipalToExactComplex(
    const EvalValue::ScalarValue& inS,
    double sqrtR,
    double sqrtI,
    EvalValue& outV) {
  if (!std::isfinite(sqrtR) || !std::isfinite(sqrtI)) {
    return false;
  }
  ExactCartesianComponent reC{};
  ExactCartesianComponent imC{};
  if (!tryExtractExactRealComponent(inS, reC) || !tryExtractExactImagComponent(inS, imC)) {
    return false;
  }
  long long arI = 0;
  long long aiI = 0;
  if (!tryExactCartesianComponentToInt64(reC, arI) || !tryExactCartesianComponentToInt64(imC, aiI)) {
    return false;
  }
  const long long rootRi = static_cast<long long>(std::round(sqrtR));
  const long long rootIi = static_cast<long long>(std::round(sqrtI));
  if (!tryVerifyComplexCartesianSquareExact(rootRi, rootIi, arI, aiI)) {
    return false;
  }
  outV = setScalarComplexFromEvalRealImagParts(makeScalarInt(rootRi), makeScalarInt(rootIi));
  return true;
}

MathParser::EvalValue MathParser::applySqrtComplexPrincipalUnary(const EvalValue::ScalarValue& inS) {
  double cr = 0.0;
  double ai = 0.0;
  scalarLoadCartesian(inS, cr, ai);
  if (!std::isfinite(cr) || !std::isfinite(ai)) {
    return makeScalarComplexFromDoubles(
        std::numeric_limits<double>::quiet_NaN(), std::numeric_limits<double>::quiet_NaN());
  }
  const double mag = std::hypot(cr, ai);
  if (mag == 0.0) {
    return makeScalarComplexFromDoubles(0.0, 0.0);
  }
  double sqrtR = 0.0;
  double sqrtI = 0.0;
  complexPrincipalSqrt(cr, ai, sqrtR, sqrtI);
  EvalValue out{};
  if (tryRefineSqrtPrincipalToExactComplex(inS, sqrtR, sqrtI, out)) {
    return out;
  }
  return makeScalarComplexFromDoubles(sqrtR, sqrtI);
}

void MathParser::applyComplexCaretPrincipalEval(
    const EvalValue::ScalarValue& lv,
    const EvalValue::ScalarValue& rv,
    EvalValue& outS) const {
  double ar = 0.0;
  double ai = 0.0;
  double br = 0.0;
  double bi = 0.0;
  scalarLoadCartesian(lv, ar, ai);
  scalarLoadCartesian(rv, br, bi);
  double powR = 0.0;
  double powI = 0.0;
  complexPowPrincipal(ar, ai, br, bi, powR, powI);
  if (tryRefinePowPrincipalToExactScalarResult(lv, rv, powR, powI, outS)) {
    return;
  }
  outS = makeScalarComplexFromDoubles(powR, powI);
}
#endif

MathParser::EvalValue MathParser::applyUnarySqrtEval(const EvalValue::ScalarValue& s) const {
#if SMARTMATH_COMPLEX_NUMBERS
  if (getSupportComplexNumbers()) {
    if (scalarHasNonzeroImaginaryPart(s)) {
      return applySqrtComplexPrincipalUnary(s);
    }
    const double x = s.scalar;
    if (std::isfinite(x) && x < 0.0) {
      EvalValue absV = applyAbsScalarValue(s);
      EvalValue magV;
      applySqrtScalarValue(absV.scalarValue, magV);
      EvalValue out{};
      setPureImaginaryFromMagnitudeScalar(out, magV.scalarValue);
      return out;
    }
  }
#endif
  EvalValue sqrtOut;
  applySqrtScalarValue(s, sqrtOut);
  return sqrtOut;
}

bool MathParser::tryGetExactSignedInt64FromScalar(const EvalValue::ScalarValue& s, long long& outI) {
  return tryGetExactSignedInt64FromScalarPolicy(
      s, outI, ExactSignedInt64Policy::AllowUIntBitReinterpret);
}

bool MathParser::tryGetExactSignedInt64NoUIntWrapFromScalar(const EvalValue::ScalarValue& s, long long& outI) {
  return tryGetExactSignedInt64FromScalarPolicy(s, outI, ExactSignedInt64Policy::NoUIntWrap);
}

bool MathParser::tryGetExactNonNegativeUInt64FromScalar(const EvalValue::ScalarValue& s, std::uint64_t& outU) {
  if (s.scalarKind == ScalarKind::Time) {
    return false;
  }
  if (scalarHasNonzeroImaginaryPart(s)) {
    return false;
  }
  EvalValue::ScalarValue sv = s;
  scalarRepairExactMetadata(sv);
  if (sv.hasExactUInt64() || sv.scalarKind == ScalarKind::UInt64) {
    outU = sv.exactUInt64;
    return true;
  }
  if (sv.hasExactInt64() && sv.exactInt64 >= 0) {
    outU = static_cast<std::uint64_t>(sv.exactInt64);
    return true;
  }
  if (sv.scalarKind == ScalarKind::Int64 && sv.exactInt64 >= 0) {
    outU = static_cast<std::uint64_t>(sv.exactInt64);
    return true;
  }
  return false;
}

bool MathParser::tryGetBothExactSignedInt64NoUIntWrapFromScalars(
    const EvalValue::ScalarValue& a,
    const EvalValue::ScalarValue& b,
    long long& outA,
    long long& outB) {
  return tryGetExactSignedInt64NoUIntWrapFromScalar(a, outA) &&
         tryGetExactSignedInt64NoUIntWrapFromScalar(b, outB);
}

bool MathParser::tryGetBothExactNonNegativeUInt64FromScalars(
    const EvalValue::ScalarValue& a,
    const EvalValue::ScalarValue& b,
    std::uint64_t& outA,
    std::uint64_t& outB) {
  return tryGetExactNonNegativeUInt64FromScalar(a, outA) &&
         tryGetExactNonNegativeUInt64FromScalar(b, outB);
}

bool MathParser::tryShiftLeftU64ExactOrMaybe(
    std::uint64_t aU,
    std::uint64_t bU,
    EvalValue& outV) {
  if (bU > 63u) return false;
  if (bU > 0u && aU > ((std::numeric_limits<std::uint64_t>::max)() >> bU)) {
    outV = makeScalarMaybeExact(std::ldexp(static_cast<double>(aU), static_cast<int>(bU)));
    return true;
  }
  outV = makeScalarUInt(aU << bU);
  return true;
}

bool MathParser::tryGetSignedInt64FromScalar(const EvalValue::ScalarValue& s, long long& outI) {
  if (s.scalarKind == ScalarKind::Time) {
    return false;
  }
  if (tryGetExactSignedInt64FromScalar(s, outI)) {
    return true;
  }
  return tryExtractExactInt64FromDoubleStrict(s.scalar, outI);
}

bool MathParser::isPureFloatingScalarPair(const EvalValue::ScalarValue& a, const EvalValue::ScalarValue& b) {
  return (a.scalarKind == ScalarKind::FloatingPoint) && (b.scalarKind == ScalarKind::FloatingPoint) &&
      !a.hasExactInt64() && !a.hasExactUInt64() && !b.hasExactInt64() && !b.hasExactUInt64() &&
      !scalarHasNonzeroImaginaryPart(a) && !scalarHasNonzeroImaginaryPart(b);
}

bool MathParser::parseUInt64FromDouble(double v, std::uint64_t& out) {
  if (!isWithinExactIntFromDoubleRange(v) || v < 0.0) {
    return false;
  }
  if (v > static_cast<double>(std::numeric_limits<std::uint64_t>::max())) {
    return false;
  }
  const std::uint64_t u = static_cast<std::uint64_t>(v);
  if (v != static_cast<double>(u)) {
    return false;
  }
  out = u;
  return true;
}

std::string MathParser::formatRationalParts(long long num, std::uint64_t den) {
  if (num == 0LL) {
    return "0";
  }
  if (den == 1u) {
    return std::to_string(num);
  }
  return std::to_string(num) + "/" + std::to_string(den);
}

bool MathParser::tryFormatRationalScalar(const EvalValue::ScalarValue& sv, std::string& outText) {
  if (!sv.hasRenderRational()) {
    return false;
  }
  const long long num = sv.exactInt64;
  const std::uint64_t den = sv.exactUInt64;
  if (den == 0u) {
    return false;
  }
  outText = formatRationalParts(num, den);
  return true;
}

#if SMARTMATH_FACTORINT
std::string formatIntPowerParts(long long baseV, std::uint64_t expV) {
  if (expV <= 1u) {
    return std::to_string(baseV);
  }
  return std::to_string(baseV) + "**" + std::to_string(expV);
}

bool MathParser::tryFormatIntPowerScalar(const EvalValue::ScalarValue& sv, std::string& outText) {
  if (!sv.hasRenderIntPower() || sv.imagExactUInt64 <= 1u) {
    return false;
  }
  outText = formatIntPowerParts(sv.imagExactInt64, sv.imagExactUInt64);
  return true;
}
#endif

#if SMARTMATH_COMPLEX_NUMBERS
bool MathParser::tryFormatComplexRationalScalar(const EvalValue::ScalarValue& sv, std::string& outText) {
  if (std::isnan(sv.scalar) || std::isnan(sv.imag)) {
    return false;
  }
  if (!sv.hasRenderRational() && !sv.hasImagRenderRational() && !sv.hasImagExactInt64()) {
    return false;
  }
  std::string rePart;
  if (sv.hasRenderRational()) {
    rePart = formatRationalParts(sv.exactInt64, sv.exactUInt64);
  } else if (sv.hasExactInt64()) {
    if (sv.exactInt64 != 0LL) {
      rePart = std::to_string(sv.exactInt64);
    }
  } else if (std::fabs(sv.scalar) >= RATIO_APPROX_EPS) {
    return false;
  }
  std::string imTail;
  if (sv.hasImagRenderRational()) {
    const std::string rp = formatRationalParts(sv.imagExactInt64, sv.imagExactUInt64);
    if (sv.imagExactUInt64 > 1u) {
      imTail = rp + "*i";
    } else if (sv.imagExactInt64 == 1LL) {
      imTail = STR_I;
    } else if (sv.imagExactInt64 == -1LL) {
      imTail = std::string("-") + STR_I;
    } else {
      imTail = rp + STR_I;
    }
  } else if (sv.hasImagExactInt64()) {
    const long long ni = sv.imagExactInt64;
    if (ni == 1LL) {
      imTail = STR_I;
    } else if (ni == -1LL) {
      imTail = std::string("-") + STR_I;
    } else if (ni != 0LL) {
      imTail = std::to_string(ni) + STR_I;
    }
  } else if (scalarHasNonzeroImaginaryPart(sv)) {
    return false;
  }
  if (imTail.empty()) {
    if (rePart.empty()) {
      rePart = "0";
    }
    outText = rePart;
    return true;
  }
  if (rePart.empty() || rePart == "0") {
    outText = imTail;
    return true;
  }
  if (!imTail.empty() && imTail[0] == '-') {
    outText = rePart + imTail;
  } else {
    outText = rePart + "+" + imTail;
  }
  return true;
}
#endif

std::string MathParser::formatScalar(const EvalValue& v, RenderBase base) const {
  const int baseCode = static_cast<int>(base);
  const double dval = v.scalarValue.scalar;
#if SMARTMATH_TIME_VALUES
  if (v.scalarValue.scalarKind == ScalarKind::Time) {
    return formatTimeCanonicalFromMs(timeTotalMsFromScalarValue(v.scalarValue));
  }
#endif
#if SMARTMATH_COMPLEX_NUMBERS
  if (getSupportComplexNumbers() && scalarHasNonzeroImaginaryPart(v.scalarValue)) {
    if (base == RenderBase::Dec) {
      std::string ratCx;
      if (tryFormatComplexRationalScalar(v.scalarValue, ratCx)) {
        return ratCx;
      }
    }
    if (base != RenderBase::Dec) {
      return formatComplexScalarWithRenderBase(v.scalarValue, base, v.hasRenderUnsigned());
    }
    return formatComplexScalarValue(v.scalarValue);
  }
#endif
  if (base == RenderBase::Dec) {
    std::string ratText;
    if (tryFormatRationalScalar(v.scalarValue, ratText)) {
      return ratText;
    }
#if SMARTMATH_FACTORINT
    std::string powText;
    if (tryFormatIntPowerScalar(v.scalarValue, powText)) {
      return powText;
    }
#endif
  }
  if (base == RenderBase::Dec) {
    if (v.scalarValue.hasExactInt64()) {
      return std::to_string(v.scalarValue.exactInt64);
    }
    if (v.scalarValue.hasExactUInt64()) {
      return std::to_string(v.scalarValue.exactUInt64);
    }
    const double p63 = std::ldexp(1.0, 63);
    if (v.scalarValue.hasDecScientificPow63High()) {
      if (dval == p63) {
        return STR_POS_POW63_DEC_TEXT;
      }
      if (dval == -p63) {
        return STR_NEG_POW63_DEC_TEXT;
      }
    }
    long long iv = 0;
    if (nearlyInt(dval, iv)) {
      // Avoid formatting tiny nonzero values as the integer "0".
      // (Complex formatting already preserves small nonzero magnitudes.)
      if (iv == 0LL && dval != 0.0) {
        return formatDoubleFast(dval);
      }
      return std::to_string(iv);
    }
    return formatDoubleFast(dval);
  }

  std::uint64_t u = 0;
  if (v.scalarValue.hasExactUInt64()) {
    u = v.scalarValue.exactUInt64;
  } else if (v.scalarValue.hasExactInt64() && v.scalarValue.exactInt64 >= 0) {
    u = static_cast<std::uint64_t>(v.scalarValue.exactInt64);
  } else if (v.scalarValue.hasExactInt64() && v.scalarValue.exactInt64 < 0) {
    const long long iv = v.scalarValue.exactInt64;
    return formatSignedMagnitudeForRenderBase(iv, baseCode, v.hasRenderUnsigned());
  } else if (!parseUInt64FromDouble(dval, u)) {
    return formatScalar(v, RenderBase::Dec);
  }

  return formatUnsignedForRenderBase(u, baseCode);
}

std::string MathParser::valueToString(const EvalValue& v, RenderBase forcedBase) const {
  if (v.kind == ValueKind::Scalar) {
    return formatScalar(v, forcedBase);
  }
  if (v.kind == ValueKind::InlineLambda) {
    return "";
  }
  std::string out = "(";
  for (std::size_t i = 0; i < v.arr.size(); ++i) {
    if (i) out += ",";
    EvalValue e = scalarFromScalarValue(v.arr[i]);
    e.setRenderUnsigned(v.hasRenderUnsigned());
    out += formatScalar(e, forcedBase);
  }
  out += ")";
  return out;
}

std::string MathParser::valueToString(const EvalValue& v) const {
  return valueToString(v, v.getRenderBase());
}

std::string MathParser::getError() const {
  return lastError_;
}

std::string MathParser::getResult() const {
  if (!lastError_.empty() || !hasResult_) {
    return "";
  }
  return valueToString(lastResult_);
}

std::string MathParser::getResultWithBase(RenderBase base) const {
  if (!lastError_.empty() || !hasResult_) {
    return "";
  }
  return valueToString(lastResult_, base);
}

std::string MathParser::getResultAsHex() const {
  return getResultWithBase(RenderBase::Hex);
}

std::string MathParser::getResultAsDec() const {
  return getResultWithBase(RenderBase::Dec);
}

std::string MathParser::getResultAsOct() const {
  return getResultWithBase(RenderBase::Oct);
}

std::string MathParser::getResultAsBin() const {
  return getResultWithBase(RenderBase::Bin);
}

MathParser::RawResult MathParser::getRawResult() const {
  if (!lastError_.empty() || !hasResult_) {
    return RawResult{};
  }
  return toRawResult(lastResult_);
}

void MathParser::setSupportComplexNumbers(bool enabled) {
#if SMARTMATH_COMPLEX_NUMBERS
  if (supportComplexNumbers_ == enabled) {
    return;
  }
  supportComplexNumbers_ = enabled;
  if (enabled) {
    setVariable(STR_I, makeImaginaryUnit());
  } else {
    variables_.erase(STR_I);
  }
#else
  (void)enabled;
  return;
#endif
}

bool MathParser::getSupportComplexNumbers() const {
#if SMARTMATH_COMPLEX_NUMBERS
  return supportComplexNumbers_;
#else
  return false;
#endif
}

void MathParser::setSupportTimeValues(bool enabled) {
#if SMARTMATH_TIME_VALUES
  if (supportTimeValues_ == enabled) {
    return;
  }
  supportTimeValues_ = enabled;
  if (enabled) {
    setVariable(STR_MILLISECOND, makeScalarTimeMs(1LL));
    setVariable(STR_SECOND, makeScalarTimeMs(1000LL));
    setVariable(STR_MINUTE, makeScalarTimeMs(60000LL));
    setVariable(STR_HOUR, makeScalarTimeMs(3600000LL));
    setVariable(STR_DAY, makeScalarTimeMs(86400000LL));
  } else {
    variables_.erase(STR_MILLISECOND);
    variables_.erase(STR_SECOND);
    variables_.erase(STR_MINUTE);
    variables_.erase(STR_HOUR);
    variables_.erase(STR_DAY);
  }
#else
  (void)enabled;
#endif
}

bool MathParser::getSupportTimeValues() const {
#if SMARTMATH_TIME_VALUES
  return supportTimeValues_;
#else
  return false;
#endif
}

void MathParser::setSupportLambdaFunctions(bool enabled) {
#if SMARTMATH_LAMBDA_FUNCTIONS
  supportLambdaFunctions_ = enabled;
#else
  (void)enabled;
#endif
}

bool MathParser::getSupportLambdaFunctions() const {
#if SMARTMATH_LAMBDA_FUNCTIONS
  return supportLambdaFunctions_;
#else
  return false;
#endif
}

void MathParser::addConst(const std::string& constName, long long intValue) {
  EvalValue v = makeScalarInt(intValue);
  setVariable(toLower(constName), v);
}

void MathParser::addConst(const std::string& constName, double dblValue) {
  EvalValue v = makeScalar(dblValue);
  setVariable(toLower(constName), v);
}

MathParser::EvalValue MathParser::makeScalar(double v) {
  EvalValue out;
  out.kind = ValueKind::Scalar;
  out.scalarValue.scalarKind = ScalarKind::FloatingPoint;
  out.scalarValue.scalar = v;
  out.scalarValue.setExactInt64Valid(false);
  out.scalarValue.exactInt64 = 0;
  out.scalarValue.setExactUInt64Valid(false);
  out.scalarValue.exactUInt64 = 0;
  out.scalarValue.setDecScientificPow63High(false);
  out.scalarValue.imag = 0.0;
  out.scalarValue.imagExactInt64 = 0;
  out.scalarValue.imagExactUInt64 = 0;
  out.scalarValue.setImagExactInt64Valid(false);
  out.scalarValue.setImagExactUInt64Valid(false);
  return out;
}

bool MathParser::scalarHasNonzeroImaginaryPart(const EvalValue::ScalarValue& s) {
  if (s.imag != 0.0) {
    return true;
  }
  return s.hasImagExactInt64() && s.imagExactInt64 != 0;
}

void MathParser::scalarClearImaginary(EvalValue::ScalarValue& s) {
  s.imag = 0.0;
  s.imagExactInt64 = 0;
  s.imagExactUInt64 = 0;
  s.setImagExactInt64Valid(false);
  s.setImagExactUInt64Valid(false);
}

double MathParser::scalarNumericReal(const EvalValue::ScalarValue& s) {
  EvalValue::ScalarValue sv = s;
  scalarRepairExactMetadata(sv);
  if (sv.hasRenderRational()) {
    if (std::isfinite(sv.scalar) && sv.scalar != 0.0) {
      return sv.scalar;
    }
    if (sv.exactUInt64 == 0u) {
      return sv.scalar;
    }
    return static_cast<double>(sv.exactInt64) / static_cast<double>(sv.exactUInt64);
  }
  if (sv.hasExactInt64()) {
    return static_cast<double>(sv.exactInt64);
  }
  if (sv.scalarKind == ScalarKind::Int64) {
    return static_cast<double>(sv.exactInt64);
  }
  if (sv.hasExactUInt64()) {
    return static_cast<double>(sv.exactUInt64);
  }
  if (sv.scalarKind == ScalarKind::UInt64) {
    return static_cast<double>(sv.exactUInt64);
  }
  return sv.scalar;
}

double MathParser::scalarNumericImag(const EvalValue::ScalarValue& s) {
  EvalValue::ScalarValue sv = s;
  scalarRepairExactMetadata(sv);
  if (sv.hasImagRenderRational()) {
    if (std::isfinite(sv.imag) && sv.imag != 0.0) {
      return sv.imag;
    }
    if (sv.imagExactUInt64 == 0u) {
      return sv.imag;
    }
    return static_cast<double>(sv.imagExactInt64) / static_cast<double>(sv.imagExactUInt64);
  }
  if (sv.hasImagExactInt64() || sv.hasImagExactUInt64()) {
    if (!imagExactMetadataMatchesFloat(sv)) {
      return sv.imag;
    }
  }
  if (sv.hasImagExactInt64()) {
    return static_cast<double>(sv.imagExactInt64);
  }
  if (sv.hasImagExactUInt64()) {
    return static_cast<double>(sv.imagExactUInt64);
  }
  return sv.imag;
}

void MathParser::scalarLoadCartesian(const EvalValue::ScalarValue& s, double& re, double& im) {
  re = scalarNumericReal(s);
  im = scalarNumericImag(s);
}

bool MathParser::imagExactMetadataMatchesFloat(const EvalValue::ScalarValue& sv) {
  long long t = 0;
  if (!tryExtractExactInt64FromDoubleStrict(sv.imag, t)) {
    return false;
  }
  if (sv.hasImagExactInt64Metadata() && t == sv.imagExactInt64) {
    return true;
  }
  if (sv.hasImagExactUInt64Metadata() && static_cast<std::uint64_t>(t) == sv.imagExactUInt64) {
    return true;
  }
  return false;
}

void MathParser::scalarClearCartesianRenderExact(EvalValue::ScalarValue& s) {
  s.setRenderRational(false);
  s.setImagRenderRational(false);
  s.setExactUInt64Valid(false);
  s.setImagExactUInt64Valid(false);
}

void MathParser::scalarApplyExactInt64Part(EvalValue::ScalarValue& s, bool imagPart, long long n) {
  if (imagPart) {
    s.imagExactInt64 = n;
    s.setImagExactInt64Valid(true);
    s.setImagExactUInt64Valid(false);
    s.setImagRenderRational(false);
  } else {
    s.exactInt64 = n;
    s.setExactInt64Valid(true);
    s.setExactUInt64Valid(false);
    s.setRenderRational(false);
  }
}

void MathParser::scalarApplyReducedRationalPart(EvalValue::ScalarValue& s, bool imagPart, long long num,
                                                std::uint64_t den) {
  long long n = num;
  std::uint64_t d = den;
  const long long g = gcdInt64(std::llabs(n), static_cast<long long>(d));
  if (g > 0) {
    n /= g;
    d = static_cast<std::uint64_t>(static_cast<long long>(d) / g);
  }
  if (imagPart) {
    s.imagExactInt64 = n;
    s.setImagExactInt64Valid(true);
    if (d == 1u) {
      s.setImagExactUInt64Valid(false);
      s.setImagRenderRational(false);
    } else {
      s.imagExactUInt64 = d;
      s.setImagExactUInt64Valid(true);
      s.setImagRenderRational(true);
    }
  } else {
    s.exactInt64 = n;
    s.setExactInt64Valid(true);
    if (d == 1u) {
      s.setExactUInt64Valid(false);
      s.setRenderRational(false);
    } else {
      s.exactUInt64 = d;
      s.setExactUInt64Valid(true);
      s.setRenderRational(true);
    }
  }
}

MathParser::RawResult::CartesianScalar MathParser::rawCartesianAssignReducedRational(long long num,
                                                                                    std::uint64_t den) {
  RawResult::CartesianScalar out;
  long long n = num;
  std::uint64_t d = den;
  const long long g = gcdInt64(std::llabs(n), static_cast<long long>(d));
  if (g > 0) {
    n /= g;
    d = static_cast<std::uint64_t>(static_cast<long long>(d) / g);
  }
  if (d == 1u) {
    out.kind = RawResult::ScalarKind::Int64;
    out.intValue = n;
  } else {
    out.kind = RawResult::ScalarKind::Rational;
    out.rational.numerator = n;
    out.rational.denominator = d;
  }
  return out;
}

#if SMARTMATH_COMPLEX_NUMBERS
MathParser::EvalValue MathParser::makeImaginaryUnit() {
  EvalValue out = makeScalarComplexFromDoubles(0.0, 1.0);
  out.scalarValue.scalarKind = ScalarKind::Int64;
  out.scalarValue.setExactInt64Valid(true);
  out.scalarValue.exactInt64 = 0;
  out.scalarValue.setExactUInt64Valid(true);
  out.scalarValue.exactUInt64 = 0;
  out.scalarValue.setImagExactInt64Valid(true);
  out.scalarValue.imagExactInt64 = 1;
  out.scalarValue.setImagExactUInt64Valid(true);
  out.scalarValue.imagExactUInt64 = 1u;
  out.scalarValue.scalar = 0.0;
  out.scalarValue.imag = 1.0;
  return out;
}

MathParser::EvalValue MathParser::makeScalarComplexFromDoubles(double re, double im) {
  if (std::isnan(re) && std::isnan(im)) {
    return makeScalar(std::numeric_limits<double>::quiet_NaN());
  }
  EvalValue out = makeScalar(re);
  out.scalarValue.setRenderRational(false);
  out.scalarValue.setImagRenderRational(false);
#if SMARTMATH_FACTORINT
  out.scalarValue.setRenderIntPower(false);
#endif
  out.scalarValue.imag = im;
  out.scalarValue.setImagExactInt64Valid(false);
  out.scalarValue.setImagExactUInt64Valid(false);
  out.scalarValue.imagExactInt64 = 0;
  out.scalarValue.imagExactUInt64 = 0;
  long long tri = 0;
  long long tii = 0;
  if (tryExtractExactInt64FromDoubleStrict(re, tri)) {
    out.scalarValue.scalarKind = ScalarKind::Int64;
    out.scalarValue.setExactInt64Valid(true);
    out.scalarValue.exactInt64 = tri;
    if (tri >= 0) {
      out.scalarValue.setExactUInt64Valid(true);
      out.scalarValue.exactUInt64 = static_cast<std::uint64_t>(tri);
    } else {
      out.scalarValue.setExactUInt64Valid(false);
      out.scalarValue.exactUInt64 = 0;
    }
  }
  if (tryExtractExactInt64FromDoubleStrict(im, tii)) {
    out.scalarValue.setImagExactInt64Valid(true);
    out.scalarValue.imagExactInt64 = tii;
    if (tii >= 0) {
      out.scalarValue.setImagExactUInt64Valid(true);
      out.scalarValue.imagExactUInt64 = static_cast<std::uint64_t>(tii);
    } else {
      out.scalarValue.setImagExactUInt64Valid(false);
      out.scalarValue.imagExactUInt64 = 0;
    }
  }
  if (!scalarHasNonzeroImaginaryPart(out.scalarValue)) {
    scalarClearImaginary(out.scalarValue);
  }
  return out;
}
#endif

void MathParser::exactCartesianComponentClear(ExactCartesianComponent& c) {
  c.hasInt = false;
  c.intV = 0;
  c.hasUInt = false;
  c.uintV = 0;
}

void MathParser::exactCartesianComponentAssignFromSignedInt64(ExactCartesianComponent& c, long long n) {
  exactCartesianComponentClear(c);
  c.hasInt = true;
  c.intV = n;
  c.hasUInt = true;
  c.uintV = static_cast<std::uint64_t>(n);
}

void MathParser::exactCartesianComponentAssignFromUInt64(ExactCartesianComponent& c, std::uint64_t u) {
  exactCartesianComponentClear(c);
  c.hasUInt = true;
  c.uintV = u;
  if (u <= static_cast<std::uint64_t>((std::numeric_limits<long long>::max)())) {
    c.hasInt = true;
    c.intV = static_cast<long long>(u);
  }
}

#if SMARTMATH_COMPLEX_NUMBERS
bool MathParser::complexNeedsPrincipalNegRealPow(double ar, double ai, double br, double bi) {
  if (ai != 0.0 || bi != 0.0) {
    return false;
  }
  if (ar >= 0.0) {
    return false;
  }
  if (!std::isfinite(ar) || !std::isfinite(br)) {
    return false;
  }
  if (std::fabs(br - std::trunc(br)) < 1e-12) {
    return false;
  }
  return true;
}

void MathParser::complexCartesianBinary(double ar, double ai, double br, double bi, char op, double& outR,
                                          double& outI) {
  switch (op) {
    case '+':
      outR = ar + br;
      outI = ai + bi;
      break;
    case '-':
      outR = ar - br;
      outI = ai - bi;
      break;
    case '*': {
      complexMultiply(ar, ai, br, bi, outR, outI);
      break;
    }
    case '/':
      complexDivide(ar, ai, br, bi, outR, outI);
      break;
    default:
      outR = std::numeric_limits<double>::quiet_NaN();
      outI = std::numeric_limits<double>::quiet_NaN();
      break;
  }
}
#endif

bool MathParser::tryExactCartesianComponentToInt64(const ExactCartesianComponent& c, long long& outI) {
  if (c.hasInt) {
    outI = c.intV;
    return true;
  }
  if (c.hasUInt && c.uintV <= static_cast<std::uint64_t>((std::numeric_limits<long long>::max)())) {
    outI = static_cast<long long>(c.uintV);
    return true;
  }
  if (!c.hasInt && !c.hasUInt) {
    outI = 0;
    return true;
  }
  return false;
}

bool MathParser::tryExtractExactCartesianComponent(bool isImag, const EvalValue::ScalarValue& sv,
                                                   ExactCartesianComponent& c) {
  if (!isImag) {
    if (sv.hasRenderRational()) {
      return false;
    }
    if (sv.hasExactInt64()) {
      exactCartesianComponentAssignFromSignedInt64(c, sv.exactInt64);
      if (sv.exactInt64 < 0 && sv.hasExactUInt64()) {
        c.hasUInt = true;
        c.uintV = sv.exactUInt64;
      }
      return true;
    }
    if (sv.hasExactUInt64()) {
      exactCartesianComponentAssignFromUInt64(c, sv.exactUInt64);
      return true;
    }
    long long t = 0;
    if (tryExtractExactInt64FromDoubleStrict(sv.scalar, t)) {
      exactCartesianComponentAssignFromSignedInt64(c, t);
      return true;
    }
    exactCartesianComponentClear(c);
    return false;
  }
#if SMARTMATH_COMPLEX_NUMBERS
  if (sv.hasImagRenderRational()) {
    return false;
  }
  if (sv.hasImagExactInt64()) {
    if (sv.imagExactInt64 == 0 && sv.imag != 0.0) {
      return false;
    }
    exactCartesianComponentAssignFromSignedInt64(c, sv.imagExactInt64);
    if (sv.imagExactInt64 < 0 && sv.hasImagExactUInt64()) {
      c.hasUInt = true;
      c.uintV = sv.imagExactUInt64;
    }
    return true;
  }
  if (sv.hasImagExactUInt64()) {
    if (sv.imagExactUInt64 == 0u && sv.imag != 0.0) {
      return false;
    }
    exactCartesianComponentAssignFromUInt64(c, sv.imagExactUInt64);
    return true;
  }
  if (imagExactMetadataMatchesFloat(sv)) {
    if (sv.hasImagExactInt64Metadata()) {
      exactCartesianComponentAssignFromSignedInt64(c, sv.imagExactInt64);
      if (sv.imagExactInt64 < 0 && sv.hasImagExactUInt64Metadata()) {
        c.hasUInt = true;
        c.uintV = sv.imagExactUInt64;
      }
      return true;
    }
    if (sv.hasImagExactUInt64Metadata()) {
      exactCartesianComponentAssignFromUInt64(c, sv.imagExactUInt64);
      return true;
    }
  }
  if (!scalarHasNonzeroImaginaryPart(sv)) {
    exactCartesianComponentClear(c);
    return true;
  }
  long long t = 0;
  if (tryExtractExactInt64FromDoubleStrict(sv.imag, t)) {
    if (t == 0 && sv.imag != 0.0) {
      return false;
    }
    exactCartesianComponentAssignFromSignedInt64(c, t);
    return true;
  }
  exactCartesianComponentClear(c);
  return false;
#else
  (void)sv;
  exactCartesianComponentClear(c);
  return false;
#endif
}

bool MathParser::tryExtractExactRealComponent(const EvalValue::ScalarValue& sv, ExactCartesianComponent& c) {
  return tryExtractExactCartesianComponent(false, sv, c);
}

#if SMARTMATH_COMPLEX_NUMBERS
bool MathParser::tryExtractExactImagComponent(const EvalValue::ScalarValue& sv, ExactCartesianComponent& c) {
  return tryExtractExactCartesianComponent(true, sv, c);
}
#endif

void MathParser::setScalarFromExactCartesianComponent(EvalValue& v, const ExactCartesianComponent& c) {
  v.kind = ValueKind::Scalar;
  if (c.hasInt) {
    v.scalarValue.scalarKind = ScalarKind::Int64;
    v.scalarValue.scalar = static_cast<double>(c.intV);
    v.scalarValue.setExactInt64Valid(true);
    v.scalarValue.exactInt64 = c.intV;
    if (c.intV >= 0) {
      v.scalarValue.setExactUInt64Valid(true);
      v.scalarValue.exactUInt64 = static_cast<std::uint64_t>(c.intV);
    } else if (c.hasUInt) {
      v.scalarValue.setExactUInt64Valid(true);
      v.scalarValue.exactUInt64 = c.uintV;
    } else {
      v.scalarValue.setExactUInt64Valid(false);
      v.scalarValue.exactUInt64 = 0;
    }
  } else if (c.hasUInt) {
    v.scalarValue.scalarKind = ScalarKind::UInt64;
    v.scalarValue.scalar = static_cast<double>(c.uintV);
    v.scalarValue.setExactUInt64Valid(true);
    v.scalarValue.exactUInt64 = c.uintV;
    if (c.uintV <= static_cast<std::uint64_t>((std::numeric_limits<long long>::max)())) {
      v.scalarValue.setExactInt64Valid(true);
      v.scalarValue.exactInt64 = static_cast<long long>(c.uintV);
    } else {
      v.scalarValue.setExactInt64Valid(false);
      v.scalarValue.exactInt64 = 0;
    }
  } else {
    v.scalarValue.scalarKind = ScalarKind::FloatingPoint;
    v.scalarValue.scalar = 0.0;
    v.scalarValue.setExactInt64Valid(false);
    v.scalarValue.setExactUInt64Valid(false);
  }
  v.scalarValue.setDecScientificPow63High(false);
  scalarClearImaginary(v.scalarValue);
}

#if SMARTMATH_COMPLEX_NUMBERS
void MathParser::scalarApplyExactImagFromCartesianComponent(EvalValue::ScalarValue& sv,
                                                          const ExactCartesianComponent& c) {
  sv.imag = 0.0;
  sv.imagExactInt64 = 0;
  sv.imagExactUInt64 = 0;
  sv.setImagExactInt64Valid(false);
  sv.setImagExactUInt64Valid(false);
  if (!c.hasInt && !c.hasUInt) {
    return;
  }
  if (c.hasInt) {
    sv.setImagExactInt64Valid(true);
    sv.imagExactInt64 = c.intV;
    sv.imag = static_cast<double>(c.intV);
    sv.setImagExactUInt64Valid(true);
    sv.imagExactUInt64 = c.hasUInt ? c.uintV : static_cast<std::uint64_t>(c.intV);
  } else {
    sv.setImagExactUInt64Valid(true);
    sv.imagExactUInt64 = c.uintV;
    sv.imag = static_cast<double>(c.uintV);
    if (c.uintV <= static_cast<std::uint64_t>((std::numeric_limits<long long>::max)())) {
      sv.setImagExactInt64Valid(true);
      sv.imagExactInt64 = static_cast<long long>(c.uintV);
    }
  }
}

void MathParser::setScalarComplexFromExactCartesian(EvalValue& v, const ExactCartesianComponent& re,
                                                    const ExactCartesianComponent& im) {
  setScalarFromExactCartesianComponent(v, re);
  v.scalarValue.imag = 0.0;
  scalarApplyExactImagFromCartesianComponent(v.scalarValue, im);
  if (!scalarHasNonzeroImaginaryPart(v.scalarValue)) {
    scalarClearImaginary(v.scalarValue);
  }
}

bool MathParser::tryApplyExactCartesianBinaryOp(bool isAdd, const ExactCartesianComponent& a,
                                                const ExactCartesianComponent& b,
                                                ExactCartesianComponent& out) {
  exactCartesianComponentClear(out);
  long long ai = 0;
  long long bi = 0;
  long long oi = 0;
  if (tryExactCartesianComponentToInt64(a, ai) && tryExactCartesianComponentToInt64(b, bi)) {
    if (isAdd) {
      if (!checkedAddLL(ai, bi, oi)) {
        return false;
      }
    } else if (!checkedSubLL(ai, bi, oi)) {
      return false;
    }
    exactCartesianComponentAssignFromSignedInt64(out, oi);
    return true;
  }
  if (isAdd && a.hasUInt && b.hasUInt) {
    std::uint64_t ou = 0;
    if (tryAddUInt64Checked(a.uintV, b.uintV, ou)) {
      exactCartesianComponentAssignFromUInt64(out, ou);
      return true;
    }
  }
  return false;
}

bool MathParser::tryAddExactCartesianComponents(const ExactCartesianComponent& a, const ExactCartesianComponent& b,
                                                ExactCartesianComponent& out) {
  return tryApplyExactCartesianBinaryOp(true, a, b, out);
}
#endif

bool MathParser::tryQuotExactInt64(long long num, long long den, long long& quo) {
  if (den == 0) {
    return false;
  }
  quo = num / den;
  return quo * den == num;
}

bool tryQuotExactUInt64(std::uint64_t num, std::uint64_t den, std::uint64_t& quo) {
  if (den == 0u) {
    return false;
  }
  quo = num / den;
  std::uint64_t prod = 0;
  if (!tryMulUInt64Checked(quo, den, prod)) {
    return false;
  }
  return prod == num;
}

namespace {

constexpr double kExactIntFloatEps = 1e-12;

bool relCloseExactIntFloat(double a, double b) {
  const double scale = std::max(1.0, std::fabs(b));
  return std::fabs(a - b) <= kExactIntFloatEps * scale;
}

bool tryRecoverReducedRationalFromFloatAbs(double fAbs, long long& outNum, long long& outDen) {
  constexpr long long kMaxB = 4096;
  double bestErr = 1e300;
  long long bestA = 0;
  long long bestB = 0;
  for (long long b = 1; b <= kMaxB; ++b) {
    long long a = static_cast<long long>(std::llround(fAbs * static_cast<double>(b)));
    if (a < 1) {
      continue;
    }
    const long long g = gcdInt64(a, b);
    a /= g;
    const long long bReduced = b / g;
    const double approx = static_cast<double>(a) / static_cast<double>(bReduced);
    const double err = std::fabs(fAbs - approx);
    if (err < bestErr) {
      bestErr = err;
      bestA = a;
      bestB = bReduced;
    }
  }
  if (bestA < 1 || bestB < 1) {
    return false;
  }
  const double scale = std::max(1.0, fAbs);
  if (bestErr > kExactIntFloatEps * scale) {
    return false;
  }
  outNum = bestA;
  outDen = bestB;
  return true;
}

}  // namespace

bool MathParser::passExactAbsFloatFactorGates(double f, bool strictAbsFAboveOne) {
  if (!std::isfinite(f) || f == 0.0) {
    return false;
  }
  if (strictAbsFAboveOne) {
    if (std::fabs(f) <= 1.0) {
      return false;
    }
  } else if (std::fabs(f) < 1.0) {
    return false;
  }
  long long tmp = 0;
  return !nearlyInt(f, tmp);
}

bool MathParser::getExactRoundedIntFromFloatResult(double r, bool boundAbsResultTo2_53, long long& outI) {
  const double rounded = std::round(r);
  if (!isWithinExactIntFromDoubleRange(rounded)) {
    return false;
  }
  const long long roundedI = static_cast<long long>(rounded);
  if (static_cast<double>(roundedI) != r || roundedI == 0) {
    return false;
  }
  if (boundAbsResultTo2_53 && std::fabs(static_cast<double>(roundedI)) > K_MAX_EXACT_INT_FROM_DOUBLE) {
    return false;
  }
  outI = roundedI;
  return true;
}

bool MathParser::verifyExactIntFloatOpResidue(
    ScalarExactKind kind,
    long long intSigned,
    std::uint64_t intUnsigned,
    double f,
    long long resultSigned,
    std::uint64_t resultUnsigned,
    bool isMultiply) {
  if (kind == ScalarExactKind::Signed) {
    if (isMultiply) {
      return relCloseExactIntFloat(
          static_cast<double>(resultSigned) - static_cast<double>(intSigned) * f, 0.0);
    }
    return relCloseExactIntFloat(
        static_cast<double>(intSigned) - static_cast<double>(resultSigned) * f, 0.0);
  }
  if (isMultiply) {
    return relCloseExactIntFloat(
        static_cast<double>(resultUnsigned) - static_cast<double>(intUnsigned) * f, 0.0);
  }
  return relCloseExactIntFloat(
      static_cast<double>(intUnsigned) - static_cast<double>(resultUnsigned) * f, 0.0);
}

bool MathParser::verifyExactIntFloatOpCrossMultiply(
    ScalarExactKind kind,
    long long intSigned,
    std::uint64_t intUnsigned,
    double floatV,
    long long resultSigned,
    std::uint64_t resultUnsigned,
    long long ratA,
    long long ratB,
    bool isMultiply) {
  if (kind == ScalarExactKind::Signed) {
    const long long signF = floatV < 0.0 ? -1 : 1;
    long long resultSignedScaled = 0;
    long long lhs = 0;
    long long rhs = 0;
    if (!checkedMulLL(resultSigned, signF, resultSignedScaled)) {
      return false;
    }
    if (isMultiply) {
      if (!checkedMulLL(resultSignedScaled, ratB, lhs) || !checkedMulLL(intSigned, ratA, rhs)) {
        return false;
      }
    } else if (!checkedMulLL(resultSignedScaled, ratA, lhs) || !checkedMulLL(intSigned, ratB, rhs)) {
      return false;
    }
    return lhs == rhs;
  }
  std::uint64_t lhs = 0;
  std::uint64_t rhs = 0;
  if (isMultiply) {
    if (!tryMulUInt64Checked(resultUnsigned, static_cast<std::uint64_t>(ratB), lhs) ||
        !tryMulUInt64Checked(intUnsigned, static_cast<std::uint64_t>(ratA), rhs)) {
      return false;
    }
  } else if (!tryMulUInt64Checked(resultUnsigned, static_cast<std::uint64_t>(ratA), lhs) ||
             !tryMulUInt64Checked(intUnsigned, static_cast<std::uint64_t>(ratB), rhs)) {
    return false;
  }
  return lhs == rhs;
}

bool MathParser::tryPromoteExactIntFromFloatOp(
    ScalarExactKind kind,
    long long intSigned,
    std::uint64_t intUnsigned,
    double f,
    double r,
    bool isMultiply,
    EvalValue& outV) {
  if (!passExactAbsFloatFactorGates(f, isMultiply)) {
    return false;
  }
  if (kind == ScalarExactKind::Signed) {
    if (std::fabs(static_cast<double>(intSigned)) > K_MAX_EXACT_INT_FROM_DOUBLE) {
      return false;
    }
  } else if (f < 0.0 || static_cast<double>(intUnsigned) > K_MAX_EXACT_INT_FROM_DOUBLE) {
    return false;
  }
  long long resultSigned = 0;
  if (!getExactRoundedIntFromFloatResult(r, isMultiply, resultSigned)) {
    return false;
  }
  std::uint64_t resultUnsigned = 0;
  if (kind == ScalarExactKind::Unsigned) {
    if (resultSigned < 0) {
      return false;
    }
    resultUnsigned = static_cast<std::uint64_t>(resultSigned);
    if (isMultiply && static_cast<double>(resultUnsigned) > K_MAX_EXACT_INT_FROM_DOUBLE) {
      return false;
    }
  }
  if (!verifyExactIntFloatOpResidue(
          kind, intSigned, intUnsigned, f, resultSigned, resultUnsigned, isMultiply)) {
    return false;
  }
  long long ratA = 0;
  long long ratB = 0;
  if (!tryRecoverReducedRationalFromFloatAbs(std::fabs(f), ratA, ratB)) {
    return false;
  }
  if (!verifyExactIntFloatOpCrossMultiply(
          kind, intSigned, intUnsigned, f, resultSigned, resultUnsigned, ratA, ratB, isMultiply)) {
    return false;
  }
  if (kind == ScalarExactKind::Signed) {
    outV = makeScalarInt(resultSigned);
  } else {
    outV = makeScalarUInt(resultUnsigned);
  }
  return true;
}

bool MathParser::tryPromoteExactDivisionByFloatDivisorSigned(
    long long intI, double f, double r, EvalValue& outV) {
  return tryPromoteExactIntFromFloatOp(ScalarExactKind::Signed, intI, 0u, f, r, false, outV);
}

bool MathParser::tryPromoteExactDivisionByFloatDivisorUnsigned(
    std::uint64_t intU, double f, double r, EvalValue& outV) {
  return tryPromoteExactIntFromFloatOp(ScalarExactKind::Unsigned, 0LL, intU, f, r, false, outV);
}

bool MathParser::tryPromoteExactMultiplicationByFloatFactorSigned(
    long long intI, double f, double r, EvalValue& outV) {
  return tryPromoteExactIntFromFloatOp(ScalarExactKind::Signed, intI, 0u, f, r, true, outV);
}

bool MathParser::tryPromoteExactMultiplicationByFloatFactorUnsigned(
    std::uint64_t intU, double f, double r, EvalValue& outV) {
  return tryPromoteExactIntFromFloatOp(ScalarExactKind::Unsigned, 0LL, intU, f, r, true, outV);
}

bool MathParser::tryApplyExactIntFloatOpFromFloatSide(
    ScalarExactKind kind,
    bool isMultiply,
    long long intSigned,
    std::uint64_t intUnsigned,
    double f,
    double r,
    EvalValue& outV) {
  if (!std::isfinite(f) || f == 0.0) {
    return false;
  }
  constexpr double kEps = 1e-12;
  long long tmp = 0;

  if (kind == ScalarExactKind::Unsigned) {
    if (isMultiply) {
      if (f <= 0.0) {
        return false;
      }
      if (f < 1.0 && !nearlyInt(f, tmp)) {
        const double inv = 1.0 / f;
        const double invRound = std::round(inv);
        if (invRound > 1.0 &&
            invRound <= static_cast<double>((std::numeric_limits<std::uint64_t>::max)()) &&
            std::fabs(inv - invRound) <= kEps * std::max(1.0, std::fabs(inv))) {
          const std::uint64_t nU = static_cast<std::uint64_t>(invRound);
          if (nU != 0u && (intUnsigned % nU) == 0u) {
            outV = makeScalarUInt(intUnsigned / nU);
            return true;
          }
        }
      }
      return tryPromoteExactMultiplicationByFloatFactorUnsigned(intUnsigned, f, r, outV);
    }
    if (std::fabs(f) < 1.0 && !nearlyInt(f, tmp)) {
      const double inv = 1.0 / f;
      const double invRound = std::round(inv);
      if (std::fabs(invRound) > 1.0 &&
          std::fabs(invRound) <= static_cast<double>((std::numeric_limits<std::uint64_t>::max)()) &&
          std::fabs(inv - invRound) <= kEps * std::max(1.0, std::fabs(inv))) {
        if (invRound > 0.0) {
          const std::uint64_t nU = static_cast<std::uint64_t>(invRound);
          std::uint64_t qU = 0;
          if (tryMulUInt64Checked(intUnsigned, nU, qU)) {
            outV = makeScalarUInt(qU);
            return true;
          }
        } else if (invRound < 0.0) {
          const std::uint64_t nU = static_cast<std::uint64_t>(-invRound);
          std::uint64_t qU = 0;
          if (tryMulUInt64Checked(intUnsigned, nU, qU) &&
              qU <= static_cast<std::uint64_t>((std::numeric_limits<long long>::max)())) {
            outV = makeScalarInt(-static_cast<long long>(qU));
            return true;
          }
        }
      }
    }
    return tryPromoteExactDivisionByFloatDivisorUnsigned(intUnsigned, f, r, outV);
  }

  if (isMultiply) {
    if (std::fabs(f) < 1.0) {
      if (nearlyInt(f, tmp)) {
        return false;
      }
      const double inv = 1.0 / f;
      const double invAbs = std::fabs(inv);
      const double invRound = std::round(invAbs);
      if (invRound <= 1.0 ||
          invRound > static_cast<double>((std::numeric_limits<long long>::max)()) ||
          std::fabs(invAbs - invRound) > kEps * std::max(1.0, invAbs)) {
        return false;
      }
      const long long nI = static_cast<long long>(invRound);
      if (nI <= 1 || (intSigned % nI) != 0) {
        return false;
      }
      long long qI = intSigned / nI;
      if (f < 0.0) {
        if (qI == (std::numeric_limits<long long>::min)()) {
          return false;
        }
        qI = -qI;
      }
      outV = makeScalarInt(qI);
      return true;
    }
    return tryPromoteExactMultiplicationByFloatFactorSigned(intSigned, f, r, outV);
  }

  if (std::fabs(f) < 1.0 && !nearlyInt(f, tmp)) {
    const double inv = 1.0 / f;
    const double invRound = std::round(inv);
    if (std::fabs(invRound) > 1.0 &&
        std::fabs(invRound) <= static_cast<double>((std::numeric_limits<long long>::max)()) &&
        std::fabs(inv - invRound) <= kEps * std::max(1.0, std::fabs(inv))) {
      const long long nI = static_cast<long long>(invRound);
      long long qI = 0;
      if (checkedMulLL(intSigned, nI, qI)) {
        outV = makeScalarInt(qI);
        return true;
      }
      if (intSigned >= 0 && nI > 0) {
        const std::uint64_t nU = static_cast<std::uint64_t>(nI);
        std::uint64_t qU = 0;
        if (tryMulUInt64Checked(static_cast<std::uint64_t>(intSigned), nU, qU)) {
          outV = makeScalarUInt(qU);
          return true;
        }
      }
    }
  }
  return tryPromoteExactDivisionByFloatDivisorSigned(intSigned, f, r, outV);
}

bool MathParser::tryApplyExactIntegerDivisionFromQuotient(
    const EvalValue::ScalarValue& leftS,
    const EvalValue::ScalarValue& rightS,
    double r,
    EvalValue& outV) const {
  if (!std::isfinite(r)) {
    return false;
  }
  std::uint64_t aU = 0;
  std::uint64_t bU = 0;
  if (tryGetExactNonNegativeUInt64FromScalar(leftS, aU) && tryGetExactNonNegativeUInt64FromScalar(rightS, bU)) {
    if (bU == 0u) {
      return false;
    }
    const double rq = std::round(r);
    if (!isWithinExactIntFromDoubleRange(rq) || rq < 0.0) {
      return false;
    }
    const std::uint64_t qU = static_cast<std::uint64_t>(rq);
    if (static_cast<double>(qU) != rq) {
      return false;
    }
    std::uint64_t prodU = 0;
    if (!tryMulUInt64Checked(qU, bU, prodU) || prodU != aU) {
      return false;
    }
    outV = makeScalarUInt(qU);
    return true;
  }
  long long aI = 0;
  long long bI = 0;
  if (tryGetExactSignedInt64NoUIntWrapFromScalar(leftS, aI) &&
      tryGetExactSignedInt64NoUIntWrapFromScalar(rightS, bI)) {
    if (bI == 0) {
      return false;
    }
    const double rq = std::round(r);
    if (!isWithinExactIntFromDoubleRange(rq)) {
      return false;
    }
    const long long qI = static_cast<long long>(rq);
    if (static_cast<double>(qI) != rq) {
      return false;
    }
    long long prodI = 0;
    if (!checkedMulLL(qI, bI, prodI) || prodI != aI) {
      return false;
    }
    outV = makeScalarInt(qI);
    return true;
  }

  std::uint64_t intU = 0;
  if (tryGetExactNonNegativeUInt64FromScalar(leftS, intU) && !scalarHasExactIntegerPayload(rightS)) {
    if (tryApplyExactIntFloatOpFromFloatSide(ScalarExactKind::Unsigned, false, 0LL, intU, rightS.scalar, r,
                                             outV)) {
      return true;
    }
  }

  long long intI = 0;
  if (tryGetExactSignedInt64NoUIntWrapFromScalar(leftS, intI) && !scalarHasExactIntegerPayload(rightS)) {
    if (tryApplyExactIntFloatOpFromFloatSide(ScalarExactKind::Signed, false, intI, 0u, rightS.scalar, r, outV)) {
      return true;
    }
  }

  return false;
}

bool MathParser::tryApplyExactIntegerMultiplicationFromProduct(
    const EvalValue::ScalarValue& leftS,
    const EvalValue::ScalarValue& rightS,
    double r,
    EvalValue& outV) const {
  if (!std::isfinite(r)) {
    return false;
  }
  std::uint64_t intU = 0;
  if (tryGetExactNonNegativeUInt64FromScalar(leftS, intU) && !scalarHasExactIntegerPayload(rightS)) {
    if (tryApplyExactIntFloatOpFromFloatSide(ScalarExactKind::Unsigned, true, 0LL, intU, rightS.scalar, r, outV)) {
      return true;
    }
  } else if (tryGetExactNonNegativeUInt64FromScalar(rightS, intU) && !scalarHasExactIntegerPayload(leftS)) {
    if (tryApplyExactIntFloatOpFromFloatSide(ScalarExactKind::Unsigned, true, 0LL, intU, leftS.scalar, r, outV)) {
      return true;
    }
  }

  long long intI = 0;
  if (tryGetExactSignedInt64NoUIntWrapFromScalar(leftS, intI) && !scalarHasExactIntegerPayload(rightS)) {
    if (tryApplyExactIntFloatOpFromFloatSide(ScalarExactKind::Signed, true, intI, 0u, rightS.scalar, r, outV)) {
      return true;
    }
  } else if (tryGetExactSignedInt64NoUIntWrapFromScalar(rightS, intI) && !scalarHasExactIntegerPayload(leftS)) {
    if (tryApplyExactIntFloatOpFromFloatSide(ScalarExactKind::Signed, true, intI, 0u, leftS.scalar, r, outV)) {
      return true;
    }
  }

  return false;
}

#if SMARTMATH_COMPLEX_NUMBERS
bool MathParser::trySubExactCartesianComponents(const ExactCartesianComponent& a, const ExactCartesianComponent& b,
                                                ExactCartesianComponent& out) {
  return tryApplyExactCartesianBinaryOp(false, a, b, out);
}

bool MathParser::tryApplyExactComplexCartesianBinary(const EvalValue::ScalarValue& leftS,
                                                   const EvalValue::ScalarValue& rightS, char op,
                                                   EvalValue& outV) {
  if (leftS.hasRenderRational() || leftS.hasImagRenderRational() || rightS.hasRenderRational() ||
      rightS.hasImagRenderRational()) {
    return false;
  }
  ExactCartesianComponent lRe{};
  ExactCartesianComponent lIm{};
  ExactCartesianComponent rRe{};
  ExactCartesianComponent rIm{};
  ExactCartesianComponent oRe{};
  ExactCartesianComponent oIm{};
  if (!tryExtractExactCartesianComponent(false, leftS, lRe) ||
      !tryExtractExactCartesianComponent(true, leftS, lIm) ||
      !tryExtractExactCartesianComponent(false, rightS, rRe) ||
      !tryExtractExactCartesianComponent(true, rightS, rIm)) {
    return false;
  }
  if (op == '+') {
    if (!tryApplyExactCartesianBinaryOp(true, lRe, rRe, oRe) ||
        !tryApplyExactCartesianBinaryOp(true, lIm, rIm, oIm)) {
      return false;
    }
    setScalarComplexFromExactCartesian(outV, oRe, oIm);
    return true;
  }
  if (op == '-') {
    if (!tryApplyExactCartesianBinaryOp(false, lRe, rRe, oRe) ||
        !tryApplyExactCartesianBinaryOp(false, lIm, rIm, oIm)) {
      return false;
    }
    setScalarComplexFromExactCartesian(outV, oRe, oIm);
    return true;
  }
  if (op == '*') {
    long long lar = 0;
    long long lai = 0;
    long long lbr = 0;
    long long lbi = 0;
    long long p1 = 0;
    long long p2 = 0;
    long long p3 = 0;
    long long p4 = 0;
    long long oreI = 0;
    long long oimI = 0;
    if (!tryExactCartesianComponentToInt64(lRe, lar) || !tryExactCartesianComponentToInt64(lIm, lai) ||
        !tryExactCartesianComponentToInt64(rRe, lbr) || !tryExactCartesianComponentToInt64(rIm, lbi) ||
        !checkedMulLL(lar, lbr, p1) || !checkedMulLL(lai, lbi, p2) || !checkedMulLL(lar, lbi, p3) ||
        !checkedMulLL(lai, lbr, p4) || !checkedSubLL(p1, p2, oreI) || !checkedAddLL(p3, p4, oimI)) {
      return false;
    }
    exactCartesianComponentAssignFromSignedInt64(oRe, oreI);
    exactCartesianComponentAssignFromSignedInt64(oIm, oimI);
    setScalarComplexFromExactCartesian(outV, oRe, oIm);
    return true;
  }
  if (op == '/') {
    long long lar = 0;
    long long lai = 0;
    long long lbr = 0;
    long long lbi = 0;
    long long p1 = 0;
    long long p2 = 0;
    long long p3 = 0;
    long long p4 = 0;
    long long p5 = 0;
    long long p6 = 0;
    long long numRe = 0;
    long long numIm = 0;
    long long denom = 0;
    long long qRe = 0;
    long long qIm = 0;
    if (!tryExactCartesianComponentToInt64(lRe, lar) || !tryExactCartesianComponentToInt64(lIm, lai) ||
        !tryExactCartesianComponentToInt64(rRe, lbr) || !tryExactCartesianComponentToInt64(rIm, lbi) ||
        !checkedMulLL(lar, lbr, p1) || !checkedMulLL(lai, lbi, p2) || !checkedAddLL(p1, p2, numRe) ||
        !checkedMulLL(lbr, lbr, p3) || !checkedMulLL(lbi, lbi, p4) || !checkedAddLL(p3, p4, denom) ||
        denom == 0 || !checkedMulLL(lai, lbr, p5) || !checkedMulLL(lar, lbi, p6) ||
        !checkedSubLL(p5, p6, numIm) || !tryQuotExactInt64(numRe, denom, qRe) ||
        !tryQuotExactInt64(numIm, denom, qIm)) {
      return false;
    }
    exactCartesianComponentAssignFromSignedInt64(oRe, qRe);
    exactCartesianComponentAssignFromSignedInt64(oIm, qIm);
    setScalarComplexFromExactCartesian(outV, oRe, oIm);
    return true;
  }
  return false;
}

MathParser::EvalValue MathParser::setScalarComplexFromEvalRealImagParts(const EvalValue& rePart,
                                                                        const EvalValue& imPart) {
  ExactCartesianComponent reC{};
  ExactCartesianComponent imC{};
  if (tryExtractExactRealComponent(rePart.scalarValue, reC) &&
      tryExtractExactRealComponent(imPart.scalarValue, imC)) {
    EvalValue out{};
    setScalarComplexFromExactCartesian(out, reC, imC);
    return out;
  }
  return makeScalarComplexFromDoubles(rePart.scalarValue.scalar, imPart.scalarValue.scalar);
}
#endif

void MathParser::setPureImaginaryFromMagnitudeScalar(EvalValue& outV, const EvalValue::ScalarValue& magSv) {
  // Copy first: callers may pass outV.scalarValue (e.g. tight-imag suffix on a literal).
  const EvalValue::ScalarValue mag = magSv;
  outV = makeScalarInt(0);
  if (scalarHasExactIntegerPayload(mag)) {
    if (mag.hasExactInt64()) {
      outV.scalarValue.setImagExactInt64Valid(true);
      outV.scalarValue.imagExactInt64 = mag.exactInt64;
      if (mag.exactInt64 >= 0) {
        outV.scalarValue.setImagExactUInt64Valid(true);
        outV.scalarValue.imagExactUInt64 = static_cast<std::uint64_t>(mag.exactInt64);
      } else {
        outV.scalarValue.setImagExactUInt64Valid(false);
        outV.scalarValue.imagExactUInt64 = 0;
      }
      outV.scalarValue.imag = static_cast<double>(mag.exactInt64);
    } else if (mag.hasExactUInt64()) {
      outV.scalarValue.setImagExactUInt64Valid(true);
      outV.scalarValue.imagExactUInt64 = mag.exactUInt64;
      outV.scalarValue.imag = static_cast<double>(mag.exactUInt64);
      if (mag.exactUInt64 <= static_cast<std::uint64_t>((std::numeric_limits<long long>::max)())) {
        outV.scalarValue.setImagExactInt64Valid(true);
        outV.scalarValue.imagExactInt64 = static_cast<long long>(mag.exactUInt64);
      } else {
        outV.scalarValue.setImagExactInt64Valid(false);
        outV.scalarValue.imagExactInt64 = 0;
      }
    } else {
      outV.scalarValue.imag = mag.scalar;
      outV.scalarValue.setImagExactInt64Valid(false);
      outV.scalarValue.setImagExactUInt64Valid(false);
    }
  } else {
    outV.scalarValue.imag = mag.scalar;
    outV.scalarValue.setImagExactInt64Valid(false);
    outV.scalarValue.setImagExactUInt64Valid(false);
    outV.scalarValue.imagExactInt64 = 0;
    outV.scalarValue.imagExactUInt64 = 0;
  }
  scalarRepairExactMetadata(outV.scalarValue);
}

bool MathParser::tryNegateExactCartesianComponent(const ExactCartesianComponent& c, ExactCartesianComponent& outC) {
  long long i = 0;
  if (tryExactCartesianComponentToInt64(c, i)) {
    if (i == (std::numeric_limits<long long>::min)()) {
      return false;
    }
    exactCartesianComponentAssignFromSignedInt64(outC, -i);
    return true;
  }
  if (c.hasUInt) {
    return false;
  }
  exactCartesianComponentAssignFromSignedInt64(outC, 0);
  return true;
}

#if SMARTMATH_COMPLEX_NUMBERS
bool MathParser::tryNegateExactComplexScalar(const EvalValue::ScalarValue& sv, EvalValue& out) {
  if (!scalarHasNonzeroImaginaryPart(sv)) {
    return false;
  }
  ExactCartesianComponent lRe{};
  ExactCartesianComponent lIm{};
  ExactCartesianComponent oRe{};
  ExactCartesianComponent oIm{};
  if (!tryExtractExactRealComponent(sv, lRe) || !tryExtractExactImagComponent(sv, lIm) ||
      !tryNegateExactCartesianComponent(lRe, oRe) || !tryNegateExactCartesianComponent(lIm, oIm)) {
    return false;
  }
  setScalarComplexFromExactCartesian(out, oRe, oIm);
  return true;
}

bool MathParser::tryFoldExactComplexCartesian(const std::vector<EvalValue>& args, char op, EvalValue& out) {
  bool gotAny = false;
  EvalValue acc{};
  const auto processScalar = [&](const EvalValue::ScalarValue& s) -> bool {
    const EvalValue item = scalarFromScalarValue(s);
    if (!gotAny) {
      acc = item;
      gotAny = true;
      return true;
    }
    EvalValue next{};
    if (!tryApplyExactComplexCartesianBinary(acc.scalarValue, item.scalarValue, op, next)) {
      return false;
    }
    acc = next;
    return true;
  };
  for (const auto& a : args) {
    if (a.kind == ValueKind::Scalar) {
      if (!processScalar(a.scalarValue)) {
        return false;
      }
    } else {
      for (const auto& item : a.arr) {
        if (!processScalar(item)) {
          return false;
        }
      }
    }
  }
  if (!gotAny) {
    return false;
  }
  out = acc;
  return true;
}

bool MathParser::tryAvgExactComplexFromSum(const EvalValue& sumV, std::size_t itemCount, EvalValue& out) {
  if (itemCount == 0) {
    return false;
  }
  const long long n = static_cast<long long>(itemCount);
  ExactCartesianComponent lRe{};
  ExactCartesianComponent lIm{};
  ExactCartesianComponent oRe{};
  ExactCartesianComponent oIm{};
  long long reI = 0;
  long long imI = 0;
  long long qRe = 0;
  long long qIm = 0;
  if (!tryExtractExactRealComponent(sumV.scalarValue, lRe) || !tryExtractExactImagComponent(sumV.scalarValue, lIm) ||
      !tryExactCartesianComponentToInt64(lRe, reI) || !tryExactCartesianComponentToInt64(lIm, imI) ||
      !tryQuotExactInt64(reI, n, qRe) || !tryQuotExactInt64(imI, n, qIm)) {
    return false;
  }
  exactCartesianComponentAssignFromSignedInt64(oRe, qRe);
  exactCartesianComponentAssignFromSignedInt64(oIm, qIm);
  setScalarComplexFromExactCartesian(out, oRe, oIm);
  return true;
}

bool MathParser::tryApplyComplexBinaryScalars(
    const EvalValue::ScalarValue& lv, const EvalValue::ScalarValue& rv, char op, EvalValue& outS) const {
  if (!getSupportComplexNumbers()) {
    return false;
  }
  if (op != '+' && op != '-' && op != '*' && op != '/' && op != '^') {
    return false;
  }
  double ar = 0.0;
  double ai = 0.0;
  double br = 0.0;
  double bi = 0.0;
  scalarLoadCartesian(lv, ar, ai);
  scalarLoadCartesian(rv, br, bi);
  const bool lIm = scalarHasNonzeroImaginaryPart(lv);
  const bool rIm = scalarHasNonzeroImaginaryPart(rv);
  if (op == '^') {
    EvalValue powSpecial;
    if (tryApplyScalarPowSpecialPaths(lv, rv, powSpecial)) {
      outS = std::move(powSpecial);
      return true;
    }
    if (lIm || rIm || complexNeedsPrincipalNegRealPow(ar, ai, br, bi)) {
      applyComplexCaretPrincipalEval(lv, rv, outS);
      return true;
    }
  }
  if (!lIm && !rIm) {
    return false;
  }
  if ((op == '+' || op == '-' || op == '*' || op == '/') &&
      tryApplyExactComplexCartesianBinary(lv, rv, op, outS)) {
    return true;
  }
  double floatR = 0.0;
  double floatI = 0.0;
  complexCartesianBinary(ar, ai, br, bi, op, floatR, floatI);
  outS = makeScalarComplexFromDoubles(floatR, floatI);
  return true;
}

std::string MathParser::formatComplexScalarValue(const EvalValue::ScalarValue& sv) const {
  std::string ratCx;
  if (tryFormatComplexRationalScalar(sv, ratCx)) {
    return ratCx;
  }
  double ar = 0.0;
  double ai = 0.0;
  scalarLoadCartesian(sv, ar, ai);
  if (std::isnan(ar) || std::isnan(ai)) {
    return STR_NAN;
  }
  EvalValue tmp = makeScalar(ar);
  tmp.scalarValue.scalarKind = sv.scalarKind;
  tmp.scalarValue.setExactInt64Valid(sv.hasExactInt64());
  tmp.scalarValue.exactInt64 = sv.exactInt64;
  tmp.scalarValue.setExactUInt64Valid(sv.hasExactUInt64());
  tmp.scalarValue.exactUInt64 = sv.exactUInt64;
  tmp.scalarValue.setDecScientificPow63High(sv.hasDecScientificPow63High());
  const std::string rePart = formatScalar(tmp, RenderBase::Dec);
  const double coeffAbs = std::fabs(ai);
  bool negUnit = false;
  std::string tail;
  if (std::isnan(ai)) {
    tail = std::string("nan*") + STR_I;
  } else if (std::isinf(ai)) {
    tail = (ai < 0.0) ? (std::string("-") + STR_INF + "*" + STR_I)
                        : (std::string(STR_INF) + "*" + STR_I);
  } else if (sv.hasImagExactInt64()) {
    const long long ii = sv.imagExactInt64;
    const long long ac = ii >= 0 ? ii : -ii;
    if (ac == 1LL) {
      tail = STR_I;
      negUnit = ii < 0;
    } else if (ii < 0) {
      tail = std::string("-") + std::to_string(ac) + STR_I;
    } else {
      tail = std::to_string(ac) + STR_I;
    }
  } else {
    long long nearlyI = 0;
    if (nearlyInt(ai, nearlyI) && nearlyI != 0) {
      const long long ac = nearlyI >= 0 ? nearlyI : -nearlyI;
      if (ac == 1LL) {
        tail = STR_I;
        negUnit = nearlyI < 0;
      } else if (nearlyI < 0) {
        tail = std::string("-") + std::to_string(ac) + STR_I;
      } else {
        tail = std::to_string(ac) + STR_I;
      }
    } else if (coeffAbs == 1.0) {
      tail = STR_I;
      negUnit = ai < 0.0;
    } else if (ai < 0.0) {
      tail = std::string("-") + formatDoubleFast(coeffAbs) + STR_I;
    } else {
      tail = formatDoubleFast(coeffAbs) + STR_I;
    }
  }
  const bool reZero = (sv.hasExactInt64() && sv.exactInt64 == 0) || (!sv.hasExactInt64() && ar == 0.0);
  return assembleComplexDecimalText(rePart, tail, negUnit, reZero);
}

std::string MathParser::assembleComplexDecimalText(
    const std::string& rePart,
    const std::string& imagTail,
    bool negUnitImag,
    bool reZero) {
  if (reZero) {
    return negUnitImag ? (std::string("-") + imagTail) : imagTail;
  }
  if (negUnitImag) {
    return rePart + "-" + imagTail;
  }
  if (!imagTail.empty() && imagTail.front() == '-') {
    return rePart + imagTail;
  }
  return rePart + "+" + imagTail;
}

std::string MathParser::formatComplexScalarWithRenderBase(
    const EvalValue::ScalarValue& sv, RenderBase base, bool asUnsigned) const {
  if (!scalarHasNonzeroImaginaryPart(sv) || base == RenderBase::Dec) {
    return formatComplexScalarValue(sv);
  }
  long long im = 0;
  if (!tryGetExactImagInt64Strict(sv, im)) {
    return formatComplexScalarValue(sv);
  }
  if (im == (std::numeric_limits<long long>::min)()) {
    return formatComplexScalarValue(sv);
  }

  EvalValue reTmp = makeScalar(0);
  double ar = 0.0;
  double ai = 0.0;
  scalarLoadCartesian(sv, ar, ai);
  (void)ai;
  reTmp.scalarValue.scalar = ar;
  reTmp.scalarValue.scalarKind = sv.scalarKind;
  reTmp.scalarValue.setExactInt64Valid(sv.hasExactInt64());
  reTmp.scalarValue.exactInt64 = sv.exactInt64;
  reTmp.scalarValue.setExactUInt64Valid(sv.hasExactUInt64());
  reTmp.scalarValue.exactUInt64 = sv.exactUInt64;
  reTmp.scalarValue.setDecScientificPow63High(sv.hasDecScientificPow63High());
  scalarClearImaginary(reTmp.scalarValue);
  reTmp.setRenderUnsigned(asUnsigned);
  std::string rePart = formatScalar(reTmp, base);

  bool negIm = false;
  long long absIm = 0;
  if (im >= 0) {
    absIm = im;
  } else {
    negIm = true;
    absIm = static_cast<long long>(-static_cast<long long>(im));
  }
  std::string tail;
  if (absIm == 1LL) {
    tail = STR_I;
  } else {
    EvalValue imTmp = makeScalarInt(absIm);
    imTmp.setRenderUnsigned(asUnsigned);
    tail = formatScalar(imTmp, base) + STR_I;
  }

  const bool reZero =
      (sv.hasExactInt64() && sv.exactInt64 == 0) || (!sv.hasExactInt64() && ar == 0.0);
  if (reZero) {
    return negIm ? (std::string("-") + tail) : tail;
  }
  return negIm ? (rePart + std::string("-") + tail) : (rePart + std::string("+") + tail);
}
#endif

MathParser::EvalValue MathParser::makeScalarMaybeExact(double v) {
  EvalValue out = makeScalar(v);
  long long asInt = 0;
  if (tryExtractExactInt64FromDoubleStrict(v, asInt)) {
    out.scalarValue.scalarKind = ScalarKind::Int64;
    out.scalarValue.setExactInt64Valid(true);
    out.scalarValue.exactInt64 = asInt;
    if (asInt >= 0) {
      out.scalarValue.setExactUInt64Valid(true);
      out.scalarValue.exactUInt64 = static_cast<std::uint64_t>(asInt);
    }
  }
  return out;
}

MathParser::EvalValue MathParser::makeScalarInt(long long v) {
  EvalValue out;
  out.kind = ValueKind::Scalar;
  out.scalarValue.scalarKind = ScalarKind::Int64;
  out.scalarValue.scalar = static_cast<double>(v);
  out.scalarValue.setExactInt64Valid(true);
  out.scalarValue.exactInt64 = v;
  out.scalarValue.setExactUInt64Valid(v >= 0);
  out.scalarValue.exactUInt64 = (v >= 0) ? static_cast<std::uint64_t>(v) : 0;
  out.scalarValue.setDecScientificPow63High(false);
  out.scalarValue.imag = 0.0;
  out.scalarValue.imagExactInt64 = 0;
  out.scalarValue.imagExactUInt64 = 0;
  out.scalarValue.setImagExactInt64Valid(false);
  out.scalarValue.setImagExactUInt64Valid(false);
  return out;
}

MathParser::EvalValue MathParser::makeScalarUInt(std::uint64_t v) {
  EvalValue out;
  out.kind = ValueKind::Scalar;
  out.scalarValue.scalarKind = ScalarKind::UInt64;
  out.scalarValue.scalar = static_cast<double>(v);
  out.scalarValue.setExactUInt64Valid(true);
  out.scalarValue.exactUInt64 = v;
  if (v <= static_cast<std::uint64_t>((std::numeric_limits<long long>::max)())) {
    out.scalarValue.setExactInt64Valid(true);
    out.scalarValue.exactInt64 = static_cast<long long>(v);
  } else {
    out.scalarValue.setExactInt64Valid(false);
    out.scalarValue.exactInt64 = 0;
  }
  out.scalarValue.setDecScientificPow63High(false);
  out.scalarValue.imag = 0.0;
  out.scalarValue.imagExactInt64 = 0;
  out.scalarValue.imagExactUInt64 = 0;
  out.scalarValue.setImagExactInt64Valid(false);
  out.scalarValue.setImagExactUInt64Valid(false);
  return out;
}

#if SMARTMATH_TIME_VALUES
MathParser::EvalValue MathParser::makeScalarTimeMs(long long totalMs) {
  EvalValue out;
  out.kind = ValueKind::Scalar;
  out.scalarValue.scalarKind = ScalarKind::Time;
  out.scalarValue.scalar = static_cast<double>(totalMs) / 1000.0;
  out.scalarValue.exactInt64 = totalMs;
  out.scalarValue.setExactInt64Valid(false);
  out.scalarValue.setExactUInt64Valid(false);
  out.scalarValue.setDecScientificPow63High(false);
  out.scalarValue.imag = 0.0;
  out.scalarValue.imagExactInt64 = 0;
  out.scalarValue.imagExactUInt64 = 0;
  out.scalarValue.setImagExactInt64Valid(false);
  out.scalarValue.setImagExactUInt64Valid(false);
  return out;
}

bool MathParser::scalarValueIsTime(const EvalValue::ScalarValue& s) const {
  if (!getSupportTimeValues()) {
    return false;
  }
  return s.scalarKind == ScalarKind::Time;
}

long long MathParser::timeTotalMsFromScalarValue(const EvalValue::ScalarValue& s) {
  return s.exactInt64;
}

bool MathParser::evalValueInvolvesTime(const EvalValue& v) const {
  if (!getSupportTimeValues()) {
    return false;
  }
  if (v.kind != ValueKind::Scalar) {
    for (const auto& item : v.arr) {
      if (scalarValueIsTime(item)) {
        return true;
      }
    }
    return false;
  }
  return scalarValueIsTime(v.scalarValue);
}
#endif

bool MathParser::evalValueHasNonzeroImaginary(const EvalValue& v) {
  if (v.kind == ValueKind::Scalar) {
    return scalarHasNonzeroImaginaryPart(v.scalarValue);
  }
  for (const auto& el : v.arr) {
    if (scalarHasNonzeroImaginaryPart(el)) {
      return true;
    }
  }
  return false;
}

#if SMARTMATH_TIME_VALUES
bool MathParser::scalarMsForCompare(const EvalValue::ScalarValue& sv, long long& outMs) const {
  if (scalarValueIsTime(sv)) {
    outMs = timeTotalMsFromScalarValue(sv);
    return true;
  }
  if (!std::isfinite(sv.scalar)) {
    return false;
  }
  outMs = roundHalfUpDoubleToLongLong(sv.scalar * 1000.0);
  return true;
}
#endif

bool MathParser::rejectBuiltinArgsWithComplexImaginary(
    EvalContext& ctx,
    const std::vector<EvalValue>& args) const {
#if SMARTMATH_COMPLEX_NUMBERS
  if (!getSupportComplexNumbers()) {
    return false;
  }
  for (const auto& v : args) {
    if (evalValueHasNonzeroImaginary(v)) {
      setParseError(ctx, ParseErrorId::IncompatibleOperands);
      return true;
    }
  }
  return false;
#else
  (void)ctx;
  (void)args;
  return false;
#endif
}

bool MathParser::cmpScalarValuesForCompare(
    EvalContext* ctx,
    const EvalValue::ScalarValue& sa,
    const EvalValue::ScalarValue& sb,
    int& cmpOut,
    CmpScalarIncompatiblePolicy policy) const {
#if SMARTMATH_TIME_VALUES
  const bool ta = scalarValueIsTime(sa);
  const bool tb = scalarValueIsTime(sb);
#else
  const bool ta = false;
  const bool tb = false;
#endif
  const bool ha = getSupportComplexNumbers() && scalarHasNonzeroImaginaryPart(sa);
  const bool hb = getSupportComplexNumbers() && scalarHasNonzeroImaginaryPart(sb);
  if ((ha || hb) && (ta || tb)) {
    if (policy == CmpScalarIncompatiblePolicy::SetError && ctx != nullptr) {
      setParseError(*ctx, ParseErrorId::IncompatibleOperands);
    }
    cmpOut = (policy == CmpScalarIncompatiblePolicy::SortUniqueReturnOne) ? 1 : 0;
    return policy == CmpScalarIncompatiblePolicy::SortUniqueReturnOne;
  }
#if SMARTMATH_TIME_VALUES
  if (ta || tb) {
    long long ams = 0;
    long long bms = 0;
    if (ta) {
      ams = timeTotalMsFromScalarValue(sa);
    } else if (!scalarMsForCompare(sa, ams)) {
      ams = 0;
    }
    if (tb) {
      bms = timeTotalMsFromScalarValue(sb);
    } else if (!scalarMsForCompare(sb, bms)) {
      bms = 0;
    }
    if (ams < bms) {
      cmpOut = -1;
    } else if (ams > bms) {
      cmpOut = 1;
    } else {
      cmpOut = 0;
    }
    return true;
  }
#endif
#if SMARTMATH_COMPLEX_NUMBERS
  if (ha || hb) {
    double ar = 0.0;
    double ai = 0.0;
    double br = 0.0;
    double bi = 0.0;
    scalarLoadCartesian(sa, ar, ai);
    scalarLoadCartesian(sb, br, bi);
    if (std::isnan(ar) || std::isnan(ai) || std::isnan(br) || std::isnan(bi)) {
      cmpOut = 1;
      return true;
    }
    if (ar < br) {
      cmpOut = -1;
    } else if (ar > br) {
      cmpOut = 1;
    } else if (ai < bi) {
      cmpOut = -1;
    } else if (ai > bi) {
      cmpOut = 1;
    } else {
      cmpOut = 0;
    }
    return true;
  }
#endif
  if (policy == CmpScalarIncompatiblePolicy::SortLess || policy == CmpScalarIncompatiblePolicy::SortGreater) {
    const bool aNan = std::isnan(sa.scalar);
    const bool bNan = std::isnan(sb.scalar);
    if (policy == CmpScalarIncompatiblePolicy::SortLess) {
      if (aNan) {
        cmpOut = bNan ? 0 : -1;
        return true;
      }
      if (bNan) {
        cmpOut = 1;
        return true;
      }
    } else {
      if (bNan) {
        cmpOut = aNan ? 0 : 1;
        return true;
      }
      if (aNan) {
        cmpOut = -1;
        return true;
      }
    }
  }
  if (sa.scalar < sb.scalar) {
    cmpOut = -1;
  } else if (sa.scalar > sb.scalar) {
    cmpOut = 1;
  } else {
    cmpOut = 0;
  }
  return true;
}

int MathParser::tryApplyGcdLcmScalars(
    const EvalValue::ScalarValue& a,
    const EvalValue::ScalarValue& b,
    bool doLcm,
    EvalValue& outV) const {
  std::uint64_t aU = 0;
  std::uint64_t bU = 0;
  if (tryGetExactNonNegativeUInt64FromScalar(a, aU) && tryGetExactNonNegativeUInt64FromScalar(b, bU)) {
    if (!doLcm) {
      outV = makeScalarUInt(gcdUInt64(aU, bU));
      return 0;
    }
    std::uint64_t lU = 0;
    if (!tryLcmUInt64(aU, bU, lU)) {
      return 2;
    }
    outV = makeScalarUInt(lU);
    return 0;
  }
  long long aI = 0;
  long long bI = 0;
  if (!tryGetSignedInt64FromScalar(a, aI) || !tryGetSignedInt64FromScalar(b, bI)) {
    return 1;
  }
  if (!doLcm) {
    outV = makeScalarInt(gcdInt64(aI, bI));
    return 0;
  }
  const long long g = gcdInt64(aI, bI);
  if (g == 0) {
    outV = makeScalarInt(0);
    return 0;
  }
  long long l = 0;
  if (!tryMulInt64Checked(aI / g, bI, l)) {
    return 2;
  }
  if (l < 0) {
    if (l == (std::numeric_limits<long long>::min)()) {
      return 2;
    }
    l = -l;
  }
  outV = makeScalarInt(l);
  return 0;
}

#if SMARTMATH_TIME_VALUES
MathParser::EvalValue MathParser::evalValueFromTimeMs(BuiltinFunctionId id, long long ms) const {
  if (id == BuiltinFunctionId::Milliseconds) {
    return makeScalarInt(ms);
  }
  if (id == BuiltinFunctionId::Seconds) {
    return makeScalarMaybeExact(static_cast<double>(ms) / 1000.0);
  }
  if (id == BuiltinFunctionId::Minutes) {
    return makeScalarMaybeExact(static_cast<double>(ms) / 60000.0);
  }
  if (id == BuiltinFunctionId::Hours) {
    return makeScalarMaybeExact(static_cast<double>(ms) / 3600000.0);
  }
  return makeScalarMaybeExact(static_cast<double>(ms) / 86400000.0);
}

MathParser::EvalValue MathParser::mapTimeUnitOverArray(
    EvalContext& ctx,
    const std::string& fnName,
    BuiltinFunctionId id,
    const EvalValue& inV) const {
  if (inV.kind != ValueKind::Array) {
    return makeScalar(0);
  }
  std::vector<EvalValue> outs;
  outs.reserve(inV.arr.size());
  for (const auto& item : inV.arr) {
    if (!scalarValueIsTime(item)) {
      setError(ctx, fnName + STR_TIME_EXPECTS_TIME_ARG);
      return makeScalar(0);
    }
    outs.push_back(evalValueFromTimeMs(id, timeTotalMsFromScalarValue(item)));
  }
  return makeArrayFromScalars(outs);
}
#endif

MathParser::EvalValue MathParser::applyGcdLcmEvalValues(const EvalValue& a, const EvalValue& b, bool doLcm, int& status) const {
  if (a.kind == ValueKind::Scalar && b.kind == ValueKind::Scalar) {
    EvalValue out = makeScalar(0);
    status = tryApplyGcdLcmScalars(a.scalarValue, b.scalarValue, doLcm, out);
    return out;
  }
  bool ok = true;
  const auto applyPair = [&](const EvalValue::ScalarValue& leftS, const EvalValue::ScalarValue& rightS, EvalValue& outS) -> bool {
    status = tryApplyGcdLcmScalars(leftS, rightS, doLcm, outS);
    return status == 0;
  };
  EvalValue result = mapBinaryBroadcast(a, b, applyPair, ok);
  if (!ok && status == 0) {
    status = 3;
  }
  return result;
}

MathParser::EvalValue MathParser::applyNcrNprEvalValues(
    const EvalValue& n,
    const EvalValue& r,
    bool doPerm,
    int& status) const {
  if (n.kind == ValueKind::Scalar && r.kind == ValueKind::Scalar) {
    long long nI = 0;
    long long rI = 0;
    long long out = 0;
    if (!tryGetSignedInt64FromScalar(n.scalarValue, nI) ||
        !tryGetSignedInt64FromScalar(r.scalarValue, rI)) {
      status = 1;
      return makeScalar(0);
    }
    if (!(doPerm ? tryComputeNprInt64(nI, rI, out) : tryComputeNcrInt64(nI, rI, out))) {
      status = 2;
      return makeScalar(0);
    }
    status = 0;
    return makeScalarInt(out);
  }
  bool ok = true;
  const auto applyPair = [&](const EvalValue::ScalarValue& nS, const EvalValue::ScalarValue& rS,
                             EvalValue& outS) -> bool {
    long long nI = 0;
    long long rI = 0;
    long long out = 0;
    if (!tryGetSignedInt64FromScalar(nS, nI) || !tryGetSignedInt64FromScalar(rS, rI)) {
      status = 1;
      return false;
    }
    if (!(doPerm ? tryComputeNprInt64(nI, rI, out) : tryComputeNcrInt64(nI, rI, out))) {
      status = 2;
      return false;
    }
    status = 0;
    outS = makeScalarInt(out);
    return true;
  };
  EvalValue result = mapBinaryBroadcast(n, r, applyPair, ok);
  if (!ok && status == 0) {
    status = 3;
  }
  return result;
}

bool MathParser::tryApplyModScalars(
    EvalContext& ctx,
    const std::string& fnName,
    const EvalValue::ScalarValue& aS,
    const EvalValue::ScalarValue& bS,
    EvalValue& outS) const {
  if (aS.hasExactUInt64() && bS.hasExactUInt64() && (!aS.hasExactInt64() || !bS.hasExactInt64())) {
    if (bS.exactUInt64 == 0u) {
      setNumericErrorInFunction(ctx, fnName);
      return false;
    }
    outS = makeScalarUInt(aS.exactUInt64 % bS.exactUInt64);
    return true;
  }
  long long a = 0;
  long long b = 0;
  if (!tryGetSignedInt64FromScalar(aS, a) || !tryGetSignedInt64FromScalar(bS, b)) {
    setIntegerValuesError(ctx, fnName);
    return false;
  }
  if (b == 0) {
    setNumericErrorInFunction(ctx, fnName);
    return false;
  }
  outS = makeScalarInt(a % b);
  return true;
}

MathParser::EvalValue MathParser::makeArrayFromScalars(const std::vector<EvalValue>& v) {
  EvalValue out;
  out.kind = ValueKind::Array;
  out.scalarValue.scalarKind = ScalarKind::FloatingPoint;
  out.arr.resize(v.size());
  for (std::size_t i = 0; i < v.size(); ++i) {
    out.arr[i] = v[i].scalarValue;
  }
  return out;
}

MathParser::RawResult::CartesianScalar MathParser::toRawCartesianScalar(const EvalValue::ScalarValue& v, bool imagPart) {
  RawResult::CartesianScalar out;
  if (!imagPart) {
#if SMARTMATH_TIME_VALUES
    if (v.scalarKind == ScalarKind::Time) {
      out.kind = RawResult::ScalarKind::Time;
      out.intValue = timeTotalMsFromScalarValue(v);
      return out;
    }
#endif
    if (v.hasRenderRational() && (v.hasExactUInt64Metadata() || v.exactUInt64 != 0)) {
      return rawCartesianAssignReducedRational(v.exactInt64, v.exactUInt64);
    }
#if SMARTMATH_FACTORINT
    if (v.hasRenderIntPower() && v.imagExactUInt64 > 1u) {
      out.kind = RawResult::ScalarKind::IntPower;
      out.rational.numerator = v.imagExactInt64;
      out.rational.denominator = v.imagExactUInt64;
      return out;
    }
#endif
    if (v.hasExactInt64()) {
      out.kind = RawResult::ScalarKind::Int64;
      out.intValue = v.exactInt64;
      return out;
    }
    if (v.hasExactUInt64()) {
      out.kind = RawResult::ScalarKind::UInt64;
      out.uintValue = v.exactUInt64;
      return out;
    }
    if (v.scalarKind == ScalarKind::Int64) {
      out.kind = RawResult::ScalarKind::Int64;
      out.intValue = v.exactInt64;
      return out;
    }
    if (v.scalarKind == ScalarKind::UInt64) {
      out.kind = RawResult::ScalarKind::UInt64;
      out.uintValue = v.exactUInt64;
      return out;
    }
    out.kind = RawResult::ScalarKind::FloatingPoint;
    out.floatingPoint = v.scalar;
    return out;
  }

  if (v.hasImagRenderRational() && (v.hasImagExactUInt64Metadata() || v.imagExactUInt64 != 0)) {
    return rawCartesianAssignReducedRational(v.imagExactInt64, v.imagExactUInt64);
  }
  if (v.hasImagExactInt64()) {
    out.kind = RawResult::ScalarKind::Int64;
    out.intValue = v.imagExactInt64;
    return out;
  }
  if (v.hasImagExactUInt64()) {
    if (v.imagExactUInt64 <= static_cast<std::uint64_t>(std::numeric_limits<std::int64_t>::max())) {
      out.kind = RawResult::ScalarKind::Int64;
      out.intValue = static_cast<std::int64_t>(v.imagExactUInt64);
    } else {
      out.kind = RawResult::ScalarKind::UInt64;
      out.uintValue = v.imagExactUInt64;
    }
    return out;
  }
  out.kind = RawResult::ScalarKind::FloatingPoint;
  out.floatingPoint = v.imag;
  return out;
}

MathParser::RawResult::Scalar MathParser::toRawScalar(const EvalValue::ScalarValue& v) {
  RawResult::Scalar out;
  if (scalarHasNonzeroImaginaryPart(v)) {
    out.kind = RawResult::ScalarKind::Complex;
    out.real = toRawCartesianScalar(v, false);
    out.imag = toRawCartesianScalar(v, true);
    return out;
  }
  out.real = toRawCartesianScalar(v, false);
  out.kind = out.real.kind;
  out.imag = RawResult::CartesianScalar{};
  return out;
}

MathParser::RawResult MathParser::toRawResult(const EvalValue& v) {
  RawResult out;
  if (v.kind == ValueKind::Scalar) {
    out.kind = RawResult::Kind::Scalar;
    out.scalar = toRawScalar(v.scalarValue);
    return out;
  }

  out.kind = RawResult::Kind::Array;
  out.array.reserve(v.arr.size());
  for (std::size_t i = 0; i < v.arr.size(); ++i) {
    out.array.push_back(toRawScalar(v.arr[i]));
  }
  return out;
}

MathParser::EvalValue MathParser::scalarFromScalarValue(const EvalValue::ScalarValue& sv) {
  EvalValue out = makeScalar(sv.scalar);
  out.scalarValue.scalarKind = sv.scalarKind;
  out.scalarValue.setExactInt64Valid(sv.hasExactInt64());
  out.scalarValue.exactInt64 = sv.exactInt64;
  out.scalarValue.setExactUInt64Valid(sv.hasExactUInt64());
  out.scalarValue.exactUInt64 = sv.exactUInt64;
  out.scalarValue.setDecScientificPow63High(sv.hasDecScientificPow63High());
  out.scalarValue.imag = sv.imag;
  out.scalarValue.imagExactInt64 = sv.imagExactInt64;
  out.scalarValue.imagExactUInt64 = sv.imagExactUInt64;
  out.scalarValue.setImagExactInt64Valid(sv.hasImagExactInt64());
  out.scalarValue.setImagExactUInt64Valid(sv.hasImagExactUInt64());
  out.scalarValue.setRenderRational(sv.hasRenderRational());
  out.scalarValue.setImagRenderRational(sv.hasImagRenderRational());
#if SMARTMATH_FACTORINT
  out.scalarValue.setRenderIntPower(sv.hasRenderIntPower());
#endif
  return out;
}

MathParser::EvalValue MathParser::scalarFromArrayAt(const EvalValue& arrV, std::size_t idx) {
  return scalarFromScalarValue(arrV.arr[idx]);
}

void MathParser::setError(EvalContext& ctx, const std::string& msg) const {
  if (!ctx.parseError) {
    ctx.parseError = true;
    ctx.errorText = msg;
  }
}

bool MathParser::flattenArgs(const std::vector<EvalValue>& args, std::vector<double>& out) {
  out.clear();
  out.reserve(forEachCallArgScalarValues(args, [](const EvalValue::ScalarValue&) {}));
  forEachCallArgScalarValues(args, [&](const EvalValue::ScalarValue& sv) { out.emplace_back(sv.scalar); });
  return !out.empty();
}

bool MathParser::flattenArgsToScalars(const std::vector<EvalValue>& args, std::vector<EvalValue>& out) {
  out.clear();
  out.reserve(forEachFlattenedEvalValue(args, [](const EvalValue&) {}));
  forEachFlattenedEvalValue(args, [&](const EvalValue& ev) { out.emplace_back(ev); });
  return !out.empty();
}

std::size_t MathParser::countFlattenedScalars(const std::vector<EvalValue>& args) {
  return forEachCallArgScalarValues(args, [](const EvalValue::ScalarValue&) {});
}

int MathParser::expandUnpackedArgs(const std::vector<EvalValue>& in, std::vector<EvalValue>& out) {
  out.clear();
  std::size_t reserveCount = 0;
  bool hasExpandMarkers = false;
  for (const auto& arg : in) {
    if (arg.hasExpandArgs()) hasExpandMarkers = true;
    if (arg.hasExpandArgs() && arg.kind == ValueKind::Array) reserveCount += arg.arr.size();
    else reserveCount += 1;
  }
  if (!hasExpandMarkers) {
    return 0;
  }
  out.reserve(reserveCount);

  for (const auto& arg : in) {
    if (!arg.hasExpandArgs()) {
      out.emplace_back(arg);
      continue;
    }
    if (arg.kind == ValueKind::Array) {
      for (const auto& item : arg.arr) {
        out.emplace_back(scalarFromScalarValue(item));
      }
    } else {
      EvalValue copy = arg;
      copy.setExpandArgs(false);
      out.emplace_back(std::move(copy));
    }
  }
  return static_cast<int>(out.size());
}

void MathParser::setVariable(const std::string& name, const EvalValue& value) {
  auto it = variables_.find(name);
  if (it != variables_.end()) {
    it->second = value;
    return;
  }
  variables_.emplace(name, value);
  ++variablesVersion_;
}

void MathParser::removeVariableByName(const std::string& name) {
  if (variables_.erase(name) != 0) {
    ++variablesVersion_;
  }
}

void MathParser::removeUserFunctionByName(const std::string& name) {
  auto it = userFunctionIndex_.find(name);
  if (it == userFunctionIndex_.end()) {
    return;
  }
  const std::size_t idx = it->second;
  const std::size_t last = userFunctions_.size() - 1;
  if (idx != last) {
    userFunctions_[idx] = std::move(userFunctions_[last]);
    userFunctionIndex_[userFunctions_[idx].name] = idx;
  }
  userFunctions_.pop_back();
  userFunctionIndex_.erase(it);
}

void MathParser::normalizeCallArgs(std::vector<EvalValue>& args) {
  std::vector<EvalValue> scratchExpandedArgs;
  if (expandUnpackedArgs(args, scratchExpandedArgs) > 0) {
    args.swap(scratchExpandedArgs);
  }
}

bool MathParser::applyBinary(double a, double b, char op, double& out) {
  switch (op) {
    case '+': out = a + b; return true;
    case '-': out = a - b; return true;
    case '*': out = a * b; return true;
    case '/':
      if (std::fabs(b) < 1e-15) {
        return false;
      }
      out = a / b;
      return true;
    case '^': out = std::pow(a, b); return true;
    default: break;
  }
  return false;
}

template <typename ScalarFn>
MathParser::EvalValue MathParser::mapUnaryEvalValue(const EvalValue& in, ScalarFn&& applyScalar) {
  if (in.kind == ValueKind::Scalar) {
    return applyScalar(in.scalarValue);
  }
  std::vector<EvalValue> outVals;
  outVals.reserve(in.arr.size());
  for (const auto& e : in.arr) {
    outVals.emplace_back(applyScalar(e));
  }
  return makeArrayFromScalars(outVals);
}

MathParser::EvalValue MathParser::negateEvalValue(const EvalValue& v) {
  if (v.kind == ValueKind::Array) {
    std::vector<EvalValue> elems;
    elems.reserve(v.arr.size());
    for (std::size_t i = 0; i < v.arr.size(); ++i) {
      EvalValue elem = scalarFromArrayAt(v, i);
      elems.emplace_back(negateEvalValue(elem));
    }
    return makeArrayFromScalars(elems);
  }
#if SMARTMATH_TIME_VALUES
  if (v.scalarValue.scalarKind == ScalarKind::Time) {
    const long long ms = timeTotalMsFromScalarValue(v.scalarValue);
    if (ms == (std::numeric_limits<long long>::min)()) {
      return makeScalarTimeMs((std::numeric_limits<long long>::max)());
    }
    return makeScalarTimeMs(-ms);
  }
#endif
#if SMARTMATH_COMPLEX_NUMBERS
  if (scalarHasNonzeroImaginaryPart(v.scalarValue)) {
    EvalValue exactNeg{};
    if (tryNegateExactComplexScalar(v.scalarValue, exactNeg)) {
      return exactNeg;
    }
    double re = 0.0;
    double im = 0.0;
    scalarLoadCartesian(v.scalarValue, re, im);
    return makeScalarComplexFromDoubles(-re, -im);
  }
#endif
  if (v.scalarValue.hasDecScientificPow63High()) {
    const double p63 = std::ldexp(1.0, 63);
    if (v.scalarValue.scalar == p63) {
      return makeScalarInt((std::numeric_limits<long long>::min)());
    }
    if (v.scalarValue.scalar == -p63) {
      return makeScalarUInt(1ull << 63);
    }
  }
  if (v.scalarValue.hasExactInt64()) {
    const long long x = v.scalarValue.exactInt64;
    if (x == (std::numeric_limits<long long>::min)()) {
      return makeScalarUInt(1ull << 63);
    }
    return makeScalarInt(-x);
  }
  if (v.scalarValue.hasExactUInt64()) {
    const std::uint64_t u = v.scalarValue.exactUInt64;
    if (u == 0) {
      return makeScalarInt(0);
    }
    if (u == (1ull << 63)) {
      return makeScalarInt((std::numeric_limits<long long>::min)());
    }
    if (u <= static_cast<std::uint64_t>((std::numeric_limits<long long>::max)())) {
      return makeScalarInt(-static_cast<long long>(u));
    }
    return makeScalar(-static_cast<double>(u));
  }
  return makeScalar(-v.scalarValue.scalar);
}

template <typename ScalarFn>
MathParser::EvalValue MathParser::mapBinaryBroadcast(
    const EvalValue& left,
    const EvalValue& right,
    ScalarFn&& scalarFn,
    bool& ok) const {
  if (left.kind == ValueKind::Scalar && right.kind == ValueKind::Scalar) {
    EvalValue outS;
    ok = scalarFn(left.scalarValue, right.scalarValue, outS);
    return ok ? outS : makeScalar(0);
  }
  if (left.kind == ValueKind::Array && right.kind == ValueKind::Array) {
    if (left.arr.size() != right.arr.size()) {
      ok = false;
      return makeScalar(0);
    }
    std::vector<EvalValue> out;
    out.reserve(left.arr.size());
    for (std::size_t i = 0; i < left.arr.size(); ++i) {
      EvalValue outS;
      if (!scalarFn(left.arr[i], right.arr[i], outS)) {
        ok = false;
        return makeScalar(0);
      }
      out.emplace_back(std::move(outS));
    }
    return makeArrayFromScalars(out);
  }
  if (left.kind == ValueKind::Array) {
    std::vector<EvalValue> out;
    out.reserve(left.arr.size());
    for (std::size_t i = 0; i < left.arr.size(); ++i) {
      EvalValue outS;
      if (!scalarFn(left.arr[i], right.scalarValue, outS)) {
        ok = false;
        return makeScalar(0);
      }
      out.emplace_back(std::move(outS));
    }
    return makeArrayFromScalars(out);
  }
  std::vector<EvalValue> out;
  out.reserve(right.arr.size());
  for (std::size_t i = 0; i < right.arr.size(); ++i) {
    EvalValue outS;
    if (!scalarFn(left.scalarValue, right.arr[i], outS)) {
      ok = false;
      return makeScalar(0);
    }
    out.emplace_back(std::move(outS));
  }
  return makeArrayFromScalars(out);
}

#if SMARTMATH_TIME_VALUES
bool MathParser::tryApplyTimeBinaryScalars(
    EvalContext& ctx,
    const EvalValue::ScalarValue& lv,
    const EvalValue::ScalarValue& rv,
    char op,
    EvalValue& outS) const {
  if (!getSupportTimeValues()) {
    return false;
  }
  const bool lt = scalarValueIsTime(lv);
  const bool rt = scalarValueIsTime(rv);
  if (!lt && !rt) {
    return false;
  }
#if SMARTMATH_COMPLEX_NUMBERS
  if (getSupportComplexNumbers()) {
    if ((!lt && scalarHasNonzeroImaginaryPart(lv)) || (!rt && scalarHasNonzeroImaginaryPart(rv))) {
      setParseError(ctx, ParseErrorId::IncompatibleOperands);
      return true;
    }
  }
#endif
  if (op != '+' && op != '-' && op != '*' && op != '/') {
    setParseError(ctx, ParseErrorId::IncompatibleOperands);
    return true;
  }
  if ((!lt && !std::isfinite(lv.scalar)) || (!rt && !std::isfinite(rv.scalar))) {
    setValidationError(ctx, STR_TIME_NON_FINITE);
    return true;
  }
  long long lms = 0;
  long long rms = 0;
  if (lt) {
    lms = timeTotalMsFromScalarValue(lv);
  } else {
    lms = roundHalfUpDoubleToLongLong(lv.scalar * 1000.0);
  }
  if (rt) {
    rms = timeTotalMsFromScalarValue(rv);
  } else {
    rms = roundHalfUpDoubleToLongLong(rv.scalar * 1000.0);
  }
  if (op == '+') {
    long long o = 0;
    if (!tryAddTimeMsChecked(lms, rms, o)) {
      setParseError(ctx, ParseErrorId::IncompatibleOperands);
      return true;
    }
    outS = makeScalarTimeMs(o);
    return true;
  }
  if (op == '-') {
    long long o = 0;
    if (!trySubTimeMsChecked(lms, rms, o)) {
      setParseError(ctx, ParseErrorId::IncompatibleOperands);
      return true;
    }
    outS = makeScalarTimeMs(o);
    return true;
  }
  if (op == '*') {
    if (lt && rt) {
      setParseError(ctx, ParseErrorId::IncompatibleOperands);
      return true;
    }
    const long long baseMs = lt ? lms : rms;
    const double mult = lt ? rv.scalar : lv.scalar;
    if (!std::isfinite(mult)) {
      setParseError(ctx, ParseErrorId::IncompatibleOperands);
      return true;
    }
    const double tMul = static_cast<double>(baseMs) * mult;
    const double kMin = static_cast<double>((std::numeric_limits<long long>::min)());
    const double kMax = static_cast<double>((std::numeric_limits<long long>::max)());
    if (tMul < kMin || tMul > kMax) {
      setParseError(ctx, ParseErrorId::IncompatibleOperands);
      return true;
    }
    outS = makeScalarTimeMs(roundHalfUpDoubleToLongLong(tMul));
    return true;
  }
  if (lt && rt) {
    if (rms == 0) {
      setParseError(ctx, ParseErrorId::IncompatibleOperands);
      return true;
    }
    const double ratio = (static_cast<double>(lms) / 1000.0) / (static_cast<double>(rms) / 1000.0);
    if (!std::isfinite(ratio)) {
      setParseError(ctx, ParseErrorId::IncompatibleOperands);
      return true;
    }
    outS = makeScalarMaybeExact(ratio);
    return true;
  }
  if (lt && !rt) {
    if (std::fabs(rv.scalar) < 1e-15) {
      setParseError(ctx, ParseErrorId::IncompatibleOperands);
      return true;
    }
    const double td = (static_cast<double>(lms) / 1000.0) / rv.scalar;
    if (!std::isfinite(td)) {
      setParseError(ctx, ParseErrorId::IncompatibleOperands);
      return true;
    }
    const double kMin = static_cast<double>((std::numeric_limits<long long>::min)());
    const double kMax = static_cast<double>((std::numeric_limits<long long>::max)());
    if (td < kMin || td > kMax) {
      setParseError(ctx, ParseErrorId::IncompatibleOperands);
      return true;
    }
    outS = makeScalarTimeMs(roundHalfUpDoubleToLongLong(td * 1000.0));
    return true;
  }
  setParseError(ctx, ParseErrorId::IncompatibleOperands);
  return true;
}
#endif

bool MathParser::tryApplyExactUInt64BinaryScalarOp(
    char op,
    std::uint64_t lu,
    std::uint64_t ru,
    EvalValue& outS) {
  if (op == '+') {
    std::uint64_t outU = 0;
    if (tryAddUInt64Checked(lu, ru, outU)) {
      outS = makeScalarUInt(outU);
      return true;
    }
  } else if (op == '-') {
    if (lu >= ru) {
      outS = makeScalarUInt(lu - ru);
      return true;
    }
  } else if (op == '*') {
    std::uint64_t outU = 0;
    if (tryMulUInt64Checked(lu, ru, outU)) {
      outS = makeScalarUInt(outU);
      return true;
    }
  } else if (op == '/') {
    std::uint64_t quoU = 0;
    if (tryQuotExactUInt64(lu, ru, quoU)) {
      outS = makeScalarUInt(quoU);
      return true;
    }
  }
  return false;
}

bool MathParser::tryCombineBinaryScalars(
    EvalContext& ctx,
    char op,
    const EvalValue::ScalarValue& lv,
    const EvalValue::ScalarValue& rv,
    EvalValue& outS) const {
#if SMARTMATH_TIME_VALUES
  if (getSupportTimeValues()) {
    EvalValue outTime;
    if (tryApplyTimeBinaryScalars(ctx, lv, rv, op, outTime)) {
      if (ctx.parseError) {
        return false;
      }
      outS = std::move(outTime);
      return true;
    }
  }
#endif
    if (op == '^') {
      EvalValue powSpecial;
      if (tryApplyScalarPowSpecialPaths(lv, rv, powSpecial)) {
        outS = std::move(powSpecial);
        return true;
      }
    }
#if SMARTMATH_COMPLEX_NUMBERS
  {
    EvalValue outComplex;
    if (tryApplyComplexBinaryScalars(lv, rv, op, outComplex)) {
      outS = std::move(outComplex);
      return true;
    }
  }
#endif
    if (isPureFloatingScalarPair(lv, rv)) {
      double outD = 0.0;
      if (!applyBinary(lv.scalar, rv.scalar, op, outD)) {
        return false;
      }
      outS = makeScalarMaybeExact(outD);
      return true;
    }

    if (lv.hasExactUInt64() && rv.hasExactUInt64()
        && (!lv.hasExactInt64() || !rv.hasExactInt64())) {
      if (tryApplyExactUInt64BinaryScalarOp(op, lv.exactUInt64, rv.exactUInt64, outS)) {
        return true;
      }
    }

    long long li = 0;
    long long ri = 0;
    const bool leftExactInt64 = tryGetExactSignedInt64NoUIntWrapFromScalar(lv, li);
    const bool rightExactInt64 = tryGetExactSignedInt64NoUIntWrapFromScalar(rv, ri);
    std::uint64_t lu = 0;
    std::uint64_t ru = 0;
    if (op == '^' && tryGetBothExactNonNegativeUInt64FromScalars(lv, rv, lu, ru)) {
      std::uint64_t outU = 0;
      if (tryPowUInt64Checked(lu, ru, outU)) {
        outS = makeScalarUInt(outU);
        return true;
      }
    }
    if (op == '^' && leftExactInt64 && li < 0 && tryGetExactNonNegativeUInt64FromScalar(rv, ru)) {
      const std::uint64_t baseMag = (li == (std::numeric_limits<long long>::min)())
        ? (1ull << 63)
        : static_cast<std::uint64_t>(-li);
      std::uint64_t powMag = 0;
      if (tryPowUInt64Checked(baseMag, ru, powMag)) {
        if ((ru & 1u) == 0u) {
          outS = makeScalarUInt(powMag);
          return true;
        }
        if (powMag <= (1ull << 63)) {
          outS = (powMag == (1ull << 63))
            ? makeScalarInt((std::numeric_limits<long long>::min)())
            : makeScalarInt(-static_cast<long long>(powMag));
          return true;
        }
      }
    }
    if (leftExactInt64 && rightExactInt64) {
      if (op == '-' && li == (std::numeric_limits<long long>::min)() && ri == 1LL) {
        outS = makeScalar(-std::ldexp(1.0, 63)); outS.scalarValue.setDecScientificPow63High(true);
        return true;
      }
      long long outI = 0;
      if (op == '+' && checkedAddLL(li, ri, outI)) {
        outS = makeScalarInt(outI);
        return true;
      }
      if (op == '-' && checkedSubLL(li, ri, outI)) {
        outS = makeScalarInt(outI);
        return true;
      }
      if (op == '*' && checkedMulLL(li, ri, outI)) {
        outS = makeScalarInt(outI);
        return true;
      }
      if (op == '^' && tryPowInt64Checked(li, ri, outI)) {
        outS = makeScalarInt(outI);
        return true;
      }
      if (op == '/' && tryQuotExactInt64(li, ri, outI)) {
        outS = makeScalarInt(outI);
        return true;
      }
      if (li >= 0 && ri >= 0) {
        lu = static_cast<std::uint64_t>(li);
        ru = static_cast<std::uint64_t>(ri);
        if (op == '+') {
          std::uint64_t outU = 0;
          if (tryAddUInt64Checked(lu, ru, outU)) {
            outS = makeScalarUInt(outU);
            return true;
          }
        } else if (op == '*') {
          std::uint64_t outU = 0;
          if (tryMulUInt64Checked(lu, ru, outU)) {
            outS = makeScalarUInt(outU);
            return true;
          }
        }
      }
    }

    lu = 0;
    ru = 0;
    if (tryGetBothExactNonNegativeUInt64FromScalars(lv, rv, lu, ru)) {
      if ((op == '+' || op == '*' || op == '/') &&
          tryApplyExactUInt64BinaryScalarOp(op, lu, ru, outS)) {
        return true;
      }
    }
    if (op == '+' && lv.hasExactInt64() && lv.exactInt64 == (std::numeric_limits<long long>::max)() && !rv.hasExactInt64() &&
        !rv.hasExactUInt64() && std::fabs(rv.scalar - 0.5) < 1e-12) {
      outS = makeScalar(std::ldexp(1.0, 63)); outS.scalarValue.setDecScientificPow63High(true);
      return true;
    }
    if (op == '+' && rv.hasExactInt64() && rv.exactInt64 == (std::numeric_limits<long long>::max)() && !lv.hasExactInt64() &&
        !lv.hasExactUInt64() && std::fabs(lv.scalar - 0.5) < 1e-12) {
      outS = makeScalar(std::ldexp(1.0, 63)); outS.scalarValue.setDecScientificPow63High(true);
      return true;
    }
#if SMARTMATH_COMPLEX_NUMBERS
    if (getSupportComplexNumbers() && (scalarHasNonzeroImaginaryPart(lv) || scalarHasNonzeroImaginaryPart(rv))) {
      return false;
    }
#endif
    double outD = 0;
    if (!applyBinary(lv.scalar, rv.scalar, op, outD)) {
      return false;
    }
    if (op == '*') {
      EvalValue exactMul{};
      if (tryApplyExactIntegerMultiplicationFromProduct(lv, rv, outD, exactMul)) {
        outS = std::move(exactMul);
        return true;
      }
    }
    if (op == '/') {
      EvalValue exactDiv{};
      if (tryApplyExactIntegerDivisionFromQuotient(lv, rv, outD, exactDiv)) {
        outS = std::move(exactDiv);
        return true;
      }
    }
    outS = makeScalarMaybeExact(outD);
    return true;
}

MathParser::EvalValue MathParser::mapBinary(EvalContext& ctx, const EvalValue& a, const EvalValue& b, char op, bool& ok) const {
  ok = true;
  return mapBinaryBroadcast(
      a,
      b,
      [this, &ctx, op](const EvalValue::ScalarValue& lv, const EvalValue::ScalarValue& rv, EvalValue& outS) -> bool {
        return tryCombineBinaryScalars(ctx, op, lv, rv, outS);
      },
      ok);
}
MathParser::EvalValue MathParser::mapBinaryBuiltinMathFunction(
    const EvalValue& left,
    const EvalValue& right,
    BuiltinFunctionId id,
    bool& ok) const {
  ok = true;
  const bool isLog = (id == BuiltinFunctionId::Log);
  double (*binaryFn)(double, double) = nullptr;
  if (!isLog) {
    if (id == BuiltinFunctionId::Atan2) {
      binaryFn = std::atan2;
    } else if (id == BuiltinFunctionId::Hypot) {
      binaryFn = [](double l, double r) {
        return std::sqrt((l * l) + (r * r));
      };
    } else {
      ok = false;
      return makeScalar(0);
    }
  }

  auto applyScalarValue = [&](const EvalValue::ScalarValue& l, const EvalValue::ScalarValue& r, EvalValue& outS) -> bool {
#if SMARTMATH_COMPLEX_NUMBERS
    if (getSupportComplexNumbers() && id == BuiltinFunctionId::Hypot &&
        (scalarHasNonzeroImaginaryPart(l) || scalarHasNonzeroImaginaryPart(r))) {
      double ar = 0.0;
      double ai = 0.0;
      double br = 0.0;
      double bi = 0.0;
      scalarLoadCartesian(l, ar, ai);
      scalarLoadCartesian(r, br, bi);
      if (!std::isfinite(ar) || !std::isfinite(ai) || !std::isfinite(br) || !std::isfinite(bi)) {
        outS = makeScalar(std::numeric_limits<double>::quiet_NaN());
        return true;
      }
      const double ml = std::hypot(ar, ai);
      const double mr = std::hypot(br, bi);
      outS = makeScalarMaybeExact(std::hypot(ml, mr));
      return true;
    }
#endif
    double outD = 0.0;
    if (isLog) {
#if SMARTMATH_COMPLEX_NUMBERS
      if (!getSupportComplexNumbers()) {
#endif
        if (l.scalar <= 0.0 || r.scalar <= 0.0 || r.scalar == 1.0) {
          return false;
        }
        outD = std::log(l.scalar) / std::log(r.scalar);
        outS = makeScalarMaybeExact(outD);
        return true;
#if SMARTMATH_COMPLEX_NUMBERS
      }
      double ar = 0.0;
      double ai = 0.0;
      double br = 0.0;
      double bi = 0.0;
      scalarLoadCartesian(l, ar, ai);
      scalarLoadCartesian(r, br, bi);
      const bool pureReal =
          !scalarHasNonzeroImaginaryPart(l) && !scalarHasNonzeroImaginaryPart(r) && ai == 0.0 && bi == 0.0;
      if (pureReal && ar > 0.0 && br > 0.0 && br != 1.0) {
        outD = std::log(ar) / std::log(br);
        outS = makeScalarMaybeExact(outD);
        return true;
      }
      if (!std::isfinite(ar) || !std::isfinite(ai) || !std::isfinite(br) || !std::isfinite(bi)) {
        outS = makeScalarComplexFromDoubles(
            std::numeric_limits<double>::quiet_NaN(), std::numeric_limits<double>::quiet_NaN());
        return true;
      }
      double lnr = 0.0;
      double lni = 0.0;
      double rnr = 0.0;
      double rni = 0.0;
      scalarPrincipalLnCartesian(ar, ai, lnr, lni);
      scalarPrincipalLnCartesian(br, bi, rnr, rni);
      const double den = rnr * rnr + rni * rni;
      if (den == 0.0) {
        outS = makeScalarComplexFromDoubles(
            std::numeric_limits<double>::quiet_NaN(), std::numeric_limits<double>::quiet_NaN());
        return true;
      }
      double outRe = (lnr * rnr + lni * rni) / den;
      double outIm = (lni * rnr - lnr * rni) / den;
      snapComplexNearZeroAxis(outRe, outIm);
      outS = makeScalarComplexFromDoubles(outRe, outIm);
      return true;
#endif
    }
    if (id == BuiltinFunctionId::Hypot) {
      EvalValue hv;
      if (tryApplyHypotExactScalars(l, r, hv)) {
        outS = hv;
        return true;
      }
    }
    outD = binaryFn(l.scalar, r.scalar);
    outS = makeScalarMaybeExact(outD);
    return true;
  };

  return mapBinaryBroadcast(left, right, applyScalarValue, ok);
}

bool MathParser::parseFunctionDefinition(
    EvalContext& ctx,
    std::string& outName,
    std::vector<std::string>& outParams,
    std::string& outExpr) const {
  const char* save = ctx.p;
  skipSpaces(ctx);

  if (!isIdentStart(*ctx.p)) {
    ctx.p = save;
    return false;
  }
  std::string fnName = consumeLowerIdentToken(ctx);
  skipSpaces(ctx);
  if (*ctx.p != '(') {
    ctx.p = save;
    return false;
  }
  ++ctx.p;
  skipSpaces(ctx);

  std::vector<std::string> params;
  if (*ctx.p != ')') {
    while (true) {
      if (!isIdentStart(*ctx.p)) {
        ctx.p = save;
        return false;
      }
      params.emplace_back(consumeLowerIdentToken(ctx));
      skipSpaces(ctx);
      if (*ctx.p == ',') {
        ++ctx.p;
        skipSpaces(ctx);
        continue;
      }
      break;
    }
  }

  if (*ctx.p != ')') {
    ctx.p = save;
    return false;
  }
  ++ctx.p;
  skipSpaces(ctx);
  // UDF: name(params)=body — single '=' only; do not treat first '=' of '==' as UDF starter.
  if (*ctx.p != '=') {
    ctx.p = save;
    return false;
  }
  // Safe: *ctx.p is '=' here, so ctx.p[1] is at worst the terminating '\0' on a c_str() buffer.
  if (ctx.p[1] == '=') {
    ctx.p = save;
    return false;
  }
  ++ctx.p;
  skipSpaces(ctx);

  outName = fnName;
  outParams = std::move(params);
  const char* exprStart = ctx.p;
  while (*ctx.p && *ctx.p != ';') {
    ++ctx.p;
  }
  outExpr = trim(std::string(exprStart, static_cast<std::size_t>(ctx.p - exprStart)));
  if (outExpr.empty()) {
    ctx.p = save;
    return false;
  }
  return true;
}

namespace {

int netRoundParenDepthBetween(const char* lo, const char* hiExclusive) {
  int d = 0;
  for (const char* q = lo; q < hiExclusive; ++q) {
    if (*q == '(') {
      ++d;
    } else if (*q == ')') {
      --d;
    }
  }
  return d;
}

}  // namespace

#if SMARTMATH_LAMBDA_FUNCTIONS
bool MathParser::lambdaBodyConsume(
    EvalContext& ctx,
    std::string& outBody,
    LambdaBodyStop stop,
    const char* sortbyLambdaKeyStart) {
  outBody.clear();
  int g = 0;
  while (*ctx.p) {
    const char c = *ctx.p;
    if (g == 0) {
      switch (stop) {
        case LambdaBodyStop::WrappedParenClose:
          if (c == ')') {
            return true;
          }
          break;
        case LambdaBodyStop::SortbyArgDelim:
          if ((c == ',' || c == ')') && sortbyLambdaKeyStart != nullptr &&
              netRoundParenDepthBetween(sortbyLambdaKeyStart, ctx.p) == 0) {
            return true;
          }
          break;
        case LambdaBodyStop::TopLevelSemicolonOrEof:
          if (c == ';') {
            return true;
          }
          break;
      }
    }
    switch (c) {
      case '(':
      case '[':
      case '{':
        ++g;
        break;
      case ')':
      case ']':
      case '}':
        --g;
        if (g < 0) {
          return false;
        }
        break;
      default:
        break;
    }
    outBody.push_back(c);
    ++ctx.p;
  }
  return stop == LambdaBodyStop::TopLevelSemicolonOrEof || stop == LambdaBodyStop::SortbyArgDelim;
}

MathParser::EvalValue MathParser::makeInlineLambdaValue(std::vector<std::string> params, std::string body) {
  EvalValue v;
  v.kind = ValueKind::InlineLambda;
  v.lambdaParams = std::move(params);
  v.lambdaBody = std::move(body);
  return v;
}

bool MathParser::tryConsumeLambdaParameterList(EvalContext& ctx, std::vector<std::string>& outParams, bool quiet) const {
  outParams.clear();
  skipSpaces(ctx);
  if (*ctx.p == '(') {
    ++ctx.p;
    skipSpaces(ctx);
    if (*ctx.p == ')') {
      ++ctx.p;
      return true;
    }
    while (true) {
      if (!isIdentStart(*ctx.p)) {
        if (!quiet) {
          setParseError(ctx, ParseErrorId::UnexpectedToken);
        }
        return false;
      }
      outParams.emplace_back(consumeLowerIdentToken(ctx));
      skipSpaces(ctx);
      if (*ctx.p == ',') {
        ++ctx.p;
        skipSpaces(ctx);
        continue;
      }
      if (*ctx.p == ')') {
        ++ctx.p;
        return true;
      }
      if (!quiet) {
        setParseError(ctx, ParseErrorId::UnexpectedToken);
      }
      return false;
    }
  }
  if (!isIdentStart(*ctx.p)) {
    return false;
  }
  outParams.emplace_back(consumeLowerIdentToken(ctx));
  skipSpaces(ctx);
  while (*ctx.p == ',') {
    ++ctx.p;
    skipSpaces(ctx);
    if (!isIdentStart(*ctx.p)) {
      if (!quiet) {
        setParseError(ctx, ParseErrorId::UnexpectedToken);
      }
      return false;
    }
    outParams.emplace_back(consumeLowerIdentToken(ctx));
    skipSpaces(ctx);
  }
  return true;
}

bool MathParser::tryParseLambdaInnerUnwrappedSuffix(
    EvalContext& ctx,
    std::vector<std::string>& outParams,
    std::string& outBody,
    LambdaBodyStop bodyStop,
    const char* sortbyLambdaKeyStart,
    bool quiet) const {
  if (!tryConsumeLambdaParameterList(ctx, outParams, quiet)) {
    return false;
  }
  skipSpaces(ctx);
  if (*ctx.p != ':') {
    return false;
  }
  ++ctx.p;
  skipSpaces(ctx);
  return lambdaBodyConsume(ctx, outBody, bodyStop, sortbyLambdaKeyStart);
}

bool MathParser::tryParseWrappedParenLambdaSuffix(
    EvalContext& ctx,
    std::vector<std::string>& outParams,
    std::string& outBody,
    LambdaBodyStop bodyStop,
    bool quiet) const {
  const char* const saveOuter = ctx.p;
  if (*ctx.p != '(') {
    return false;
  }
  if (!peekRhsMayBeLambdaSyntaxAt(ctx.p + 1)) {
    return false;
  }
  ++ctx.p;
  if (!tryParseLambdaInnerUnwrappedSuffix(ctx, outParams, outBody, bodyStop, nullptr, quiet)) {
    ctx.p = saveOuter;
    return false;
  }
  skipSpaces(ctx);
  if (*ctx.p != ')') {
    ctx.p = saveOuter;
    return false;
  }
  ++ctx.p;
  skipSpaces(ctx);
  return true;
}

bool MathParser::tryParseLambdaRhsAfterEquals(
    EvalContext& ctx, std::vector<std::string>& outParams, std::string& outExpr) const {
  const char* const saveOuter = ctx.p;
  if (!peekRhsMayBeLambdaSyntaxAt(saveOuter)) {
    return false;
  }
  skipSpaces(ctx);
  if (tryParseWrappedParenLambdaSuffix(
          ctx, outParams, outExpr, LambdaBodyStop::WrappedParenClose, true)) {
    if (*ctx.p == '\0' || *ctx.p == ';') {
      return true;
    }
    ctx.p = saveOuter;
    outParams.clear();
    outExpr.clear();
  }
  skipSpaces(ctx);
  if (!tryParseLambdaInnerUnwrappedSuffix(
          ctx, outParams, outExpr, LambdaBodyStop::TopLevelSemicolonOrEof, nullptr, true)) {
    return false;
  }
  skipSpaces(ctx);
  return *ctx.p == '\0' || *ctx.p == ';';
}

std::unique_ptr<MathParser::Expr> MathParser::makeUnarySortbyInlineLambdaExpr(
    EvalContext& ctx,
    std::vector<std::string>&& params,
    std::string&& body) {
  if (params.size() != 1U) {
    setValidationError(ctx, STR_SORTBY_EXPECTS_UNARY_FUNCTION);
    return nullptr;
  }
  if (trim(body).empty()) {
    setValidationError(ctx, STR_FUNCTION_BODY_IS_EMPTY);
    return nullptr;
  }
  auto lit = std::make_unique<Expr>();
  lit->tag = Expr::Tag::Literal;
  lit->literalValue = makeInlineLambdaValue(std::move(params), std::move(body));
  return lit;
}

std::unique_ptr<MathParser::Expr> MathParser::tryParseUnarySortbyLambdaOrFunctionRef(EvalContext& ctx) {
  ctx.parseError = false;
  ctx.errorText.clear();
  const char* const save = ctx.p;
  skipSpaces(ctx);

  if (*ctx.p == '(' && peekRhsMayBeLambdaSyntaxAt(ctx.p + 1)) {
    std::vector<std::string> params;
    std::string body;
    if (tryParseWrappedParenLambdaSuffix(ctx, params, body, LambdaBodyStop::WrappedParenClose, true)) {
      if (*ctx.p == ',' || *ctx.p == ')' || *ctx.p == '\0') {
        return makeUnarySortbyInlineLambdaExpr(ctx, std::move(params), std::move(body));
      }
    }
  }

  ctx.p = save;
  skipSpaces(ctx);
  if (!peekRhsMayBeLambdaSyntaxAt(ctx.p)) {
    return parseSortbyFunctionRef(ctx);
  }

  std::vector<std::string> params;
  std::string body;
  if (!tryParseLambdaInnerUnwrappedSuffix(
          ctx, params, body, LambdaBodyStop::SortbyArgDelim, save, true)) {
    ctx.p = save;
    return parseSortbyFunctionRef(ctx);
  }
  skipSpaces(ctx);
  if (*ctx.p == ',' || *ctx.p == ')' || *ctx.p == '\0') {
    return makeUnarySortbyInlineLambdaExpr(ctx, std::move(params), std::move(body));
  }
  ctx.p = save;
  return parseSortbyFunctionRef(ctx);
}

MathParser::EvalValue MathParser::evalInlineLambdaCall(
    EvalContext& ctx,
    const std::vector<std::string>& paramNames,
    const std::string& bodyExpr,
    std::vector<EvalValue>&& args) {
  static const std::string kLambda = "lambda";
  if (args.size() != paramNames.size()) {
    setExactArgCountError(ctx, kLambda, paramNames.size());
    return makeScalar(0);
  }
  if (ctx.evalDepth >= kMaxEvalDepth) {
    setParseError(ctx, ParseErrorId::MaxEvaluationDepthReached);
    return makeScalar(0);
  }

  std::unordered_map<std::string, EvalValue> localVars;
  localVars.reserve(paramNames.size());
  for (std::size_t i = 0; i < paramNames.size(); ++i) {
    localVars.emplace(paramNames[i], std::move(args[i]));
  }

  const std::string trimmedBody = trim(bodyExpr);
  if (trimmedBody.empty() || udfBodyIsEmptyTupleLiteral(trimmedBody)) {
    setValidationError(ctx, STR_FUNCTION_BODY_IS_EMPTY);
    return makeScalar(0);
  }

  std::vector<AstStatement> compiledBody;
  if (!tryCompileExprToProgram(trimmedBody, ctx, nullptr, nullptr, true, compiledBody)) {
    return makeScalar(0);
  }

  EvalContext sub;
  sub.evalDepth = ctx.evalDepth + 1;
  sub.userFunctionCallStack = ctx.userFunctionCallStack;
  sub.suppressBareFunctionTailHint = true;
  EvalValue v = runCompiledProgram(sub, compiledBody, &localVars, false);
  if (!finishSubEvalFromCompiledProgram(ctx, sub)) {
    return makeScalar(0);
  }
  return v;
}
#endif

std::unique_ptr<MathParser::Expr> MathParser::parseSortbyKeyArg(EvalContext& ctx) {
#if SMARTMATH_LAMBDA_FUNCTIONS
  if (getSupportLambdaFunctions()) {
    return tryParseUnarySortbyLambdaOrFunctionRef(ctx);
  }
#endif
  return parseSortbyKeyArgFunctionRefOnly(ctx);
}

std::unique_ptr<MathParser::Expr> MathParser::parseSortbyKeyArgFunctionRefOnly(EvalContext& ctx) {
  ctx.parseError = false;
  ctx.errorText.clear();
  skipSpaces(ctx);
#if SMARTMATH_LAMBDA_FUNCTIONS
  if (peekRhsMayBeLambdaSyntaxAt(ctx.p)) {
      setParseError(ctx, ParseErrorId::UnexpectedToken);
      return nullptr;
    }
#endif
  return parseSortbyFunctionRef(ctx);
}

bool MathParser::isTightImagUnitSuffixAt(const char* p) {
  if (*p != 'i' && *p != 'I') {
    return false;
  }
  const unsigned char next = static_cast<unsigned char>(p[1]);
  return !(std::isalnum(next) || next == '_');
}

std::unique_ptr<MathParser::Expr> MathParser::wrapExprWithTightImagSuffixIfPresent(
    EvalContext& ctx,
    std::unique_ptr<Expr> expr) {
#if SMARTMATH_COMPLEX_NUMBERS
  if (!getSupportComplexNumbers() || !expr || !isTightImagUnitSuffixAt(ctx.p)) {
    return expr;
  }
  ++ctx.p;
  if (expr->tag == Expr::Tag::Literal) {
    setPureImaginaryFromMagnitudeScalar(expr->literalValue, expr->literalValue.scalarValue);
    return expr;
  }
  auto imagLit = std::make_unique<Expr>();
  imagLit->tag = Expr::Tag::Literal;
  imagLit->literalValue = makeImaginaryUnit();
  return makeBinaryExpr(std::move(expr), std::move(imagLit), Expr::BinaryOp::Mul, false);
#else
  return expr;
#endif
}

std::unique_ptr<MathParser::Expr> MathParser::parsePrimary(EvalContext& ctx) {
  skipSpaces(ctx);
  std::unique_ptr<Expr> prim;
  if (*ctx.p == '(') {
    prim = parsePrimaryParenthesized(ctx);
  } else if (isNumericLiteralStart(*ctx.p)) {
    return parsePrimaryNumericLiteral(ctx);
  } else if (isIdentStart(*ctx.p)) {
    prim = parsePrimaryIdentifierOrCall(ctx);
  } else {
    setParseError(ctx, ParseErrorId::UnexpectedToken);
    return nullptr;
  }
  if (ctx.parseError || !prim) {
    return nullptr;
  }
  return wrapExprWithTightImagSuffixIfPresent(ctx, std::move(prim));
}

std::unique_ptr<MathParser::Expr> MathParser::parsePrimaryParenthesized(EvalContext& ctx) {
  ++ctx.p;
  skipSpaces(ctx);
  if (*ctx.p == ')') {
    ++ctx.p;
    auto emptyArr = std::make_unique<Expr>();
    emptyArr->tag = Expr::Tag::ArrayOrParens;
    return emptyArr;
  }
  std::vector<std::unique_ptr<Expr>> values;
  if (!parseParenthesizedExprList(ctx, values)) {
    return nullptr;
  }
  if (*ctx.p != ')') {
    setMissingClosingParenLikeError(ctx);
    return nullptr;
  }
  ++ctx.p;
  if (values.size() == 1) {
    return std::move(values[0]);
  }
  auto arr = std::make_unique<Expr>();
  arr->tag = Expr::Tag::ArrayOrParens;
  arr->elements = std::move(values);
  return arr;
}

namespace {

enum class NumericLiteralRoute { Plain, ColonTime, CompactTime };

bool isDecimalRadixPrefixedAt(const char* p) {
  if (p[0] != '0') {
    return false;
  }
  const char px = static_cast<char>(std::tolower(static_cast<unsigned char>(p[1])));
  return px == 'x' || px == 'b' || px == 'o';
}

#if SMARTMATH_TIME_VALUES
bool scanColonTimeLiteralEnd(const char* p0, const char*& outEnd) {
  if (*p0 < '0' || *p0 > '9' || isDecimalRadixPrefixedAt(p0)) {
    return false;
  }
  const char* q = p0;
  bool hasColon = false;
  while ((*q >= '0' && *q <= '9') || *q == ':' || *q == '.') {
    if (*q == ':') {
      hasColon = true;
    }
    ++q;
  }
  if (!hasColon) {
    return false;
  }
  outEnd = q;
  return true;
}

bool peekCompactTimeSuffixAfterDigitRun(const char* digitEnd) {
  const char* r = digitEnd;
  while (*r == ' ' || *r == '\t') {
    ++r;
  }
  if (*r == 'm' && r[1] == 's') {
    return true;
  }
  if (*r == 'd' || *r == 'h' || *r == 's') {
    return true;
  }
  if (*r == 'm') {
    return true;
  }
  return false;
}
#endif

const char* advancePastOptionalFraction(const char* q) {
  if (*q == '.') {
    ++q;
    while (*q >= '0' && *q <= '9') {
      ++q;
    }
  }
  return q;
}

NumericLiteralRoute classifyNumericLiteralRoute(const char* p, int arrayIndexBracketDepth = 0) {
  if (*p < '0' || *p > '9') {
    return NumericLiteralRoute::Plain;
  }
  if (isDecimalRadixPrefixedAt(p)) {
    return NumericLiteralRoute::Plain;
  }
  const char* q = p;
  while (*q >= '0' && *q <= '9') {
    ++q;
  }
  q = advancePastOptionalFraction(q);
  if (*q == 'e' || *q == 'E') {
    ++q;
    if (*q == '+' || *q == '-') {
      ++q;
    }
    while (*q >= '0' && *q <= '9') {
      ++q;
    }
  }
  if (*q != ':') {
    const char* digitEnd = p;
    while (*digitEnd >= '0' && *digitEnd <= '9') {
      ++digitEnd;
    }
#if SMARTMATH_TIME_VALUES
    if (peekCompactTimeSuffixAfterDigitRun(digitEnd)) {
      return NumericLiteralRoute::CompactTime;
    }
#endif
    return NumericLiteralRoute::Plain;
  }
  // Inside ``[...]``, top-level ``:`` is slice syntax, not a time literal.
  if (arrayIndexBracketDepth > 0) {
    return NumericLiteralRoute::Plain;
  }
#if SMARTMATH_TIME_VALUES
  const char* colonEnd = nullptr;
  if (scanColonTimeLiteralEnd(p, colonEnd)) {
    return NumericLiteralRoute::ColonTime;
  }
#endif
  return NumericLiteralRoute::Plain;
}

const char* scanDecimalNumericLiteralEnd(const char* p) {
  if (!p || !*p) {
    return p;
  }
  if (*p == '+') {
    ++p;
  }
  const char* q = p;
  bool anyDigit = false;
  while (*q >= '0' && *q <= '9') {
    anyDigit = true;
    ++q;
  }
  if (*q == '.') {
    const char* afterFrac = advancePastOptionalFraction(q);
    if (afterFrac != q) {
      anyDigit = true;
      q = afterFrac;
    }
  }
  if (*q == 'e' || *q == 'E') {
    const char* e = q + 1;
    if (*e == '+' || *e == '-') {
      ++e;
    }
    if (*e >= '0' && *e <= '9') {
      while (*e >= '0' && *e <= '9') {
        ++e;
      }
      q = e;
    }
  }
  return anyDigit ? q : p;
}

}  // namespace

#if SMARTMATH_TIME_VALUES
bool MathParser::tryParseScalarTimeLiteral(EvalContext& ctx, EvalValue& out) const {
  if (!getSupportTimeValues()) {
    return false;
  }
  if (ctx.arrayIndexBracketDepth > 0) {
    return false;
  }
  const char* p0 = ctx.p;
  const char* q = nullptr;
  if (!scanColonTimeLiteralEnd(p0, q)) {
    return false;
  }
  const char* timeErr = nullptr;
  long long ms = 0;
  if (!parseTimeLiteralStringToMs(p0, q, ms, timeErr)) {
    if (timeErr != nullptr) {
      setValidationError(ctx, timeErr);
    }
    return false;
  }
  ctx.p = q;
  out = makeScalarTimeMs(ms);
  return true;
}

bool MathParser::tryParseCompactSuffixTimeLiteral(EvalContext& ctx, EvalValue& out) const {
  if (!getSupportTimeValues()) {
    return false;
  }
  const char* const pSave = ctx.p;
  if (*ctx.p < '0' || *ctx.p > '9' || isDecimalRadixPrefixedAt(ctx.p)) {
    return false;
  }
  long long totalMs = 0;
  int lastUnitRank = -1;
  int comps = 0;
  while (true) {
    if (*ctx.p < '0' || *ctx.p > '9') {
      break;
    }
    std::uint64_t uv = 0;
    while (*ctx.p >= '0' && *ctx.p <= '9') {
      const int dig = *ctx.p - '0';
      if (uv > ((std::numeric_limits<std::uint64_t>::max)() - static_cast<std::uint64_t>(dig)) / 10u) {
        ctx.p = pSave;
        if (comps > 0) {
          setValidationError(ctx, STR_TIME_LITERAL_INVALID_SEGMENT);
        }
        return false;
      }
      uv = uv * 10u + static_cast<std::uint64_t>(dig);
      ++ctx.p;
    }
    skipSpaces(ctx);
    int ur = -1;
    long long fac = 0;
    if (ctx.p[0] == 'm' && ctx.p[1] == 's') {
      ur = 4;
      fac = 1;
      ctx.p += 2;
    } else if (ctx.p[0] == 'd') {
      ur = 0;
      fac = 86400000;
      ++ctx.p;
    } else if (ctx.p[0] == 'h') {
      ur = 1;
      fac = 3600000;
      ++ctx.p;
    } else if (ctx.p[0] == 'm') {
      ur = 2;
      fac = 60000;
      ++ctx.p;
    } else if (ctx.p[0] == 's') {
      ur = 3;
      fac = 1000;
      ++ctx.p;
    } else {
      if (comps == 0) {
        ctx.p = pSave;
        return false;
      }
      ctx.p = pSave;
      setValidationError(ctx, STR_TIME_COMPACT_EXPECTED_UNIT);
      return false;
    }
    if (ur <= lastUnitRank) {
      ctx.p = pSave;
      setValidationError(ctx, STR_TIME_COMPACT_UNIT_ORDER);
      return false;
    }
    lastUnitRank = ur;
    if (uv > static_cast<std::uint64_t>((std::numeric_limits<long long>::max)())) {
      ctx.p = pSave;
      setValidationError(ctx, STR_TIME_LITERAL_INVALID_SEGMENT);
      return false;
    }
    const long long cf = static_cast<long long>(uv);
    long long dlt = 0;
    if (!tryProductCoeffUnitMs(cf, fac, dlt)) {
      ctx.p = pSave;
      setValidationError(ctx, STR_TIME_LITERAL_INVALID_SEGMENT);
      return false;
    }
    if (!tryAddTimeMsChecked(totalMs, dlt, totalMs)) {
      ctx.p = pSave;
      setValidationError(ctx, STR_TIME_LITERAL_INVALID_SEGMENT);
      return false;
    }
    ++comps;
    skipSpaces(ctx);
  }
  if (comps <= 0) {
    ctx.p = pSave;
    return false;
  }
  skipSpaces(ctx);
  const unsigned char c2 = static_cast<unsigned char>(*ctx.p);
  if (c2 >= static_cast<unsigned char>('a') && c2 <= static_cast<unsigned char>('z')) {
    const auto peekKeywordNoConsume = [](const char* p, const char* kw) -> bool {
      std::size_t n = 0;
      while (kw[n] != '\0') {
        const unsigned char a = static_cast<unsigned char>(p[n]);
        const unsigned char b = static_cast<unsigned char>(kw[n]);
        if (a == '\0' || std::tolower(a) != std::tolower(b)) {
          return false;
        }
        ++n;
      }
      const char t = p[n];
      return !(std::isalnum(static_cast<unsigned char>(t)) || t == '_');
    };
    const bool kwOk = peekKeywordNoConsume(ctx.p, STR_AND) || peekKeywordNoConsume(ctx.p, STR_OR) ||
        peekKeywordNoConsume(ctx.p, STR_NOT);
    if (!kwOk) {
      ctx.p = pSave;
      setValidationError(ctx, STR_TIME_COMPACT_INVALID_SUFFIX);
      return false;
    }
  }
  out = makeScalarTimeMs(totalMs);
  return true;
}
#endif

std::unique_ptr<MathParser::Expr> MathParser::parsePrimaryNumericLiteral(EvalContext& ctx) {
  auto emitNumericLiteral = [&](EvalValue litV) -> std::unique_ptr<Expr> {
    auto lit = std::make_unique<Expr>();
    lit->tag = Expr::Tag::Literal;
    lit->literalValue = std::move(litV);
    return lit;
  };
  auto finishNumericLiteral = [&](std::unique_ptr<Expr> lit) -> std::unique_ptr<Expr> {
    return wrapExprWithTightImagSuffixIfPresent(ctx, std::move(lit));
  };

  const char* parsedEnd = nullptr;
  std::uint64_t parsedUInt = 0;
  if (ctx.p[0] == '0') {
    const char px = static_cast<char>(std::tolower(static_cast<unsigned char>(ctx.p[1])));
    unsigned int radix = 0U;
    if (px == 'x') {
      radix = 16U;
    } else if (px == 'b') {
      radix = 2U;
    } else if (px == 'o') {
      radix = 8U;
    }
    if (radix != 0U) {
      if (!tryParsePrefixedUIntLiteral(ctx.p, px, radix, parsedEnd, parsedUInt)) {
        setInvalidPrefixedLiteralError(ctx, px);
        return nullptr;
      }
      ctx.p = parsedEnd;
      if (parsedUInt > static_cast<std::uint64_t>((std::numeric_limits<long long>::max)())) {
        return finishNumericLiteral(emitNumericLiteral(makeScalarUInt(parsedUInt)));
      }
      return finishNumericLiteral(emitNumericLiteral(makeScalarInt(static_cast<long long>(parsedUInt))));
    }
  }
#if SMARTMATH_TIME_VALUES
  if (getSupportTimeValues()) {
    switch (classifyNumericLiteralRoute(ctx.p, ctx.arrayIndexBracketDepth)) {
      case NumericLiteralRoute::ColonTime: {
        EvalValue timeLit;
        if (tryParseScalarTimeLiteral(ctx, timeLit)) {
          auto lit = std::make_unique<Expr>();
          lit->tag = Expr::Tag::Literal;
          lit->literalValue = std::move(timeLit);
          return lit;
        }
        if (ctx.parseError) {
          return nullptr;
        }
        break;
      }
      case NumericLiteralRoute::CompactTime: {
        EvalValue compactLit;
        if (tryParseCompactSuffixTimeLiteral(ctx, compactLit)) {
          auto lit = std::make_unique<Expr>();
          lit->tag = Expr::Tag::Literal;
          lit->literalValue = std::move(compactLit);
          return lit;
        }
        if (ctx.parseError) {
          return nullptr;
        }
        break;
      }
      case NumericLiteralRoute::Plain:
        break;
    }
  }
#endif
  const char* const numStart = ctx.p;
  const char* const numEnd = scanDecimalNumericLiteralEnd(numStart);
  if (numEnd == numStart) {
    if (*numStart == '.') {
      setParseError(ctx, ParseErrorId::UnexpectedToken);
    } else {
      setParseError(ctx, ParseErrorId::InvalidNumericLiteral);
    }
    return nullptr;
  }

  if (tryParseInputNumberAsInteger(numStart, numEnd, parsedUInt)) {
    ctx.p = numEnd;
    if (parsedUInt <= static_cast<std::uint64_t>((std::numeric_limits<long long>::max)())) {
      return finishNumericLiteral(emitNumericLiteral(makeScalarInt(static_cast<long long>(parsedUInt))));
    }
    return finishNumericLiteral(emitNumericLiteral(makeScalarUInt(parsedUInt)));
  }

  char* end = nullptr;
  const double d = std::strtod(numStart, &end);
  if (end == numStart) {
    if (*numStart == '.') {
      setParseError(ctx, ParseErrorId::UnexpectedToken);
    } else {
      setParseError(ctx, ParseErrorId::InvalidNumericLiteral);
    }
    return nullptr;
  }
  if ((*end == 'e' || *end == 'E')) {
    const char* pExp = end + 1;
    if (*pExp == '+' || *pExp == '-') {
      ++pExp;
    }
    if (!std::isdigit(static_cast<unsigned char>(*pExp))) {
      setParseError(ctx, ParseErrorId::UnexpectedToken);
      return nullptr;
    }
  }
  ctx.p = end;
  return finishNumericLiteral(emitNumericLiteral(makeScalar(d)));
}

std::unique_ptr<MathParser::Expr> MathParser::parsePrimaryIdentifierOrCall(EvalContext& ctx) {
  const char* identStart = ctx.p;
  if (!isIdentStart(*ctx.p)) {
    setParseError(ctx, ParseErrorId::UnexpectedToken);
    return nullptr;
  }
  const char* identEnd = ctx.p + 1;
  while (isIdentChar(*identEnd)) {
    ++identEnd;
  }
  ctx.p = identEnd;
  skipSpaces(ctx);
  if (*ctx.p != '(') {
    if (!MathParser::identMayBeBareBuiltinName(identStart, identEnd) && userFunctionIndex_.empty()
        && ctx.compilingUserFunctionParams == nullptr) {
      auto v = std::make_unique<Expr>();
      v->tag = Expr::Tag::Variable;
      assignLowerIdentFromRange(v->name, identStart, identEnd);
      return v;
    }
    std::string ident;
    assignLowerIdentFromRange(ident, identStart, identEnd);
    if (MathParser::identMayBeBareBuiltinName(ident) && trySetMissingFunctionCallError(ctx, ident, identStart)) {
      return nullptr;
    }
    if ((!userFunctionIndex_.empty() || ctx.compilingUserFunctionParams != nullptr)
        && trySetBareUserFunctionNameError(ctx, ident, identStart)) {
      return nullptr;
    }
    auto v = std::make_unique<Expr>();
    v->tag = Expr::Tag::Variable;
    v->name = std::move(ident);
    return v;
  }
  std::string ident;
  assignLowerIdentFromRange(ident, identStart, identEnd);
  ++ctx.p;
  skipSpaces(ctx);
  if (trySetIncompleteOpenedFunctionCallHint(ctx, ident)) {
    return nullptr;
  }
  std::vector<std::unique_ptr<Expr>> args;
  const bool isSortbyCall = (ident == STR_SORTBY);
  if (isSortbyCall) {
    if (!parseSortbyCallArguments(ctx, args)) {
      return nullptr;
    }
  } else if (!parseParenthesizedExprList(ctx, args)) {
    return nullptr;
  }
  if (*ctx.p != ')') {
    setMissingClosingParenLikeError(ctx);
    return nullptr;
  }
  ++ctx.p;
  auto call = std::make_unique<Expr>();
  call->tag = Expr::Tag::Call;
  call->name = std::move(ident);
  BuiltinFunctionId id = BuiltinFunctionId::Count;
  if (tryGetBuiltinFunctionId(call->name, id)) {
    call->builtinFunctionId = id;
  }
  call->elements = std::move(args);
  return call;
}

std::unique_ptr<MathParser::Expr> MathParser::parseUnary(EvalContext& ctx) {
  skipSpaces(ctx);
  if (*ctx.p == '+') {
    ++ctx.p;
    return parseUnary(ctx);
  }
  if (*ctx.p == '-') {
    ++ctx.p;
    auto inner = parseUnary(ctx);
    if (hasExprParseFailure(ctx, inner)) {
      return nullptr;
    }
    auto u = std::make_unique<Expr>();
    u->tag = Expr::Tag::Unary;
    u->unaryOp = '-';
    u->child = std::move(inner);
    return u;
  }
  if (consumeKeyword(ctx, opName(OperatorNameId::Not).c_str())) {
    auto inner = parseLogicalNot(ctx);
    if (hasExprParseFailure(ctx, inner)) {
      return nullptr;
    }
    auto u = std::make_unique<Expr>();
    u->tag = Expr::Tag::Unary;
    u->unaryOp = 'N';
    u->child = std::move(inner);
    return u;
  }
  if (*ctx.p == '!') {
    ++ctx.p;
    auto inner = parseUnary(ctx);
    if (hasExprParseFailure(ctx, inner)) {
      return nullptr;
    }
    auto u = std::make_unique<Expr>();
    u->tag = Expr::Tag::Unary;
    u->unaryOp = 'N';
    u->child = std::move(inner);
    return u;
  }
  if (*ctx.p == '~') {
    ++ctx.p;
    auto inner = parseUnary(ctx);
    if (hasExprParseFailure(ctx, inner)) {
      return nullptr;
    }
    auto u = std::make_unique<Expr>();
    u->tag = Expr::Tag::Unary;
    u->unaryOp = '~';
    u->child = std::move(inner);
    return u;
  }
  auto prim = parsePower(ctx);
  if (hasExprParseFailure(ctx, prim)) {
    return nullptr;
  }
  while (true) {
    skipSpaces(ctx);
    if (*ctx.p == '[') {
      ++ctx.p;
      ++ctx.arrayIndexBracketDepth;
      skipSpaces(ctx);
      if (*ctx.p == ']') {
        --ctx.arrayIndexBracketDepth;
        setParseError(ctx, ParseErrorId::MissingIndex);
        return nullptr;
      }

      bool hasStart = false;
      bool hasStop = false;
      bool hasStep = false;
      std::unique_ptr<Expr> startExpr;
      std::unique_ptr<Expr> stopExpr;
      std::unique_ptr<Expr> stepExpr;

      if (*ctx.p == ':') {
        hasStart = false;
        ++ctx.p;
      } else {
        startExpr = parseExpression(ctx);
        if (ctx.parseError || !startExpr) {
          --ctx.arrayIndexBracketDepth;
          return nullptr;
        }
        hasStart = true;
        skipSpaces(ctx);
        if (*ctx.p == ':') {
          ++ctx.p;
        } else if (*ctx.p == ']') {
          ++ctx.p;
          --ctx.arrayIndexBracketDepth;
          auto w = std::make_unique<Expr>();
          w->tag = Expr::Tag::Index;
          w->left = std::move(prim);
          w->right = std::move(startExpr);
          prim = std::move(w);
          continue;
        } else {
          --ctx.arrayIndexBracketDepth;
          setParseError(ctx, ParseErrorId::MissingClosingBracket);
          return nullptr;
        }
      }

      // Slice form: [start:stop:step] with any bound omitted.
      skipSpaces(ctx);
      if (*ctx.p == ']') {
        hasStop = false;
        hasStep = false;
      } else if (*ctx.p == ':') {
        hasStop = false;
        ++ctx.p;
        skipSpaces(ctx);
        if (*ctx.p == ']') {
          hasStep = false;
        } else {
          stepExpr = parseExpression(ctx);
          if (ctx.parseError || !stepExpr) {
            --ctx.arrayIndexBracketDepth;
            return nullptr;
          }
          hasStep = true;
          skipSpaces(ctx);
          if (*ctx.p != ']') {
            --ctx.arrayIndexBracketDepth;
            setParseError(ctx, ParseErrorId::MissingClosingBracket);
            return nullptr;
          }
        }
      } else {
        stopExpr = parseExpression(ctx);
        if (ctx.parseError || !stopExpr) {
          --ctx.arrayIndexBracketDepth;
          return nullptr;
        }
        hasStop = true;
        skipSpaces(ctx);
        if (*ctx.p == ':') {
          ++ctx.p;
          skipSpaces(ctx);
          if (*ctx.p == ']') {
            hasStep = false;
          } else {
            stepExpr = parseExpression(ctx);
            if (ctx.parseError || !stepExpr) {
              --ctx.arrayIndexBracketDepth;
              return nullptr;
            }
            hasStep = true;
            skipSpaces(ctx);
            if (*ctx.p != ']') {
              --ctx.arrayIndexBracketDepth;
              setParseError(ctx, ParseErrorId::MissingClosingBracket);
              return nullptr;
            }
          }
        } else if (*ctx.p == ']') {
          hasStep = false;
        } else {
          --ctx.arrayIndexBracketDepth;
          setParseError(ctx, ParseErrorId::MissingClosingBracket);
          return nullptr;
        }
      }

      ++ctx.p;
      --ctx.arrayIndexBracketDepth;
      auto w = std::make_unique<Expr>();
      w->tag = Expr::Tag::Slice;
      w->left = std::move(prim);
      w->elements.resize(3);
      if (hasStart) {
        w->elements[0] = std::move(startExpr);
      }
      if (hasStop) {
        w->elements[1] = std::move(stopExpr);
      }
      if (hasStep) {
        w->elements[2] = std::move(stepExpr);
      }
      prim = std::move(w);
      continue;
    }
    if (isPercentPostfixContext(ctx.p)) {
      ++ctx.p;
      auto w = std::make_unique<Expr>();
      w->tag = Expr::Tag::PostfixPercent;
      w->child = std::move(prim);
      prim = std::move(w);
      continue;
    }
    break;
  }
  return prim;
}

std::unique_ptr<MathParser::Expr> MathParser::parsePower(EvalContext& ctx) {
  auto left = parsePrimary(ctx);
  if (ctx.parseError || !left) {
    return nullptr;
  }
  skipSpaces(ctx);
  std::unique_ptr<Expr> out;
  // AST parser uses '**' for exponentiation.
  if (ctx.p[0] == '*' && ctx.p[1] == '*') {
    ctx.p += 2;
    auto right = parseUnary(ctx);
    if (ctx.parseError || !right) {
      return nullptr;
    }
    out = std::make_unique<Expr>();
    out->tag = Expr::Tag::Binary;
    out->binaryOp = Expr::BinaryOp::Pow;
    out->left = std::move(left);
    out->right = std::move(right);
  } else {
    out = std::move(left);
  }
  return out;
}

std::unique_ptr<MathParser::Expr> MathParser::parseMulDivMod(EvalContext& ctx) {
  return parseLeftAssocBinary(ctx, &MathParser::parseUnary, &MathParser::tryConsumeMulDivModOp);
}

std::unique_ptr<MathParser::Expr> MathParser::parseShift(EvalContext& ctx) {
  return parseLeftAssocBinary(ctx, &MathParser::parseAddSub, &MathParser::tryConsumeShiftOp);
}

std::unique_ptr<MathParser::Expr> MathParser::parseBitAnd(EvalContext& ctx) {
  return parseLeftAssocBinary(ctx, &MathParser::parseShift, &MathParser::tryConsumeBitAndOp);
}

std::unique_ptr<MathParser::Expr> MathParser::parseBitXor(EvalContext& ctx) {
  return parseLeftAssocBinary(ctx, &MathParser::parseBitAnd, &MathParser::tryConsumeBitXorOp);
}

std::unique_ptr<MathParser::Expr> MathParser::parseBitOr(EvalContext& ctx) {
  return parseLeftAssocBinary(ctx, &MathParser::parseBitXor, &MathParser::tryConsumeBitOrOp);
}

std::unique_ptr<MathParser::Expr> MathParser::parseAddSub(EvalContext& ctx) {
  return parseLeftAssocBinary(ctx, &MathParser::parseMulDivMod, &MathParser::tryConsumeAddSubOp, true);
}

std::unique_ptr<MathParser::Expr> MathParser::parseCompare(EvalContext& ctx) {
  return parseLeftAssocBinary(ctx, &MathParser::parseBitOr, &MathParser::tryConsumeCompareOp);
}

std::unique_ptr<MathParser::Expr> MathParser::parseLogicalNot(EvalContext& ctx) {
  skipSpaces(ctx);
  if (consumeKeyword(ctx, opName(OperatorNameId::Not).c_str())) {
    auto inner = parseLogicalNot(ctx);
    if (ctx.parseError || !inner) {
      return nullptr;
    }
    auto u = std::make_unique<Expr>();
    u->tag = Expr::Tag::Unary;
    u->unaryOp = 'N';
    u->child = std::move(inner);
    return u;
  }
  return parseCompare(ctx);
}

std::unique_ptr<MathParser::Expr> MathParser::parseAnd(EvalContext& ctx) {
  return parseLeftAssocBinary(ctx, &MathParser::parseLogicalNot, &MathParser::tryConsumeLogicalAndOp);
}

std::unique_ptr<MathParser::Expr> MathParser::parseOr(EvalContext& ctx) {
  return parseLeftAssocBinary(ctx, &MathParser::parseAnd, &MathParser::tryConsumeLogicalOrOp);
}

std::unique_ptr<MathParser::Expr> MathParser::parseExpression(EvalContext& ctx) {
  skipSpaces(ctx);
  ctx.start = ctx.p;
  return parseOr(ctx);
}

bool MathParser::tryAppendTrailingFormatterSugarStatement(
    EvalContext& ctx,
    std::vector<AstStatement>& out,
    bool& handled) {
  handled = false;
  if (out.empty() || !isIdentStart(*ctx.p)) {
    return true;
  }
  const char* sugarStart = ctx.p;
  const char* q = ctx.p + 1;
  while (isIdentChar(*q)) {
    ++q;
  }
  const std::string ident = toLower(std::string(sugarStart, static_cast<std::size_t>(q - sugarStart)));
  if (!isTrailingFormatterFunctionName(ident)) {
    return true;
  }
  const char* sugarEnd = q;
  while (*sugarEnd != '\0' && std::isspace(static_cast<unsigned char>(*sugarEnd))) {
    ++sugarEnd;
  }
  bool sugarOk = false;
  if (*sugarEnd == ';' || *sugarEnd == '\0') {
    sugarOk = true;
  } else if (*sugarEnd == '(') {
    const char* r = sugarEnd + 1;
    while (*r != '\0' && std::isspace(static_cast<unsigned char>(*r))) {
      ++r;
    }
    if (*r == ')') {
      ++r;
      while (*r != '\0' && std::isspace(static_cast<unsigned char>(*r))) {
        ++r;
      }
      if (*r == ';' || *r == '\0') {
        sugarOk = true;
        sugarEnd = r;
      }
    }
  }
  if (!sugarOk) {
    return true;
  }
  auto callExpr = std::make_unique<Expr>();
  callExpr->tag = Expr::Tag::Call;
  callExpr->name = ident;
  BuiltinFunctionId id = BuiltinFunctionId::Count;
  if (tryGetBuiltinFunctionId(callExpr->name, id)) {
    callExpr->builtinFunctionId = id;
  }
  auto ansExpr = std::make_unique<Expr>();
  ansExpr->tag = Expr::Tag::Variable;
  ansExpr->name = STR_ANS;
  callExpr->elements.emplace_back(std::move(ansExpr));
  AstStatement st;
  st.kind = AstStatement::Kind::Expr;
  st.expr = std::move(callExpr);
  out.emplace_back(std::move(st));
  ctx.p = sugarEnd;
  handled = true;
  return true;
}

bool MathParser::tryCompileSingleExpressionProgram(EvalContext& ctx, std::vector<AstStatement>& out) {
  out.clear();
  ctx.compilingUserFunctionParams = nullptr;
  skipSpaces(ctx);
  if (*ctx.p == '\0') {
    return false;
  }
  if (*ctx.p == ';') {
    setError(ctx, STR_EMPTY_STATEMENT);
    return false;
  }
  if (std::strchr(ctx.p, ';') != nullptr) {
    return false;
  }
  if (peekIdentFollowedByChar(ctx.p, '(')) {
    return false;
  }
  bool statementHandled = false;
  if (!tryAppendTrailingFormatterSugarStatement(ctx, out, statementHandled)) {
    return false;
  }
  if (!statementHandled && !tryAppendAssignOrExpressionStatement(ctx, out)) {
    return false;
  }
  skipSpaces(ctx);
  if (!ctx.parseError && *ctx.p != '\0') {
    setParseError(ctx, ParseErrorId::UnexpectedInput);
    return false;
  }
  return !ctx.parseError && !out.empty();
}

bool MathParser::parseProgram(EvalContext& ctx, std::vector<AstStatement>& out) {
  out.clear();
  std::unique_ptr<std::unordered_map<std::string, std::vector<std::string>>> compilingUserFunctionParams;
  ctx.compilingUserFunctionParams = nullptr;
  while (true) {
    skipSpaces(ctx);
    if (*ctx.p == '\0') {
      break;
    }
    if (*ctx.p == ';') {
      setError(ctx, STR_EMPTY_STATEMENT);
      ctx.compilingUserFunctionParams = nullptr;
      return false;
    }
    const std::size_t stmtBefore = out.size();
    const char* save = ctx.p;
    if (peekIdentFollowedByChar(save, '(') && tryAppendFunctionDefinitionStatement(ctx, out)) {
      // handled above
    } else {
      ctx.p = save;
      skipSpaces(ctx);
      bool statementHandled = false;
      if (!tryAppendTrailingFormatterSugarStatement(ctx, out, statementHandled)) {
        ctx.compilingUserFunctionParams = nullptr;
        return false;
      }
      if (!statementHandled) {
        if (!tryAppendAssignOrExpressionStatement(ctx, out)) {
          ctx.compilingUserFunctionParams = nullptr;
          return false;
        }
      }
    }
    if (out.size() > stmtBefore) {
      if (out.back().kind == AstStatement::Kind::FunDef) {
        if (!compilingUserFunctionParams) {
          compilingUserFunctionParams = std::make_unique<std::unordered_map<std::string, std::vector<std::string>>>();
          ctx.compilingUserFunctionParams = compilingUserFunctionParams.get();
        }
        (*compilingUserFunctionParams)[out.back().fun.name] = out.back().fun.params;
      } else if (out.back().kind == AstStatement::Kind::Assign && compilingUserFunctionParams) {
        compilingUserFunctionParams->erase(out.back().assignName);
      }
    }
    if (consumeProgramStatementSeparator(ctx)) continue;
    if (*ctx.p == '\0') break;
    ctx.compilingUserFunctionParams = nullptr;
    return false;
  }
  ctx.compilingUserFunctionParams = nullptr;
  return true;
}

bool MathParser::exprIsDirectPostfixPercent(const Expr& e) {
  if (e.tag == Expr::Tag::PostfixPercent) {
    return true;
  }
  if (e.tag == Expr::Tag::Unary && e.child) {
    return exprIsDirectPostfixPercent(*e.child);
  }
  if (e.tag == Expr::Tag::ArrayOrParens && e.elements.size() == 1U && e.elements[0]) {
    return exprIsDirectPostfixPercent(*e.elements[0]);
  }
  return false;
}

bool MathParser::exprIsScalarOnly(const Expr& e) const {
  switch (e.tag) {
    case Expr::Tag::Literal:
      if (e.literalValue.kind == ValueKind::InlineLambda) {
        return false;
      }
#if SMARTMATH_COMPLEX_NUMBERS
      if (e.literalValue.kind == ValueKind::Scalar &&
          scalarHasNonzeroImaginaryPart(e.literalValue.scalarValue)) {
        return false;
      }
#endif
      return true;
    case Expr::Tag::Variable:
      return true;
    case Expr::Tag::Unary:
    case Expr::Tag::PostfixPercent:
      return e.child ? exprIsScalarOnly(*e.child) : false;
    case Expr::Tag::Binary:
      return e.left && e.right && exprIsScalarOnly(*e.left) && exprIsScalarOnly(*e.right);
    case Expr::Tag::ArrayOrParens:
      if (e.elements.size() != 1U) return false;
      return exprIsScalarOnly(*e.elements[0]);
    case Expr::Tag::Call:
    case Expr::Tag::Index:
    case Expr::Tag::Slice:
    case Expr::Tag::FunctionRef:
      return false;
    default:
      return false;
  }
}

bool MathParser::programIsScalarOnly(const std::vector<AstStatement>& program) const {
  for (const auto& st : program) {
    if (st.kind == AstStatement::Kind::FunDef) {
      return false;
    }
    if (!st.expr || !exprIsScalarOnly(*st.expr)) {
      return false;
    }
  }
  return true;
}

void MathParser::bindExprVariableRefs(
    Expr& e,
    bool hasAssignments,
    const std::unordered_map<std::string, EvalValue>& variables) {
  if (e.tag == Expr::Tag::Variable) {
    if (hasAssignments) {
      e.boundVariable = nullptr;
      return;
    }
    auto it = variables.find(e.name);
    e.boundVariable = (it != variables.end()) ? &it->second : nullptr;
    return;
  }
  if (e.child) {
    bindExprVariableRefs(*e.child, hasAssignments, variables);
  }
  if (e.left) {
    bindExprVariableRefs(*e.left, hasAssignments, variables);
  }
  if (e.right) {
    bindExprVariableRefs(*e.right, hasAssignments, variables);
  }
  for (auto& item : e.elements) {
    if (item) {
      bindExprVariableRefs(*item, hasAssignments, variables);
    }
  }
}

void MathParser::bindCompiledVariableRefs(CompiledProgramState& state) {
  for (auto& st : state.program) {
    if (st.expr) {
      bindExprVariableRefs(*st.expr, state.hasAssignments, variables_);
    }
  }
  state.boundVariablesVersion = variablesVersion_;
}

MathParser::EvalValue MathParser::evalUnaryExpr(
    const Expr& e,
    EvalContext& ctx,
    const std::unordered_map<std::string, EvalValue>* scopedVars,
    bool scalarOnlyMode) {
  const auto evalChild = [&](const Expr& childExpr) -> EvalValue {
    return scalarOnlyMode ? evalExprScalar(childExpr, ctx, scopedVars) : evalExpr(childExpr, ctx, scopedVars);
  };

  if (e.unaryOp == '+') {
    return evalChild(*e.child);
  }

  EvalValue v = evalChild(*e.child);
  if (ctx.parseError) {
    return v;
  }

  if (e.unaryOp == '-') {
    return negateEvalValue(v);
  }
  if (e.unaryOp == 'N') {
    return makeScalarInt(isTruthy(v) ? 0LL : 1LL);
  }
  if (e.unaryOp == '~') {
#if SMARTMATH_COMPLEX_NUMBERS
    if (getSupportComplexNumbers() && MathParser::evalValueHasNonzeroImaginary(v)) {
      setParseError(ctx, ParseErrorId::IncompatibleOperands);
      return makeScalar(0);
    }
#endif
    if (!scalarOnlyMode && v.kind != ValueKind::Scalar) {
      setParseError(ctx, ParseErrorId::BitwiseIntegerOperands);
      return makeScalar(0);
    }
    long long iv = 0;
    if (!tryGetSignedInt64FromScalar(v.scalarValue, iv)) {
      setParseError(ctx, ParseErrorId::BitwiseIntegerOperands);
      return makeScalar(0);
    }
    return makeScalarInt(~iv);
  }

  setParseError(ctx, ParseErrorId::InternalUnaryOp);
  return makeScalar(0);
}

MathParser::EvalValue MathParser::evalMappedBinaryOp(
    EvalContext& ctx,
    Expr::BinaryOp op,
    const EvalValue& left,
    const EvalValue& right) const {
  const char opChar = (op == Expr::BinaryOp::Pow) ? '^' :
                      (op == Expr::BinaryOp::Mul) ? '*' :
                      (op == Expr::BinaryOp::Div) ? '/' :
                      (op == Expr::BinaryOp::Add) ? '+' : '-';
#if SMARTMATH_TIME_VALUES
  if (rejectNumericBinaryPowWithTime(ctx, left, right, opChar)) {
    return makeScalar(0);
  }
#endif
  bool ok = false;
  EvalValue out = mapBinary(ctx, left, right, opChar, ok);
  if (ok) {
    return out;
  }

  if (ctx.parseError) {
    return makeScalar(0);
  }

  if (op == Expr::BinaryOp::Pow) {
    setParseError(ctx, ParseErrorId::NumericErrorInPowerOperation);
  } else if (left.kind == ValueKind::Array && right.kind == ValueKind::Array && left.arr.size() != right.arr.size()) {
    setParseError(ctx, ParseErrorId::IncompatibleOperands);
  } else {
    setParseError(ctx, ParseErrorId::NumericErrorInExpression);
  }
  return makeScalar(0);
}

MathParser::EvalValue MathParser::evalInt64BinaryOp(
    EvalContext& ctx,
    const EvalValue& left,
    const EvalValue& right,
    Expr::BinaryOp op) const {
  const bool isModulo = (op == Expr::BinaryOp::Modulo);
  if (rejectInt64BinaryOperands(ctx, left, right, isModulo)) {
    return makeScalar(0);
  }
  const auto setIntegerOperandError = [&]() {
    if (isModulo) setParseError(ctx, ParseErrorId::ModuloIntegerOperands);
    else setParseError(ctx, ParseErrorId::BitwiseIntegerOperands);
  };
  const auto applyOp = [&](const EvalValue::ScalarValue& lv, const EvalValue::ScalarValue& rv, EvalValue& outS) -> bool {
    if (op == Expr::BinaryOp::ShiftLeft) {
      std::uint64_t aU = 0;
      std::uint64_t bU = 0;
      if (tryGetBothExactNonNegativeUInt64FromScalars(lv, rv, aU, bU)) {
        if (!tryShiftLeftU64ExactOrMaybe(aU, bU, outS)) {
          setParseError(ctx, ParseErrorId::IncompatibleOperands);
          return false;
        }
        return true;
      }
    }

    if (lv.hasExactUInt64() && rv.hasExactUInt64()) {
      const std::uint64_t aU = lv.exactUInt64;
      const std::uint64_t bU = rv.exactUInt64;
      if (isModulo) {
        if (bU == 0) {
          setParseError(ctx, ParseErrorId::IncompatibleOperands);
          return false;
        }
        outS = makeScalarUInt(aU % bU);
        return true;
      }
      if ((op == Expr::BinaryOp::ShiftLeft || op == Expr::BinaryOp::ShiftRight) && bU > 63u) {
        setParseError(ctx, ParseErrorId::IncompatibleOperands);
        return false;
      }
      if (op == Expr::BinaryOp::ShiftLeft) {
        if (!tryShiftLeftU64ExactOrMaybe(aU, bU, outS)) {
          setParseError(ctx, ParseErrorId::IncompatibleOperands);
          return false;
        }
        return true;
      }
      if (op == Expr::BinaryOp::ShiftRight) {
        outS = makeScalarUInt(aU >> bU);
        return true;
      }
      if (op == Expr::BinaryOp::BitAnd) {
        outS = makeScalarUInt(aU & bU);
        return true;
      }
      if (op == Expr::BinaryOp::BitOr) {
        outS = makeScalarUInt(aU | bU);
        return true;
      }
      if (op == Expr::BinaryOp::BitXor) {
        outS = makeScalarUInt(aU ^ bU);
        return true;
      }
    }

    long long a = 0, b = 0;
    if (!tryGetSignedInt64FromScalar(lv, a) || !tryGetSignedInt64FromScalar(rv, b)) return false;
    if (isModulo) {
      if (b == 0) {
        setParseError(ctx, ParseErrorId::IncompatibleOperands);
        return false;
      }
      outS = makeScalarInt(a % b);
      return true;
    }
    if ((op == Expr::BinaryOp::ShiftLeft || op == Expr::BinaryOp::ShiftRight) && (b < 0 || b > 63)) {
      setParseError(ctx, ParseErrorId::IncompatibleOperands);
      return false;
    }
    const long long out = (op == Expr::BinaryOp::BitAnd) ? (a & b) :
                          (op == Expr::BinaryOp::BitOr) ? (a | b) :
                          (op == Expr::BinaryOp::BitXor) ? (a ^ b) :
                          (op == Expr::BinaryOp::ShiftLeft) ? bitwiseShiftLeftDefined(a, static_cast<unsigned int>(b)) :
                                                               bitwiseShiftRightDefined(a, static_cast<unsigned int>(b));
    outS = makeScalarInt(out);
    return true;
  };
  if (left.kind == ValueKind::Array && right.kind == ValueKind::Array && left.arr.size() != right.arr.size()) {
    setParseError(ctx, ParseErrorId::IncompatibleOperands);
    return makeScalar(0);
  }
  bool ok = true;
  EvalValue out = mapBinaryBroadcast(left, right, applyOp, ok);
  if (!ok && !ctx.parseError) {
    setIntegerOperandError();
    return makeScalar(0);
  }
  return out;
}

bool MathParser::isComparisonBinaryOp(Expr::BinaryOp op) {
  return op == Expr::BinaryOp::CmpLt || op == Expr::BinaryOp::CmpGt ||
         op == Expr::BinaryOp::CmpLe || op == Expr::BinaryOp::CmpGe ||
         op == Expr::BinaryOp::CmpEq || op == Expr::BinaryOp::CmpNe;
}

bool MathParser::evalComparisonByOp(Expr::BinaryOp op, int cmp) {
  switch (op) {
    case Expr::BinaryOp::CmpLt: return cmp < 0;
    case Expr::BinaryOp::CmpGt: return cmp > 0;
    case Expr::BinaryOp::CmpLe: return cmp <= 0;
    case Expr::BinaryOp::CmpGe: return cmp >= 0;
    case Expr::BinaryOp::CmpEq: return cmp == 0;
    case Expr::BinaryOp::CmpNe: return cmp != 0;
    default: return false;
  }
}

bool MathParser::evalComparisonTruthWhenUnorderedNan(Expr::BinaryOp op) {
  switch (op) {
    case Expr::BinaryOp::CmpNe:
      return true;
    default:
      return false;
  }
}

MathParser::EvalValue MathParser::evalExprScalar(
    const Expr& e,
    EvalContext& ctx,
    const std::unordered_map<std::string, EvalValue>* scopedVars) {
  switch (e.tag) {
    case Expr::Tag::Literal:
      return e.literalValue;
    case Expr::Tag::Variable: {
      EvalValue out = makeScalar(0);
      if (!tryResolveVariableValue(e, scopedVars, out)) {
        if (handleUnknownIdentifier(ctx, e.name, ctx.unknownVarsText)) {
          return makeScalar(0);
        }
        return makeScalar(0);
      }
      if (out.kind != ValueKind::Scalar) {
        setParseError(ctx, ParseErrorId::ScalarOnlyExpressionEncounteredNon);
        return makeScalar(0);
      }
      return out;
    }
    case Expr::Tag::Unary:
      return evalUnaryExpr(e, ctx, scopedVars, true);
    case Expr::Tag::PostfixPercent: {
      EvalValue v = evalExprScalar(*e.child, ctx, scopedVars);
      if (ctx.parseError) return v;
#if SMARTMATH_TIME_VALUES
      if (scalarValueIsTime(v.scalarValue)) {
        setParseError(ctx, ParseErrorId::IncompatibleOperands);
        return makeScalar(0);
      }
#endif
#if SMARTMATH_COMPLEX_NUMBERS
      if (getSupportComplexNumbers() && scalarHasNonzeroImaginaryPart(v.scalarValue)) {
        setParseError(ctx, ParseErrorId::IncompatibleOperands);
        return makeScalar(0);
      }
#endif
      v.scalarValue.scalar /= 100.0;
      v.scalarValue.setExactInt64Valid(false);
      v.scalarValue.setExactUInt64Valid(false);
      v.scalarValue.scalarKind = ScalarKind::FloatingPoint;
      return v;
    }
    case Expr::Tag::ArrayOrParens:
      return evalExprScalar(*e.elements[0], ctx, scopedVars);
    case Expr::Tag::Binary: {
      EvalValue l = evalExprScalar(*e.left, ctx, scopedVars);
      if (ctx.parseError) return l;
      EvalValue r = evalExprScalar(*e.right, ctx, scopedVars);
      if (ctx.parseError) return r;
      switch (e.binaryOp) {
        case Expr::BinaryOp::LogicalOr:
          return makeScalarInt((isTruthy(l) || isTruthy(r)) ? 1LL : 0LL);
        case Expr::BinaryOp::LogicalAnd:
          return makeScalarInt((isTruthy(l) && isTruthy(r)) ? 1LL : 0LL);
        case Expr::BinaryOp::Modulo: {
#if SMARTMATH_COMPLEX_NUMBERS
          if (getSupportComplexNumbers() &&
              (MathParser::evalValueHasNonzeroImaginary(l) || MathParser::evalValueHasNonzeroImaginary(r))) {
            setParseError(ctx, ParseErrorId::IncompatibleOperands);
            return makeScalar(0);
          }
#endif
          if (l.scalarValue.hasExactUInt64() && r.scalarValue.hasExactUInt64() &&
              (!l.scalarValue.hasExactInt64() || !r.scalarValue.hasExactInt64())) {
            if (r.scalarValue.exactUInt64 == 0u) {
              setParseError(ctx, ParseErrorId::IncompatibleOperands);
              return makeScalar(0);
            }
            return makeScalarUInt(l.scalarValue.exactUInt64 % r.scalarValue.exactUInt64);
          }
          long long a = 0, b = 0;
          if (!tryGetSignedInt64FromScalar(l.scalarValue, a) || !tryGetSignedInt64FromScalar(r.scalarValue, b)) {
            setParseError(ctx, ParseErrorId::ModuloIntegerOperands);
            return makeScalar(0);
          }
          if (b == 0) {
            setParseError(ctx, ParseErrorId::IncompatibleOperands);
            return makeScalar(0);
          }
          return makeScalarInt(a % b);
        }
        case Expr::BinaryOp::BitAnd:
        case Expr::BinaryOp::BitOr:
        case Expr::BinaryOp::BitXor:
        case Expr::BinaryOp::ShiftLeft:
        case Expr::BinaryOp::ShiftRight: {
#if SMARTMATH_COMPLEX_NUMBERS
          if (getSupportComplexNumbers() &&
              (MathParser::evalValueHasNonzeroImaginary(l) || MathParser::evalValueHasNonzeroImaginary(r))) {
            setParseError(ctx, ParseErrorId::IncompatibleOperands);
            return makeScalar(0);
          }
#endif
          if (e.binaryOp == Expr::BinaryOp::ShiftLeft) {
            std::uint64_t aU = 0;
            std::uint64_t bU = 0;
            if (tryGetBothExactNonNegativeUInt64FromScalars(l.scalarValue, r.scalarValue, aU, bU)) {
              EvalValue outS;
              if (!tryShiftLeftU64ExactOrMaybe(aU, bU, outS)) {
                setParseError(ctx, ParseErrorId::IncompatibleOperands);
                return makeScalar(0);
              }
              return outS;
            }
          }
          if (l.scalarValue.hasExactUInt64() && r.scalarValue.hasExactUInt64() &&
              (!l.scalarValue.hasExactInt64() || !r.scalarValue.hasExactInt64())) {
            const std::uint64_t aU = l.scalarValue.exactUInt64;
            const std::uint64_t bU = r.scalarValue.exactUInt64;
            if (bU > 63u) {
              setParseError(ctx, ParseErrorId::IncompatibleOperands);
              return makeScalar(0);
            }
            if (e.binaryOp == Expr::BinaryOp::BitAnd) return makeScalarUInt(aU & bU);
            if (e.binaryOp == Expr::BinaryOp::BitOr) return makeScalarUInt(aU | bU);
            if (e.binaryOp == Expr::BinaryOp::BitXor) return makeScalarUInt(aU ^ bU);
            if (e.binaryOp == Expr::BinaryOp::ShiftLeft) {
              EvalValue outS;
              if (!tryShiftLeftU64ExactOrMaybe(aU, bU, outS)) {
                setParseError(ctx, ParseErrorId::IncompatibleOperands);
                return makeScalar(0);
              }
              return outS;
            }
            return makeScalarUInt(aU >> bU);
          }
          long long a = 0, b = 0;
          if (!tryGetSignedInt64FromScalar(l.scalarValue, a) || !tryGetSignedInt64FromScalar(r.scalarValue, b)) {
            setParseError(ctx, ParseErrorId::BitwiseIntegerOperands);
            return makeScalar(0);
          }
          if ((e.binaryOp == Expr::BinaryOp::ShiftLeft || e.binaryOp == Expr::BinaryOp::ShiftRight) &&
              (b < 0 || b > 63)) {
            setParseError(ctx, ParseErrorId::IncompatibleOperands);
            return makeScalar(0);
          }
          if (e.binaryOp == Expr::BinaryOp::BitAnd) return makeScalarInt(a & b);
          if (e.binaryOp == Expr::BinaryOp::BitOr) return makeScalarInt(a | b);
          if (e.binaryOp == Expr::BinaryOp::BitXor) return makeScalarInt(a ^ b);
          if (e.binaryOp == Expr::BinaryOp::ShiftLeft) return makeScalarInt(bitwiseShiftLeftDefined(a, static_cast<unsigned int>(b)));
          return makeScalarInt(bitwiseShiftRightDefined(a, static_cast<unsigned int>(b)));
        }
        case Expr::BinaryOp::CmpLt:
        case Expr::BinaryOp::CmpGt:
        case Expr::BinaryOp::CmpLe:
        case Expr::BinaryOp::CmpGe:
        case Expr::BinaryOp::CmpEq:
        case Expr::BinaryOp::CmpNe: {
#if SMARTMATH_TIME_VALUES
          const bool lTime = scalarValueIsTime(l.scalarValue);
          const bool rTime = scalarValueIsTime(r.scalarValue);
#else
          const bool lTime = false;
          const bool rTime = false;
#endif
          const bool lIm = getSupportComplexNumbers() && scalarHasNonzeroImaginaryPart(l.scalarValue);
          const bool rIm = getSupportComplexNumbers() && scalarHasNonzeroImaginaryPart(r.scalarValue);
          const bool orderingOp = e.binaryOp == Expr::BinaryOp::CmpLt || e.binaryOp == Expr::BinaryOp::CmpLe ||
                                  e.binaryOp == Expr::BinaryOp::CmpGt || e.binaryOp == Expr::BinaryOp::CmpGe;
          if (orderingOp && (lIm || rIm)) {
            setParseError(ctx, ParseErrorId::IncompatibleOperands);
            return makeScalar(0);
          }
          if ((lIm || rIm) && (lTime || rTime)) {
            setParseError(ctx, ParseErrorId::IncompatibleOperands);
            return makeScalar(0);
          }
          if (!lTime && !rTime) {
#if SMARTMATH_COMPLEX_NUMBERS
            if (getSupportComplexNumbers() && (lIm || rIm)) {
              double lr = 0.0;
              double li = 0.0;
              double rr = 0.0;
              double ri = 0.0;
              scalarLoadCartesian(l.scalarValue, lr, li);
              scalarLoadCartesian(r.scalarValue, rr, ri);
              if (std::isnan(lr) || std::isnan(li) || std::isnan(rr) || std::isnan(ri)) {
                return makeScalarInt(evalComparisonTruthWhenUnorderedNan(e.binaryOp) ? 1LL : 0LL);
              }
            } else
#endif
            {
              const double ls = l.scalarValue.scalar;
              const double rs = r.scalarValue.scalar;
              if (std::isnan(ls) || std::isnan(rs)) {
                return makeScalarInt(evalComparisonTruthWhenUnorderedNan(e.binaryOp) ? 1LL : 0LL);
              }
            }
          }
          int cmp = 0;
          if (!cmpScalarValuesForCompare(
                  &ctx, l.scalarValue, r.scalarValue, cmp, CmpScalarIncompatiblePolicy::SetError)) {
            return makeScalar(0);
          }
          return makeScalarInt(evalComparisonByOp(e.binaryOp, cmp) ? 1LL : 0LL);
        }
        case Expr::BinaryOp::Pow:
          return evalMappedBinaryOp(ctx, e.binaryOp, l, r);
        case Expr::BinaryOp::Mul:
        case Expr::BinaryOp::Div:
        case Expr::BinaryOp::Add:
        case Expr::BinaryOp::Sub: {
          if ((e.binaryOp == Expr::BinaryOp::Add || e.binaryOp == Expr::BinaryOp::Sub) &&
              e.rhsIsDirectPostfixPercent) {
            r.scalarValue.scalar = l.scalarValue.scalar * r.scalarValue.scalar;
            r.scalarValue.setExactInt64Valid(false);
            r.scalarValue.setExactUInt64Valid(false);
            r.scalarValue.scalarKind = ScalarKind::FloatingPoint;
          }

#if SMARTMATH_COMPLEX_NUMBERS
          if (getSupportComplexNumbers() &&
              (scalarHasNonzeroImaginaryPart(l.scalarValue) ||
               scalarHasNonzeroImaginaryPart(r.scalarValue))) {
            return evalMappedBinaryOp(ctx, e.binaryOp, l, r);
          }
#endif

          if (!e.rhsIsDirectPostfixPercent &&
              (e.binaryOp == Expr::BinaryOp::Mul || e.binaryOp == Expr::BinaryOp::Add ||
               e.binaryOp == Expr::BinaryOp::Sub)) {
            long long li = 0, ri = 0;
            if (tryGetBothExactSignedInt64NoUIntWrapFromScalars(l.scalarValue, r.scalarValue, li, ri)) {
              long long outI = 0;
              const bool okInt =
                  (e.binaryOp == Expr::BinaryOp::Mul) ? checkedMulLL(li, ri, outI) :
                  (e.binaryOp == Expr::BinaryOp::Add) ? checkedAddLL(li, ri, outI) :
                                                         checkedSubLL(li, ri, outI);
              if (okInt) {
                return makeScalarInt(outI);
              }
              if (li >= 0 && ri >= 0) {
                const std::uint64_t lu = static_cast<std::uint64_t>(li);
                const std::uint64_t ru = static_cast<std::uint64_t>(ri);
                if (e.binaryOp == Expr::BinaryOp::Add) {
                  std::uint64_t outU = 0;
                  if (tryAddUInt64Checked(lu, ru, outU)) {
                    return makeScalarUInt(outU);
                  }
                } else if (e.binaryOp == Expr::BinaryOp::Mul) {
                  std::uint64_t outU = 0;
                  if (tryMulUInt64Checked(lu, ru, outU)) {
                    return makeScalarUInt(outU);
                  }
                }
              }
            }
            std::uint64_t lu = 0;
            std::uint64_t ru = 0;
            if (tryGetBothExactNonNegativeUInt64FromScalars(l.scalarValue, r.scalarValue, lu, ru)) {
              if (e.binaryOp == Expr::BinaryOp::Add) {
                std::uint64_t outU = 0;
                if (tryAddUInt64Checked(lu, ru, outU)) {
                  return makeScalarUInt(outU);
                }
              } else if (e.binaryOp == Expr::BinaryOp::Mul) {
                std::uint64_t outU = 0;
                if (tryMulUInt64Checked(lu, ru, outU)) {
                  return makeScalarUInt(outU);
                }
              }
            }
          }

          return evalMappedBinaryOp(ctx, e.binaryOp, l, r);
        }
        default:
          setParseError(ctx, ParseErrorId::InternalBinaryOp);
          return makeScalar(0);
      }
    }
    default:
      setParseError(ctx, ParseErrorId::InternalEvalError);
      return makeScalar(0);
  }
}

MathParser::EvalValue MathParser::evalExpr(
    const Expr& e,
    EvalContext& ctx,
    const std::unordered_map<std::string, EvalValue>* scopedVars) {
  const auto tryGetArrayIndexIntFromScalarValue = [](const EvalValue::ScalarValue& sIn, long long& outI) -> bool {
    if (scalarHasNonzeroImaginaryPart(sIn)) {
      return false;
    }
    EvalValue::ScalarValue s = sIn;
    scalarRepairExactMetadata(s);
    if (s.scalarKind == ScalarKind::Int64) {
      outI = s.exactInt64;
      return true;
    }
    if (s.scalarKind == ScalarKind::UInt64) {
      if (s.exactUInt64 > static_cast<std::uint64_t>((std::numeric_limits<long long>::max)())) {
        return false;
      }
      outI = static_cast<long long>(s.exactUInt64);
      return true;
    }
    return false;
  };

  switch (e.tag) {
    case Expr::Tag::Literal:
      return e.literalValue;
    case Expr::Tag::Variable: {
      EvalValue out = makeScalar(0);
      if (tryResolveVariableValue(e, scopedVars, out)) {
        return out;
      }
      if (handleUnknownIdentifier(ctx, e.name, ctx.unknownVarsText)) {
        return makeScalar(0);
      }
      return makeScalar(0);
    }
    case Expr::Tag::FunctionRef: {
      EvalValue out;
      out.kind = ValueKind::FunctionRef;
      out.funcRefName = e.name;
      return out;
    }
    case Expr::Tag::Unary:
      return evalUnaryExpr(e, ctx, scopedVars, false);
    case Expr::Tag::PostfixPercent: {
      EvalValue v = evalExpr(*e.child, ctx, scopedVars);
      if (ctx.parseError) {
        return v;
      }
      if (v.kind != ValueKind::Scalar) {
        setParseError(ctx, ParseErrorId::PercentageRequiresScalarValue);
        return makeScalar(0);
      }
#if SMARTMATH_TIME_VALUES
      if (scalarValueIsTime(v.scalarValue)) {
        setParseError(ctx, ParseErrorId::IncompatibleOperands);
        return makeScalar(0);
      }
#endif
#if SMARTMATH_COMPLEX_NUMBERS
      if (getSupportComplexNumbers() && scalarHasNonzeroImaginaryPart(v.scalarValue)) {
        setParseError(ctx, ParseErrorId::IncompatibleOperands);
        return makeScalar(0);
      }
#endif
      v.scalarValue.scalar /= 100.0;
      v.scalarValue.setExactInt64Valid(false);
      v.scalarValue.setExactUInt64Valid(false);
      v.scalarValue.scalarKind = ScalarKind::FloatingPoint;
      return v;
    }
    case Expr::Tag::ArrayOrParens: {
      std::vector<EvalValue> flatVals;
      flatVals.reserve(e.elements.size());
      for (const auto& ch : e.elements) {
        EvalValue a = evalExpr(*ch, ctx, scopedVars);
        if (ctx.parseError) {
          return makeScalar(0);
        }
        if (a.kind == ValueKind::Scalar) {
          flatVals.emplace_back(std::move(a));
        } else {
          if (a.arr.empty()) {
            continue;
          }
          const std::size_t oldSize = flatVals.size();
          flatVals.reserve(oldSize + a.arr.size());
          for (const auto& item : a.arr) {
            flatVals.emplace_back(scalarFromScalarValue(item));
          }
        }
      }
      if (flatVals.empty()) {
        return makeArrayFromScalars({});
      }
#if SMARTMATH_TIME_VALUES
      bool anyTimeArr = false;
      bool anyNonTimeArr = false;
      for (const auto& fv : flatVals) {
        if (fv.kind != ValueKind::Scalar) {
          continue;
        }
        if (scalarValueIsTime(fv.scalarValue)) {
          anyTimeArr = true;
        } else {
          anyNonTimeArr = true;
        }
      }
      if (anyTimeArr && anyNonTimeArr) {
        setError(ctx, constructTimeArrayMixedError("array literal", false));
        return makeScalar(0);
      }
#endif
      return makeArrayFromScalars(flatVals);
    }
    case Expr::Tag::Call: {
      std::vector<EvalValue> args;
      args.reserve(e.elements.size());
      for (const auto& ch : e.elements) {
        args.emplace_back(evalExpr(*ch, ctx, scopedVars));
        if (ctx.parseError) {
          return makeScalar(0);
        }
      }
      return evalFunctionCall(ctx, e.name, std::move(args), e.builtinFunctionId, scopedVars);
    }
    case Expr::Tag::Index: {
      EvalValue base = evalExpr(*e.left, ctx, scopedVars);
      if (ctx.parseError) return base;
      EvalValue idxv = evalExpr(*e.right, ctx, scopedVars);
      if (ctx.parseError) return idxv;
      if (idxv.kind != ValueKind::Scalar) {
        setParseError(ctx, ParseErrorId::ArrayIndexMustBeScalar);
        return makeScalar(0);
      }
      long long idx = 0;
      if (!tryGetArrayIndexIntFromScalarValue(idxv.scalarValue, idx)) {
        setParseError(ctx, ParseErrorId::ArrayIndexMustBeInteger);
        return makeScalar(0);
      }
      bool formalParamIndexBase = (base.kind == ValueKind::UdfFormalValidationDummy);
      if (!formalParamIndexBase && scopedVars && e.left && e.left->tag == Expr::Tag::Variable) {
        const auto sit = scopedVars->find(e.left->name);
        formalParamIndexBase =
            (sit != scopedVars->end() && sit->second.kind == ValueKind::UdfFormalValidationDummy);
      }
      if (formalParamIndexBase) {
        // UDF body validation only: argument length is unknown; bounds are checked at call time.
        auto pit = variables_.find(STR_FORMAL_VALIDATION_PROBE);
        if (pit != variables_.end() && pit->second.kind == ValueKind::Scalar) {
          return pit->second;
        }
        return makeScalarInt(0);
      }
      if (base.kind != ValueKind::Array) {
        setParseError(ctx, ParseErrorId::IndexingRequiresArray);
        return makeScalar(0);
      }
      const long long n = static_cast<long long>(base.arr.size());
      const long long realIdx = (idx >= 0) ? idx : (n + idx);
      if (realIdx < 0 || realIdx >= n) {
        setParseError(ctx, ParseErrorId::ArrayIndexOutOfRange);
        return makeScalar(0);
      }
      return scalarFromArrayAt(base, static_cast<std::size_t>(realIdx));
    }
    case Expr::Tag::Slice: {
      EvalValue base = evalExpr(*e.left, ctx, scopedVars);
      if (ctx.parseError) return base;

      const auto evalOptionalBound = [&](std::size_t slot, bool& hasBound, long long& outBound) -> bool {
        hasBound = (slot < e.elements.size() && e.elements[slot] != nullptr);
        if (!hasBound) {
          outBound = 0;
          return true;
        }
        EvalValue bv = evalExpr(*e.elements[slot], ctx, scopedVars);
        if (ctx.parseError) return false;
        if (bv.kind != ValueKind::Scalar) {
          setParseError(ctx, ParseErrorId::ArrayIndexMustBeScalar);
          return false;
        }
        if (!tryGetArrayIndexIntFromScalarValue(bv.scalarValue, outBound)) {
          setParseError(ctx, ParseErrorId::ArrayIndexMustBeInteger);
          return false;
        }
        return true;
      };

      bool hasStart = false;
      bool hasStop = false;
      bool hasStep = false;
      long long startIdx = 0;
      long long stopIdx = 0;
      long long stepIdx = 1;
      if (!evalOptionalBound(0, hasStart, startIdx) || ctx.parseError) return makeScalar(0);
      if (!evalOptionalBound(1, hasStop, stopIdx) || ctx.parseError) return makeScalar(0);
      if (!evalOptionalBound(2, hasStep, stepIdx) || ctx.parseError) return makeScalar(0);
      if (!hasStep) {
        stepIdx = 1;
      }
      if (stepIdx == 0) {
        setValidationError(ctx, STR_SLICE_STEP_MUST_NOT_BE_ZERO);
        return makeScalar(0);
      }

      bool formalParamSliceBase = (base.kind == ValueKind::UdfFormalValidationDummy);
      if (!formalParamSliceBase && scopedVars && e.left && e.left->tag == Expr::Tag::Variable) {
        const auto sit = scopedVars->find(e.left->name);
        formalParamSliceBase =
            (sit != scopedVars->end() && sit->second.kind == ValueKind::UdfFormalValidationDummy);
      }
      if (formalParamSliceBase) {
        // UDF body validation: 1-element probe array so chained indexing still parses.
        auto pit = variables_.find(STR_FORMAL_VALIDATION_PROBE);
        EvalValue probe = (pit != variables_.end() && pit->second.kind == ValueKind::Scalar)
                              ? pit->second
                              : makeScalarInt(1);
        return makeArrayFromScalars(std::vector<EvalValue>{probe});
      }
      if (base.kind != ValueKind::Array) {
        setParseError(ctx, ParseErrorId::IndexingRequiresArray);
        return makeScalar(0);
      }

      const long long length = static_cast<long long>(base.arr.size());
      const bool stepNeg = (stepIdx < 0);
      long long normStart = 0;
      long long normStop = 0;
      if (!hasStart) {
        normStart = stepNeg ? (length - 1) : 0;
      } else {
        normStart = startIdx;
        if (normStart < 0) normStart += length;
        if (normStart < 0) {
          normStart = stepNeg ? -1 : 0;
        } else if (normStart >= length) {
          normStart = stepNeg ? (length - 1) : length;
        }
      }
      if (!hasStop) {
        normStop = stepNeg ? -1 : length;
      } else {
        normStop = stopIdx;
        if (normStop < 0) normStop += length;
        if (normStop < 0) {
          normStop = stepNeg ? -1 : 0;
        } else if (normStop >= length) {
          normStop = stepNeg ? (length - 1) : length;
        }
      }

      long long outCount = 0;
      if (stepIdx > 0) {
        if (normStart < normStop) {
          outCount = (normStop - normStart + stepIdx - 1) / stepIdx;
        }
      } else {
        if (normStart > normStop) {
          outCount = (normStart - normStop - stepIdx - 1) / (-stepIdx);
        }
      }
      if (outCount <= 0) {
        return makeArrayFromScalars({});
      }
      std::vector<EvalValue> out;
      out.reserve(static_cast<std::size_t>(outCount));
      long long cur = normStart;
      for (long long i = 0; i < outCount; ++i) {
        out.push_back(scalarFromArrayAt(base, static_cast<std::size_t>(cur)));
        cur += stepIdx;
      }
      return makeArrayFromScalars(out);
    }
    case Expr::Tag::Binary: {
      const auto returnIntegerOperandError = [&](const char* msg) -> EvalValue {
        setValidationError(ctx, msg);
        return makeScalar(0);
      };
      enum class BinaryEvalStatus {
        Ok,
        LeftError,
        RightError
      };
      const auto evalBinaryOperands = [&](EvalValue& leftOut, EvalValue& rightOut) -> BinaryEvalStatus {
        leftOut = evalExpr(*e.left, ctx, scopedVars);
        if (ctx.parseError) {
          return BinaryEvalStatus::LeftError;
        }
        rightOut = evalExpr(*e.right, ctx, scopedVars);
        return ctx.parseError ? BinaryEvalStatus::RightError : BinaryEvalStatus::Ok;
      };
      if (e.binaryOp == Expr::BinaryOp::LogicalOr) {
        EvalValue l = makeScalar(0), r = makeScalar(0);
        const BinaryEvalStatus evalStatus = evalBinaryOperands(l, r);
        if (evalStatus != BinaryEvalStatus::Ok) {
          return (evalStatus == BinaryEvalStatus::LeftError) ? l : r;
        }
        return makeScalarInt((isTruthy(l) || isTruthy(r)) ? 1LL : 0LL);
      }
      if (e.binaryOp == Expr::BinaryOp::LogicalAnd) {
        EvalValue l = makeScalar(0), r = makeScalar(0);
        const BinaryEvalStatus evalStatus = evalBinaryOperands(l, r);
        if (evalStatus != BinaryEvalStatus::Ok) {
          return (evalStatus == BinaryEvalStatus::LeftError) ? l : r;
        }
        return makeScalarInt((isTruthy(l) && isTruthy(r)) ? 1LL : 0LL);
      }
      if (e.binaryOp == Expr::BinaryOp::Modulo || e.binaryOp == Expr::BinaryOp::BitAnd ||
          e.binaryOp == Expr::BinaryOp::BitOr || e.binaryOp == Expr::BinaryOp::BitXor ||
          e.binaryOp == Expr::BinaryOp::ShiftLeft || e.binaryOp == Expr::BinaryOp::ShiftRight) {
        EvalValue l = makeScalar(0), r = makeScalar(0);
        if (evalBinaryOperands(l, r) != BinaryEvalStatus::Ok) {
          return makeScalar(0);
        }
        return evalInt64BinaryOp(ctx, l, r, e.binaryOp);
      }
      if (isComparisonBinaryOp(e.binaryOp)) {
        EvalValue l = evalExpr(*e.left, ctx, scopedVars);
        if (ctx.parseError) return makeScalar(0);
        EvalValue r = evalExpr(*e.right, ctx, scopedVars);
        if (ctx.parseError) return makeScalar(0);
#if SMARTMATH_COMPLEX_NUMBERS
        if (getSupportComplexNumbers()) {
          const auto valueHasComplex = [](const EvalValue& vv) -> bool {
            if (vv.kind == ValueKind::Scalar) {
              return scalarHasNonzeroImaginaryPart(vv.scalarValue);
            }
            for (const auto& el : vv.arr) {
              if (scalarHasNonzeroImaginaryPart(el)) {
                return true;
              }
            }
            return false;
          };
          const bool orderingOp = e.binaryOp == Expr::BinaryOp::CmpLt || e.binaryOp == Expr::BinaryOp::CmpLe ||
                                  e.binaryOp == Expr::BinaryOp::CmpGt || e.binaryOp == Expr::BinaryOp::CmpGe;
          if (orderingOp && (valueHasComplex(l) || valueHasComplex(r))) {
            setParseError(ctx, ParseErrorId::IncompatibleOperands);
            return makeScalar(0);
          }
#if SMARTMATH_TIME_VALUES
          if ((valueHasComplex(l) || valueHasComplex(r)) &&
              (evalValueInvolvesTime(l) || evalValueInvolvesTime(r))) {
            setParseError(ctx, ParseErrorId::IncompatibleOperands);
            return makeScalar(0);
          }
#endif
        }
#endif
        if (l.kind == ValueKind::Scalar && r.kind == ValueKind::Scalar) {
#if SMARTMATH_TIME_VALUES
          const bool lTime = scalarValueIsTime(l.scalarValue);
          const bool rTime = scalarValueIsTime(r.scalarValue);
          if (!lTime && !rTime)
#endif
          {
#if SMARTMATH_COMPLEX_NUMBERS
            if (getSupportComplexNumbers()) {
              double lr = 0.0;
              double li = 0.0;
              double rr = 0.0;
              double ri = 0.0;
              scalarLoadCartesian(l.scalarValue, lr, li);
              scalarLoadCartesian(r.scalarValue, rr, ri);
              if (std::isnan(lr) || std::isnan(li) || std::isnan(rr) || std::isnan(ri)) {
                return makeScalarInt(evalComparisonTruthWhenUnorderedNan(e.binaryOp) ? 1LL : 0LL);
              }
            } else
#endif
            {
              const double ls = l.scalarValue.scalar;
              const double rs = r.scalarValue.scalar;
              if (std::isnan(ls) || std::isnan(rs)) {
                return makeScalarInt(evalComparisonTruthWhenUnorderedNan(e.binaryOp) ? 1LL : 0LL);
              }
            }
          }
        }
        int cmp = 0;
        if (!tryLexicographicCompareEvalValues(ctx, l, r, cmp)) {
          return makeScalar(0);
        }
        return makeScalarInt(evalComparisonByOp(e.binaryOp, cmp) ? 1LL : 0LL);
      }
      EvalValue l = evalExpr(*e.left, ctx, scopedVars);
      if (ctx.parseError) {
        return l;
      }
      EvalValue r = evalExpr(*e.right, ctx, scopedVars);
      if (ctx.parseError) {
        return r;
      }
      if ((e.binaryOp == Expr::BinaryOp::Add || e.binaryOp == Expr::BinaryOp::Sub) &&
          e.rhsIsDirectPostfixPercent &&
          l.kind == ValueKind::Scalar && r.kind == ValueKind::Scalar) {
        r.scalarValue.scalar = l.scalarValue.scalar * r.scalarValue.scalar;
        r.scalarValue.setExactInt64Valid(false);
        r.scalarValue.setExactUInt64Valid(false);
        r.scalarValue.scalarKind = ScalarKind::FloatingPoint;
      }
      if (e.binaryOp == Expr::BinaryOp::Pow) {
        return evalMappedBinaryOp(ctx, e.binaryOp, l, r);
      }
      if (e.binaryOp == Expr::BinaryOp::Mul || e.binaryOp == Expr::BinaryOp::Div ||
          e.binaryOp == Expr::BinaryOp::Add || e.binaryOp == Expr::BinaryOp::Sub) {
        return evalMappedBinaryOp(ctx, e.binaryOp, l, r);
      }
      setParseError(ctx, ParseErrorId::InternalBinaryOp);
      return makeScalar(0);
    }
    default:
      break;
  }
  setParseError(ctx, ParseErrorId::InternalEvalError);
  return makeScalar(0);
}

MathParser::EvalValue MathParser::runCompiledProgram(
    EvalContext& ctx,
    const std::vector<AstStatement>& program,
    const std::unordered_map<std::string, EvalValue>* scopedVars,
    bool scalarOnlyMode) {
  EvalValue out = makeScalarInt(0);
  const auto evalIntoOut = [&](const Expr& ex) {
    out = scalarOnlyMode ? evalExprScalar(ex, ctx, scopedVars) : evalExpr(ex, ctx, scopedVars);
  };
  for (const auto& st : program) {
    if (st.kind == AstStatement::Kind::FunDef) {
      const std::string udfErr =
          getUserFunctionDefinitionErrorText(st.fun.name, st.fun.params, st.fun.expr, true);
      if (!udfErr.empty()) {
        setError(ctx, udfErr);
        return out;
      }
      upsertUserFunction(UserFunction{st.fun.name, st.fun.params, st.fun.expr});
      out = makeScalarInt(0);
      setVariable(STR_ANS, out);
    } else if (st.kind == AstStatement::Kind::Assign) {
      evalIntoOut(*st.expr);
      if (ctx.parseError) {
        return out;
      }
      removeUserFunctionByName(st.assignName);
      setVariable(st.assignName, out);
      setVariable(STR_ANS, out);
    } else {
      evalIntoOut(*st.expr);
      if (ctx.parseError) {
        return out;
      }
      setVariable(STR_ANS, out);
    }
  }
  return out;
}

bool MathParser::argsContainNonFinite(const std::vector<EvalValue>& args) {
  bool found = false;
  forEachCallArgScalarValues(args, [&](const EvalValue::ScalarValue& s) {
    if (!found && !std::isfinite(s.scalar)) {
      found = true;
    }
  });
  return found;
}

MathParser::BuiltinFlags MathParser::getBuiltinFlags(BuiltinFunctionId id) {
  if (id >= BuiltinFunctionId::Count) {
    return BuiltinFlags::None;
  }
  return kBuiltinMeta[static_cast<std::size_t>(id)].flags;
}

MathParser::BuiltinCategory MathParser::getBuiltinCategory(BuiltinFunctionId id) {
  if (id >= BuiltinFunctionId::Count) {
    return BuiltinCategory::Unknown;
  }
  return kBuiltinMeta[static_cast<std::size_t>(id)].category;
}

void MathParser::applyBuiltinFormatRenderMeta(BuiltinFunctionId id, EvalValue& out) {
  if (id == BuiltinFunctionId::Hex || id == BuiltinFunctionId::Uhex) {
    out.setRenderBase(RenderBase::Hex);
  } else if (id == BuiltinFunctionId::Oct || id == BuiltinFunctionId::Uoct) {
    out.setRenderBase(RenderBase::Oct);
  } else {
    out.setRenderBase(RenderBase::Bin);
  }
  out.setRenderUnsigned(
      id == BuiltinFunctionId::Uhex || id == BuiltinFunctionId::Uoct || id == BuiltinFunctionId::Ubin);
}

bool MathParser::hasBuiltinFlag(BuiltinFunctionId id, BuiltinFlags flag) {
  const unsigned combined = static_cast<unsigned>(getBuiltinFlags(id));
  const unsigned mask = static_cast<unsigned>(flag);
  return (combined & mask) == mask;
}

bool MathParser::getBuiltinArity(BuiltinFunctionId id, uint8_t& minArgs, uint8_t& maxArgs) {
  if (id >= BuiltinFunctionId::Count) {
    return false;
  }
  if (hasBuiltinFlag(id, BuiltinFlags::Unary)) {
    minArgs = 1;
    maxArgs = 1;
    return true;
  }
  const auto& row = kBuiltinMeta[static_cast<std::size_t>(id)];
  if (row.minArgs == MathParser::kBuiltinMetaArityUnset) {
    return false;
  }
  minArgs = row.minArgs;
  maxArgs = row.maxArgs;
  return true;
}

bool MathParser::validateCallArity(
    EvalContext& ctx,
    const std::string& fnName,
    uint8_t minArgs,
    uint8_t maxArgs,
    const std::size_t argc) const {
  if (minArgs == maxArgs) {
    if (argc != minArgs) {
      setExactArgCountError(ctx, fnName, minArgs);
      return false;
    }
    return true;
  }
  if (argc < minArgs) {
    setAtLeastNArgsError(ctx, fnName, minArgs);
    return false;
  }
  if (maxArgs != kBuiltinArityUnbounded && argc > maxArgs) {
    setExactArgCountError(ctx, fnName, maxArgs);
    return false;
  }
  return true;
}

bool MathParser::validateBuiltinCallArity(
    EvalContext& ctx,
    const std::string& fnName,
    BuiltinFunctionId id,
    const std::vector<EvalValue>& args) const {
  uint8_t minArgs = 0;
  uint8_t maxArgs = 0;
  if (!getBuiltinArity(id, minArgs, maxArgs)) {
    return true;
  }
  return validateCallArity(ctx, fnName, minArgs, maxArgs, args.size());
}

bool MathParser::validateIntegerRepresentableArgs(
    EvalContext& ctx,
    const std::string& fnName,
    const std::vector<EvalValue>& args,
    bool allowNonFiniteForFormat) const {
  const auto validateScalar = [&](const EvalValue::ScalarValue& s) -> bool {
    if (!std::isfinite(s.scalar)) {
      return allowNonFiniteForFormat;
    }
    if (!(s.hasExactInt64() || s.hasExactUInt64())) {
      long long signedV = 0;
      if (!tryGetSignedInt64FromScalar(s, signedV)) {
        setIntegerValuesError(ctx, fnName);
        return false;
      }
    }
#if SMARTMATH_COMPLEX_NUMBERS
    if (getSupportComplexNumbers() && scalarHasNonzeroImaginaryPart(s)) {
      double ai = 0.0;
      double ar = 0.0;
      scalarLoadCartesian(s, ar, ai);
      (void)ar;
      if (!std::isfinite(ai)) {
        if (!allowNonFiniteForFormat) {
          setIntegerValuesError(ctx, fnName);
          return false;
        }
        return true;
      }
      if (s.hasImagExactInt64()) {
        return true;
      }
      long long tmp = 0;
      if (!tryExtractExactInt64FromDoubleStrict(ai, tmp)) {
        setIntegerValuesError(ctx, fnName);
        return false;
      }
    }
#endif
    return true;
  };
  bool ok = true;
  forEachCallArgScalarValues(args, [&](const EvalValue::ScalarValue& s) {
    if (!ok) {
      return;
    }
    if (!validateScalar(s)) {
      ok = false;
    }
  });
  if (!ok) {
    setIntegerValuesError(ctx, fnName);
    return false;
  }
  return true;
}

bool MathParser::validateBuiltinArgs(
    EvalContext& ctx,
    const std::string& fnName,
    BuiltinFunctionId id,
    const std::vector<EvalValue>& args) const {
  if (hasBuiltinFlag(id, BuiltinFlags::Format)) {
    return validateIntegerRepresentableArgs(ctx, fnName, args, true);
  }
  if (hasBuiltinFlag(id, BuiltinFlags::IntegerOnly)) {
    return validateIntegerRepresentableArgs(ctx, fnName, args, false);
  }
  if (hasBuiltinFlag(id, BuiltinFlags::NonCalculating)) {
    return true;
  }
  if (hasBuiltinFlag(id, BuiltinFlags::FiniteRequired) && argsContainNonFinite(args)) {
    setNumericErrorInFunction(ctx, fnName);
    return false;
  }
  return true;
}

MathParser::EvalValue MathParser::builtinUnpack(EvalContext& ctx, const std::vector<EvalValue>& args) const {
  const auto markExpanded = [](EvalValue v) {
    v.setExpandArgs(true);
    return v;
  };
  if (args.size() == 1) {
    return markExpanded(args[0]);
  }
  std::vector<EvalValue> elems;
  if (!flattenArgsToScalars(args, elems)) {
    setAtLeastOneArgError(ctx, getFunctionName(BuiltinFunctionId::Unpack));
    return makeScalar(0);
  }
  return markExpanded(makeArrayFromScalars(elems));
}

MathParser::EvalValue MathParser::builtinLen(
    EvalContext& /*ctx*/,
    const std::vector<EvalValue>& args) const {
  return makeScalarInt(static_cast<long long>(countAggregateArgScalarItems(args)));
}

MathParser::EvalValue MathParser::builtinRepeat(
    EvalContext& ctx,
    const std::string& fnName,
    const std::vector<EvalValue>& args) const {
  const auto copySingleArgToFlat = [](const EvalValue& a, std::vector<EvalValue>& out) {
    out.clear();
    if (a.kind == ValueKind::Scalar) {
      out.emplace_back(a);
      return;
    }
    out.reserve(a.arr.size());
    for (const auto& item : a.arr) {
      out.emplace_back(scalarFromScalarValue(item));
    }
  };

  std::vector<EvalValue> vals;
  copySingleArgToFlat(args[0], vals);
  if (vals.empty()) {
    setAtLeastOneArgError(ctx, fnName);
    return makeScalar(0);
  }

  const EvalValue& countArg = args[1];
  if (countArg.kind == ValueKind::Array
#if SMARTMATH_TIME_VALUES
      || evalValueInvolvesTime(countArg)
#endif
#if SMARTMATH_COMPLEX_NUMBERS
      || (getSupportComplexNumbers() && evalValueHasNonzeroImaginary(countArg))
#endif
  ) {
    setIntegerValuesError(ctx, fnName);
    return makeScalar(0);
  }

  long long n = 0;
  if (!tryGetSignedInt64FromScalar(countArg.scalarValue, n)) {
    setIntegerValuesError(ctx, fnName);
    return makeScalar(0);
  }
  if (n <= 0) {
    setPositiveIntegerError(ctx, fnName);
    return makeScalar(0);
  }

  const std::size_t blockSize = vals.size();
  std::uint64_t estimated = 0;
  if (!checkedMulU64(static_cast<std::uint64_t>(blockSize), static_cast<std::uint64_t>(n), estimated)) {
    setTooManyValuesError(ctx, fnName);
    return makeScalar(0);
  }
  if (!acceptProducedItemCount(ctx, fnName, estimated)) {
    return makeScalar(0);
  }

  const std::size_t totalSize = static_cast<std::size_t>(estimated);
  std::vector<EvalValue> out;
  out.reserve(totalSize);
  for (long long rep = 0; rep < n; ++rep) {
    for (const auto& item : vals) {
      out.push_back(item);
    }
  }
  return makeArrayFromScalars(out);
}

MathParser::EvalValue MathParser::builtinRange(
    EvalContext& ctx,
    const std::string& fnName,
    const std::vector<EvalValue>& args) const {
  const auto tryGetRangeIntegerScalarArg = [this](const EvalValue& v, long long& outN) -> bool {
    if (v.kind == ValueKind::Array) {
      return false;
    }
#if SMARTMATH_TIME_VALUES
    if (evalValueInvolvesTime(v)) {
      return false;
    }
#endif
#if SMARTMATH_COMPLEX_NUMBERS
    if (getSupportComplexNumbers() && evalValueHasNonzeroImaginary(v)) {
      return false;
    }
#endif
    return tryGetSignedInt64FromScalar(v.scalarValue, outN);
  };

  long long startV = 0;
  long long stopV = 0;
  long long stepV = 1;
  if (!tryGetRangeIntegerScalarArg(args[0], startV) || !tryGetRangeIntegerScalarArg(args[1], stopV)) {
    setIntegerValuesError(ctx, fnName);
    return makeScalar(0);
  }
  if (args.size() >= 3U) {
    if (!tryGetRangeIntegerScalarArg(args[2], stepV)) {
      setIntegerValuesError(ctx, fnName);
      return makeScalar(0);
    }
  }

  if (stepV == 0) {
    setRangeStepZeroError(ctx, fnName);
    return makeScalar(0);
  }
  if ((stepV > 0 && startV >= stopV) || (stepV < 0 && startV <= stopV)) {
    setRangeEmptyError(ctx, fnName);
    return makeScalar(0);
  }

  const std::uint64_t estimated = estimateRangeItemCount(startV, stopV, stepV);
  if (estimated == 0ull) {
    setRangeEmptyError(ctx, fnName);
    return makeScalar(0);
  }
  if (!acceptProducedItemCount(ctx, fnName, estimated)) {
    return makeScalar(0);
  }

  const std::size_t count = static_cast<std::size_t>(estimated);
  std::vector<EvalValue> out;
  out.reserve(count);
  long long cur = startV;
  for (std::size_t i = 0; i < count; ++i) {
    out.push_back(makeScalarInt(cur));
    if (i + 1U < count) {
      cur += stepV;
    }
  }
  return makeArrayFromScalars(out);
}


namespace {

double medianFromDoubleFlatInPlace(std::vector<double>& flat) {
  const std::size_t n = flat.size();
  const std::size_t mid = n / 2U;
  std::nth_element(flat.begin(), flat.begin() + static_cast<std::ptrdiff_t>(mid), flat.end());
  const double upper = flat[mid];
  if (n % 2U == 1U) {
    return upper;
  }
  const double lower = *std::max_element(flat.begin(), flat.begin() + static_cast<std::ptrdiff_t>(mid));
  return (lower + upper) / 2.0;
}

}  // namespace

std::size_t MathParser::countAggregateArgScalarItems(const std::vector<EvalValue>& args) const {
  std::size_t itemCount = 0;
  forEachCallArgScalarValues(args, [&](const EvalValue::ScalarValue&) { ++itemCount; });
  return itemCount;
}

MathParser::EvalValue MathParser::aggregateFoldByMode(
    EvalContext& ctx,
    const std::string& fnName,
    BuiltinFunctionId id,
    const std::vector<EvalValue>& args,
    AggregateFoldMode mode) const {
  switch (mode) {
#if SMARTMATH_COMPLEX_NUMBERS
  case AggregateFoldMode::Complex: {
  const auto forEachArgScalarValue = [&](const auto& fn) -> std::size_t {
    return forEachCallArgScalarValues(args, fn);
  };
  if (args.empty()) {
    setAtLeastOneArgError(ctx, fnName);
    return makeScalar(0);
  }
  if (args.size() == 1 && args[0].kind == ValueKind::Scalar) {
    return args[0];
  }
  const std::size_t itemCount = countAggregateArgScalarItems(args);
  if (itemCount == 0) {
    setAtLeastOneArgError(ctx, fnName);
    return makeScalar(0);
  }
  const char foldOp = (id == BuiltinFunctionId::Product) ? '*' : '+';
  EvalValue folded{};
  if (tryFoldExactComplexCartesian(args, foldOp, folded)) {
    if (id == BuiltinFunctionId::Avg || id == BuiltinFunctionId::Mean) {
      EvalValue avg{};
      if (tryAvgExactComplexFromSum(folded, itemCount, avg)) {
        return avg;
      }
      double ar = 0.0;
      double ai = 0.0;
      scalarLoadCartesian(folded.scalarValue, ar, ai);
      ar /= static_cast<double>(itemCount);
      ai /= static_cast<double>(itemCount);
      return makeScalarComplexFromDoubles(ar, ai);
    }
    return folded;
  }
  double ar = 0.0;
  double ai = 0.0;
  if (id == BuiltinFunctionId::Product) {
    ar = 1.0;
    ai = 0.0;
  }
  double br = 0.0;
  double bi = 0.0;
  bool stopComplexProduct = false;
  forEachArgScalarValue([&](const EvalValue::ScalarValue& s) {
    if (stopComplexProduct) {
      return;
    }
    scalarLoadCartesian(s, br, bi);
    if (id == BuiltinFunctionId::Product) {
      complexMultiply(ar, ai, br, bi, ar, ai);
      // allow `Inf + N*i` or `N + Inf*i`
      if (!std::isfinite(ar) && !std::isfinite(ai)) {
        stopComplexProduct = true;
      }
    } else {
      ar += br;
      ai += bi;
    }
  });
  if (id == BuiltinFunctionId::Avg || id == BuiltinFunctionId::Mean) {
    ar /= static_cast<double>(itemCount);
    ai /= static_cast<double>(itemCount);
  }
  return makeScalarComplexFromDoubles(ar, ai);
  }
#endif
#if SMARTMATH_TIME_VALUES
  case AggregateFoldMode::Time: {
  const auto forEachArgScalarValue = [&](const auto& fn) -> std::size_t {
    return forEachCallArgScalarValues(args, fn);
  };
  if (id == BuiltinFunctionId::Median) {
    std::vector<double> flatMs;
    bool okM = true;
    forEachArgScalarValue([&](const EvalValue::ScalarValue& s) {
      if (!okM) {
        return;
      }
      if (!scalarValueIsTime(s)) {
        setError(ctx, constructTimeArrayMixedError(fnName));
        okM = false;
        return;
      }
      flatMs.push_back(static_cast<double>(timeTotalMsFromScalarValue(s)));
    });
    if (!okM) {
      return makeScalar(0);
    }
    if (flatMs.empty()) {
      setAtLeastOneArgError(ctx, fnName);
      return makeScalar(0);
    }
    return makeScalarTimeMs(roundHalfUpDoubleToLongLong(medianFromDoubleFlatInPlace(flatMs)));
  }

  long long accMs = 0;
  bool accInit = false;
  std::size_t itemCount = 0;
  bool okAgg = true;
  forEachArgScalarValue([&](const EvalValue::ScalarValue& s) {
    if (!okAgg) {
      return;
    }
    if (!scalarValueIsTime(s)) {
      setError(ctx, constructTimeArrayMixedError(fnName));
      okAgg = false;
      return;
    }
    const long long curMs = timeTotalMsFromScalarValue(s);
    if (id == BuiltinFunctionId::Min) {
      if (!accInit || curMs < accMs) {
        accMs = curMs;
      }
      accInit = true;
    } else if (id == BuiltinFunctionId::Max) {
      if (!accInit || curMs > accMs) {
        accMs = curMs;
      }
      accInit = true;
    } else {
      long long nacc = 0;
      if (!tryAddTimeMsChecked(accMs, curMs, nacc)) {
        setParseError(ctx, ParseErrorId::IncompatibleOperands);
        okAgg = false;
        return;
      }
      accMs = nacc;
    }
    itemCount += 1;
  });
  if (!okAgg) {
    return makeScalar(0);
  }
  if (itemCount == 0) {
    setAtLeastOneArgError(ctx, fnName);
    return makeScalar(0);
  }
  if (id == BuiltinFunctionId::Avg || id == BuiltinFunctionId::Mean) {
    const double avgD = static_cast<double>(accMs) / static_cast<double>(itemCount);
    return makeScalarTimeMs(roundHalfUpDoubleToLongLong(avgD));
  }
  return makeScalarTimeMs(accMs);
  }
#endif
  case AggregateFoldMode::Real:
  default: {
  const auto forEachArgScalar = [&](const auto& fn) -> std::size_t {
    return forEachCallArgScalarValues(args, [&](const EvalValue::ScalarValue& sv) { fn(sv.scalar); });
  };
  const auto forEachArgScalarValue = [&](const auto& fn) -> std::size_t {
    return forEachCallArgScalarValues(args, fn);
  };
  if (id == BuiltinFunctionId::Sum || id == BuiltinFunctionId::Product || id == BuiltinFunctionId::Min ||
      id == BuiltinFunctionId::Max || id == BuiltinFunctionId::Avg || id == BuiltinFunctionId::Mean) {
    if (args.size() == 1 && args[0].kind == ValueKind::Scalar) {
      return args[0];
    }
    if (id == BuiltinFunctionId::Sum || id == BuiltinFunctionId::Product || id == BuiltinFunctionId::Min ||
        id == BuiltinFunctionId::Max) {
      bool allNnU = true;
      bool allSwI = true;
      bool anyNonInteger = false;
      const auto classify = [&](const EvalValue::ScalarValue& s) {
        std::uint64_t u = 0;
        long long i = 0;
        const bool nnU = tryGetExactNonNegativeUInt64FromScalar(s, u);
        const bool swI = tryGetExactSignedInt64NoUIntWrapFromScalar(s, i);
        if (!nnU && !swI) {
          anyNonInteger = true;
          return;
        }
        if (!nnU) {
          allNnU = false;
        }
        if (!swI) {
          allSwI = false;
        }
      };
      const std::size_t nn = forEachArgScalarValue(classify);
      if (nn == 0U) {
        setAtLeastOneArgError(ctx, fnName);
        return makeScalar(0);
      }
      if (!anyNonInteger && id == BuiltinFunctionId::Product) {
        bool ok = true;
        std::uint64_t productMag = 1;
        bool isNegativeProduct = false;
        forEachArgScalarValue([&](const EvalValue::ScalarValue& s) {
          if (!ok) {
            return;
          }
          std::uint64_t termMag = 0;
          if (!tryGetExactNonNegativeUInt64FromScalar(s, termMag)) {
            long long termI = 0;
            if (!tryGetExactSignedInt64NoUIntWrapFromScalar(s, termI)) {
              ok = false;
              return;
            }
            if (termI < 0) {
              isNegativeProduct = !isNegativeProduct;
              termMag = (termI == (std::numeric_limits<long long>::min)())
                ? (1ull << 63)
                : static_cast<std::uint64_t>(-termI);
            } else {
              termMag = static_cast<std::uint64_t>(termI);
            }
          }
          std::uint64_t next = 0;
          if (!tryMulUInt64Checked(productMag, termMag, next)) {
            ok = false;
            return;
          }
          productMag = next;
        });
        if (ok) {
          if (isNegativeProduct) {
            if (productMag <= (1ull << 63)) {
              return (productMag == (1ull << 63))
                ? makeScalarInt((std::numeric_limits<long long>::min)())
                : makeScalarInt(-static_cast<long long>(productMag));
            }
          } else {
            return makeScalarUInt(productMag);
          }
        }
      } else if (!anyNonInteger && allNnU) {
        bool ok = true;
        std::uint64_t accU = 0;
        bool hasMm = false;
        std::uint64_t mmU = 0;
        const bool wantMin = (id == BuiltinFunctionId::Min);
        forEachArgScalarValue([&](const EvalValue::ScalarValue& s) {
          std::uint64_t u = 0;
          (void)tryGetExactNonNegativeUInt64FromScalar(s, u);
          if (id == BuiltinFunctionId::Sum) {
            std::uint64_t next = 0;
            if (!tryAddUInt64Checked(accU, u, next)) {
              ok = false;
            }
            accU = next;
          } else if (!hasMm || (wantMin ? u < mmU : u > mmU)) {
            mmU = u;
            hasMm = true;
          }
        });
        if (ok) {
          if (id == BuiltinFunctionId::Sum) {
            return makeScalarUInt(accU);
          }
          return makeScalarUInt(mmU);
        }
      } else if (!anyNonInteger && allSwI) {
        bool ok = true;
        long long accI = 0;
        bool hasMm = false;
        long long mmI = 0;
        const bool wantMin = (id == BuiltinFunctionId::Min);
        forEachArgScalarValue([&](const EvalValue::ScalarValue& s) {
          long long i = 0;
          (void)tryGetExactSignedInt64NoUIntWrapFromScalar(s, i);
          if (id == BuiltinFunctionId::Sum) {
            long long next = 0;
            if (!tryAddInt64Checked(accI, i, next)) {
              ok = false;
            }
            accI = next;
          } else if (!hasMm || (wantMin ? i < mmI : i > mmI)) {
            mmI = i;
            hasMm = true;
          }
        });
        if (ok) {
          if (id == BuiltinFunctionId::Sum) {
            return makeScalarInt(accI);
          }
          return makeScalarInt(mmI);
        }
      }
    }
    if (id == BuiltinFunctionId::Min || id == BuiltinFunctionId::Max) {
      const bool wantMin = (id == BuiltinFunctionId::Min);
      const auto hasExactIntegerMetadata = [](const EvalValue::ScalarValue& sv) {
        if (sv.hasRenderRational()) {
          return false;
        }
        return sv.hasExactInt64() || sv.hasExactUInt64();
      };
      const auto shouldReplace = [&](double v, double winnerV, bool hasWinner) {
        if (!hasWinner) {
          return true;
        }
        return wantMin ? (v < winnerV) : (v > winnerV);
      };
      const auto shouldPreferOnTie = [&](const EvalValue::ScalarValue& winnerSv,
                                       const EvalValue::ScalarValue& candidateSv) {
        if (!hasExactIntegerMetadata(candidateSv)) {
          return false;
        }
        if (hasExactIntegerMetadata(winnerSv)) {
          return false;
        }
        return true;
      };
      EvalValue::ScalarValue winnerSv{};
      double winnerV = 0.0;
      bool hasWinner = false;
      bool hasNonNanWinner = false;
      std::size_t itemCount = 0;
      forEachArgScalarValue([&](const EvalValue::ScalarValue& s) {
        ++itemCount;
        const double v = scalarNumericReal(s);
        if (std::isnan(v)) {
          if (!hasNonNanWinner && !hasWinner) {
            winnerSv = s;
            winnerV = v;
            hasWinner = true;
          }
          return;
        }
        if (!hasWinner || std::isnan(winnerV) || shouldReplace(v, winnerV, true)) {
          winnerSv = s;
          winnerV = v;
          hasWinner = true;
          hasNonNanWinner = true;
        } else if (v == winnerV) {
          if (shouldPreferOnTie(winnerSv, s)) {
            winnerSv = s;
          }
        }
      });
      if (itemCount == 0) {
        setAtLeastOneArgError(ctx, fnName);
        return makeScalar(0);
      }
      return scalarFromScalarValue(winnerSv);
    }
    double acc = 0.0;
    std::size_t n = 0;
    if (id == BuiltinFunctionId::Product) {
      acc = 1.0;
      bool stopProduct = false;
      n = forEachArgScalar([&](double v) {
        if (stopProduct) {
          return;
        }
        acc *= v;
        if (!std::isfinite(acc)) {
          stopProduct = true;
        }
      });
    } else {
      acc = 0.0;
      n = forEachArgScalar([&](double v) { acc += v; });
    }

    if (n == 0) {
      setAtLeastOneArgError(ctx, fnName);
        return makeScalar(0);
    }
    if (id == BuiltinFunctionId::Avg || id == BuiltinFunctionId::Mean) {
      acc /= static_cast<double>(n);
    }
    return makeScalarMaybeExact(acc);
  }

  switch (id) {
    case BuiltinFunctionId::Median: {
      if (args.size() == 1U) {
        const EvalValue& single = args.front();
        if (single.kind == ValueKind::Scalar) {
          return single;
        }
        if (single.arr.size() == 1U) {
          return scalarFromArrayAt(single, 0);
        }
      }
      std::vector<double> flat;
      if (!flattenArgs(args, flat)) {
        setAtLeastOneArgError(ctx, fnName);
        return makeScalar(0);
      }
      return makeScalar(medianFromDoubleFlatInPlace(flat));
    }
    case BuiltinFunctionId::Variance:
    case BuiltinFunctionId::Stddev: {
#if SMARTMATH_TIME_VALUES
      if (std::any_of(args.begin(), args.end(), [this](const EvalValue& v) { return evalValueInvolvesTime(v); })) {
        setParseError(ctx, ParseErrorId::IncompatibleOperands);
        return makeScalar(0);
      }
#endif
      // Welford single-pass accumulation for better stability and fewer passes.
      double mean = 0.0;
      double m2 = 0.0;
      std::size_t n = 0;
      forEachArgScalar([&](double v) {
        ++n;
        const double delta = v - mean;
        mean += delta / static_cast<double>(n);
        const double delta2 = v - mean;
        m2 += delta * delta2;
      });
      if (n == 0) {
        setAtLeastOneArgError(ctx, fnName);
        return makeScalar(0);
      }
      double var = m2 / static_cast<double>(n);
      if (id == BuiltinFunctionId::Variance) {
        return makeScalar(var);
      }
      return makeScalar(std::sqrt(var));
    }
    default:
      break;
  }
  setParseError(ctx, ParseErrorId::InternalAggregateBuiltin);
  return makeScalar(0);
  }
  }
  //unreachable code:
  //setParseError(ctx, ParseErrorId::InternalAggregateBuiltin);
  //return makeScalar(0);
}


MathParser::EvalValue MathParser::builtinAggregateFamily(
    EvalContext& ctx,
    const std::string& fnName,
    BuiltinFunctionId id,
    const std::vector<EvalValue>& args) const {
#if SMARTMATH_TIME_VALUES
  const bool anyTimeAgg =
      std::any_of(args.begin(), args.end(), [this](const EvalValue& v) { return evalValueInvolvesTime(v); });
  if (anyTimeAgg && id == BuiltinFunctionId::Product) {
    setParseError(ctx, ParseErrorId::IncompatibleOperands);
    return makeScalar(0);
  }
#else
  const bool anyTimeAgg = false;
#endif
#if SMARTMATH_COMPLEX_NUMBERS
  if (getSupportComplexNumbers()) {
    const bool anyComplex = std::any_of(
        args.begin(), args.end(),
        [](const EvalValue& v) { return MathParser::evalValueHasNonzeroImaginary(v); });
    if (anyComplex) {
#if SMARTMATH_TIME_VALUES
      if (anyTimeAgg) {
        setParseError(ctx, ParseErrorId::IncompatibleOperands);
        return makeScalar(0);
      }
#endif
      if (id == BuiltinFunctionId::Min || id == BuiltinFunctionId::Max || id == BuiltinFunctionId::Median ||
          id == BuiltinFunctionId::Variance || id == BuiltinFunctionId::Stddev) {
        setParseError(ctx, ParseErrorId::IncompatibleOperands);
        return makeScalar(0);
      }
      if (id == BuiltinFunctionId::Sum || id == BuiltinFunctionId::Product || id == BuiltinFunctionId::Avg ||
          id == BuiltinFunctionId::Mean) {
        return aggregateFoldByMode(ctx, fnName, id, args, AggregateFoldMode::Complex);
      }
    }
  }
#endif
#if SMARTMATH_TIME_VALUES
  if (anyTimeAgg) {
    if (id == BuiltinFunctionId::Variance || id == BuiltinFunctionId::Stddev) {
      setParseError(ctx, ParseErrorId::IncompatibleOperands);
      return makeScalar(0);
    }
    if (id == BuiltinFunctionId::Sum || id == BuiltinFunctionId::Min || id == BuiltinFunctionId::Max ||
        id == BuiltinFunctionId::Avg || id == BuiltinFunctionId::Mean || id == BuiltinFunctionId::Median) {
      return aggregateFoldByMode(ctx, fnName, id, args, AggregateFoldMode::Time);
    }
  }
#endif
  return aggregateFoldByMode(ctx, fnName, id, args, AggregateFoldMode::Real);
}

bool MathParser::isSortbyIneligibleBuiltin(BuiltinFunctionId id) {
  switch (id) {
    case BuiltinFunctionId::Rand:
    case BuiltinFunctionId::Pow:
    case BuiltinFunctionId::Atan2:
    case BuiltinFunctionId::Hypot:
    case BuiltinFunctionId::Gcd:
    case BuiltinFunctionId::Lcm:
    case BuiltinFunctionId::Ncr:
    case BuiltinFunctionId::Npr:
    case BuiltinFunctionId::Mod:
    case BuiltinFunctionId::Clamp:
    case BuiltinFunctionId::Log:
    case BuiltinFunctionId::Random:
      return true;
    default:
      return false;
  }
}

bool MathParser::isSortbyEligibleFunctionName(
    const MathParser& parser,
    const std::string& funcName,
    std::string& outErr) {
  const std::string lowFn = toLower(funcName);
  BuiltinFunctionId fnId = BuiltinFunctionId::Count;
  if (tryGetBuiltinFunctionId(lowFn, fnId)) {
    if (isSortbyIneligibleBuiltin(fnId)) {
      outErr = STR_SORTBY_EXPECTS_UNARY_FUNCTION;
      return false;
    }
    return true;
  }
  const UserFunction* uf = parser.findUserFunction(lowFn);
  if (!uf) {
    if (parser.variables_.find(lowFn) != parser.variables_.end()) {
      outErr = STR_SORTBY_EXPECTS_UNARY_FUNCTION;
      return false;
    }
    outErr = std::string(STR_UNKNOWN_FUNCTION_COLON) + funcName;
    return false;
  }
  if (uf->params.size() != 1U) {
    outErr = STR_SORTBY_EXPECTS_UNARY_FUNCTION;
    return false;
  }
  return true;
}

bool MathParser::tryLexicographicCompareEvalValues(
    EvalContext& ctx,
    const EvalValue& a,
    const EvalValue& b,
    int& cmpOut) const {
#if SMARTMATH_COMPLEX_NUMBERS
  if (!getSupportComplexNumbers()) {
    if (evalValueHasNonzeroImaginary(a) || evalValueHasNonzeroImaginary(b)) {
      setParseError(ctx, ParseErrorId::IncompatibleOperands);
      cmpOut = 0;
      return false;
    }
  }
#endif
  if (a.kind == ValueKind::Scalar && b.kind == ValueKind::Scalar) {
    if (!cmpScalarValuesForCompare(
            &ctx, a.scalarValue, b.scalarValue, cmpOut, CmpScalarIncompatiblePolicy::SetError)) {
      return false;
    }
    return !ctx.parseError;
  }
  const auto refAt = [](const EvalValue& v, std::size_t i) -> const EvalValue::ScalarValue& {
    return (v.kind == ValueKind::Scalar) ? v.scalarValue : v.arr[i];
  };
  const std::size_t na = (a.kind == ValueKind::Scalar) ? 1U : a.arr.size();
  const std::size_t nb = (b.kind == ValueKind::Scalar) ? 1U : b.arr.size();
  const std::size_t n = (na < nb) ? na : nb;
  for (std::size_t i = 0; i < n; ++i) {
    int c = 0;
    if (!cmpScalarValuesForCompare(
            &ctx, refAt(a, i), refAt(b, i), c, CmpScalarIncompatiblePolicy::SetError)) {
      cmpOut = 0;
      return false;
    }
    if (ctx.parseError) {
      cmpOut = 0;
      return false;
    }
    if (c != 0) {
      cmpOut = c;
      return true;
    }
  }
  if (na < nb) {
    cmpOut = -1;
  } else if (na > nb) {
    cmpOut = 1;
  } else {
    cmpOut = 0;
  }
  return true;
}

bool MathParser::sortbyStableSortIndicesFromKeys(
    EvalContext& ctx,
    const std::vector<EvalValue>& sortKeys,
    std::vector<int>& orderIdx) const {
  const int count = static_cast<int>(sortKeys.size());
  orderIdx.resize(static_cast<std::size_t>(count));
  for (int i = 0; i < count; ++i) {
    orderIdx[static_cast<std::size_t>(i)] = i;
  }
  if (count <= 1) {
    return true;
  }
  for (int i = 1; i < count; ++i) {
    int j = i;
    while (j > 0) {
      int cmp = 0;
      if (!tryLexicographicCompareEvalValues(
              ctx,
              sortKeys[static_cast<std::size_t>(orderIdx[static_cast<std::size_t>(j)])],
              sortKeys[static_cast<std::size_t>(orderIdx[static_cast<std::size_t>(j - 1)])],
              cmp)) {
        return false;
      }
      if (cmp < 0) {
        const int t = orderIdx[static_cast<std::size_t>(j)];
        orderIdx[static_cast<std::size_t>(j)] = orderIdx[static_cast<std::size_t>(j - 1)];
        orderIdx[static_cast<std::size_t>(j - 1)] = t;
        --j;
      } else {
        break;
      }
    }
  }
  return true;
}

MathParser::EvalValue MathParser::makeRationalReduced(long long num, std::uint64_t den) {
  EvalValue out = makeScalarMaybeExact(static_cast<double>(num) / static_cast<double>(den));
  scalarApplyReducedRationalPart(out.scalarValue, false, num, den);
  return out;
}

namespace {

inline double ratioApproxAbsNumerator(double v, long long p, std::uint64_t q) {
  return std::fabs(v * static_cast<double>(q) - static_cast<double>(p));
}

inline bool ratioApproxErrLess(double n1, std::uint64_t q1, double n2, std::uint64_t q2) {
  return n1 * static_cast<double>(q2) < n2 * static_cast<double>(q1);
}

struct RatioRationalSearch {
  const double v;
  long long bestP = 0;
  std::uint64_t bestQ = 1;
  double bestAbsNum = 1e300;

  explicit RatioRationalSearch(double vIn) : v(vIn) {}

  bool consider(long long p, std::uint64_t q) {
    if (q == 0 || q > static_cast<std::uint64_t>(RATIO_MAX_DENOMINATOR)) {
      return false;
    }
    const double absNum = ratioApproxAbsNumerator(v, p, q);
    if (ratioApproxErrLess(absNum, q, bestAbsNum, bestQ)) {
      bestAbsNum = absNum;
      bestP = p;
      bestQ = q;
      return absNum == 0.0;
    }
    return false;
  }

  bool semiconvergentAbsNumLess(long long p1, long long q1, long long p0, long long q0, long long h1,
                                long long h2) const {
    const long long qh1 = q1 + h1 * q0;
    const long long qh2 = q1 + h2 * q0;
    if (qh1 <= 0 || qh2 <= 0) {
      return false;
    }
    const long long ph1 = p1 + h1 * p0;
    const long long ph2 = p1 + h2 * p0;
    const double n1 = ratioApproxAbsNumerator(v, ph1, static_cast<std::uint64_t>(qh1));
    const double n2 = ratioApproxAbsNumerator(v, ph2, static_cast<std::uint64_t>(qh2));
    return ratioApproxErrLess(n1, static_cast<std::uint64_t>(qh1), n2, static_cast<std::uint64_t>(qh2));
  }

  void scanSemiconvergentRange(long long p1, long long q1, long long p0, long long q0, long long hMax) {
    if (hMax < 0) {
      return;
    }
    if (hMax <= RATIO_SEMICONV_LINEAR_THRESH) {
      for (long long h = 0; h <= hMax; ++h) {
        if (consider(p1 + h * p0, static_cast<std::uint64_t>(q1 + h * q0))) {
          return;
        }
      }
      return;
    }
    long long lo = 0;
    long long hi = hMax;
    if (consider(p1, static_cast<std::uint64_t>(q1))) {
      return;
    }
    while (hi - lo > 2) {
      const long long span = hi - lo;
      const long long m1 = lo + span / 3;
      const long long m2 = hi - span / 3;
      if (semiconvergentAbsNumLess(p1, q1, p0, q0, m1, m2)) {
        hi = m2;
      } else {
        lo = m1;
      }
    }
    for (long long h = lo; h <= hi; ++h) {
      if (consider(p1 + h * p0, static_cast<std::uint64_t>(q1 + h * q0))) {
        return;
      }
    }
  }
};

bool tryExactPower10Rational(double v, RatioRationalSearch& search) {
  for (int k = 1; k <= RATIO_MAX_POWER10_EXP; ++k) {
    const std::uint64_t denPow = pow10_u64[k];
    const double scaled = v * static_cast<double>(denPow);
    const long long n = static_cast<long long>(std::llround(scaled));
    if (n == 0) {
      continue;
    }
    const long long nAbs = std::llabs(n);
    double scaleErr = RATIO_APPROX_EPS * std::fabs(scaled);
    if (scaleErr < RATIO_APPROX_EPS) {
      scaleErr = RATIO_APPROX_EPS;
    }
    if (std::fabs(scaled - static_cast<double>(n)) > scaleErr) {
      continue;
    }
    const long long denPowLong = static_cast<long long>(denPow);
    const long long g = gcdInt64(nAbs, denPowLong);
    const long long candNum = n / g;
    const std::uint64_t candDen = static_cast<std::uint64_t>(denPowLong / g);
    if (candDen == 0) {
      continue;
    }
    const double absNum = ratioApproxAbsNumerator(v, candNum, candDen);
    if (ratioApproxErrLess(absNum, candDen, search.bestAbsNum, search.bestQ)) {
      search.bestAbsNum = absNum;
      search.bestP = candNum;
      search.bestQ = candDen;
      if (absNum == 0.0) {
        return true;
      }
    }
  }
  return false;
}

}  // namespace

bool MathParser::tryApproximateRational(double x, long long& num, std::uint64_t& den) {
  if (std::isnan(x) || std::isinf(x)) {
    return false;
  }
  const bool neg = x < 0.0;
  const double v = std::fabs(x);
  double scale = v;
  if (scale < 1.0) {
    scale = 1.0;
  }
  const double tol = RATIO_APPROX_EPS * scale;
  if (v < tol) {
    num = 0;
    den = 1;
    return true;
  }
  RatioRationalSearch search(v);
  long long p0 = 0;
  long long p1 = 1;
  long long q0 = 1;
  long long q1 = 0;
  double frac = v;
  for (int i = 0; i < 64; ++i) {
    double ipart = 0.0;
    const double t = std::modf(frac, &ipart);
    const long long a = static_cast<long long>(ipart);
    const long long p = a * p1 + p0;
    const long long q = a * q1 + q0;
    if (q > RATIO_MAX_DENOMINATOR) {
      if (q1 > 0 && q0 > 0) {
        const long long hMax = (RATIO_MAX_DENOMINATOR - q1) / q0;
        search.scanSemiconvergentRange(p1, q1, p0, q0, hMax);
      }
      break;
    }
    if (search.consider(p, static_cast<std::uint64_t>(q))) {
      break;
    }
    if (t <= RATIO_APPROX_EPS) {
      break;
    }
    frac = 1.0 / t;
    p0 = p1;
    p1 = p;
    q0 = q1;
    q1 = q;
  }
  if (v <= 1.0 / static_cast<double>(RATIO_MAX_DENOMINATOR)) {
    tryExactPower10Rational(v, search);
  }
  if (search.bestP == 0) {
    num = 0;
    den = 1;
    return true;
  }
  const long long g = gcdInt64(std::llabs(search.bestP), static_cast<long long>(search.bestQ));
  num = search.bestP / g;
  den = static_cast<std::uint64_t>(static_cast<long long>(search.bestQ) / g);
  if (neg) {
    num = -num;
  }
  return true;
}

bool MathParser::tryBuiltinRatioScalar(EvalContext& ctx, const EvalValue::ScalarValue& sv, EvalValue& outV) const {
#if SMARTMATH_TIME_VALUES
  if (scalarValueIsTime(sv)) {
    ctx.parseError = true;
    return false;
  }
#endif
  double ar = 0.0;
  double ai = 0.0;
  scalarLoadCartesian(sv, ar, ai);
  if (std::isnan(ar) || std::isnan(ai)) {
    if (std::isnan(ar)) {
      if (std::isnan(ai) || std::fabs(ai) < RATIO_APPROX_EPS) {
        outV = makeScalar(std::numeric_limits<double>::quiet_NaN());
        return true;
      }
    }
    ctx.parseError = true;
    return false;
  }
  if (!scalarHasNonzeroImaginaryPart(sv)) {
    if (std::isinf(ar)) {
      outV = makeScalar(ar);
      return true;
    }
    if (std::isinf(ai)) {
      ctx.parseError = true;
      return false;
    }
  }
#if SMARTMATH_COMPLEX_NUMBERS
  if (getSupportComplexNumbers() && scalarHasNonzeroImaginaryPart(sv)) {
    long long numR = 0;
    std::uint64_t denR = 1;
    long long numI = 0;
    std::uint64_t denI = 1;
    ExactCartesianComponent reC;
    double aiScale = std::fabs(ai);
    if (aiScale < 1.0) {
      aiScale = 1.0;
    }
    double arScale = std::fabs(ar);
    if (arScale < 1.0) {
      arScale = 1.0;
    }
    outV = scalarFromScalarValue(sv);
    if (sv.hasRenderRational() || sv.hasImagRenderRational()) {
      return true;
    }
    scalarClearCartesianRenderExact(outV.scalarValue);
    if (sv.hasImagExactInt64() && imagExactMetadataMatchesFloat(sv)) {
      scalarApplyExactInt64Part(outV.scalarValue, true, sv.imagExactInt64);
    } else if (std::fabs(ai) >= RATIO_APPROX_EPS * aiScale) {
      if (!tryApproximateRational(ai, numI, denI)) {
        ctx.parseError = true;
        return false;
      }
      scalarApplyReducedRationalPart(outV.scalarValue, true, numI, denI);
    }
    if (tryExtractExactRealComponent(sv, reC) && tryExactCartesianComponentToInt64(reC, numR)) {
      scalarApplyExactInt64Part(outV.scalarValue, false, numR);
    } else if (std::fabs(ar) >= RATIO_APPROX_EPS * arScale) {
      if (!tryApproximateRational(ar, numR, denR)) {
        ctx.parseError = true;
        return false;
      }
      scalarApplyReducedRationalPart(outV.scalarValue, false, numR, denR);
    } else if (sv.hasExactInt64()) {
      scalarApplyExactInt64Part(outV.scalarValue, false, sv.exactInt64);
    }
    return true;
  }
#endif
  if (sv.hasExactInt64()) {
    outV = makeScalarInt(sv.exactInt64);
    return true;
  }
  if (sv.hasExactUInt64()) {
    outV = makeScalarUInt(sv.exactUInt64);
    return true;
  }
  double arScale0 = std::fabs(ar);
  if (arScale0 < 1.0) {
    arScale0 = 1.0;
  }
  if (std::fabs(ar) < RATIO_APPROX_EPS * arScale0) {
    outV = makeScalar(0.0);
    return true;
  }
  long long nearInt = 0;
  if (tryExtractExactInt64FromDoubleStrict(ar, nearInt)) {
    outV = makeScalarInt(nearInt);
    return true;
  }
  long long num = 0;
  std::uint64_t den = 1;
  if (!tryApproximateRational(ar, num, den)) {
    ctx.parseError = true;
    return false;
  }
  if (den == 1u) {
    outV = makeScalarInt(num);
  } else {
    outV = makeRationalReduced(num, den);
  }
  return true;
}

MathParser::EvalValue MathParser::builtinRatio(
    EvalContext& ctx,
    const std::string& fnName,
    const std::vector<EvalValue>& args) const {
  EvalContext sub;
  std::vector<EvalValue> flat;
  if (!flattenArgsToScalars({args[0]}, flat)) {
    setAtLeastOneArgError(ctx, fnName);
    return makeScalar(0);
  }
  if (flat.size() == 1U && args[0].kind == ValueKind::Scalar) {
    EvalValue out = makeScalar(0);
    if (!tryBuiltinRatioScalar(sub, flat[0].scalarValue, out)) {
      if (sub.parseError) {
        setParseError(ctx, ParseErrorId::IncompatibleOperands);
      }
      return makeScalar(0);
    }
    return out;
  }
  std::vector<EvalValue> wrapped;
  wrapped.reserve(flat.size());
  for (const auto& item : flat) {
    EvalValue tmp = makeScalar(0);
    if (!tryBuiltinRatioScalar(sub, item.scalarValue, tmp)) {
      if (sub.parseError) {
        setParseError(ctx, ParseErrorId::IncompatibleOperands);
      }
      return makeScalar(0);
    }
    wrapped.push_back(scalarFromScalarValue(tmp.scalarValue));
  }
  return makeArrayFromScalars(wrapped);
}

MathParser::EvalValue MathParser::sortbyInvokeKeyFunction(
    EvalContext& ctx,
    const std::string& funcName,
    const EvalValue::ScalarValue& elem,
    const std::unordered_map<std::string, EvalValue>* scopedVars) {
  std::vector<EvalValue> args;
  args.push_back(scalarFromScalarValue(elem));
  BuiltinFunctionId fnId = BuiltinFunctionId::Count;
  if (!tryGetBuiltinFunctionId(toLower(funcName), fnId)) {
    return evalUserFunctionCall(ctx, funcName, args, scopedVars);
  }
  if (getBuiltinCategory(fnId) == BuiltinCategory::Sortby ||
      getBuiltinCategory(fnId) == BuiltinCategory::Ratio) {
    setValidationError(ctx, STR_SORTBY_EXPECTS_UNARY_FUNCTION);
    return makeScalar(0);
  }
  EvalValue out = evalFunctionCall(ctx, funcName, std::move(args), fnId, scopedVars);
  if (ctx.parseError) {
    return out;
  }
  if (out.kind == ValueKind::Scalar || out.kind == ValueKind::Array) {
    return out;
  }
  setValidationError(ctx, STR_SORTBY_EXPECTS_UNARY_FUNCTION);
  return makeScalar(0);
}

MathParser::EvalValue MathParser::builtinSortby(
    EvalContext& ctx,
    const std::string& fnName,
    const std::vector<EvalValue>& args,
    const std::unordered_map<std::string, EvalValue>* scopedVars) {
  const bool keyIsLambda = (args[1].kind == ValueKind::InlineLambda);
  const bool keyIsRef = (args[1].kind == ValueKind::FunctionRef);
  if (!keyIsLambda && !keyIsRef) {
    if (args[1].kind == ValueKind::Array) {
      setValidationError(ctx, STR_SORTBY_EXPECTS_ONE_FUNCTION);
    } else {
      setValidationError(ctx, STR_SORTBY_EXPECTS_UNARY_FUNCTION);
    }
    return makeScalar(0);
  }
  if (keyIsRef) {
    std::string errText;
    if (!isSortbyEligibleFunctionName(*this, args[1].funcRefName, errText)) {
      setValidationError(ctx, errText.c_str());
      return makeScalar(0);
    }
  }
  std::vector<EvalValue> flat;
  if (!flattenArgsToScalars({args[0]}, flat)) {
    setAtLeastOneArgError(ctx, fnName);
    return makeScalar(0);
  }
  std::vector<EvalValue::ScalarValue> vals;
  vals.reserve(flat.size());
  for (const auto& item : flat) {
    vals.push_back(item.scalarValue);
  }
  std::vector<EvalValue> keys;
  keys.reserve(vals.size());
  for (std::size_t i = 0; i < vals.size(); ++i) {
    EvalValue keyV;
#if SMARTMATH_LAMBDA_FUNCTIONS
    if (keyIsLambda) {
      std::vector<EvalValue> lamArgs;
      lamArgs.push_back(scalarFromScalarValue(vals[i]));
      keyV = evalInlineLambdaCall(ctx, args[1].lambdaParams, args[1].lambdaBody, std::move(lamArgs));
    } else
#endif
    {
      keyV = sortbyInvokeKeyFunction(ctx, args[1].funcRefName, vals[i], scopedVars);
    }
    if (ctx.parseError) {
      return makeScalar(0);
    }
    if (keyV.kind == ValueKind::FunctionRef || keyV.kind == ValueKind::InlineLambda ||
        (keyV.kind != ValueKind::Scalar && keyV.kind != ValueKind::Array)) {
      setValidationError(ctx, STR_SORTBY_KEY_MUST_BE_SCALAR_OR_ARRAY);
      return makeScalar(0);
    }
    keys.push_back(keyV);
  }
  std::vector<int> order;
  if (!sortbyStableSortIndicesFromKeys(ctx, keys, order)) {
    return makeScalar(0);
  }
  std::vector<EvalValue::ScalarValue> sorted(vals.size());
  for (std::size_t i = 0; i < vals.size(); ++i) {
    sorted[i] = vals[static_cast<std::size_t>(order[i])];
  }
  std::vector<EvalValue> wrapped;
  wrapped.reserve(sorted.size());
  for (const auto& s : sorted) {
    wrapped.push_back(scalarFromScalarValue(s));
  }
  return makeArrayFromScalars(wrapped);
}

std::unique_ptr<MathParser::Expr> MathParser::parseSortbyFunctionRef(EvalContext& ctx) {
  skipSpaces(ctx);
  if (*ctx.p == '(') {
    auto parsed = parseExpression(ctx);
    if (ctx.parseError || !parsed) {
      return nullptr;
    }
    if (parsed->tag == Expr::Tag::ArrayOrParens) {
      setValidationError(ctx, STR_SORTBY_EXPECTS_ONE_FUNCTION);
      return nullptr;
    }
    setValidationError(ctx, STR_SORTBY_EXPECTS_UNARY_FUNCTION);
    return nullptr;
  }
  if (!isIdentStart(*ctx.p)) {
    setValidationError(ctx, STR_SORTBY_EXPECTS_UNARY_FUNCTION);
    return nullptr;
  }
  const char* start = ctx.p;
  ++ctx.p;
  while (isIdentChar(*ctx.p)) {
    ++ctx.p;
  }
  const std::string nameText(start, static_cast<std::size_t>(ctx.p - start));
  skipSpaces(ctx);
  if (*ctx.p == ':') {
    setParseError(ctx, ParseErrorId::UnexpectedToken);
    return nullptr;
  }
  if (*ctx.p == '(') {
    setValidationError(ctx, STR_SORTBY_EXPECTS_ONE_FUNCTION);
    return nullptr;
  }
  auto ref = std::make_unique<Expr>();
  ref->tag = Expr::Tag::FunctionRef;
  ref->name = toLower(nameText);
  return ref;
}

bool MathParser::parseSortbyCallArguments(EvalContext& ctx, std::vector<std::unique_ptr<Expr>>& outArgs) {
  skipSpaces(ctx);
  if (*ctx.p == ',') {
    setParseError(ctx, ParseErrorId::UnexpectedComma);
    return false;
  }
  if (*ctx.p == ')') {
    setValidationError(ctx, STR_SORTBY_EXPECTS_UNARY_FUNCTION);
    return false;
  }
  auto first = parseExpression(ctx);
  if (ctx.parseError || !first) {
    return false;
  }
  outArgs.push_back(std::move(first));
  skipSpaces(ctx);
  bool hasComma = false;
  if (!tryConsumeCommaArgSeparator(ctx, hasComma)) {
    return false;
  }
  if (!hasComma) {
    setValidationError(ctx, STR_SORTBY_EXPECTS_UNARY_FUNCTION);
    return false;
  }
  auto keyArg = parseSortbyKeyArg(ctx);
  if (ctx.parseError || !keyArg) {
    return false;
  }
  outArgs.push_back(std::move(keyArg));
  skipSpaces(ctx);
  if (!tryConsumeCommaArgSeparator(ctx, hasComma)) {
    return false;
  }
  if (hasComma) {
    setValidationError(ctx, STR_SORTBY_EXPECTS_ONE_FUNCTION);
    return false;
  }
  return true;
}

MathParser::EvalValue MathParser::builtinSortFamily(
    EvalContext& ctx,
    const std::string& fnName,
    BuiltinFunctionId id,
    const std::vector<EvalValue>& args) const {
  auto keyForUnique = [](double v) -> std::uint64_t {
    std::uint64_t bits = 0;
    static_assert(sizeof(bits) == sizeof(v));
    std::memcpy(&bits, &v, sizeof(bits));
    // Canonicalize signed zero to preserve equality semantics of (0 == -0).
    if ((bits << 1) == 0) bits = 0;
    return bits;
  };
  auto dedupUniqueInPlace = [&](std::vector<EvalValue>& values) {
#if SMARTMATH_COMPLEX_NUMBERS
    bool useComplexDedup = false;
    if (getSupportComplexNumbers()) {
      for (const auto& ev : values) {
        if (scalarHasNonzeroImaginaryPart(ev.scalarValue)) {
          useComplexDedup = true;
          break;
        }
      }
    }
    if (useComplexDedup) {
      std::vector<EvalValue> out;
      out.reserve(values.size());
      for (const auto& ev : values) {
        bool seen = false;
        for (const auto& prev : out) {
          int cmpDup = 0;
          if (cmpScalarValuesForCompare(
                  nullptr,
                  ev.scalarValue,
                  prev.scalarValue,
                  cmpDup,
                  CmpScalarIncompatiblePolicy::SortUniqueReturnOne) &&
              cmpDup == 0) {
            seen = true;
            break;
          }
        }
        if (!seen) {
          out.push_back(ev);
        }
      }
      values = std::move(out);
      return;
    }
#endif
    std::unordered_set<std::uint64_t> seen;
    seen.reserve(values.size() * 2 + 1);
    std::size_t writePos = 0;
    for (std::size_t readPos = 0; readPos < values.size(); ++readPos) {
      const double v = values[readPos].scalarValue.scalar;
      if (std::isnan(v)) {
        values[writePos++] = values[readPos]; // NaN never equals NaN in existing semantics; keep all.
        continue;
      }
      if (seen.insert(keyForUnique(v)).second) {
        values[writePos++] = values[readPos];
      }
    }
    values.resize(writePos);
  };

#if SMARTMATH_COMPLEX_NUMBERS
  if (getSupportComplexNumbers() && id == BuiltinFunctionId::Sort &&
      std::any_of(args.begin(), args.end(), [](const EvalValue& v) { return MathParser::evalValueHasNonzeroImaginary(v); })) {
    setParseError(ctx, ParseErrorId::IncompatibleOperands);
    return makeScalar(0);
  }
#endif
#if SMARTMATH_TIME_VALUES
  if (id == BuiltinFunctionId::Sort && getSupportTimeValues() &&
      std::any_of(args.begin(), args.end(), [this](const EvalValue& v) { return evalValueInvolvesTime(v); })) {
    bool okTimeSort = true;
    forEachCallArgScalarValues(args, [&](const EvalValue::ScalarValue& s) {
      if (!okTimeSort) {
        return;
      }
      if (!scalarValueIsTime(s)) {
        setError(ctx, constructTimeArrayMixedError(fnName));
        okTimeSort = false;
      }
    });
    if (!okTimeSort) {
      return makeScalar(0);
    }
  }
#endif
  const auto copyArrayScalars = [&](const EvalValue& a, std::vector<EvalValue>& out) {
    out.clear();
    out.reserve(a.arr.size());
    for (const auto& item : a.arr) {
      out.emplace_back(scalarFromScalarValue(item));
    }
  };
  const auto copySingleArgToFlat = [&](const EvalValue& a, std::vector<EvalValue>& out) {
    out.clear();
    if (a.kind == ValueKind::Scalar) {
      out.emplace_back(a);
      return;
    }
    copyArrayScalars(a, out);
  };
  const auto sortScalarsInPlace = [this](std::vector<EvalValue>& values) {
    std::sort(values.begin(), values.end(), [this](const EvalValue& lhs, const EvalValue& rhs) {
      int cmp = 0;
      cmpScalarValuesForCompare(
          nullptr, lhs.scalarValue, rhs.scalarValue, cmp, CmpScalarIncompatiblePolicy::SortLess);
      return cmp < 0;
    });
  };
  if (args.size() == 1) {
    const EvalValue& a = args[0];
    if (a.kind == ValueKind::Scalar) {
      return makeArrayFromScalars(std::vector<EvalValue>{a});
    }
    std::vector<EvalValue> out;
    copySingleArgToFlat(a, out);
    if (out.empty()) {
      setAtLeastOneArgError(ctx, fnName);
      return makeScalar(0);
    }
    if (id == BuiltinFunctionId::Sort) {
      sortScalarsInPlace(out);
      return makeArrayFromScalars(out);
    }
    if (id == BuiltinFunctionId::Reverse) {
      std::reverse(out.begin(), out.end());
      return makeArrayFromScalars(out);
    }
    if (id == BuiltinFunctionId::Unique) {
      dedupUniqueInPlace(out);
      return makeArrayFromScalars(out);
    }
  }

  std::vector<EvalValue> flat;
  if (!flattenArgsToScalars(args, flat)) {
    setAtLeastOneArgError(ctx, fnName);
    return makeScalar(0);
  }
  if (id == BuiltinFunctionId::Sort) {
    sortScalarsInPlace(flat);
  } else if (id == BuiltinFunctionId::Reverse) {
    std::reverse(flat.begin(), flat.end());
  } else {
    dedupUniqueInPlace(flat);
  }
  return makeArrayFromScalars(flat);
}

MathParser::EvalValue MathParser::builtinTimeUnit(
    EvalContext& ctx,
    const std::string& fnName,
    BuiltinFunctionId id,
    const std::vector<EvalValue>& args) const {
  if (!getSupportTimeValues()) {
    setParseError(ctx, ParseErrorId::IncompatibleOperands);
    return makeScalar(0);
  }
#if SMARTMATH_TIME_VALUES
  const EvalValue& a0 = args[0];
  if (a0.kind == ValueKind::Scalar) {
    if (!scalarValueIsTime(a0.scalarValue)) {
      setError(ctx, fnName + STR_TIME_EXPECTS_TIME_ARG);
      return makeScalar(0);
    }
    return evalValueFromTimeMs(id, timeTotalMsFromScalarValue(a0.scalarValue));
  }
  if (a0.kind == ValueKind::Array) {
    return mapTimeUnitOverArray(ctx, fnName, id, a0);
  }
  setError(ctx, fnName + STR_TIME_EXPECTS_TIME_ARG);
#endif
  return makeScalar(0);
}

MathParser::EvalValue MathParser::builtinBaseFormat(
    EvalContext& ctx,
    const std::string& fnName,
    BuiltinFunctionId id,
    const std::vector<EvalValue>& args) const {
  EvalValue out;
  if (args.size() == 1) {
    out = args[0];
  } else {
    std::vector<EvalValue> elems;
    if (!flattenArgsToScalars(args, elems)) {
      setAtLeastOneArgError(ctx, fnName);
      return makeScalar(0);
    }
    out = makeArrayFromScalars(elems);
  }
  applyBuiltinFormatRenderMeta(id, out);
  return out;
}

MathParser::EvalValue MathParser::builtinPow(EvalContext& ctx, const std::vector<EvalValue>& args) const {
  const std::string fnName = getFunctionName(BuiltinFunctionId::Pow);
#if SMARTMATH_TIME_VALUES
  if (evalValueInvolvesTime(args[0]) || evalValueInvolvesTime(args[1])) {
    setParseError(ctx, ParseErrorId::IncompatibleOperands);
    return makeScalar(0);
  }
#endif
  bool ok = false;
  EvalValue out = mapBinary(ctx, args[0], args[1], '^', ok);
  if (!ok) {
    if (!ctx.parseError) {
      setBinaryBuiltinBroadcastFailure(ctx, fnName, args[0], args[1], 2);
    }
    return makeScalar(0);
  }
  return out;
}

MathParser::EvalValue MathParser::builtinScalarBinaryFamily(
    EvalContext& ctx,
    const std::string& fnName,
    BuiltinFunctionId id,
    const std::vector<EvalValue>& args) const {
  if (id == BuiltinFunctionId::Atan2) {
    return builtinMapBinaryTwoArg(ctx, fnName, id, args, true, false);
  }
  if (id == BuiltinFunctionId::Hypot) {
    return builtinMapBinaryTwoArg(ctx, fnName, id, args, false, true);
  }

  if (id == BuiltinFunctionId::Clamp) {
    if (args[1].kind != ValueKind::Scalar || args[2].kind != ValueKind::Scalar) {
      setScalarMinMaxError(ctx, fnName);
      return makeScalar(0);
    }
    return builtinApplyClamp(ctx, args[0], args[1], args[2]);
  }

  if (id == BuiltinFunctionId::Gcd || id == BuiltinFunctionId::Lcm || id == BuiltinFunctionId::Ncr ||
      id == BuiltinFunctionId::Npr) {
    if (rejectBuiltinArgsWithComplexImaginary(ctx, args)) {
      return makeScalar(0);
    }
    int pairStatus = 0;
    EvalValue out;
    if (id == BuiltinFunctionId::Gcd || id == BuiltinFunctionId::Lcm) {
      out = applyGcdLcmEvalValues(args[0], args[1], id == BuiltinFunctionId::Lcm, pairStatus);
    } else {
      out = applyNcrNprEvalValues(args[0], args[1], id == BuiltinFunctionId::Npr, pairStatus);
    }
    if (pairStatus != 0) {
      setBinaryBuiltinBroadcastFailure(ctx, fnName, args[0], args[1], pairStatus);
      return makeScalar(0);
    }
    return out;
  }

  const bool hasNonScalarArg = std::any_of(args.begin(), args.end(), [](const EvalValue& v) {
    return v.kind != ValueKind::Scalar;
  });
  if (hasNonScalarArg) {
    setScalarValuesError(ctx, fnName);
    return makeScalar(0);
  }
  switch (id) {
    case BuiltinFunctionId::Random:
      return makeScalar(
          args[0].scalarValue.scalar + (args[1].scalarValue.scalar - args[0].scalarValue.scalar) *
              randomUnitScalar());
    default:
      break;
  }
  setParseError(ctx, ParseErrorId::InternalScalarBinaryBuiltin);
  return makeScalar(0);
}

MathParser::EvalValue MathParser::builtinRand(EvalContext& /*ctx*/, const std::vector<EvalValue>& /*args*/) const {
  return makeScalar(randomUnitScalar());
}

MathParser::EvalValue MathParser::builtinModCall(EvalContext& ctx, const std::vector<EvalValue>& args) const {
  const std::string fnName = getFunctionName(BuiltinFunctionId::Mod);
  if (rejectBuiltinArgsWithComplexImaginary(ctx, args)) {
    return makeScalar(0);
  }
  bool ok = true;
  const auto applyPair = [&](const EvalValue::ScalarValue& aS, const EvalValue::ScalarValue& bS,
                             EvalValue& outS) -> bool {
    return tryApplyModScalars(ctx, fnName, aS, bS, outS);
  };
  EvalValue out = mapBinaryBroadcast(args[0], args[1], applyPair, ok);
  if (!ok) {
    if (ctx.errorText.empty() && !ctx.parseError) {
      setBinaryBuiltinBroadcastFailure(ctx, fnName, args[0], args[1], 2);
    }
    return makeScalar(0);
  }
  return out;
}

MathParser::EvalValue MathParser::builtinFactorial(
    EvalContext& ctx,
    const std::string& fnName,
    const std::vector<EvalValue>& args) const {
#if SMARTMATH_COMPLEX_NUMBERS
  if (getSupportComplexNumbers()) {
    return builtinUnaryMath(ctx, fnName, BuiltinFunctionId::Fact, args);
  }
#endif
  long long n = 0;
  if (args[0].kind != ValueKind::Scalar || !nearlyInt(args[0].scalarValue.scalar, n) || n < 0) {
    setNonNegativeIntegerError(ctx, fnName);
    return makeScalar(0);
  }
  if (n <= 20) {
    return makeScalarInt(kFactorialTable21[n]);
  }
  return makeScalar(factorialScalarFloatFromInt(n));
}


MathParser::EvalValue MathParser::builtinDegRad(
    EvalContext& ctx,
    const std::string& fnName,
    BuiltinFunctionId id,
    const std::vector<EvalValue>& args) const {
#if SMARTMATH_COMPLEX_NUMBERS
  if (getSupportComplexNumbers()) {
    for (const auto& a : args) {
      if (evalValueHasNonzeroImaginary(a)) {
        setParseError(ctx, ParseErrorId::IncompatibleOperands);
        return makeScalar(0);
      }
    }
  }
#endif
  const bool toDeg = (id == BuiltinFunctionId::Deg);
  if (args.size() == 1) {
    return mapUnaryEvalValue(args[0], [&](const EvalValue::ScalarValue& s) -> EvalValue {
      const double x = s.scalar;
      return makeScalarMaybeExact(toDeg ? (x * 180.0 / kPi) : (x * kPi / 180.0));
    });
  }
  std::vector<EvalValue> inVals;
  if (!flattenArgsToScalars(args, inVals)) {
    setAtLeastOneArgError(ctx, fnName);
    return makeScalar(0);
  }
  std::vector<EvalValue> outVals;
  outVals.reserve(inVals.size());
  for (const auto& item : inVals) {
    const double x = item.scalarValue.scalar;
    outVals.emplace_back(makeScalarMaybeExact(toDeg ? (x * 180.0 / kPi) : (x * kPi / 180.0)));
  }
  return makeArrayFromScalars(outVals);
}

MathParser::EvalValue MathParser::applyAbsScalarValue(const EvalValue::ScalarValue& s) {
  const double x = s.scalar;
  if (std::isnan(x)) {
    return makeScalar(std::numeric_limits<double>::quiet_NaN());
  }
  if (s.hasExactUInt64() && !s.hasExactInt64()) {
    return makeScalarUInt(s.exactUInt64);
  }
  if (s.hasExactInt64()) {
    if (s.exactInt64 >= 0) {
      return makeScalarInt(s.exactInt64);
    }
    if (s.exactInt64 == (std::numeric_limits<long long>::min)()) {
      return makeScalarUInt(1ull << 63);
    }
    return makeScalarInt(-s.exactInt64);
  }
  if (s.hasExactUInt64()) {
    return makeScalarUInt(s.exactUInt64);
  }
  return makeScalarMaybeExact(std::fabs(x));
}

MathParser::EvalValue MathParser::calcRoundingFn(BuiltinFunctionId id, const EvalValue::ScalarValue& s)
{
  const double x = s.scalar;
  if (!std::isfinite(x))
    return makeScalar(x);

  if (s.hasExactUInt64())
    return makeScalarUInt(s.exactUInt64);

  if (s.hasExactInt64())
    return makeScalarInt(s.exactInt64);

  if (x > K_MAX_EXACT_INT_FROM_DOUBLE || x < -K_MAX_EXACT_INT_FROM_DOUBLE)
  {
    if (x <= (std::numeric_limits<long long>::max)() && x >= (std::numeric_limits<long long>::min)())
      return makeScalarInt(static_cast<long long>(x));

    const double p64 = std::ldexp(1.0, 64);
    if (x >= 0.0 && x < p64 && x == std::trunc(x)) {
      const std::uint64_t u = static_cast<std::uint64_t>(x);
      if (static_cast<double>(u) == x)
        return makeScalarUInt(u);
    }

    switch (id) {
      case BuiltinFunctionId::Floor: return makeScalar(std::floor(x));
      case BuiltinFunctionId::Ceil: return makeScalar(std::ceil(x));
      case BuiltinFunctionId::Trunc:
      case BuiltinFunctionId::Int: return makeScalar(std::trunc(x));
      case BuiltinFunctionId::Round: return makeScalar(std::round(x));
      default: return makeScalar(x);
    }
  }

  long long rounded = 0;
  switch (id) {
    case BuiltinFunctionId::Floor:
      rounded = static_cast<long long>(std::floor(x));
      break;
    case BuiltinFunctionId::Ceil:
      rounded = static_cast<long long>(std::ceil(x));
      break;
    case BuiltinFunctionId::Trunc:
    case BuiltinFunctionId::Int:
      rounded = static_cast<long long>(std::trunc(x));
      break;
    case BuiltinFunctionId::Round:
      rounded = static_cast<long long>(std::round(x));
      break;
    default:
      break;
  }
  return makeScalarInt(rounded);
}

#if SMARTMATH_COMPLEX_NUMBERS
MathParser::EvalValue MathParser::calcRoundingFnCartesian(
    BuiltinFunctionId rid,
    const EvalValue::ScalarValue& sv,
    double ar,
    double ai) {
  EvalValue::ScalarValue svRe = sv;
  svRe.scalar = ar;
  scalarClearImaginary(svRe);
  EvalValue::ScalarValue svIm{};
  svIm.scalar = ai;
  scalarClearImaginary(svIm);
  if (sv.hasImagExactInt64()) {
    svIm.setExactInt64Valid(true);
    svIm.exactInt64 = sv.imagExactInt64;
    if (sv.hasImagExactUInt64()) {
      svIm.setExactUInt64Valid(true);
      svIm.exactUInt64 = sv.imagExactUInt64;
    }
  }
  const EvalValue outRe = calcRoundingFn(rid, svRe);
  const EvalValue outIm = calcRoundingFn(rid, svIm);
  return setScalarComplexFromEvalRealImagParts(outRe, outIm);
}
#endif

bool MathParser::tryUnaryComplexBuiltinSupport(
    BuiltinFunctionId id,
    const EvalValue::ScalarValue& sv,
    EvalValue& out) const {
  const bool isCxComponentUnary =
      id == BuiltinFunctionId::Real || id == BuiltinFunctionId::Imag || id == BuiltinFunctionId::Phase ||
      id == BuiltinFunctionId::Polar || id == BuiltinFunctionId::Cart || id == BuiltinFunctionId::Conj;
#if SMARTMATH_COMPLEX_NUMBERS
  if (!getSupportComplexNumbers() && !isCxComponentUnary) {
    return false;
  }
#else
  if (!isCxComponentUnary) {
    return false;
  }
#endif
  double ar = 0.0;
  double ai = 0.0;
  loadUnaryOperandCartesian(sv, ar, ai);
#if SMARTMATH_COMPLEX_NUMBERS
  if (!getSupportComplexNumbers() && scalarHasNonzeroImaginaryPart(sv)) {
    return false;
  }
#else
  if (scalarHasNonzeroImaginaryPart(sv)) {
    return false;
  }
#endif
  switch (id) {
    case BuiltinFunctionId::Real: {
      ExactCartesianComponent cRe{};
      if (tryExtractExactRealComponent(sv, cRe)) {
        setScalarFromExactCartesianComponent(out, cRe);
      } else {
        out = makeScalarMaybeExact(ar);
        scalarClearImaginary(out.scalarValue);
      }
      return true;
    }
    case BuiltinFunctionId::Imag: {
#if SMARTMATH_COMPLEX_NUMBERS
      if (getSupportComplexNumbers()) {
      // Keep behavior consistent with complex formatting:
      // if the complex value is stored in a "render rational" imag form,
      // `imag(z)` must return that component (not a snapped/rounded float 0).
      if (sv.hasImagRenderRational()) {
        const long long num = sv.imagExactInt64;
        const std::uint64_t den = sv.imagExactUInt64;
        if (den == 0u) {
          out = makeScalar(0.0);
          scalarClearImaginary(out.scalarValue);
          return true;
        }
        if (den == 1u) {
          ExactCartesianComponent cIm{};
          if (num < 0) {
            exactCartesianComponentAssignFromSignedInt64(cIm, num);
          } else {
            exactCartesianComponentAssignFromSignedInt64(cIm, num);
          }
          setScalarFromExactCartesianComponent(out, cIm);
          return true;
        }
        const long double imOut = static_cast<long double>(num) / static_cast<long double>(den);
        out = makeScalar(static_cast<double>(imOut));
        scalarClearImaginary(out.scalarValue);
        return true;
      }

      // Prefer correct "exact imag" when it's nonzero; otherwise, use the
      // floating component loaded from the scalar (used by complex formatting).
      if (sv.hasImagExactInt64() && sv.imagExactInt64 != 0LL) {
        ExactCartesianComponent cIm{};
        exactCartesianComponentAssignFromSignedInt64(cIm, sv.imagExactInt64);
        setScalarFromExactCartesianComponent(out, cIm);
        return true;
      }
      if (sv.hasImagExactUInt64() && sv.imagExactUInt64 != 0u) {
        ExactCartesianComponent cIm{};
        exactCartesianComponentAssignFromUInt64(cIm, sv.imagExactUInt64);
        setScalarFromExactCartesianComponent(out, cIm);
        return true;
      }

      out = makeScalar(ai); // (float) imag component
      scalarClearImaginary(out.scalarValue);
      return true;
      }
#endif
      out = makeScalar(ai);
      scalarClearImaginary(out.scalarValue);
      return true;
    }
    case BuiltinFunctionId::Phase:
      out = makeScalarMaybeExact(calcAtan2Basic(ai, ar));
      return true;
    case BuiltinFunctionId::Polar: {
      std::vector<EvalValue> elems;
      elems.emplace_back(makeScalarMaybeExact(calcHypotBasic(ar, ai)));
      elems.emplace_back(makeScalarMaybeExact(calcAtan2Basic(ai, ar)));
      out = makeArrayFromScalars(elems);
      return true;
    }
    case BuiltinFunctionId::Cart: {
#if SMARTMATH_COMPLEX_NUMBERS
      if (getSupportComplexNumbers()) {
        ExactCartesianComponent cReCart{};
        ExactCartesianComponent cImCart{};
        if (tryExtractExactRealComponent(sv, cReCart) && tryExtractExactImagComponent(sv, cImCart)) {
          setScalarComplexFromExactCartesian(out, cReCart, cImCart);
        } else {
          out = makeScalarComplexFromDoubles(ar, ai);
        }
        return true;
      }
#endif
      out = makeScalarMaybeExact(ar);
      scalarClearImaginary(out.scalarValue);
      return true;
    }
    case BuiltinFunctionId::Conj: {
#if SMARTMATH_COMPLEX_NUMBERS
      if (getSupportComplexNumbers()) {
        ExactCartesianComponent cRe{};
        ExactCartesianComponent cIm{};
        ExactCartesianComponent cImConj{};
        if (tryExtractExactRealComponent(sv, cRe) && tryExtractExactImagComponent(sv, cIm)) {
          if (scalarHasNonzeroImaginaryPart(sv) && !tryNegateExactCartesianComponent(cIm, cImConj)) {
            out = makeScalarComplexFromDoubles(ar, -ai);
          } else if (scalarHasNonzeroImaginaryPart(sv)) {
            setScalarComplexFromExactCartesian(out, cRe, cImConj);
          } else {
            setScalarComplexFromExactCartesian(out, cRe, cIm);
          }
        } else {
          out = makeScalarComplexFromDoubles(ar, -ai);
        }
        return true;
      }
#endif
      out = makeScalarMaybeExact(ar);
      scalarClearImaginary(out.scalarValue);
      return true;
    }
#if SMARTMATH_COMPLEX_NUMBERS
    case BuiltinFunctionId::Int:
    case BuiltinFunctionId::Trunc:
    case BuiltinFunctionId::Floor:
    case BuiltinFunctionId::Ceil:
    case BuiltinFunctionId::Round:
      out = calcRoundingFnCartesian(id, sv, ar, ai);
      return true;
    case BuiltinFunctionId::Frac: {
      EvalValue::ScalarValue svFracRe = sv;
      svFracRe.scalar = ar;
      scalarClearImaginary(svFracRe);
      EvalValue::ScalarValue svFracIm{};
      svFracIm.scalar = ai;
      scalarClearImaginary(svFracIm);
      if (sv.hasImagExactInt64()) {
        svFracIm.setExactInt64Valid(true);
        svFracIm.exactInt64 = sv.imagExactInt64;
        if (sv.hasImagExactUInt64()) {
          svFracIm.setExactUInt64Valid(true);
          svFracIm.exactUInt64 = sv.imagExactUInt64;
        }
      }
      const EvalValue fracIntRe = calcRoundingFn(BuiltinFunctionId::Int, svFracRe);
      const EvalValue fracIntIm = calcRoundingFn(BuiltinFunctionId::Int, svFracIm);
      EvalValue fRe = makeScalar(ar - fracIntRe.scalarValue.scalar);
      EvalValue fIm = makeScalar(ai - fracIntIm.scalarValue.scalar);
      out = setScalarComplexFromEvalRealImagParts(fRe, fIm);
      return true;
    }
    case BuiltinFunctionId::Abs:
      if (!scalarHasNonzeroImaginaryPart(sv)) {
        out = applyAbsScalarValue(sv);
        return true;
      }
      if (!std::isfinite(ar) || !std::isfinite(ai)) {
        out = makeScalar(std::numeric_limits<double>::quiet_NaN());
      } else {
        out = makeScalarMaybeExact(calcHypotBasic(ar, ai));
      }
      return true;
    case BuiltinFunctionId::Sign: {
      const double mag = calcHypotBasic(ar, ai);
      if (mag == 0.0 || !std::isfinite(mag) || !std::isfinite(ar) || !std::isfinite(ai)) {
        out = makeScalarComplexFromDoubles(0.0, 0.0);
      } else {
        out = makeScalarComplexFromDoubles(ar / mag, ai / mag);
      }
      return true;
    }
    case BuiltinFunctionId::Fact: {
      if (ai == 0.0 && !scalarHasNonzeroImaginaryPart(sv)) {
        long long n = 0;
        if (nearlyInt(sv.scalar, n) && n >= 0) {
          if (n <= 20) {
            out = makeScalarInt(kFactorialTable21[n]);
            return true;
          }
          out = makeScalar(factorialScalarFloatFromInt(n));
          return true;
        }
      }
      double gr = 0.0;
      double gi = 0.0;
      complexGamma(ar + 1.0, ai, gr, gi);
      out = makeScalarComplexFromDoubles(gr, gi);
      return true;
    }
#endif
    default:
      return false;
  }
}

MathParser::EvalValue MathParser::builtinPolarCart(
    EvalContext& ctx,
    const std::string& fnName,
    BuiltinFunctionId id,
    const std::vector<EvalValue>& args) const {
  const auto tryNormalizeUnaryScalarInput = [&](const EvalValue& inV, EvalValue& outScalarV) -> bool {
    if (inV.kind == ValueKind::Scalar) {
      outScalarV = inV;
      return true;
    }
    if (inV.kind == ValueKind::Array && inV.arr.size() == 1U) {
      outScalarV = scalarFromScalarValue(inV.arr[0]);
      return true;
    }
    return false;
  };
  const auto cartFromPolarScalars = [&](double rMag, double rAng) -> EvalValue {
#if SMARTMATH_COMPLEX_NUMBERS
    if (!getSupportComplexNumbers() && std::fabs(rAng) > 1e-14) {
      setParseError(ctx, ParseErrorId::IncompatibleOperands);
      return makeScalar(0);
    }
    return makeScalarComplexFromDoubles(rMag * calcCos(rAng), rMag * calcSin(rAng));
#else
    if (std::fabs(rAng) > 1e-14) {
      setParseError(ctx, ParseErrorId::IncompatibleOperands);
      return makeScalar(0);
    }
    return makeScalarMaybeExact(rMag);
#endif
  };

  if (id == BuiltinFunctionId::Polar) {
    EvalValue polarIn = makeScalar(0);
    if (!tryNormalizeUnaryScalarInput(args[0], polarIn)) {
      setParseError(ctx, ParseErrorId::IncompatibleOperands);
      return makeScalar(0);
    }
    EvalValue outS = makeScalar(0);
    if (!tryUnaryComplexBuiltinSupport(BuiltinFunctionId::Polar, polarIn.scalarValue, outS)) {
      setNumericErrorInFunction(ctx, fnName);
      return makeScalar(0);
    }
    return outS;
  }

  if (args.size() == 1U) {
    const EvalValue& cartIn = args[0];
    if (cartIn.kind == ValueKind::Scalar) {
      EvalValue outS = makeScalar(0);
      if (!tryUnaryComplexBuiltinSupport(BuiltinFunctionId::Cart, cartIn.scalarValue, outS)) {
        setNumericErrorInFunction(ctx, fnName);
        return makeScalar(0);
      }
      return outS;
    }
    if (cartIn.kind == ValueKind::Array) {
      const std::size_t nCart = cartIn.arr.size();
      if (nCart == 1U) {
        return cartFromPolarScalars(cartIn.arr[0].scalar, 0.0);
      }
      if (nCart == 2U) {
        return cartFromPolarScalars(cartIn.arr[0].scalar, cartIn.arr[1].scalar);
      }
    }
    setParseError(ctx, ParseErrorId::IncompatibleOperands);
    return makeScalar(0);
  }
  if (args.size() == 2U) {
    if (args[0].kind != ValueKind::Scalar || args[1].kind != ValueKind::Scalar) {
      setParseError(ctx, ParseErrorId::IncompatibleOperands);
      return makeScalar(0);
    }
    if (scalarHasNonzeroImaginaryPart(args[0].scalarValue) ||
        scalarHasNonzeroImaginaryPart(args[1].scalarValue)) {
      setParseError(ctx, ParseErrorId::IncompatibleOperands);
      return makeScalar(0);
    }
    return cartFromPolarScalars(args[0].scalarValue.scalar, args[1].scalarValue.scalar);
  }
  setParseError(ctx, ParseErrorId::IncompatibleOperands);
  return makeScalar(0);
}

#if SMARTMATH_COMPLEX_NUMBERS
bool MathParser::applyUnaryComplexMathOverlay(
    BuiltinFunctionId id,
    const EvalValue::ScalarValue& sv,
    bool& trigOutOfRange,
    bool& realUnaryOk,
    EvalValue& outV) const {
  if (!getSupportComplexNumbers()) {
    return false;
  }
  const double nanv = std::numeric_limits<double>::quiet_NaN();
  const double x = sv.scalar;
  if (scalarHasNonzeroImaginaryPart(sv)) {
    double arCx = 0.0;
    double aiCx = 0.0;
    loadUnaryOperandCartesian(sv, arCx, aiCx);
    if (id == BuiltinFunctionId::Sqrt || id == BuiltinFunctionId::Sqr) {
      if (!std::isfinite(arCx) || !std::isfinite(aiCx)) {
        outV = makeScalarComplexFromDoubles(nanv, nanv);
        return true;
      }
      if (id == BuiltinFunctionId::Sqr) {
        outV = makeScalarComplexFromDoubles(arCx * arCx - aiCx * aiCx, 2.0 * arCx * aiCx);
        return true;
      }
      outV = applyUnarySqrtEval(sv);
      return true;
    }
    if (id == BuiltinFunctionId::Exp) {
      if (!std::isfinite(arCx) || !std::isfinite(aiCx)) {
        outV = makeScalarComplexFromDoubles(nanv, nanv);
        return true;
      }
      if (!isTrigRadiansInRange(aiCx)) {
        trigOutOfRange = true;
        outV = makeScalar(0);
        return true;
      }
      const double ea = std::exp(arCx);
      double pr = ea * calcCos(aiCx);
      double pi = ea * calcSin(aiCx);
      snapComplexNearZeroAxis(pr, pi);
      outV = makeScalarComplexFromDoubles(pr, pi);
      return true;
    }
    if (id == BuiltinFunctionId::Ln || id == BuiltinFunctionId::Log10) {
      double loR = 0.0;
      double loI = 0.0;
      scalarPrincipalLnCartesian(arCx, aiCx, loR, loI);
      if (id == BuiltinFunctionId::Ln) {
        outV = makeScalarComplexFromDoubles(loR, loI);
      } else {
        outV = makeScalarComplexFromDoubles(loR * kInvLn10, loI * kInvLn10);
      }
      return true;
    }
    if (isComplexUnaryTrigBuiltin(id)) {
      double or_ = 0.0;
      double oi = 0.0;
      if (!complexUnaryTrigCartesian(id, arCx, aiCx, or_, oi)) {
        trigOutOfRange = true;
        outV = makeScalar(0);
        return true;
      }
      snapComplexNearZeroAxis(or_, oi);
      outV = makeScalarComplexFromDoubles(or_, oi);
      return true;
    }
    realUnaryOk = false;
    outV = makeScalar(0);
    return true;
  }
  if ((id == BuiltinFunctionId::Ln || id == BuiltinFunctionId::Log10) && std::isfinite(x)) {
    if (x < 0.0) {
      if (id == BuiltinFunctionId::Ln) {
        double loR = std::log(-x);
        double loI = kPi;
        snapComplexNearZeroAxis(loR, loI);
        outV = makeScalarComplexFromDoubles(loR, loI);
      } else {
        double loR = std::log10(-x);
        double loI = kPi * kInvLn10;
        snapComplexNearZeroAxis(loR, loI);
        outV = makeScalarComplexFromDoubles(loR, loI);
      }
      return true;
    }
    if (x == 0.0) {
      double loR = 0.0;
      double loI = 0.0;
      scalarPrincipalLnCartesian(0.0, 0.0, loR, loI);
      if (id == BuiltinFunctionId::Ln) {
        outV = makeScalarComplexFromDoubles(loR, loI);
      } else {
        outV = makeScalarComplexFromDoubles(loR * kInvLn10, loI * kInvLn10);
      }
      return true;
    }
  }
  return false;
}
#endif

MathParser::EvalValue MathParser::applyUnaryBuiltin(
    BuiltinFunctionId id,
    const EvalValue::ScalarValue& sv,
    bool& trigOutOfRange,
    bool& realUnaryOk) const {
  EvalValue cxOut = makeScalar(0);
  if (tryUnaryComplexBuiltinSupport(id, sv, cxOut)) {
    return cxOut;
  }

  const double x = sv.scalar;
#if SMARTMATH_COMPLEX_NUMBERS
  EvalValue overlayOut = makeScalar(0);
  if (applyUnaryComplexMathOverlay(id, sv, trigOutOfRange, realUnaryOk, overlayOut)) {
    return overlayOut;
  }
#endif

  switch (kUnaryScalarKind[static_cast<std::size_t>(id)]) {
    case UnaryScalarKind::PrimaryTrig:
      if (!isTrigRadiansInRange(x)) {
        trigOutOfRange = true;
        return makeScalar(0);
      }
      if (id == BuiltinFunctionId::Sin) {
        return makeScalar(calcSin(x));
      }
      if (id == BuiltinFunctionId::Cos) {
        return makeScalar(calcCos(x));
      }
      return makeScalar(calcTan(x));
    case UnaryScalarKind::InverseTrig:
      if (id == BuiltinFunctionId::Asin) {
        return makeScalar(std::asin(x));
      }
      if (id == BuiltinFunctionId::Acos) {
        return makeScalar(std::acos(x));
      }
      return makeScalar(std::atan(x));
    case UnaryScalarKind::HyperbolicTrig:
      if (id == BuiltinFunctionId::Sinh) {
        return makeScalar(std::sinh(x));
      }
      if (id == BuiltinFunctionId::Cosh) {
        return makeScalar(std::cosh(x));
      }
      return makeScalar(std::tanh(x));
    case UnaryScalarKind::InverseHyperbolicTrig:
      if (id == BuiltinFunctionId::Acosh) {
        return makeScalar(std::acosh(x));
      }
      if (id == BuiltinFunctionId::Asinh) {
        return makeScalar(std::asinh(x));
      }
      return makeScalar(std::atanh(x));
    case UnaryScalarKind::Exp:
      return makeScalarMaybeExact(std::exp(x));
    case UnaryScalarKind::Ln:
      return makeScalarMaybeExact(std::log(x));
    case UnaryScalarKind::Log10:
      return makeScalarMaybeExact(std::log10(x));
    case UnaryScalarKind::Sqrt:
      return applyUnarySqrtEval(sv);
    case UnaryScalarKind::Sqr: {
      EvalValue sqrOut;
      if (tryApplySqrExactScalar(sv, sqrOut)) {
        return sqrOut;
      }
      return makeScalar(x * x);
    }
    case UnaryScalarKind::Rounding:
      return calcRoundingFn(id, sv);
    case UnaryScalarKind::Frac:
      return makeScalarMaybeExact(x - std::trunc(x));
    case UnaryScalarKind::Abs:
      return applyAbsScalarValue(sv);
    case UnaryScalarKind::Sign:
      if (sv.hasExactInt64()) {
        return makeScalarInt((sv.exactInt64 > 0) ? 1LL : ((sv.exactInt64 < 0) ? -1LL : 0LL));
      }
      if (sv.hasExactUInt64()) {
        return makeScalarInt((sv.exactUInt64 == 0u) ? 0LL : 1LL);
      }
      return makeScalarInt((x > 0.0) ? 1LL : ((x < 0.0) ? -1LL : 0LL));
    case UnaryScalarKind::Deg:
      return makeScalarMaybeExact(x * 180.0 / kPi);
    case UnaryScalarKind::Rad:
      return makeScalarMaybeExact(x * kPi / 180.0);
    case UnaryScalarKind::None:
    default:
      realUnaryOk = false;
      break;
  }
  return makeScalar(0);
}

MathParser::EvalValue MathParser::builtinUnaryMath(
    EvalContext& ctx,
    const std::string& fnName,
    BuiltinFunctionId id,
    const std::vector<EvalValue>& args) const {
#if SMARTMATH_TIME_VALUES
  if (evalValueInvolvesTime(args[0])) {
    setParseError(ctx, ParseErrorId::IncompatibleOperands);
    return makeScalar(0);
  }
#endif
  bool realUnaryOk = true;
  bool trigOutOfRange = false;
  const EvalValue result = mapUnaryEvalValue(args[0], [&](const EvalValue::ScalarValue& s) -> EvalValue {
    return applyUnaryBuiltin(id, s, trigOutOfRange, realUnaryOk);
  });
  if (trigOutOfRange) {
    setNumericErrorInFunction(ctx, fnName);
    return makeScalar(0);
  }
  if (!realUnaryOk) {
    setParseError(ctx, ParseErrorId::InternalUnaryMathBuiltin);
    return makeScalar(0);
  }
  return result;
}

MathParser::EvalValue::ScalarValue MathParser::clampScalarValue(
    const EvalValue::ScalarValue& s,
    const EvalValue::ScalarValue& minSv,
    const EvalValue::ScalarValue& maxSv) {
  const double v = scalarNumericReal(s);
  const double minS = minSv.scalar;
  const double maxS = maxSv.scalar;

  if (std::isnan(v)) {
    if (!std::isnan(minS)) {
      return minSv;
    }
    return s;
  }

  if (!std::isnan(minS) && v <= minS) {
    return minSv;
  }

  if (!std::isnan(maxS) && v > maxS) {
    return maxSv;
  }

  if (std::isnan(maxS) && !std::isnan(minS) && v > minS) {
    return makeScalar(std::numeric_limits<double>::quiet_NaN()).scalarValue;
  }

  return s;
}

MathParser::EvalValue MathParser::builtinApplyClamp(
    EvalContext& ctx,
    const EvalValue& valueV,
    const EvalValue& minV,
    const EvalValue& maxV) const {
  if (minV.kind != ValueKind::Scalar || maxV.kind != ValueKind::Scalar) {
    setScalarMinMaxError(ctx, getFunctionName(BuiltinFunctionId::Clamp));
    return makeScalar(0);
  }
#if SMARTMATH_TIME_VALUES
  if (evalValueInvolvesTime(valueV) || scalarValueIsTime(minV.scalarValue) || scalarValueIsTime(maxV.scalarValue)) {
    setParseError(ctx, ParseErrorId::IncompatibleOperands);
    return makeScalar(0);
  }
#endif
#if SMARTMATH_COMPLEX_NUMBERS
  if (getSupportComplexNumbers() &&
      (MathParser::evalValueHasNonzeroImaginary(valueV) || MathParser::evalValueHasNonzeroImaginary(minV) ||
          MathParser::evalValueHasNonzeroImaginary(maxV))) {
    setParseError(ctx, ParseErrorId::IncompatibleOperands);
    return makeScalar(0);
  }
#endif
  if (valueV.kind != ValueKind::Scalar && valueV.kind != ValueKind::Array) {
    setNumericErrorInFunction(ctx, getFunctionName(BuiltinFunctionId::Clamp));
    return makeScalar(0);
  }
  const EvalValue::ScalarValue& minSv = minV.scalarValue;
  const EvalValue::ScalarValue& maxSv = maxV.scalarValue;
  return mapUnaryEvalValue(valueV, [&minSv, &maxSv](const EvalValue::ScalarValue& s) -> EvalValue {
    return scalarFromScalarValue(clampScalarValue(s, minSv, maxSv));
  });
}

MathParser::EvalValue MathParser::builtinLog(EvalContext& ctx, const std::vector<EvalValue>& args) const {
  return builtinMapBinaryTwoArg(
      ctx, getFunctionName(BuiltinFunctionId::Log), BuiltinFunctionId::Log, args, false, true);
}

MathParser::EvalValue MathParser::evalUserFunctionCall(
    EvalContext& ctx,
    const std::string& fnName,
    const std::vector<EvalValue>& args,
    const std::unordered_map<std::string, EvalValue>* /*scopedVars*/) {
  UserFunction* uf = findUserFunction(fnName);
  if (!uf) {
    appendUniqueName(ctx.unknownFuncsText, fnName);
    return makeScalar(0);
  }
  if (args.size() != uf->params.size()) {
    setExactArgCountError(ctx, fnName, uf->params.size());
    return makeScalar(0);
  }
  if (ctx.evalDepth >= kMaxEvalDepth) {
    setParseError(ctx, ParseErrorId::MaxEvaluationDepthReached);
    return makeScalar(0);
  }

  std::vector<std::string> localUserFunctionCallStack;
  std::vector<std::string>* const callStack =
      ctx.userFunctionCallStack ? ctx.userFunctionCallStack : &localUserFunctionCallStack;

  for (const std::string& active : *callStack) {
    if (active == fnName) {
      setRecursiveUserFunctionCallError(ctx, fnName);
      return makeScalar(0);
    }
  }
  if (callStack->size() >= static_cast<std::size_t>(kMaxEvalDepth)) {
    setParseError(ctx, ParseErrorId::UserFunctionCallStackOverflow);
    return makeScalar(0);
  }
  callStack->push_back(fnName);
  struct UserFnCallPopGuard {
    std::vector<std::string>& stack;
    explicit UserFnCallPopGuard(std::vector<std::string>& s) : stack(s) {}
    ~UserFnCallPopGuard() { stack.pop_back(); }
  } popGuard{*callStack};

  std::unordered_map<std::string, EvalValue> localVars;
  localVars.reserve(uf->params.size());
  for (std::size_t i = 0; i < uf->params.size(); ++i) {
    localVars.emplace(uf->params[i], args[i]);
  }

  if (!uf->compiledProgramReady) {
    std::vector<AstStatement> compiledBody;
    if (!tryCompileExprToProgram(uf->expr, ctx, callStack, fnName.c_str(), false, compiledBody)) {
      return makeScalar(0);
    }
    uf->compiledProgram = std::move(compiledBody);
    uf->compiledProgramReady = true;
  }

  EvalContext sub;
  sub.evalDepth = ctx.evalDepth + 1;
  sub.userFunctionCallStack = ctx.userFunctionCallStack;
  EvalValue v = runCompiledProgram(sub, uf->compiledProgram, &localVars, false);
  if (!finishSubEvalFromCompiledProgram(ctx, sub)) {
    return makeScalar(0);
  }
  return v;
}

MathParser::EvalValue MathParser::evalFunctionCall(
    EvalContext& ctx,
    const std::string& fnName,
    std::vector<EvalValue>&& args,
    BuiltinFunctionId preboundId,
    const std::unordered_map<std::string, EvalValue>* scopedVars) {
  normalizeCallArgs(args);

  BuiltinFunctionId id = preboundId;
  if (id == BuiltinFunctionId::Count && !tryGetBuiltinFunctionId(fnName, id)) {
    return evalUserFunctionCall(ctx, fnName, args, scopedVars);
  }
  if (!validateBuiltinArgs(ctx, fnName, id, args)) {
    return makeScalar(0);
  }
  if (!validateBuiltinCallArity(ctx, fnName, id, args)) {
    return makeScalar(0);
  }

  switch (getBuiltinCategory(id)) {
    case BuiltinCategory::Sortby:
      return builtinSortby(ctx, fnName, args, scopedVars);
    case BuiltinCategory::Ratio:
      return builtinRatio(ctx, fnName, args);
    case BuiltinCategory::TimeUnit:
      return builtinTimeUnit(ctx, fnName, id, args);
    case BuiltinCategory::ArrayTransform:
      if (id == BuiltinFunctionId::Unpack) {
        return builtinUnpack(ctx, args);
      }
      if (id == BuiltinFunctionId::Len) {
        return builtinLen(ctx, args);
      }
      return builtinSortFamily(ctx, fnName, id, args);
    case BuiltinCategory::Repeat:
      return builtinRepeat(ctx, fnName, args);
    case BuiltinCategory::Range:
      return builtinRange(ctx, fnName, args);
    case BuiltinCategory::Aggregate:
      return builtinAggregateFamily(ctx, fnName, id, args);
    case BuiltinCategory::BaseFormat:
      return builtinBaseFormat(ctx, fnName, id, args);
    case BuiltinCategory::Pow:
      return builtinPow(ctx, args);
    case BuiltinCategory::ScalarBinary:
      return builtinScalarBinaryFamily(ctx, fnName, id, args);
    case BuiltinCategory::Rand:
      return builtinRand(ctx, args);
    case BuiltinCategory::Mod:
      return builtinModCall(ctx, args);
    case BuiltinCategory::Factorial:
      return builtinFactorial(ctx, fnName, args);
    case BuiltinCategory::Factorint:
#if SMARTMATH_FACTORINT
      return builtinFactorint(ctx, fnName, args);
#else
      setParseError(ctx, ParseErrorId::IncompatibleOperands);
      return makeScalar(0);
#endif
    case BuiltinCategory::DegRad:
      return builtinDegRad(ctx, fnName, id, args);
    case BuiltinCategory::Log:
      return builtinLog(ctx, args);
    case BuiltinCategory::UnaryMath:
      return builtinUnaryMath(ctx, fnName, id, args);
    case BuiltinCategory::PolarCart:
      return builtinPolarCart(ctx, fnName, id, args);
    case BuiltinCategory::Unknown:
      break;
  }
  appendUniqueName(ctx.unknownFuncsText, fnName);
  return makeScalar(0);
}

bool MathParser::prepareCompileParseSource(
    const std::string& mathExpression,
    EvalContext& ctx,
    std::string& ownedParseBuffer,
    bool& strippedTrailingSemicolon) {
  ownedParseBuffer.clear();
  ctx.sourceExpr.clear();
  strippedTrailingSemicolon = false;
  const char* text = mathExpression.data();
  const std::size_t len = mathExpression.size();
  const bool hadLineComment = parseSourceHasLineComment(text, len);
  if (hadLineComment || parseSourceNeedsTrailingSemicolonStrip(text, len)) {
    if (hadLineComment) {
      ownedParseBuffer = stripLineComments(mathExpression);
    } else {
      ownedParseBuffer = mathExpression;
    }
    strippedTrailingSemicolon = stripTrailingSemicolonsForTopLevelInput(ownedParseBuffer);
    if (ownedParseBuffer.empty()) {
      return false;
    }
    ctx.sourceExpr = ownedParseBuffer;
    ctx.p = ownedParseBuffer.c_str();
  } else {
    if (len == 0) {
      return false;
    }
    ctx.p = mathExpression.c_str();
  }
  ctx.start = ctx.p;
  return true;
}

bool MathParser::compile(const std::string& mathExpression) {
  resetCompileState();

  EvalContext ctx;
  std::string ownedParseBuffer;
  bool strippedTrailingSemicolon = false;
  if (!prepareCompileParseSource(mathExpression, ctx, ownedParseBuffer, strippedTrailingSemicolon)) {
    if (parseSourceHasLineComment(mathExpression.data(), mathExpression.size())) {
      compiled_.program.clear();
      compiled_.hasAssignments = false;
      compiled_.scalarOnly = true;
      compiled_.ready = true;
      compiled_.boundVariablesVersion = variablesVersion_;
      return true;
    }
    if (!mathExpression.empty()) {
      lastError_ = STR_EMPTY_STATEMENT;
    }
    return false;
  }
  // Path C only: parseProgram leaves ';' in ctx.p (Path B); suppress applies after strip above.
  ctx.suppressBareFunctionTailHint = strippedTrailingSemicolon;

  std::vector<AstStatement> program;
  if (!tryCompileSingleExpressionProgram(ctx, program) && !parseProgram(ctx, program)) {
    if (!ctx.parseError) {
      setParseError(ctx, ParseErrorId::ParseFailed);
    }
    lastError_ = ctx.errorText;
    return false;
  }
  skipSpaces(ctx);
  if (!ctx.parseError && *ctx.p != '\0') {
    setParseError(ctx, ParseErrorId::UnexpectedInput);
  }
  if (ctx.parseError) {
    lastError_ = ctx.errorText;
    compiled_.program.clear();
    return false;
  }
  compiled_.program = std::move(program);
  compiled_.hasAssignments = std::any_of(
      compiled_.program.begin(),
      compiled_.program.end(),
      [](const AstStatement& st) { return st.kind == AstStatement::Kind::Assign; });
  bindCompiledVariableRefs(compiled_);
  compiled_.scalarOnly = programIsScalarOnly(compiled_.program);
  compiled_.ready = true;
  return true;
}

void MathParser::evaluate() {
  resetEvaluateState();

  EvalContext ctx;
  std::vector<std::string> userFunctionCallStack;
  ctx.userFunctionCallStack = &userFunctionCallStack;
  if (!prepareEvaluate(ctx)) {
    return;
  }
  if (compiled_.program.empty()) {
    return;
  }
  EvalValue out = runCompiledProgram(ctx, compiled_.program, nullptr, compiled_.scalarOnly);
  finalizeEvaluate(ctx, std::move(out));
}

void MathParser::resetCompileState() {
  hasResult_ = false;
  lastError_.clear();
  compiled_.program.clear();
  compiled_.ready = false;
  compiled_.scalarOnly = false;
  compiled_.hasAssignments = false;
  compiled_.boundVariablesVersion = static_cast<std::size_t>(-1);
}

void MathParser::resetEvaluateState() {
  hasResult_ = false;
  lastError_.clear();
  lastResult_ = makeScalar(0);
}

bool MathParser::prepareEvaluate(EvalContext& ctx) {
  ctx.start = nullptr;
  ctx.p = nullptr;
  if (!compiled_.ready) {
    lastError_ = STR_NOTHING_COMPILED_SEMICOLON_CALL_COMPILE_PAR;
    return false;
  }
  if (compiled_.boundVariablesVersion != variablesVersion_) {
    bindCompiledVariableRefs(compiled_);
  }
  return true;
}

bool MathParser::finalizeEvaluate(EvalContext& ctx, EvalValue&& out) {
  if (ctx.parseError) {
    lastError_ = ctx.errorText;
    return false;
  }
  if (trySetUnknownNameError(ctx)) {
    return false;
  }
  lastResult_ = std::move(out);
  hasResult_ = true;
  return true;
}

void MathParser::parseAndEvaluate(const std::string& mathExpression) {
  if (!compile(mathExpression)) {
    return;
  }
  evaluate();
}

MathParser::RawResult MathParser::parseAndEvaluateRaw(const std::string& mathExpression) {
  if (!compile(mathExpression)) {
    return RawResult{};
  }
  evaluate();
  if (!lastError_.empty() || !hasResult_) {
    return RawResult{};
  }
  return toRawResult(lastResult_);
}

std::string MathParser::addUserFunction(const std::string& mathExpression) {
  EvalContext ctx;
  ctx.sourceExpr = stripLineComments(mathExpression);
  const bool nonemptyAfterLineComment = !ctx.sourceExpr.empty();
  stripTrailingSemicolonsForTopLevelInput(ctx.sourceExpr);
  if (ctx.sourceExpr.empty()) {
    if (nonemptyAfterLineComment) {
      return STR_EMPTY_STATEMENT;
    }
  }
  ctx.p = ctx.sourceExpr.c_str();
  ctx.start = ctx.p;

  std::string fnName;
  std::vector<std::string> fnParams;
  std::string fnExpr;
  if (!parseFunctionDefinition(ctx, fnName, fnParams, fnExpr)) {
    return STR_INVALID_USER_FUNCTION_EXPRESSION;
  }

  skipSpaces(ctx);
  if (*ctx.p != '\0') {
    return STR_UNEXPECTED_CONTENT_AFTER_FUNCTION_DEFINITION;
  }

  const std::string udfErr = getUserFunctionDefinitionErrorText(fnName, fnParams, fnExpr, true);
  if (!udfErr.empty()) return udfErr;

  upsertUserFunction(UserFunction{fnName, fnParams, fnExpr});
  return "";
}

MathParser::UserFunction* MathParser::findUserFunction(const std::string& fnName) {
  auto it = userFunctionIndex_.find(fnName);
  if (it == userFunctionIndex_.end()) {
    return nullptr;
  }
  return &userFunctions_[it->second];
}

const MathParser::UserFunction* MathParser::findUserFunction(const std::string& fnName) const {
  auto it = userFunctionIndex_.find(fnName);
  if (it == userFunctionIndex_.end()) {
    return nullptr;
  }
  return &userFunctions_[it->second];
}

void MathParser::upsertUserFunction(UserFunction uf) {
  removeVariableByName(uf.name);
  uf.compiledProgram.clear();
  uf.compiledProgramReady = false;
  auto it = userFunctionIndex_.find(uf.name);
  if (it != userFunctionIndex_.end()) {
    userFunctions_[it->second] = std::move(uf);
    return;
  }
  userFunctions_.push_back(std::move(uf));
  userFunctionIndex_[userFunctions_.back().name] = userFunctions_.size() - 1;
}
