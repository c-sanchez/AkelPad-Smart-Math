#include "MathParser.hpp"
#include "MathParserInternal.hpp"

using smartmath_internal::gcdUInt;
using smartmath_internal::isWithinExactIntFromFloatRange;
using smartmath_internal::tryExtractExactSIntFromFloatStrict;
using smartmath_internal::makeArrayFromScalars;
using smartmath_internal::makeInlineLambdaValue;
using smartmath_internal::makeScalar;
using smartmath_internal::makeScalarInt;
using smartmath_internal::makeScalarMaybeExact;
using smartmath_internal::makeScalarUInt;
using smartmath_internal::makeScalarUIntPreferInt;
using smartmath_internal::scalarClearImaginary;
using smartmath_internal::tryApplyRealScalarPowNegFractional;
using smartmath_internal::tryApplyScalarPowSpecialPaths;
using smartmath_internal::calcRoundingFn;
using smartmath_internal::nearlyInt;
using smartmath_internal::parseUIntFromFloat;
using smartmath_internal::tryGetExactSignedIntFromScalarPolicy;
using smartmath_internal::tryGetExactSignedIntFromScalar;
using smartmath_internal::tryGetExactSignedIntNoUIntWrapFromScalar;
using smartmath_internal::scalarHasExactIntegerPayload;
using smartmath_internal::TryMulExactSIntSquare;
using smartmath_internal::TryGetExactSignedIntNoUIntWrapScalarStrict;
using smartmath_internal::ArrayIndexIntKind;
using smartmath_internal::classifyArrayIndexInt;
using smartmath_internal::tryGetArrayIndexIntFromScalarValue;
using smartmath_internal::normalizePythonSliceBounds;
using smartmath_internal::estimateSliceItemCount;
using smartmath_internal::tryResolveArrayIndex;
using smartmath_internal::tryGetExactNonNegativeUIntFromScalar;
using smartmath_internal::tryGetSignedIntFromScalar;
using smartmath_internal::tryGetExactSignMagnitudeFromScalar;
using smartmath_internal::makeScalarFromSignMagnitude;
using smartmath_internal::uintMagnitudeOfSInt;
using smartmath_internal::sintFromNegatedMagnitude;
using smartmath_internal::tryGetBothExactSignedIntNoUIntWrapFromScalars;
using smartmath_internal::tryGetBothExactNonNegativeUIntFromScalars;
using smartmath_internal::tryShiftLeftU64ExactOrMaybe;
using smartmath_internal::TryQuotExactSInt;
using smartmath_internal::tryMulUIntChecked;
using smartmath_internal::tryPowUIntChecked;
using smartmath_internal::passExactAbsFloatFactorGates;
using smartmath_internal::getExactRoundedIntFromFloatResult;
using smartmath_internal::verifyExactIntFloatOpResidue;
using smartmath_internal::verifyExactIntFloatOpCrossMultiply;
using smartmath_internal::tryPromoteExactIntFromFloatOp;
using smartmath_internal::tryPromoteExactDivisionByFloatDivisorSigned;
using smartmath_internal::tryPromoteExactDivisionByFloatDivisorUnsigned;
using smartmath_internal::tryPromoteExactMultiplicationByFloatFactorSigned;
using smartmath_internal::tryPromoteExactMultiplicationByFloatFactorUnsigned;
using smartmath_internal::tryApplyExactIntFloatOpFromFloatSide;
using smartmath_internal::scalarApplyReducedRationalPart;
using smartmath_internal::applySqrtScalarValue;
using smartmath_internal::tryApplySqrExactScalar;
using smartmath_internal::tryApplyHypotExactScalars;
using smartmath_internal::tryApplyPowExactScalars;
using smartmath_internal::tryApplyPowExactScalarsByKind;
using smartmath_internal::makeRationalReduced;
using smartmath_internal::negateEvalValue;
using smartmath_internal::applyAbsScalarValue;
using smartmath_internal::exactCartesianComponentClear;
using smartmath_internal::ExactCartesianComponentAssignFromSignedInt;
using smartmath_internal::ExactCartesianComponentAssignFromUInt;
using smartmath_internal::tryExtractExactRealComponent;
using smartmath_internal::setScalarFromExactCartesianComponent;
using smartmath_internal::expandUnpackedArgs;
using smartmath_internal::toRawResult;
using smartmath_internal::valueToString;

#if SMARTMATH_TIME_VALUES
using smartmath_internal::makeScalarTimeMs;
using smartmath_internal::timeTotalMsFromScalarValue;
#endif

#if SMARTMATH_COMPLEX_NUMBERS
using smartmath_internal::scalarClearCartesianRenderExact;
using smartmath_internal::tryGetExactImagSIntStrict;
using smartmath_internal::ScalarApplyExactSIntPart;
using smartmath_internal::TryExactCartesianComponentToSInt;
using smartmath_internal::tryExtractExactCartesianComponent;
using smartmath_internal::tryNegateExactCartesianComponent;
using smartmath_internal::setPureImaginaryFromMagnitudeScalar;
using smartmath_internal::makeImaginaryUnit;
using smartmath_internal::makeScalarComplexFromFloats;
using smartmath_internal::tryRefinePowPrincipalToExactScalarResult;
using smartmath_internal::tryVerifyComplexCartesianSquareExact;
using smartmath_internal::tryRefineSqrtPrincipalToExactComplex;
using smartmath_internal::applySqrtComplexPrincipalUnary;
using smartmath_internal::complexNeedsPrincipalNegRealPow;
using smartmath_internal::setScalarComplexFromEvalRealImagParts;
using smartmath_internal::tryApplyExactComplexCartesianBinary;
using smartmath_internal::tryFoldExactComplexCartesian;
using smartmath_internal::tryAvgExactComplexFromSum;
using smartmath_internal::complexUnaryTrigCartesian;
using smartmath_internal::scalarSplitCartesianParts;
using smartmath_internal::calcRoundingFnCartesian;
using smartmath_internal::complexCartesianBinary;
using smartmath_internal::tryExtractExactImagComponent;
using smartmath_internal::tryApplyExactCartesianBinaryOp;
using smartmath_internal::scalarApplyExactImagFromCartesianComponent;
using smartmath_internal::setScalarComplexFromExactCartesian;
using smartmath_internal::tryNegateExactComplexScalar;
#endif

#include <algorithm>
#include <cassert>
#include <cmath>
#include <cctype>
#include <cstdio>
#include <cstring>
#include <cstdlib>
#include <limits>
#include <numeric>
#include <unordered_set>
#include <utility>

#define BF smartmath_internal::BuiltinFlags
#define BHK smartmath_internal::BuiltinHintKind
#define BC smartmath_internal::BuiltinCategory

// Numeric type configuration (cpp/MathParserTypes.hpp): every width-specific type and limit comes from here.
using smartmath::SmFloat;
using smartmath::SmInt;
using smartmath::SmUInt;
using smartmath::floatTolerance;
using smartmath::isFloatInSIntRange;
using smartmath::isFloatInUIntRange;
using smartmath::kFloatDisplayDigits;
using smartmath::kFloatSIntBound;
using smartmath::kFloatUIntBound;
using smartmath::kHugeSentinel;
using smartmath::kHypotDirectMax;
using smartmath::kHypotDirectMin;
using smartmath::kIntBits;
using smartmath::kIntMax;
using smartmath::kIntMin;
using smartmath::kIntMinMag;
using smartmath::kMaxExactIntFromFloat;
using smartmath::kMaxFactorialArg;
using smartmath::kMaxPow10ExponentOfInt;
using smartmath::kMaxShift;
using smartmath::kPow10UIntCount;
using smartmath::kRadixDigitBufLen;
using smartmath::kTolLost0;
using smartmath::kTolLost1;
using smartmath::kTolLost2;
using smartmath::kTolLost3;
using smartmath::kTolLost6;
using smartmath::kUIntDigits10;
using smartmath::kUIntMax;
using smartmath::factorialInt;
using smartmath::pow10UInt;

namespace {

constexpr SmFloat K_UNIT_ROOT_EXPONENT_TOL = kTolLost0;
constexpr SmFloat K_HYPOT_DIRECT_MAX = kHypotDirectMax;
constexpr SmFloat K_HYPOT_DIRECT_MIN = kHypotDirectMin;
constexpr SmFloat kPi = static_cast<SmFloat>(3.1415926535897932384626433832795L);
constexpr SmFloat kInvLn10 = static_cast<SmFloat>(0.43429448190325182046017448968723710L);  // 1 / ln(10)
constexpr SmFloat kTrigMaxAbsRadians = kMaxExactIntFromFloat;  // 2^mantissa (2^53 for binary64)
constexpr int kMaxEvalDepth = 128;
constexpr std::uint8_t kBuiltinMetaArityUnset = 254;
constexpr std::uint8_t kBuiltinArityUnbounded = 255;

using smartmath_internal::BuiltinFunctionId;
using smartmath_internal::UnaryScalarKind;

const smartmath_internal::BuiltinMetaRow kBuiltinMeta[] = {
    { BF::NonCalculating, 0, 0, BHK::EmptyPar, BC::Rand },  // Rand
    { BF::FiniteRequired, 2, 2, BHK::MinMax, BC::ScalarBinary },  // Random
    { BF::Format | BF::NonCalculating | BF::TrailingFormatter, 1, kBuiltinArityUnbounded, BHK::DotDotDot, BC::BaseFormat },  // Bin
    { BF::Format | BF::NonCalculating | BF::TrailingFormatter, 1, kBuiltinArityUnbounded, BHK::DotDotDot, BC::BaseFormat },  // Hex
    { BF::Format | BF::NonCalculating | BF::TrailingFormatter, 1, kBuiltinArityUnbounded, BHK::DotDotDot, BC::BaseFormat },  // Oct
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
    { BF::TrailingFormatter, 1, kBuiltinArityUnbounded, BHK::DotDotDot, BC::DegRad },  // Deg
    { BF::TrailingFormatter, 1, kBuiltinArityUnbounded, BHK::DotDotDot, BC::DegRad },  // Rad
    { BF::None, 1, kBuiltinArityUnbounded, BHK::DotDotDot, BC::Aggregate },  // Sum
    { BF::None, 1, kBuiltinArityUnbounded, BHK::DotDotDot, BC::Aggregate },  // Median
    { BF::None, 1, kBuiltinArityUnbounded, BHK::DotDotDot, BC::Aggregate },  // Variance
    { BF::None, 1, kBuiltinArityUnbounded, BHK::DotDotDot, BC::Aggregate },  // Stddev
    { BF::NonCalculating, 1, kBuiltinArityUnbounded, BHK::DotDotDot, BC::ArrayTransform },  // Sort
    { BF::NonCalculating, 2, 2, BHK::ArrayFunc, BC::Sortby },  // Sortby
    { BF::Unary, 1, 1, BHK::Value, BC::Ratio },  // Ratio
    { BF::NonCalculating, 2, 3, BHK::StartStopStep, BC::Range },  // Range
    { BF::NonCalculating, 1, kBuiltinArityUnbounded, BHK::DotDotDot, BC::ArrayTransform },  // Reverse
    { BF::NonCalculating, 1, kBuiltinArityUnbounded, BHK::DotDotDot, BC::ArrayTransform },  // Unique
    { BF::NonCalculating, 1, kBuiltinArityUnbounded, BHK::DotDotDot, BC::ArrayTransform },  // Unpack
    { BF::NonCalculating, 0, kBuiltinArityUnbounded, BHK::DotDotDot, BC::ArrayTransform },  // Flat
    { BF::NonCalculating, 2, 2, BHK::XN, BC::Repeat },  // Repeat
    { BF::NonCalculating, 0, kBuiltinArityUnbounded, BHK::X, BC::ArrayTransform },  // Len
    { BF::Unary | BF::IntegerOnly, 1, 1, BHK::N, BC::Factorial },  // Fact
#if SMARTMATH_FACTORINT
    { BF::Unary | BF::IntegerOnly, 1, 1, BHK::N, BC::Factorint },  // Factorint
#endif
    { BF::None, 1, kBuiltinArityUnbounded, BHK::DotDotDot, BC::Aggregate },  // Avg
    { BF::None, 1, kBuiltinArityUnbounded, BHK::DotDotDot, BC::Aggregate },  // Mean
    { BF::IntegerOnly, 2, 2, BHK::ValueDivisor, BC::Mod },  // Mod
    { BF::None, 3, 3, BHK::ValueMinMax, BC::ScalarBinary },  // Clamp
    { BF::None, 2, 2, BHK::XY, BC::ScalarBinary },  // Hypot
    { BF::IntegerOnly, 2, 2, BHK::AB, BC::ScalarBinary },  // Gcd
    { BF::IntegerOnly, 2, 2, BHK::AB, BC::ScalarBinary },  // Lcm
    { BF::IntegerOnly, 2, 2, BHK::AB, BC::ScalarBinary },  // Ncr
    { BF::IntegerOnly, 2, 2, BHK::AB, BC::ScalarBinary },  // Npr
    { BF::None, 1, kBuiltinArityUnbounded, BHK::DotDotDot, BC::Aggregate },  // Product
    { BF::None, 1, kBuiltinArityUnbounded, BHK::DotDotDot, BC::Aggregate },  // Min
    { BF::None, 1, kBuiltinArityUnbounded, BHK::DotDotDot, BC::Aggregate },  // Max
    { BF::Format | BF::NonCalculating | BF::TrailingFormatter, 1, kBuiltinArityUnbounded, BHK::DotDotDot, BC::BaseFormat },  // Uhex
    { BF::Format | BF::NonCalculating | BF::TrailingFormatter, 1, kBuiltinArityUnbounded, BHK::DotDotDot, BC::BaseFormat },  // Uoct
    { BF::Format | BF::NonCalculating | BF::TrailingFormatter, 1, kBuiltinArityUnbounded, BHK::DotDotDot, BC::BaseFormat },  // Ubin
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
    { BF::NonCalculating, 2, 2, BHK::ArrayFunc, BC::Filter },  // Filter
    { BF::None, 3, 3, BHK::Iif, BC::Iif },  // Iif
    { BF::None, 2, 2, BHK::IsInstance, BC::IsInstance },  // IsInstance
    { BF::Format | BF::NonCalculating | BF::TrailingFormatter, 1, kBuiltinArityUnbounded, BHK::DotDotDot, BC::BaseFormat },  // Dec
    { BF::Format | BF::NonCalculating | BF::TrailingFormatter, 1, kBuiltinArityUnbounded, BHK::DotDotDot, BC::BaseFormat },  // Udec
};
#undef BC
#undef BHK
#undef BF

static_assert(
  sizeof(kBuiltinMeta) / sizeof(kBuiltinMeta[0]) ==
    static_cast<std::size_t>(BuiltinFunctionId::Count),
  "kBuiltinMeta size mismatch");

constexpr UnaryScalarKind kUnaryScalarKind[static_cast<std::size_t>(BuiltinFunctionId::Count)] = {
  /* Rand..Atan2 */ UnaryScalarKind::None, UnaryScalarKind::None, UnaryScalarKind::None, UnaryScalarKind::None,
  UnaryScalarKind::None, UnaryScalarKind::None, UnaryScalarKind::None,
  /* Sin..Tan */ UnaryScalarKind::PrimaryTrig, UnaryScalarKind::PrimaryTrig, UnaryScalarKind::PrimaryTrig,
  /* Asin..Atan */ UnaryScalarKind::InverseTrig, UnaryScalarKind::InverseTrig, UnaryScalarKind::InverseTrig,
  /* Sinh..Tanh */ UnaryScalarKind::HyperbolicTrig, UnaryScalarKind::HyperbolicTrig, UnaryScalarKind::HyperbolicTrig,
  /* Acosh..Atanh */ UnaryScalarKind::InverseHyperbolicTrig, UnaryScalarKind::InverseHyperbolicTrig, UnaryScalarKind::InverseHyperbolicTrig,
  /* Exp */ UnaryScalarKind::Exp,
  /* Log */ UnaryScalarKind::None,
  /* Ln Log10 Sqrt Sqr */ UnaryScalarKind::Ln, UnaryScalarKind::Log10, UnaryScalarKind::Sqrt, UnaryScalarKind::Sqr,
  /* Int Frac Abs */ UnaryScalarKind::Rounding, UnaryScalarKind::Frac, UnaryScalarKind::Abs,
  /* Floor Ceil Trunc Round */ UnaryScalarKind::Rounding, UnaryScalarKind::Rounding, UnaryScalarKind::Rounding, UnaryScalarKind::Rounding,
  /* Sign Deg Rad */ UnaryScalarKind::Sign, UnaryScalarKind::Deg, UnaryScalarKind::Rad,
  /* Sum..Len (14) */ UnaryScalarKind::None, UnaryScalarKind::None, UnaryScalarKind::None, UnaryScalarKind::None,
  UnaryScalarKind::None, UnaryScalarKind::None, UnaryScalarKind::None, UnaryScalarKind::None,
  UnaryScalarKind::None, UnaryScalarKind::None, UnaryScalarKind::None, UnaryScalarKind::None,
  UnaryScalarKind::None, UnaryScalarKind::None,
  /* Fact */ UnaryScalarKind::None,
#if SMARTMATH_FACTORINT
  /* Factorint */ UnaryScalarKind::None,
#endif
  /* Avg..Conj */ UnaryScalarKind::None, UnaryScalarKind::None,
  UnaryScalarKind::None, UnaryScalarKind::None, UnaryScalarKind::None, UnaryScalarKind::None,
  UnaryScalarKind::None, UnaryScalarKind::None, UnaryScalarKind::None, UnaryScalarKind::None,
  UnaryScalarKind::None, UnaryScalarKind::None, UnaryScalarKind::None, UnaryScalarKind::None,
  UnaryScalarKind::None, UnaryScalarKind::None, UnaryScalarKind::None, UnaryScalarKind::None,
  UnaryScalarKind::None, UnaryScalarKind::None, UnaryScalarKind::None, UnaryScalarKind::None,
  UnaryScalarKind::None, UnaryScalarKind::None, UnaryScalarKind::None, UnaryScalarKind::None,
  /* Filter */ UnaryScalarKind::None,
  /* Iif */ UnaryScalarKind::None,
  /* IsInstance */ UnaryScalarKind::None,
  /* Dec Udec */ UnaryScalarKind::None, UnaryScalarKind::None,
};

static_assert(
  sizeof(kUnaryScalarKind) / sizeof(kUnaryScalarKind[0]) ==
    static_cast<std::size_t>(BuiltinFunctionId::Count),
  "kUnaryScalarKind size mismatch");

bool isTrigRadiansInRange(SmFloat radians) {
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

// ``name [ ... ] =`` at the start of a statement. The matching ``]`` closes the
// opening bracket; a following ``==`` stays a comparison.
bool peekIdentFollowedByIndexThenAssignEquals(const char* p) {
  p = skipAsciiSpacesPtr(endOfIdentTokenPtr(p));
  if (*p != '[') {
    return false;
  }
  int depth = 0;
  while (*p != '\0') {
    if (*p == '[') {
      ++depth;
    } else if (*p == ']') {
      --depth;
      if (depth == 0) {
        ++p;
        p = skipAsciiSpacesPtr(p);
        return *p == '=' && p[1] != '=';
      }
    }
    ++p;
  }
  return false;
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

bool isMultipleOf(SmFloat x, SmFloat x_mult)
{
  if (x_mult == 0.0) {
    return false;
  }
  const SmFloat q = x / x_mult;
  if (!std::isfinite(q)) {
    return false;
  }
  const SmFloat n = std::round(q);
  return std::fabs(q - n) <= kTolLost6;
}

bool tryTrigHalfPiQuotient(SmFloat x, SmInt& outK) {
  constexpr SmFloat kHalfPi = kPi / 2.0;
  if (!isMultipleOf(x, kHalfPi)) {
    return false;
  }
  const SmFloat q = std::round(x / kHalfPi);
  if (!std::isfinite(q) || std::fabs(q) >= kTrigMaxAbsRadians) {
    return false;
  }
  outK = static_cast<SmInt>(q);
  return true;
}

bool tryTrigQuarterPiQuotient(SmFloat x, SmInt& outK) {
  constexpr SmFloat kQuarterPi = kPi / 4.0;
  if (!isMultipleOf(x, kQuarterPi)) {
    return false;
  }
  const SmFloat q = std::round(x / kQuarterPi);
  if (!std::isfinite(q) || std::fabs(q) >= kTrigMaxAbsRadians) {
    return false;
  }
  outK = static_cast<SmInt>(q);
  return true;
}

SmFloat calcSin(SmFloat x)
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
  SmInt k = 0;
  if (tryTrigHalfPiQuotient(x, k)) {
    SmInt r = k % 4;
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

SmFloat calcCos(SmFloat x)
{
  if (!std::isfinite(x)) {
    return std::cos(x);
  }
  if (!isMultipleOf(x, kPi)) {
    SmInt k = 0;
    if (tryTrigHalfPiQuotient(x, k) && (k % 2) != 0) {
      return 0.0;
    }
  }
  return std::cos(x);
}

SmFloat calcTan(SmFloat x)
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
  SmInt k = 0;
  if (tryTrigHalfPiQuotient(x, k) && (k % 2) != 0) {
    return (k > 0) ? std::numeric_limits<SmFloat>::infinity()
                   : -std::numeric_limits<SmFloat>::infinity();
  }
  if (tryTrigQuarterPiQuotient(x, k) && (k % 2) != 0) {
    SmInt r = k % 4;
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

SmFloat calcAtan2Basic(SmFloat y, SmFloat x) {
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

SmFloat calcHypotBasic(SmFloat x, SmFloat y) {
  SmFloat ax = std::fabs(x);
  SmFloat ay = std::fabs(y);
  if (ax < ay) std::swap(ax, ay);
  if (ay == 0.0 || !std::isfinite(ax)) return ax + ay;
  // Squares stay in range here; outside it scale by the larger part to avoid overflow/underflow.
  if (ax <= K_HYPOT_DIRECT_MAX && ay >= K_HYPOT_DIRECT_MIN) return std::sqrt((ax * ax) + (ay * ay));
  const SmFloat t = ay / ax;
  return ax * std::sqrt(1.0 + t * t);
}

bool fractionalPowerIsOddUnitRoot(SmFloat p, SmInt& outN) {
  if (p <= 0.0 || p >= 1.0) {
    return false;
  }
  const SmFloat inv = 1.0 / p;
  // Any odd n whose reciprocal is still the closest double. 2^53 is the last exact integer double.
  if (inv < 2.0 || inv > kMaxExactIntFromFloat) {
    return false;
  }
  outN = smartmath::roundFloatToSInt(inv);
  if (outN < 2) {
    return false;
  }
  // Only the double nearest 1/n qualifies; 0.33333334 is not a cube root.
  if (std::fabs(p * static_cast<SmFloat>(outN) - 1.0) > K_UNIT_ROOT_EXPONENT_TOL) {
    return false;
  }
  return (outN % 2) != 0;
}

bool fractionalPowerResolveRootDegree(SmFloat p, SmInt& outN) {
  if (fractionalPowerIsOddUnitRoot(p, outN)) {
    return true;
  }
  if (p == 0.5) {
    outN = 2;
    return true;
  }
  return false;
}

bool tryGetSmallIntegerExponent(SmFloat p, SmInt& outN) {
  if (!(p >= 0.0 && p <= static_cast<SmFloat>(kMaxShift)) || p != std::floor(p)) {
    return false;
  }
  outN = static_cast<SmInt>(p);
  return true;
}

bool tryRoundFloatToSInt(SmFloat r, SmInt& out) {
  const SmFloat rounded = std::round(r);
  if (!(std::fabs(rounded) < kFloatSIntBound)) {
    return false;
  }
  out = static_cast<SmInt>(rounded);
  return true;
}

bool tryRoundFloatToUInt(SmFloat r, SmUInt& out) {
  const SmFloat rounded = std::round(r);
  if (!(rounded >= 0.0 && rounded < kFloatUIntBound)) {
    return false;
  }
  out = static_cast<SmUInt>(rounded);
  return true;
}

#if SMARTMATH_COMPLEX_NUMBERS
void snapComplexNearZeroAxis(SmFloat& zr, SmFloat& zi) {
  if (!std::isfinite(zr) || !std::isfinite(zi)) {
    return;
  }
  constexpr SmFloat kRel = kTolLost2;
  if (std::fabs(zr) <= kRel * std::fmax(1.0, std::fabs(zi))) {
    zr = 0.0;
  }
  if (std::fabs(zi) <= kRel * std::fmax(1.0, std::fabs(zr))) {
    zi = 0.0;
  }
}

void complexCartesianPrincipalNthRoot(SmFloat ar, SmFloat ai, SmFloat invN, SmFloat& outR, SmFloat& outI) {
  const SmFloat nanv = std::numeric_limits<SmFloat>::quiet_NaN();
  if (!std::isfinite(ar) || !std::isfinite(ai)) {
    outR = nanv;
    outI = nanv;
    return;
  }
  const SmFloat mag = std::hypot(ar, ai);
  if (mag == 0.0) {
    outR = 0.0;
    outI = 0.0;
    return;
  }
  const SmFloat angN = calcAtan2Basic(ai, ar) * invN;
  SmFloat rmN = 0.0;
  if (std::fabs(invN - 0.5) < kTolLost3) {
    rmN = std::sqrt(mag);
  } else {
    rmN = std::exp(std::log(mag) * invN);
  }
  outR = rmN * calcCos(angN);
  outI = rmN * calcSin(angN);
  snapComplexNearZeroAxis(outR, outI);
}

void scalarPrincipalLnCartesian(SmFloat ar, SmFloat ai, SmFloat& outRe, SmFloat& outIm) {
  const SmFloat nanv = std::numeric_limits<SmFloat>::quiet_NaN();
  if (!std::isfinite(ar) || !std::isfinite(ai)) {
    outRe = nanv;
    outIm = nanv;
    return;
  }
  const SmFloat mag = std::hypot(ar, ai);
  if (mag == 0.0) {
    outRe = -std::numeric_limits<SmFloat>::infinity();
    outIm = 0.0;
    return;
  }
  constexpr SmFloat kAxisEps = kTolLost0;
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

void complexPowPrincipal(SmFloat ar, SmFloat ai, SmFloat br, SmFloat bi, SmFloat& outRe, SmFloat& outIm) {
  const SmFloat nanv = std::numeric_limits<SmFloat>::quiet_NaN();
  if (!std::isfinite(ar) || !std::isfinite(ai) || !std::isfinite(br) || !std::isfinite(bi)) {
    outRe = nanv;
    outIm = nanv;
    return;
  }
  const SmFloat mag = std::hypot(ar, ai);
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
  constexpr SmFloat kEps = kTolLost1;
  if (std::fabs(bi) < kEps) {
    SmInt nRootFrac = 0;
    if (fractionalPowerResolveRootDegree(br, nRootFrac)) {
      complexCartesianPrincipalNthRoot(ar, ai, 1.0 / static_cast<SmFloat>(nRootFrac), outRe, outIm);
      return;
    }
    if (std::fabs(br - std::trunc(br)) < kTolLost3) {
      const SmInt n = static_cast<SmInt>(std::trunc(br));
      if (n >= -256 && n <= 256) {
        if (n >= 0) {
          SmFloat cr = 1.0;
          SmFloat ci = 0.0;
          for (SmInt k = 0; k < n; ++k) {
            const SmFloat nr = cr * ar - ci * ai;
            const SmFloat ni = cr * ai + ci * ar;
            cr = nr;
            ci = ni;
          }
          outRe = cr;
          outIm = ci;
        } else {
          SmFloat pr = 0.0;
          SmFloat pi = 0.0;
          SmFloat cr = 1.0;
          SmFloat ci = 0.0;
          for (SmInt k = 0; k < -n; ++k) {
            const SmFloat nr = cr * ar - ci * ai;
            const SmFloat ni = cr * ai + ci * ar;
            cr = nr;
            ci = ni;
          }
          pr = cr;
          pi = ci;
          const SmFloat den = pr * pr + pi * pi;
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
  const SmFloat loR = std::log(mag);
  const SmFloat loI = calcAtan2Basic(ai, ar);
  const SmFloat powRe = br * loR - bi * loI;
  const SmFloat powIm = br * loI + bi * loR;
  if (!isTrigRadiansInRange(powIm)) {
    outRe = std::numeric_limits<SmFloat>::quiet_NaN();
    outIm = std::numeric_limits<SmFloat>::quiet_NaN();
    return;
  }
  outRe = std::exp(powRe) * calcCos(powIm);
  outIm = std::exp(powRe) * calcSin(powIm);
  snapComplexNearZeroAxis(outRe, outIm);
}

void complexDivide(SmFloat numR, SmFloat numI, SmFloat denR, SmFloat denI, SmFloat& outR, SmFloat& outI) {
  const SmFloat den = denR * denR + denI * denI;
  if (den == 0.0) {
    outR = std::numeric_limits<SmFloat>::quiet_NaN();
    outI = std::numeric_limits<SmFloat>::quiet_NaN();
  } else {
    outR = (numR * denR + numI * denI) / den;
    outI = (numI * denR - numR * denI) / den;
  }
}

void complexMultiply(SmFloat ar, SmFloat ai, SmFloat br, SmFloat bi, SmFloat& outR, SmFloat& outI) {
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

bool complexExpCartesian(SmFloat ar, SmFloat ai, SmFloat& outR, SmFloat& outI) {
  if (!isTrigRadiansInRange(ai)) {
    return false;
  }
  const SmFloat ea = std::exp(ar);
  outR = ea * calcCos(ai);
  outI = ea * calcSin(ai);
  return true;
}

void complexPrincipalSqrt(SmFloat ar, SmFloat ai, SmFloat& outR, SmFloat& outI) {
  complexCartesianPrincipalNthRoot(ar, ai, 0.5, outR, outI);
}

void complexGamma(SmFloat zr, SmFloat zi, SmFloat& outR, SmFloat& outI) {
  const SmFloat nanv = std::numeric_limits<SmFloat>::quiet_NaN();
  if (!std::isfinite(zr) || !std::isfinite(zi)) {
    outR = nanv;
    outI = nanv;
    return;
  }
  static constexpr SmFloat kLanczosC[] = {
      0.99999999999980993,
      676.5203681218851,
      -1259.1392167224028,
      771.32342877765313,
      -176.61502916214059,
      12.507343278686905,
      -0.13857109526572012,
      9.9843695780195716e-6,
      1.5056327351493116e-7};
  constexpr SmFloat kG = 7.0;
  const SmFloat zmr = zr - 1.0;
  const SmFloat zmi = zi;
  SmFloat xRe = kLanczosC[0];
  SmFloat xIm = 0.0;
  for (int i = 1; i <= 8; ++i) {
    SmFloat quotR = 0.0;
    SmFloat quotI = 0.0;
    complexDivide(kLanczosC[i], 0.0, zmr + static_cast<SmFloat>(i), zmi, quotR, quotI);
    xRe += quotR;
    xIm += quotI;
  }
  const SmFloat tRe = zmr + kG + 0.5;
  const SmFloat tIm = zmi;
  SmFloat lnRe = 0.0;
  SmFloat lnIm = 0.0;
  scalarPrincipalLnCartesian(tRe, tIm, lnRe, lnIm);
  const SmFloat pwRe = zmr + 0.5;
  const SmFloat pwIm = zmi;
  SmFloat lnPowRe = 0.0;
  SmFloat lnPowIm = 0.0;
  complexMultiply(pwRe, pwIm, lnRe, lnIm, lnPowRe, lnPowIm);
  SmFloat powRe = 0.0;
  SmFloat powIm = 0.0;
  if (!complexExpCartesian(lnPowRe, lnPowIm, powRe, powIm)) {
    outR = nanv;
    outI = nanv;
    return;
  }
  SmFloat expNegPrR = 0.0;
  SmFloat expNegPrI = 0.0;
  if (!complexExpCartesian(-tRe, -tIm, expNegPrR, expNegPrI)) {
    outR = nanv;
    outI = nanv;
    return;
  }
  const SmFloat scale = std::sqrt(2.0 * kPi);
  SmFloat prodRe = 0.0;
  SmFloat prodIm = 0.0;
  complexMultiply(scale * powRe, scale * powIm, expNegPrR, expNegPrI, prodRe, prodIm);
  complexMultiply(prodRe, prodIm, xRe, xIm, outR, outI);
  snapComplexNearZeroAxis(outR, outI);
}
#endif

constexpr const char* STR_NAN = "nan";
constexpr const char* STR_NEG_INF = "-inf";
constexpr const char* STR_INF = "inf";
constexpr const char* STR_HEX_DIGITS_LOWER = "0123456789abcdef";
constexpr const char* STR_HEX_DIGITS_UPPER = "0123456789ABCDEF";
constexpr const char* STR_COMMA = ", ";
constexpr const char* STR_PI = "pi";
constexpr const char* STR_E = "e";
constexpr const char* STR_I = "i";
constexpr const char* STR_ANS = "ans";
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
constexpr const char* STR_FILTER = "filter";
constexpr const char* STR_IIF = "iif";
constexpr const char* STR_PAR_CONDITION_COMMA_TRUE_COMMA_FALSE = "(condition, expr_if_true, expr_if_false)";
constexpr const char* STR_ISINSTANCE = "isinstance";
constexpr const char* STR_PAR_VALUE_COMMA_TYPE = "(value, type)";
constexpr const char* STR_FLOAT = "float";
constexpr const char* STR_TIME = "time";
constexpr const char* STR_COMPLEX = "complex";
constexpr const char* STR_ARRAY = "array";
constexpr const char* STR_LIST = "list";
constexpr const char* STR_UNKNOWN_TYPE_NAME_COLON = "unknown type name: ";
constexpr const char* STR_PAR_EXPECTS_A_TYPE_NAME = "() expects a type name";
constexpr const char* STR_RATIO = "ratio";
constexpr const char* STR_RANGE = "range";
constexpr const char* STR_SORTED = "sorted";
constexpr const char* STR_SORTBY_EXPECTS_ONE_FUNCTION = "sortby expects exactly one function";
constexpr const char* STR_SORTBY_EXPECTS_UNARY_FUNCTION = "sortby expects a function that takes 1 parameter";
constexpr const char* STR_SORTBY_KEY_MUST_BE_SCALAR_OR_ARRAY = "sortby key function must return a scalar or an array";
constexpr const char* STR_FILTER_EXPECTS_ONE_FUNCTION = "filter expects exactly one function";
constexpr const char* STR_FILTER_EXPECTS_UNARY_FUNCTION = "filter expects a function that takes 1 parameter";
constexpr const char* STR_FILTER_PREDICATE_MUST_BE_SCALAR_OR_ARRAY = "filter predicate must return a scalar or an array";
constexpr const char* STR_PAR_ARRAY_COMMA_FUNC = "(array, func)";
constexpr SmFloat RATIO_APPROX_EPS = kTolLost1;
constexpr SmInt RATIO_MAX_DENOMINATOR = 10000000LL;
constexpr int RATIO_MAX_POWER10_EXP = kMaxPow10ExponentOfInt;
constexpr int RATIO_SEMICONV_LINEAR_THRESH = 64;
// Cap on continued-fraction terms; the denominator limit stops the expansion much earlier.
constexpr int RATIO_MAX_CF_TERMS = 64;
constexpr const char* STR_REVERSE = "reverse";
constexpr const char* STR_REVERSED = "reversed";
constexpr const char* STR_UNIQUE = "unique";
constexpr const char* STR_UNPACK = "unpack";
constexpr const char* STR_FLAT = "flat";
constexpr const char* STR_REPEAT = "repeat";
constexpr const char* STR_LEN = "len";
constexpr const char* STR_LENGTH = "length";
constexpr const char* STR_FACT = "fact";
#if SMARTMATH_FACTORINT
constexpr const char* STR_FACTORINT = "factorint";
#endif
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
constexpr const char* STR_DEC = "dec";
constexpr const char* STR_UDEC = "udec";
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
constexpr const char* STR_PAR_PRODUCES_TOO_MANY_VALUES = "() produces too many values";
constexpr int BUILTIN_MAX_PRODUCED_ITEMS = 1000;
constexpr const char* STR_PAR_DOTDOTDOT = "(...)";
constexpr const char* STR_UNEXPECTED_TOKEN = "unexpected token";
constexpr const char* STR_INCOMPATIBLE_OPERANDS = "incompatible operands";
constexpr const char* STR_EXPLICIT_FLOAT_FRAC_SUFFIX = ".0";
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
constexpr const char* STR_RESERVED_CONSTANT_NAME_COLON = "reserved constant name: ";
constexpr const char* STR_RESERVED_FUNCTION_NAME_COLON = "reserved function name: ";
constexpr const char* STR_RESERVED_FUNCTION_PARAM_NAME_COLON = "reserved function name can't be a param name: ";
constexpr const char* STR_RESERVED_BUILTIN_VARIABLE_NAME_COLON = "reserved built-in variable name: ";
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
constexpr const char* STR_TIME_LITERAL_EMPTY_SEGMENT = "time literal: empty segment between colons";
constexpr const char* STR_TIME_LITERAL_INVALID_SEGMENT = "time literal: invalid segment";
constexpr const char* STR_TIME_COMPACT_EXPECTED_UNIT = "compact time literal: expected unit suffix";
constexpr const char* STR_TIME_COMPACT_UNIT_ORDER = "compact time literal: unit order or duplicate unit";
constexpr const char* STR_TIME_COMPACT_INVALID_SUFFIX = "compact time literal: invalid suffix";
constexpr const char* STR_TIME_NON_FINITE = "time value: non-finite operand";
constexpr const char* STR_TIME_ARRAY_MIXED = "time values cannot be mixed with non-time values";
constexpr const char* STR_TIME_EXPECTS_TIME_ARG = "() expects a time value";
constexpr const char* STR_INDEXING_REQUIRES_AN_ARRAY_VALUE = "indexing requires an array value";
constexpr const char* STR_ARRAY_INDEX_MUST_BE_A_SCALAR = "array index must be a scalar integer";
constexpr const char* STR_ARRAY_INDEX_MUST_BE_AN_INTEGER = "array index must be an integer";
constexpr const char* STR_ARRAY_INDEX_IS_OUT_OF_RANGE = "array index is out of range";
constexpr const char* STR_CANNOT_ASSIGN_TO_AN_ARRAY_SLICE = "cannot assign to an array slice";
constexpr const char* STR_SLICE_STEP_MUST_NOT_BE_ZERO = "slice step must not be zero";
constexpr const char* STR_PAR_EXPECTS_AT_LEAST = "() expects at least ";
constexpr const char* STR_ARGUMENT = " argument";
constexpr const char* STR_ARGUMENTS = " arguments";
constexpr const char* STR_INTERNAL_ERROR_IN_AGGREGATE_BUILTIN = "internal error in aggregate builtin";
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
constexpr const char* STR_DEFINED = "defined ";
constexpr const char* STR_USER_DEFINED_FUNCTION_COLON = "user-defined function: ";
constexpr const char* STR_SEMICOLON_UNKNOWN_FUNCTION_COLON = "; unknown function: ";
constexpr const char* STR_UNKNOWN_FUNCTION_COLON = "unknown function: ";
constexpr const char* STR_INVALID_USER_FUNCTION_EXPRESSION = "invalid user function expression";
constexpr const char* STR_UNEXPECTED_CONTENT_AFTER_FUNCTION_DEFINITION = "unexpected content after function definition";


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

SmFloat randomUnitScalar() {
  return static_cast<SmFloat>(std::rand()) / static_cast<SmFloat>(RAND_MAX);
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

bool tryParsePrefixedUIntLiteral(const char* p, char prefixLower, unsigned int radix, const char*& outEnd, SmUInt& outValue) {
  if (p[0] != '0') {
    return false;
  }
  const char px = static_cast<char>(std::tolower(static_cast<unsigned char>(p[1])));
  if (px != prefixLower) {
    return false;
  }

  const char* cur = p + 2;
  SmUInt u = 0;
  int digits = 0;
  while (*cur) {
    const int d = parseDigitForRadix(*cur, radix);
    if (d < 0) {
      break;
    }
    const SmUInt ud = static_cast<SmUInt>(d);
    const SmUInt base = static_cast<SmUInt>(radix);
    if (u > ((kUIntMax - ud) / base)) {
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
inline bool digitsToUint(const unsigned char* digits, int digitCount, SmUInt& outValue) {
  SmUInt value = 0;
  for (int i = 0; i < digitCount; ++i) {
    const unsigned char digit = digits[i];
    if (value > (kUIntMax - digit) / 10u) {
      return false;
    }
    value = value * 10 + digit;
  }
  outValue = value;
  return true;
}

bool tryParseInputNumberAsInteger(const char* numStart, const char* numEnd, SmUInt& outValue) {

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
    if (significantDigits + adjust > kUIntDigits10) {
      return false;
    }

    SmUInt value = 0;
    if (!digitsToUint(digits, storedDigitCount, value)) {
      return false;
    }

    if (adjust > 0) {
      if (value > kUIntMax / pow10UInt(adjust)) {
        return false;
      }
      value *= pow10UInt(adjust);
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

  SmUInt value = 0;
  if (!digitsToUint(digits, resultDigitCount, value)) {
    return false;
  }

  outValue = value;
  return true;
}

// Callers must not pass operands whose gcd is 2^63 (e.g. kIntMin with 0); use gcdUInt on magnitudes.
SmInt GcdSInt(SmInt a, SmInt b) {
  return static_cast<SmInt>(gcdUInt(uintMagnitudeOfSInt(a), uintMagnitudeOfSInt(b)));
}

bool tryLcmUInt(SmUInt a, SmUInt b, SmUInt& out) {
  if (a == 0u || b == 0u) {
    out = 0u;
    return true;
  }
  const SmUInt g = gcdUInt(a, b);
  const SmUInt a1 = a / g;
  if (a1 > (kUIntMax / b)) {
    return false;
  }
  out = a1 * b;
  return true;
}

bool tryAddUIntChecked(SmUInt a, SmUInt b, SmUInt& out) {
  const SmUInt s = a + b;
  if (s < a) {
    return false;
  }
  out = s;
  return true;
}

bool tryAddSIntChecked(SmInt a, SmInt b, SmInt& out) {
  if ((b > 0 && a > kIntMax - b) ||
      (b < 0 && a < kIntMin - b)) {
    return false;
  }
  out = a + b;
  return true;
}

bool tryMulSIntChecked(SmInt a, SmInt b, SmInt& out) {
  if (a == 0 || b == 0) {
    out = 0;
    return true;
  }
  if (a == -1) {
    if (b == kIntMin) return false;
    out = -b;
    return true;
  }
  if (b == -1) {
    if (a == kIntMin) return false;
    out = -a;
    return true;
  }
  if (a > 0) {
    if (b > 0) {
      if (a > kIntMax / b) return false;
    } else {
      if (b < kIntMin / a) return false;
    }
  } else {
    if (b > 0) {
      if (a < kIntMin / b) return false;
    } else {
      if (a != 0 && b < kIntMax / a) return false;
    }
  }
  out = a * b;
  return true;
}

bool tryPowSIntChecked(SmInt base, SmInt exp, SmInt& out) {
  if (exp < 0) {
    return false;
  }
  SmInt r = 1;
  SmInt b = base;
  SmInt e = exp;
  while (e > 0) {
    if ((e & 1) != 0 && !tryMulSIntChecked(r, b, r)) {
      return false;
    }
    e >>= 1;
    if (e > 0 && !tryMulSIntChecked(b, b, b)) {
      return false;
    }
  }
  out = r;
  return true;
}

bool tryQuotExactUInt(SmUInt num, SmUInt den, SmUInt& quo) {
  if (den == 0u) {
    return false;
  }
  quo = num / den;
  SmUInt prod = 0;
  if (!tryMulUIntChecked(quo, den, prod)) {
    return false;
  }
  return prod == num;
}

bool tryComputeNprUInt(SmUInt n, SmUInt r, SmUInt& out) {
  if (r > n) return false;
  SmUInt acc = 1;
  for (SmUInt i = 0; i < r; ++i) {
    if (!tryMulUIntChecked(acc, n - i, acc)) return false;
  }
  out = acc;
  return true;
}

bool tryComputeNcrUInt(SmUInt n, SmUInt r, SmUInt& out) {
  if (r > n) return false;
  SmUInt rEff = r;
  if ((n - rEff) < rEff) rEff = (n - rEff);
  SmUInt acc = 1;
  for (SmUInt i = 1; i <= rEff; ++i) {
    SmUInt num = (n - rEff) + i;
    SmUInt den = i;
    const SmUInt g1 = gcdUInt(num, den);
    num /= g1;
    den /= g1;
    const SmUInt g2 = gcdUInt(acc, den);
    acc /= g2;
    den /= g2;
    if (den != 1u) return false;
    if (!tryMulUIntChecked(acc, num, acc)) return false;
  }
  out = acc;
  return true;
}


SmFloat factorialScalarFloatFromInt(SmInt n) {
  SmFloat d = static_cast<SmFloat>(factorialInt(kMaxFactorialArg));
  for (SmInt i = kMaxFactorialArg + 1; i <= n; ++i) {
    d *= static_cast<SmFloat>(i);
    if (!std::isfinite(d)) {
      break;
    }
  }
  return d;
}

std::string formatScientific16(SmFloat v) {
  // kFloatDisplayDigits (16 for double) significant digits, trailing mantissa zeros removed, exponent at least 3 digits.
  char buf[kFloatDisplayDigits + 32];
  const int n = smartmath::snprintfFloatE(buf, sizeof(buf), kFloatDisplayDigits - 1, v);
  if (n <= 0 || n >= static_cast<int>(sizeof(buf))) return "0";
  std::string text(buf, static_cast<std::size_t>(n));
  for (char& ch : text) {
    if (ch == 'E') ch = 'e';
  }
  const std::size_t expPos = text.find('e');
  if (expPos == std::string::npos) return text;
  std::string mantissa = text.substr(0, expPos);
  const std::size_t dotPos = mantissa.find('.');
  if (dotPos != std::string::npos) {
    while (!mantissa.empty() && mantissa.back() == '0') mantissa.pop_back();
    if (!mantissa.empty() && mantissa.back() == '.') mantissa.pop_back();
  }
  std::string exponent = text.substr(expPos + 1);
  char expSign = '+';
  std::size_t expIndex = 0;
  if (!exponent.empty() && (exponent[0] == '+' || exponent[0] == '-')) {
    expSign = exponent[0];
    expIndex = 1;
  }
  while (expIndex < exponent.size() && exponent[expIndex] == '0') ++expIndex;
  std::string expDigits = (expIndex < exponent.size()) ? exponent.substr(expIndex) : std::string();
  if (expDigits.empty()) expDigits = "0";
  constexpr std::size_t kMinExpDigits = 3;
  if (expDigits.size() < kMinExpDigits) {
    expDigits.insert(0, kMinExpDigits - expDigits.size(), '0');
  }
  return mantissa + "e" + expSign + expDigits;
}

// Decimal display of the float just above 2^(kIntBits-1) ("9.223372036854778e+018" for double): the deliberate
// display of the +-2^(kIntBits-1) float results flagged with DecScientificPow63High.
const std::string& posPow63DecText() {
  static const std::string text =
      formatScientific16(std::nextafter(kFloatSIntBound, std::numeric_limits<SmFloat>::infinity()));
  return text;
}

const std::string& negPow63DecText() {
  static const std::string text = "-" + posPow63DecText();
  return text;
}

bool generalFormatHidesFraction(const std::string& text) {
  return text.find('.') == std::string::npos && text.find('e') == std::string::npos &&
         text.find('E') == std::string::npos;
}

std::string formatFloatFast(SmFloat v) {
  // General format, kFloatDisplayDigits (16 for double) significant digits, exponent at least 3 digits (Basic str() parity).
  // A non-exact float whose general form has no fraction is marked so it is not read as an exact int:
  // magnitude above 2^53 uses scientific form; a smaller magnitude keeps the digits and appends ".0".
  if (std::isnan(v)) return STR_NAN;
  if (std::isinf(v)) return (v < 0.0) ? STR_NEG_INF : STR_INF;
  if (v == 0.0) return "0";
  char buf[kFloatDisplayDigits + 32];
  const int n = smartmath::snprintfFloatG(buf, sizeof(buf), kFloatDisplayDigits, v);
  if (n <= 0 || n >= static_cast<int>(sizeof(buf))) return "0";
  const char* expMark = std::strchr(buf, 'e');
  std::string text;
  if (expMark == nullptr) {
    text.assign(buf, static_cast<std::size_t>(n));
  } else {
    const char* expDigits = expMark + 2;
    const std::size_t expDigitCount = std::strlen(expDigits);
    constexpr std::size_t kMinExpDigits = 3;
    if (expDigitCount >= kMinExpDigits) {
      text.assign(buf, static_cast<std::size_t>(n));
    } else {
      text.assign(buf, static_cast<std::size_t>(expDigits - buf));
      text.append(kMinExpDigits - expDigitCount, '0');
      text.append(expDigits);
    }
  }
  if (!generalFormatHidesFraction(text)) return text;
  if (std::fabs(v) > kMaxExactIntFromFloat) return formatScientific16(v);
  text += STR_EXPLICIT_FLOAT_FRAC_SUFFIX;
  return text;
}

std::string formatUnsignedBase(SmUInt u, unsigned base, const char* prefix, bool uppercase) {
  static constexpr const char* kDigitsLo = STR_HEX_DIGITS_LOWER;
  static constexpr const char* kDigitsHi = STR_HEX_DIGITS_UPPER;
  const char* digits = uppercase ? kDigitsHi : kDigitsLo;

  char tmp[kRadixDigitBufLen];
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
  if (baseCode == static_cast<int>(smartmath_internal::RenderBase::PlainDec) ||
      baseCode == static_cast<int>(smartmath_internal::RenderBase::Dec)) {
    return 10U;
  }
  return 2U;
}

const char* renderBasePrefix(int baseCode) {
  if (baseCode == 16) return STR_PREFIX_HEX;
  if (baseCode == 8) return STR_PREFIX_OCT;
  if (baseCode == static_cast<int>(smartmath_internal::RenderBase::PlainDec) ||
      baseCode == static_cast<int>(smartmath_internal::RenderBase::Dec)) {
    return "";
  }
  return STR_PREFIX_BIN;
}

bool renderBaseUppercase(int baseCode) {
  return baseCode == 16;
}

std::string formatUnsignedForRenderBase(SmUInt u, int baseCode) {
  return formatUnsignedBase(u, renderBaseRadix(baseCode), renderBasePrefix(baseCode), renderBaseUppercase(baseCode));
}

std::string formatSignedMagnitudeForRenderBase(
    SmInt iv,
    int baseCode,
    bool asUnsigned) {
  if (asUnsigned) {
    return formatUnsignedForRenderBase(static_cast<SmUInt>(iv), baseCode);
  }
  SmUInt mag = 0;
  if (iv < 0) {
    if (iv == kIntMin) {
      mag = kIntMinMag;
    } else {
      mag = static_cast<SmUInt>(-iv);
    }
    return "-" + formatUnsignedForRenderBase(mag, baseCode);
  }
  return formatUnsignedForRenderBase(static_cast<SmUInt>(iv), baseCode);
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

bool checkedAddLL(SmInt a, SmInt b, SmInt& out) {
  return tryAddSIntChecked(a, b, out);
}

bool checkedSubLL(SmInt a, SmInt b, SmInt& out) {
  if ((b < 0 && a > kIntMax + b) ||
      (b > 0 && a < kIntMin + b)) {
    return false;
  }
  out = a - b;
  return true;
}

bool checkedMulLL(SmInt a, SmInt b, SmInt& out) {
  return tryMulSIntChecked(a, b, out);
}

SmInt bitwiseShiftLeftDefined(SmInt value, unsigned int shiftCount) {
  const SmUInt u = static_cast<SmUInt>(value);
  const SmUInt shifted = u << shiftCount;
  if (shifted <= static_cast<SmUInt>(kIntMax)) {
    return static_cast<SmInt>(shifted);
  }
  const SmUInt magnitude = (~shifted) + 1u;
  if (magnitude == kIntMinMag) {
    return kIntMin;
  }
  return -static_cast<SmInt>(magnitude);
}

SmInt bitwiseShiftRightDefined(SmInt value, unsigned int shiftCount) {
  // Arithmetic right shift; the shift count is in [0, kMaxShift]. Signed `>>` is implementation-defined
  // (not undefined) and arithmetic on every supported compiler; MathParserTypes.hpp static_asserts it.
  return value >> shiftCount;
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

std::string pad2DigitsLl(SmInt n) {
  if (n < 0) {
    n = 0;
  }
  if (n >= 100) {
    return smartmath::intToString(n);
  }
  if (n >= 10) {
    return std::string(1, static_cast<char>('0' + (n / 10))) + std::string(1, static_cast<char>('0' + (n % 10)));
  }
  return std::string("0") + std::string(1, static_cast<char>('0' + (n % 10)));
}

#if SMARTMATH_TIME_VALUES
bool tryRoundHalfUpFloatToSInt(SmFloat x, SmInt& out) {
  if (!std::isfinite(x)) {
    return false;
  }
  const SmFloat r = (x >= 0.0) ? std::floor(x + 0.5) : -std::floor(-x + 0.5);
  if (r < -kFloatSIntBound || r >= kFloatSIntBound) {
    return false;
  }
  out = static_cast<SmInt>(r);
  return true;
}

// For values already known to be in SmInt range; returns 0 otherwise.
SmInt roundHalfUpFloatToSInt(SmFloat x) {
  SmInt r = 0;
  return tryRoundHalfUpFloatToSInt(x, r) ? r : 0;
}

SmInt secondsFieldToMsRounded(SmInt wholeSec, const std::string& fracDigits) {
  SmFloat d = static_cast<SmFloat>(wholeSec);
  if (!fracDigits.empty()) {
    const std::string s = std::string("0.") + fracDigits;
    char* end = nullptr;
    const SmFloat fd = smartmath::parseFloatPrefix(s.c_str(), &end);
    if (end == s.c_str() || !std::isfinite(fd)) {
      return 0;
    }
    d += fd;
  }
  return roundHalfUpFloatToSInt(d * 1000.0);
}

bool tryAddTimeMsChecked(SmInt a, SmInt b, SmInt& outMs) {
  return tryRoundHalfUpFloatToSInt(static_cast<SmFloat>(a) + static_cast<SmFloat>(b), outMs);
}

bool trySubTimeMsChecked(SmInt a, SmInt b, SmInt& outMs) {
  return tryRoundHalfUpFloatToSInt(static_cast<SmFloat>(a) - static_cast<SmFloat>(b), outMs);
}

bool tryProductCoeffUnitMs(SmInt coeff, SmInt factor, SmInt& outDelta) {
  if (coeff == 0 || factor == 0) {
    outDelta = 0;
    return true;
  }
  if (coeff > 0 && factor > 0) {
    if (coeff > kIntMax / factor) {
      return false;
    }
  } else {
    return false;
  }
  outDelta = coeff * factor;
  return true;
}

/** Parse `[litStart,litEnd)` as MM:SS, HH:MM:SS, or DD:HH:MM:SS (optional fractional last field). */
bool parseTimeLiteralStringToMs(const char* litStart, const char* litEnd, SmInt& outMs, const char*& err) {
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

  SmInt d = 0, h = 0, m = 0;
  SmInt lastWhole = 0;
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
    SmUInt uv = 0;
    for (const char* t = segPtr; t < wholeEndPtr; ++t) {
      const int dig = *t - '0';
      if (uv > (kUIntMax - static_cast<SmUInt>(dig)) / 10u) {
        err = STR_TIME_LITERAL_INVALID_SEGMENT;
        return false;
      }
      uv = uv * 10u + static_cast<SmUInt>(dig);
    }
    if (uv > static_cast<SmUInt>(kIntMax)) {
      err = STR_TIME_LITERAL_INVALID_SEGMENT;
      return false;
    }
    lastWhole = static_cast<SmInt>(uv);
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

  const SmInt secMs = secondsFieldToMsRounded(lastWhole, fracPart);
  SmFloat t = 0.0;
  switch (segCount) {
    case 2:
      t = static_cast<SmFloat>(m) * 60000.0 + static_cast<SmFloat>(secMs);
      break;
    case 3:
      t = (static_cast<SmFloat>(h) * 3600.0 + static_cast<SmFloat>(m) * 60.0) * 1000.0 + static_cast<SmFloat>(secMs);
      break;
    case 4:
      t = ((static_cast<SmFloat>(d) * 24.0 + static_cast<SmFloat>(h)) * 3600.0 + static_cast<SmFloat>(m) * 60.0) * 1000.0 +
          static_cast<SmFloat>(secMs);
      break;
    default:
      err = STR_TIME_LITERAL_INVALID_SEGMENT;
      return false;
  }
  if (!std::isfinite(t)) {
    err = STR_TIME_LITERAL_INVALID_SEGMENT;
    return false;
  }
  if (!tryRoundHalfUpFloatToSInt(t, outMs)) {
    err = STR_TIME_LITERAL_INVALID_SEGMENT;
    return false;
  }
  return true;
}

std::string formatTimeCanonicalFromMs(SmInt totalMs) {
  const bool neg = (totalMs < 0);
  SmUInt rU = 0;
  if (neg) {
    if (totalMs == kIntMin) {
      rU = kIntMinMag;
    } else {
      rU = static_cast<SmUInt>(-totalMs);
    }
  } else {
    rU = static_cast<SmUInt>(totalMs);
  }
  SmInt msPart = static_cast<SmInt>(rU % 1000ull);
  rU /= 1000ull;
  SmInt sPart = static_cast<SmInt>(rU % 60ull);
  rU /= 60ull;
  SmInt mPart = static_cast<SmInt>(rU % 60ull);
  rU /= 60ull;
  SmInt hPart = static_cast<SmInt>(rU % 24ull);
  rU /= 24ull;
  const SmUInt dPart = rU;
  std::string body;
  if (dPart > 0ull) {
    body = smartmath::uintToString(dPart) + ":" + pad2DigitsLl(hPart) + ":" + pad2DigitsLl(mPart) + ":" + pad2DigitsLl(sPart);
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

std::string toLower(std::string s) {
  std::transform(s.begin(), s.end(), s.begin(), [](unsigned char c) { return static_cast<char>(std::tolower(c)); });
  return s;
}

std::string trim(const std::string& s) {
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

bool bodyIsEffectivelyEmpty(const std::string& bodyExpr) {
  const std::string t = trim(bodyExpr);
  if (t.empty()) {
    return true;
  }
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

bool isNumericLiteralStart(char c) {
  return std::isdigit(static_cast<unsigned char>(c)) || c == '.';
}

void skipSpaces(const char*& p) {
  while (*p && std::isspace(static_cast<unsigned char>(*p))) {
    ++p;
  }
}

std::string consumeLowerIdentToken(const char*& p) {
  const char* start = p++;
  while (localIsIdentChar(*p)) {
    ++p;
  }
  const char* const end = p;
  std::string out;
  assignLowerIdentFromRange(out, start, end);
  return out;
}

bool consumeKeyword(const char*& p, const char* kw) {
  skipSpaces(p);
  std::size_t n = 0;
  while (kw[n] != '\0') {
    const unsigned char a = static_cast<unsigned char>(p[n]);
    const unsigned char b = static_cast<unsigned char>(kw[n]);
    if (a == '\0' || std::tolower(a) != std::tolower(b)) return false;
    ++n;
  }
  if (localIsIdentChar(p[n])) return false;
  p += n;
  return true;
}

#if SMARTMATH_COMPLEX_NUMBERS
bool isTightImagUnitSuffixAt(const char* p) {
  if (*p != 'i' && *p != 'I') {
    return false;
  }
  const unsigned char next = static_cast<unsigned char>(p[1]);
  return !(std::isalnum(next) || next == '_');
}
#endif

bool tryPowFloatGuess(SmInt valueInt, SmFloat baseScalar, SmFloat p, SmFloat& outR) {
  if (valueInt >= 0 || p <= 0.0 || p >= 1.0) {
    outR = std::pow(baseScalar, p);
    return std::isfinite(outR);
  }
  SmInt nRoot = 0;
  if (!fractionalPowerIsOddUnitRoot(p, nRoot)) {
    outR = std::pow(baseScalar, p);
    return std::isfinite(outR);
  }
  const SmFloat magD = (valueInt == kIntMin)
    ? kFloatSIntBound
    : static_cast<SmFloat>(-valueInt);
  const SmFloat rootMag = std::pow(magD, p);
  if (!std::isfinite(rootMag)) {
    return false;
  }
  outR = -rootMag;
  return true;
}

constexpr SmFloat kExactIntFloatEps = kTolLost3;

bool relCloseExactIntFloat(SmFloat a, SmFloat b) {
  const SmFloat scale = std::max(static_cast<SmFloat>(1), std::fabs(b));
  return std::fabs(a - b) <= kExactIntFloatEps * scale;
}

bool tryRecoverReducedRationalFromFloatAbs(SmFloat fAbs, SmInt& outNum, SmInt& outDen) {
  constexpr SmInt kMaxB = 4096;
  SmFloat bestErr = kHugeSentinel;
  SmInt bestA = 0;
  SmInt bestB = 0;
  for (SmInt b = 1; b <= kMaxB; ++b) {
    SmInt a = smartmath::roundFloatToSInt(fAbs * static_cast<SmFloat>(b));
    if (a < 1) {
      continue;
    }
    const SmInt g = GcdSInt(a, b);
    a /= g;
    const SmInt bReduced = b / g;
    const SmFloat approx = static_cast<SmFloat>(a) / static_cast<SmFloat>(bReduced);
    const SmFloat err = std::fabs(fAbs - approx);
    if (err < bestErr) {
      bestErr = err;
      bestA = a;
      bestB = bReduced;
    }
  }
  if (bestA < 1 || bestB < 1) {
    return false;
  }
  const SmFloat scale = std::max(static_cast<SmFloat>(1), fAbs);
  if (bestErr > kExactIntFloatEps * scale) {
    return false;
  }
  outNum = bestA;
  outDen = bestB;
  return true;
}

using smartmath_internal::BuiltinCategory;
using smartmath_internal::BuiltinFlags;
using smartmath_internal::BuiltinFunctionId;
using smartmath_internal::BuiltinHintKind;
using smartmath_internal::CallErrorId;
using smartmath_internal::EvalContext;
using smartmath_internal::EvalValue;
using smartmath_internal::Expr;
using smartmath_internal::OperatorNameId;
using smartmath_internal::ParseErrorId;
using smartmath_internal::RenderBase;
using smartmath_internal::ScalarExactKind;
using smartmath_internal::ScalarKind;
using smartmath_internal::ValueKind;

// value * 2^shiftCount: exact SmInt when it fits, otherwise a float (same as the SmUInt overflow path).
EvalValue makeScalarShiftLeftSInt(SmInt value, unsigned int shiftCount) {
  const SmInt lo = bitwiseShiftRightDefined(kIntMin, shiftCount);
  const SmInt hi = bitwiseShiftRightDefined(kIntMax, shiftCount);
  if (value < lo || value > hi) {
    return makeScalarMaybeExact(std::ldexp(static_cast<SmFloat>(value), static_cast<int>(shiftCount)));
  }
  return makeScalarInt(bitwiseShiftLeftDefined(value, shiftCount));
}

EvalValue scalarFromScalarValue(const EvalValue::ScalarValue& sv) {
  EvalValue out = makeScalar(sv.scalar);
  out.scalarValue.scalarKind = sv.scalarKind;
  out.scalarValue.setExactSIntValid(sv.hasExactSInt());
  out.scalarValue.exactSInt = sv.exactSInt;
  out.scalarValue.setExactUIntValid(sv.hasExactUInt());
  out.scalarValue.exactUInt = sv.exactUInt;
  out.scalarValue.setDecScientificPow63High(sv.hasDecScientificPow63High());
  out.scalarValue.imag = sv.imag;
  out.scalarValue.imagExactSInt = sv.imagExactSInt;
  out.scalarValue.imagExactUInt = sv.imagExactUInt;
  out.scalarValue.setImagExactSIntValid(sv.hasImagExactSInt());
  out.scalarValue.setImagExactUIntValid(sv.hasImagExactUInt());
  out.scalarValue.setRenderRational(sv.hasRenderRational());
  out.scalarValue.setImagRenderRational(sv.hasImagRenderRational());
#if SMARTMATH_FACTORINT
  out.scalarValue.setRenderIntPower(sv.hasRenderIntPower());
#endif
  return out;
}

EvalValue scalarFromArrayAt(const EvalValue& arrV, std::size_t idx) {
  return scalarFromScalarValue(arrV.arr[idx]);
}

void scalarRepairExactMetadata(EvalValue::ScalarValue& sv) {
  switch (sv.scalarKind) {
    case ScalarKind::SInt:
      sv.setExactSIntValid(true);
      if (sv.exactSInt >= 0) {
        sv.setExactUIntValid(true);
        sv.exactUInt = static_cast<SmUInt>(sv.exactSInt);
      } else if (sv.exactUInt != 0u) {
        sv.setExactUIntValid(true);
      } else {
        sv.setExactUIntValid(false);
        sv.exactUInt = 0;
      }
      break;
    case ScalarKind::UInt:
      sv.setExactUIntValid(true);
      if (sv.exactUInt <= static_cast<SmUInt>(kIntMax)) {
        sv.setExactSIntValid(true);
        sv.exactSInt = static_cast<SmInt>(sv.exactUInt);
      } else {
        sv.setExactSIntValid(false);
        sv.exactSInt = 0;
      }
      break;
    default:
      if (sv.hasRenderRational()) {
        sv.setExactSIntValid(true);
        sv.setExactUIntValid(true);
      } else {
        sv.setExactSIntValid(false);
        sv.setExactUIntValid(false);
      }
      break;
  }
#if SMARTMATH_COMPLEX_NUMBERS
#if SMARTMATH_FACTORINT
  if (!sv.hasRenderIntPower()) {
#endif
    if (!sv.hasImagExactSInt() && sv.imagExactSInt != 0) {
      SmInt tIm = 0;
      if (tryExtractExactSIntFromFloatStrict(sv.imag, tIm) && tIm == sv.imagExactSInt) {
        sv.setImagExactSIntValid(true);
        if (sv.imagExactSInt >= 0) {
          sv.setImagExactUIntValid(true);
          sv.imagExactUInt = static_cast<SmUInt>(sv.imagExactSInt);
        }
      }
    } else if (!sv.hasImagExactUInt() && sv.imagExactUInt != 0u) {
      SmInt tIm = 0;
      if (tryExtractExactSIntFromFloatStrict(sv.imag, tIm) &&
          static_cast<SmUInt>(tIm) == sv.imagExactUInt) {
        sv.setImagExactUIntValid(true);
        if (sv.imagExactUInt <= static_cast<SmUInt>(kIntMax)) {
          sv.setImagExactSIntValid(true);
          sv.imagExactSInt = static_cast<SmInt>(sv.imagExactUInt);
        }
      }
    }
#if SMARTMATH_FACTORINT
  }
#endif
#endif
}

SmFloat scalarNumericReal(const EvalValue::ScalarValue& s) {
  EvalValue::ScalarValue sv = s;
  scalarRepairExactMetadata(sv);
  if (sv.hasRenderRational()) {
    if (std::isfinite(sv.scalar) && sv.scalar != 0.0) {
      return sv.scalar;
    }
    if (sv.exactUInt == 0u) {
      return sv.scalar;
    }
    return static_cast<SmFloat>(sv.exactSInt) / static_cast<SmFloat>(sv.exactUInt);
  }
  if (sv.hasExactSInt()) {
    return static_cast<SmFloat>(sv.exactSInt);
  }
  if (sv.scalarKind == ScalarKind::SInt) {
    return static_cast<SmFloat>(sv.exactSInt);
  }
  if (sv.hasExactUInt()) {
    return static_cast<SmFloat>(sv.exactUInt);
  }
  if (sv.scalarKind == ScalarKind::UInt) {
    return static_cast<SmFloat>(sv.exactUInt);
  }
  return sv.scalar;
}

#if SMARTMATH_COMPLEX_NUMBERS
bool imagExactMetadataMatchesFloat(const EvalValue::ScalarValue& sv) {
  SmInt t = 0;
  if (!tryExtractExactSIntFromFloatStrict(sv.imag, t)) {
    return false;
  }
  if (sv.hasImagExactSIntMetadata() && t == sv.imagExactSInt) {
    return true;
  }
  if (sv.hasImagExactUIntMetadata() && static_cast<SmUInt>(t) == sv.imagExactUInt) {
    return true;
  }
  return false;
}

SmFloat scalarNumericImag(const EvalValue::ScalarValue& s) {
  EvalValue::ScalarValue sv = s;
  scalarRepairExactMetadata(sv);
  if (sv.hasImagRenderRational()) {
    if (std::isfinite(sv.imag) && sv.imag != 0.0) {
      return sv.imag;
    }
    if (sv.imagExactUInt == 0u) {
      return sv.imag;
    }
    return static_cast<SmFloat>(sv.imagExactSInt) / static_cast<SmFloat>(sv.imagExactUInt);
  }
  if (sv.hasImagExactSInt() || sv.hasImagExactUInt()) {
    if (!imagExactMetadataMatchesFloat(sv)) {
      return sv.imag;
    }
  }
  if (sv.hasImagExactSInt()) {
    return static_cast<SmFloat>(sv.imagExactSInt);
  }
  if (sv.hasImagExactUInt()) {
    return static_cast<SmFloat>(sv.imagExactUInt);
  }
  return sv.imag;
}

bool scalarHasNonzeroImaginaryPart(const EvalValue::ScalarValue& s) {
  if (s.imag != 0.0) {
    return true;
  }
  return s.hasImagExactSInt() && s.imagExactSInt != 0;
}

bool evalValueHasNonzeroImaginary(const EvalValue& v) {
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
#endif

void scalarLoadCartesian(const EvalValue::ScalarValue& s, SmFloat& re, SmFloat& im) {
  re = scalarNumericReal(s);
#if SMARTMATH_COMPLEX_NUMBERS
  im = scalarNumericImag(s);
#else
  im = 0.0;
#endif
}

bool isTruthy(const EvalValue& v) {
  if (v.kind == ValueKind::Array) {
    return !v.arr.empty();
  }
  const EvalValue::ScalarValue& sv = v.scalarValue;
  SmFloat re = 0.0;
  SmFloat im = 0.0;
  scalarLoadCartesian(sv, re, im);
  if (std::isnan(re) || std::isnan(im)) {
    return false;
  }
  return std::fabs(re) > 0.0 || std::fabs(im) > 0.0;
}

bool isPureFloatingScalarPair(const EvalValue::ScalarValue& a, const EvalValue::ScalarValue& b) {
  if (a.scalarKind != ScalarKind::FloatingPoint || b.scalarKind != ScalarKind::FloatingPoint ||
      a.hasExactSInt() || a.hasExactUInt() || b.hasExactSInt() || b.hasExactUInt()) {
    return false;
  }
#if SMARTMATH_COMPLEX_NUMBERS
  if (scalarHasNonzeroImaginaryPart(a) || scalarHasNonzeroImaginaryPart(b)) {
    return false;
  }
#endif
  return true;
}

bool applyBinary(SmFloat a, SmFloat b, char op, SmFloat& out) {
  switch (op) {
    case '+': out = a + b; return true;
    case '-': out = a - b; return true;
    case '*': out = a * b; return true;
    case '/':
      if (b == 0.0) {
        if (std::isnan(a) || a == 0.0) {
          out = std::numeric_limits<SmFloat>::quiet_NaN();
        } else {
          const bool neg = std::signbit(a) != std::signbit(b);
          out = neg ? -std::numeric_limits<SmFloat>::infinity() : std::numeric_limits<SmFloat>::infinity();
        }
        return true;
      }
      out = a / b;
      return true;
    case '^': out = std::pow(a, b); return true;
    default: break;
  }
  return false;
}

template <typename ScalarFn>
EvalValue mapUnaryEvalValue(const EvalValue& in, ScalarFn&& applyScalar) {
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

template <typename Visitor>
std::size_t forEachCallArgScalarValues(const std::vector<EvalValue>& args, Visitor&& visit) {
  std::size_t count = 0;
  for (const auto& a : args) {
    if (a.kind == ValueKind::Scalar) {
      visit(a.scalarValue);
      ++count;
    } else {
      for (const auto& item : a.arr) {
        visit(item);
        ++count;
      }
    }
  }
  return count;
}

template <typename Visitor>
std::size_t forEachFlattenedEvalValue(const std::vector<EvalValue>& args, Visitor&& visit) {
  std::size_t count = 0;
  for (const auto& a : args) {
    if (a.kind == ValueKind::Scalar) {
      visit(a);
      ++count;
    } else {
      for (const auto& item : a.arr) {
        visit(scalarFromScalarValue(item));
        ++count;
      }
    }
  }
  return count;
}

std::size_t countFlattenedScalars(const std::vector<EvalValue>& args) {
  return forEachCallArgScalarValues(args, [](const EvalValue::ScalarValue&) {});
}

bool flattenArgs(const std::vector<EvalValue>& args, std::vector<SmFloat>& out) {
  out.clear();
  out.reserve(countFlattenedScalars(args));
  forEachCallArgScalarValues(args, [&](const EvalValue::ScalarValue& sv) { out.emplace_back(sv.scalar); });
  return !out.empty();
}

bool flattenArgsToScalars(const std::vector<EvalValue>& args, std::vector<EvalValue>& out) {
  out.clear();
  out.reserve(forEachFlattenedEvalValue(args, [](const EvalValue&) {}));
  forEachFlattenedEvalValue(args, [&](const EvalValue& ev) { out.emplace_back(ev); });
  return !out.empty();
}

bool TryApplyExactUIntBinaryScalarOp(
    char op,
    SmUInt lu,
    SmUInt ru,
    EvalValue& outS) {
  if (op == '+') {
    SmUInt outU = 0;
    if (tryAddUIntChecked(lu, ru, outU)) {
      outS = makeScalarUInt(outU);
      return true;
    }
  } else if (op == '-') {
    if (lu >= ru) {
      outS = makeScalarUInt(lu - ru);
      return true;
    }
  } else if (op == '*') {
    SmUInt outU = 0;
    if (tryMulUIntChecked(lu, ru, outU)) {
      outS = makeScalarUInt(outU);
      return true;
    }
  } else if (op == '/') {
    SmUInt quoU = 0;
    if (tryQuotExactUInt(lu, ru, quoU)) {
      outS = makeScalarUInt(quoU);
      return true;
    }
  }
  return false;
}

bool argsContainNonFinite(const std::vector<EvalValue>& args) {
  bool found = false;
  forEachCallArgScalarValues(args, [&](const EvalValue::ScalarValue& s) {
    if (!found && !std::isfinite(s.scalar)) {
      found = true;
    }
  });
  return found;
}

EvalValue::ScalarValue clampScalarValue(
    const EvalValue::ScalarValue& s,
    const EvalValue::ScalarValue& minSv,
    const EvalValue::ScalarValue& maxSv) {
  const SmFloat v = scalarNumericReal(s);
  const SmFloat minS = minSv.scalar;
  const SmFloat maxS = maxSv.scalar;

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
    return makeScalar(std::numeric_limits<SmFloat>::quiet_NaN()).scalarValue;
  }

  return s;
}

bool isIneligibleUnaryBuiltin(BuiltinFunctionId id) {
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
    case BuiltinFunctionId::Iif:
    case BuiltinFunctionId::IsInstance:
      return true;
    default:
      return false;
  }
}

#if SMARTMATH_COMPLEX_NUMBERS
bool isComplexUnaryTrigBuiltin(BuiltinFunctionId id) {
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
#endif

const std::vector<std::string>& functionNames() {
  // refer to: BuiltinFunctionId
  static const std::vector<std::string> kNames = {
      STR_RAND,         STR_RANDOM,   STR_BIN,       STR_HEX,    STR_OCT,     STR_POW,
      STR_ATAN2,        STR_SIN,      STR_COS,       STR_TAN,    STR_ASIN,    STR_ACOS,
      STR_ATAN,         STR_SINH,     STR_COSH,      STR_TANH,   STR_ACOSH,   STR_ASINH,
      STR_ATANH,        STR_EXP,      STR_LOG,       STR_LN,     STR_LOG10,   STR_SQRT,
      STR_SQR,          STR_INT,      STR_FRAC,      STR_ABS,    STR_FLOOR,   STR_CEIL,
      STR_TRUNC,        STR_ROUND,    STR_SIGN,      STR_DEG,    STR_RAD,     STR_SUM,
      STR_MEDIAN,       STR_VARIANCE, STR_STDDEV,    STR_SORT,   STR_SORTBY,  STR_RATIO,
      STR_RANGE,        STR_REVERSE,  STR_UNIQUE,    STR_UNPACK, STR_FLAT,    STR_REPEAT,
      STR_LEN,          STR_FACT,
#if SMARTMATH_FACTORINT
      STR_FACTORINT,
#endif
                        STR_AVG,      STR_MEAN,    STR_MOD,
      STR_CLAMP,        STR_HYPOT,    STR_GCD,       STR_LCM,    STR_NCR,     STR_NPR,
      STR_PRODUCT,      STR_MIN,      STR_MAX,       STR_UHEX,   STR_UOCT,    STR_UBIN,
      STR_MILLISECONDS, STR_SECONDS,  STR_MINUTES,   STR_HOURS,  STR_DAYS,    STR_REAL,
      STR_IMAG,         STR_PHASE,    STR_POLAR,     STR_CART,   STR_CONJ,          STR_FILTER,
      STR_IIF,      STR_ISINSTANCE, STR_DEC,       STR_UDEC
  };
  return kNames;
}

const std::unordered_map<std::string, BuiltinFunctionId>& functionNameToId() {
  static const std::unordered_map<std::string, BuiltinFunctionId> kByName = [] {
    std::unordered_map<std::string, BuiltinFunctionId> m;
    m.reserve(functionNames().size() + 9u);
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

const std::vector<std::string>& operatorNames() {
  // refer to: OperatorNameId
  static const std::vector<std::string> kNames = {STR_NOT, STR_AND, STR_OR};
  return kNames;
}

const std::string& getFunctionName(BuiltinFunctionId id) {
  return functionNames().at(static_cast<std::size_t>(id));
}

const std::string& opName(OperatorNameId id) {
  return operatorNames().at(static_cast<std::size_t>(id));
}

bool tryGetBuiltinFunctionId(const std::string& nameText, BuiltinFunctionId& outId) {
  const auto& byName = functionNameToId();
  auto it = byName.find(nameText);
  if (it == byName.end()) {
    return false;
  }
  outId = it->second;
  return true;
}

BuiltinHintKind getBuiltinHintKind(BuiltinFunctionId id) {
  if (id >= BuiltinFunctionId::Count) {
    return BuiltinHintKind::None;
  }
  return kBuiltinMeta[static_cast<std::size_t>(id)].hintKind;
}

std::string getBuiltinFunctionMissingCallHint(BuiltinFunctionId id) {
  const BuiltinHintKind kind = getBuiltinHintKind(id);
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
    case BuiltinHintKind::Iif:
      return fnName + STR_PAR_CONDITION_COMMA_TRUE_COMMA_FALSE;
    case BuiltinHintKind::IsInstance:
      return fnName + STR_PAR_VALUE_COMMA_TYPE;
    default: return "";
  }
}

bool isOpKeyword(const std::string& nameText, OperatorNameId id) {
  return nameText == opName(id);
}

bool isLogicalBinaryOperatorKeyword(const std::string& nameText) {
  return isOpKeyword(nameText, OperatorNameId::And) || isOpKeyword(nameText, OperatorNameId::Or);
}

bool isReservedFunctionName(const std::string& nameText) {
  BuiltinFunctionId id = BuiltinFunctionId::Count;
  return tryGetBuiltinFunctionId(nameText, id)
      || isOpKeyword(nameText, OperatorNameId::Not)
      || isLogicalBinaryOperatorKeyword(nameText);
}

BuiltinCategory getBuiltinCategory(BuiltinFunctionId id) {
  if (id >= BuiltinFunctionId::Count) {
    return BuiltinCategory::Unknown;
  }
  return kBuiltinMeta[static_cast<std::size_t>(id)].category;
}

BuiltinFlags getBuiltinFlags(BuiltinFunctionId id) {
  if (id >= BuiltinFunctionId::Count) {
    return BuiltinFlags::None;
  }
  return kBuiltinMeta[static_cast<std::size_t>(id)].flags;
}

bool hasBuiltinFlag(BuiltinFunctionId id, BuiltinFlags flag) {
  const unsigned combined = static_cast<unsigned>(getBuiltinFlags(id));
  const unsigned mask = static_cast<unsigned>(flag);
  return (combined & mask) == mask;
}

bool isTrailingFormatterFunctionName(const std::string& nameText) {
  BuiltinFunctionId id = BuiltinFunctionId::Count;
  if (!tryGetBuiltinFunctionId(nameText, id)) {
    return false;
  }
  return hasBuiltinFlag(id, BuiltinFlags::TrailingFormatter);
}

std::string buildUnknownVariableErrorText(const std::string& unknownVarsText) {
  return STR_UNKNOWN_VARIABLE_COLON + unknownVarsText;
}

std::string buildUnknownFunctionErrorText(const std::string& unknownFuncsText) {
  return STR_UNKNOWN_FUNCTION_COLON + unknownFuncsText;
}

void appendUnknownFunctionErrorText(std::string& errorText, const std::string& unknownFuncsText) {
  errorText += STR_SEMICOLON_UNKNOWN_FUNCTION_COLON + unknownFuncsText;
}

std::string buildUnknownNameErrorText(
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

bool exprIsDirectPostfixPercent(const Expr& e) {
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

std::unique_ptr<Expr> makeBinaryExpr(
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

std::unique_ptr<Expr> makePostfixPercentExpr(std::unique_ptr<Expr> child) {
  auto w = std::make_unique<Expr>();
  w->tag = Expr::Tag::PostfixPercent;
  w->child = std::move(child);
  return w;
}

bool isComparisonBinaryOp(Expr::BinaryOp op) {
  return op == Expr::BinaryOp::CmpLt || op == Expr::BinaryOp::CmpGt ||
         op == Expr::BinaryOp::CmpLe || op == Expr::BinaryOp::CmpGe ||
         op == Expr::BinaryOp::CmpEq || op == Expr::BinaryOp::CmpNe;
}

bool evalComparisonByOp(Expr::BinaryOp op, int cmp) {
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

bool evalComparisonTruthWhenUnorderedNan(Expr::BinaryOp op) {
  switch (op) {
    case Expr::BinaryOp::CmpNe:
      return true;
    default:
      return false;
  }
}

void applyBuiltinFormatRenderMeta(BuiltinFunctionId id, EvalValue& out) {
  if (id == BuiltinFunctionId::Hex || id == BuiltinFunctionId::Uhex) {
    out.setRenderBase(RenderBase::Hex);
  } else if (id == BuiltinFunctionId::Oct || id == BuiltinFunctionId::Uoct) {
    out.setRenderBase(RenderBase::Oct);
  } else if (id == BuiltinFunctionId::Dec || id == BuiltinFunctionId::Udec) {
    out.setRenderBase(RenderBase::PlainDec);
  } else {
    out.setRenderBase(RenderBase::Bin);
  }
  out.setRenderUnsigned(
      id == BuiltinFunctionId::Uhex || id == BuiltinFunctionId::Uoct || id == BuiltinFunctionId::Ubin ||
      id == BuiltinFunctionId::Udec);
}

void clearBaseFormatIntPowerDisplay(EvalValue& v) {
#if SMARTMATH_FACTORINT
  if (v.kind == ValueKind::Array) {
    for (std::size_t i = 0; i < v.arr.size(); ++i) {
      v.arr[i].setRenderIntPower(false);
    }
  } else if (v.kind == ValueKind::Scalar) {
    v.scalarValue.setRenderIntPower(false);
  }
#else
  (void)v;
#endif
}

bool getBuiltinArity(BuiltinFunctionId id, uint8_t& minArgs, uint8_t& maxArgs) {
  if (id >= BuiltinFunctionId::Count) {
    return false;
  }
  if (hasBuiltinFlag(id, BuiltinFlags::Unary)) {
    minArgs = 1;
    maxArgs = 1;
    return true;
  }
  const auto& row = kBuiltinMeta[static_cast<std::size_t>(id)];
  if (row.minArgs == kBuiltinMetaArityUnset) {
    return false;
  }
  minArgs = row.minArgs;
  maxArgs = row.maxArgs;
  return true;
}

bool tryApplyExactIntegerDivisionFromQuotient(
    const EvalValue::ScalarValue& leftS,
    const EvalValue::ScalarValue& rightS,
    SmFloat r,
    EvalValue& outV) {
  if (!std::isfinite(r)) {
    return false;
  }
  SmUInt aU = 0;
  SmUInt bU = 0;
  if (tryGetExactNonNegativeUIntFromScalar(leftS, aU) && tryGetExactNonNegativeUIntFromScalar(rightS, bU)) {
    if (bU == 0u) {
      return false;
    }
    const SmFloat rq = std::round(r);
    if (!isWithinExactIntFromFloatRange(rq) || rq < 0.0) {
      return false;
    }
    const SmUInt qU = static_cast<SmUInt>(rq);
    if (static_cast<SmFloat>(qU) != rq) {
      return false;
    }
    SmUInt prodU = 0;
    if (!tryMulUIntChecked(qU, bU, prodU) || prodU != aU) {
      return false;
    }
    outV = makeScalarUInt(qU);
    return true;
  }
  SmInt aI = 0;
  SmInt bI = 0;
  if (tryGetExactSignedIntNoUIntWrapFromScalar(leftS, aI) &&
      tryGetExactSignedIntNoUIntWrapFromScalar(rightS, bI)) {
    if (bI == 0) {
      return false;
    }
    const SmFloat rq = std::round(r);
    if (!isWithinExactIntFromFloatRange(rq)) {
      return false;
    }
    const SmInt qI = static_cast<SmInt>(rq);
    if (static_cast<SmFloat>(qI) != rq) {
      return false;
    }
    SmInt prodI = 0;
    if (!checkedMulLL(qI, bI, prodI) || prodI != aI) {
      return false;
    }
    outV = makeScalarInt(qI);
    return true;
  }

  SmUInt intU = 0;
  if (tryGetExactNonNegativeUIntFromScalar(leftS, intU) && !scalarHasExactIntegerPayload(rightS)) {
    if (tryApplyExactIntFloatOpFromFloatSide(ScalarExactKind::Unsigned, false, 0LL, intU, rightS.scalar, r,
                                             outV)) {
      return true;
    }
  }

  SmInt intI = 0;
  if (tryGetExactSignedIntNoUIntWrapFromScalar(leftS, intI) && !scalarHasExactIntegerPayload(rightS)) {
    if (tryApplyExactIntFloatOpFromFloatSide(ScalarExactKind::Signed, false, intI, 0u, rightS.scalar, r, outV)) {
      return true;
    }
  }

  return false;
}

bool tryApplyExactIntegerMultiplicationFromProduct(
    const EvalValue::ScalarValue& leftS,
    const EvalValue::ScalarValue& rightS,
    SmFloat r,
    EvalValue& outV) {
  if (!std::isfinite(r)) {
    return false;
  }
  SmUInt intU = 0;
  if (tryGetExactNonNegativeUIntFromScalar(leftS, intU) && !scalarHasExactIntegerPayload(rightS)) {
    if (tryApplyExactIntFloatOpFromFloatSide(ScalarExactKind::Unsigned, true, 0LL, intU, rightS.scalar, r, outV)) {
      return true;
    }
  } else if (tryGetExactNonNegativeUIntFromScalar(rightS, intU) && !scalarHasExactIntegerPayload(leftS)) {
    if (tryApplyExactIntFloatOpFromFloatSide(ScalarExactKind::Unsigned, true, 0LL, intU, leftS.scalar, r, outV)) {
      return true;
    }
  }

  SmInt intI = 0;
  if (tryGetExactSignedIntNoUIntWrapFromScalar(leftS, intI) && !scalarHasExactIntegerPayload(rightS)) {
    if (tryApplyExactIntFloatOpFromFloatSide(ScalarExactKind::Signed, true, intI, 0u, rightS.scalar, r, outV)) {
      return true;
    }
  } else if (tryGetExactSignedIntNoUIntWrapFromScalar(rightS, intI) && !scalarHasExactIntegerPayload(leftS)) {
    if (tryApplyExactIntFloatOpFromFloatSide(ScalarExactKind::Signed, true, intI, 0u, leftS.scalar, r, outV)) {
      return true;
    }
  }

  return false;
}

int tryApplyGcdLcmScalars(
    const EvalValue::ScalarValue& a,
    const EvalValue::ScalarValue& b,
    bool doLcm,
    EvalValue& outV) {
  SmUInt aU = 0;
  SmUInt bU = 0;
  if (tryGetExactNonNegativeUIntFromScalar(a, aU) && tryGetExactNonNegativeUIntFromScalar(b, bU)) {
    if (aU == 0u && bU == 0u) {
      return 2;
    }
    if (!doLcm) {
      outV = makeScalarUInt(gcdUInt(aU, bU));
      return 0;
    }
    SmUInt lU = 0;
    if (!tryLcmUInt(aU, bU, lU)) {
      const SmUInt g = gcdUInt(aU, bU);
      outV = makeScalar(static_cast<SmFloat>(aU / g) * static_cast<SmFloat>(bU));
      return 0;
    }
    outV = makeScalarUInt(lU);
    return 0;
  }
  bool aNeg = false;
  bool bNeg = false;
  if (!tryGetExactSignMagnitudeFromScalar(a, aNeg, aU) || !tryGetExactSignMagnitudeFromScalar(b, bNeg, bU)) {
    return 1;
  }
  if (aU == 0u && bU == 0u) {
    return 2;
  }
  if (!doLcm) {
    outV = makeScalarUIntPreferInt(gcdUInt(aU, bU));
    return 0;
  }
  SmUInt lU = 0;
  if (!tryLcmUInt(aU, bU, lU)) {
    const SmUInt g = gcdUInt(aU, bU);
    outV = makeScalar(static_cast<SmFloat>(aU / g) * static_cast<SmFloat>(bU));
    return 0;
  }
  outV = makeScalarUIntPreferInt(lU);
  return 0;
}

template <typename ScalarFn>
EvalValue mapBinaryBroadcast(
    const EvalValue& left,
    const EvalValue& right,
    ScalarFn&& scalarFn,
    bool& ok) {
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

EvalValue applyGcdLcmEvalValues(const EvalValue& a, const EvalValue& b, bool doLcm, int& status) {
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

void ddTwoSum(SmFloat a, SmFloat b, SmFloat& hi, SmFloat& lo) {
  hi = a + b;
  const SmFloat bb = hi - a;
  lo = (a - (hi - bb)) + (b - bb);
}

// SM_FIXED_WIDTH_BEGIN: exact product for IEEE binary64 (53-bit mantissa) using 64-bit integer arithmetic.
// Fixed widths on purpose; used only when SmFloat has a 53-bit mantissa (see ddMulExact).
void ddMulExactBinary64(double a, double b, double& hi, double& lo) {
  if (a == 0.0 || b == 0.0 || !std::isfinite(a) || !std::isfinite(b)) {
    hi = a * b;
    lo = 0.0;
    return;
  }
  const double sign = ((a < 0.0) != (b < 0.0)) ? -1.0 : 1.0;
  int ea = 0;
  int eb = 0;
  const double af = std::frexp(std::fabs(a), &ea);
  const double bf = std::frexp(std::fabs(b), &eb);
  const std::uint64_t ai = static_cast<std::uint64_t>(std::ldexp(af, 53));
  const std::uint64_t bi = static_cast<std::uint64_t>(std::ldexp(bf, 53));
  const std::uint64_t a0 = ai & 0xffffffffu;
  const std::uint64_t a1 = ai >> 32;
  const std::uint64_t b0 = bi & 0xffffffffu;
  const std::uint64_t b1 = bi >> 32;
  const std::uint64_t p00 = a0 * b0;
  const std::uint64_t p01 = a0 * b1;
  const std::uint64_t p10 = a1 * b0;
  const std::uint64_t p11 = a1 * b1;
  const std::uint64_t mid = (p00 >> 32) + (p01 & 0xffffffffu) + (p10 & 0xffffffffu);
  const std::uint64_t plo = (p00 & 0xffffffffu) | ((mid & 0xffffffffu) << 32);
  const std::uint64_t phi = p11 + (p01 >> 32) + (p10 >> 32) + (mid >> 32);
  auto bitLength64 = [](std::uint64_t x) {
    int n = 0;
    if ((x >> 32) != 0u) { n += 32; x >>= 32; }
    if ((x >> 16) != 0u) { n += 16; x >>= 16; }
    if ((x >> 8) != 0u) { n += 8; x >>= 8; }
    if ((x >> 4) != 0u) { n += 4; x >>= 4; }
    if ((x >> 2) != 0u) { n += 2; x >>= 2; }
    if ((x >> 1) != 0u) { n += 2; }
    else if (x != 0u) { n += 1; }
    return n;
  };
  const int bits = (phi != 0u) ? bitLength64(phi) + 64 : bitLength64(plo);
  int shift = bits > 53 ? bits - 53 : 0;
  std::uint64_t top = (shift == 0) ? plo : ((plo >> shift) | (phi << (64 - shift)));
  if (shift > 0) {
    const bool roundBit = ((plo >> (shift - 1)) & 1ull) != 0ull;
    const bool sticky = shift > 1 && (plo & ((1ull << (shift - 1)) - 1ull)) != 0ull;
    if (roundBit && (sticky || (top & 1ull) != 0ull)) {
      top += 1ull;
      if (top == (1ull << 53)) {
        top >>= 1;
        shift += 1;
      }
    }
  }
  const std::uint64_t subLo = top << shift;
  const std::uint64_t subHi = (shift == 0) ? 0ull : (top >> (64 - shift));
  const std::uint64_t remLo = plo - subLo;
  const std::uint64_t borrow = (plo < subLo) ? 1ull : 0ull;
  const std::uint64_t remHi = phi - subHi - borrow;
  long long remS = 0;
  if (remHi == 0ull) {
    remS = static_cast<long long>(remLo);
  } else if (remLo != 0ull) {
    remS = -static_cast<long long>((~remLo) + 1ull);
  }
  const int exp = ea + eb - 106 + shift;
  hi = std::ldexp(sign * static_cast<double>(top), exp);
  lo = (remS == 0) ? 0.0 : std::ldexp(sign * static_cast<double>(remS), exp - shift);
}
// SM_FIXED_WIDTH_END

// Exact product for any binary SmFloat (Veltkamp split + Dekker two-product): hi = RN(a*b), hi + lo == a*b.
void ddMulExactDekker(SmFloat a, SmFloat b, SmFloat& hi, SmFloat& lo) {
  hi = a * b;
  lo = static_cast<SmFloat>(0);
  if (a == 0 || b == 0 || !std::isfinite(a) || !std::isfinite(b) || !std::isfinite(hi)) {
    return;
  }
  const SmFloat splitFactor = smartmath::pow2Float((smartmath::kFloatMantDigits + 1) / 2) + static_cast<SmFloat>(1);
  const SmFloat ca = splitFactor * a;
  const SmFloat cb = splitFactor * b;
  if (!std::isfinite(ca) || !std::isfinite(cb)) {
    return;
  }
  const SmFloat ah = ca - (ca - a);
  const SmFloat al = a - ah;
  const SmFloat bh = cb - (cb - b);
  const SmFloat bl = b - bh;
  const SmFloat err = (((ah * bh - hi) + ah * bl) + al * bh) + al * bl;
  if (std::isfinite(err)) {
    lo = err;
  }
}

// Dispatch is a template so that the binary64 branch is only instantiated when SmFloat has a 53-bit mantissa.
template <typename F>
void ddMulExactBy(F a, F b, F& hi, F& lo, std::true_type) {
  double hd = 0.0;  // SM_FIXED_WIDTH
  double ld = 0.0;  // SM_FIXED_WIDTH
  ddMulExactBinary64(static_cast<double>(a), static_cast<double>(b), hd, ld);  // SM_FIXED_WIDTH
  hi = static_cast<F>(hd);
  lo = static_cast<F>(ld);
}

template <typename F>
void ddMulExactBy(F a, F b, F& hi, F& lo, std::false_type) {
  ddMulExactDekker(a, b, hi, lo);
}

void ddMulExact(SmFloat a, SmFloat b, SmFloat& hi, SmFloat& lo) {
  ddMulExactBy(a, b, hi, lo, std::integral_constant<bool, smartmath::kFloatHasBinary64Mantissa>());
}

void ddScale(SmFloat& hi, SmFloat& lo, SmFloat b) {
  SmFloat p1 = 0.0;
  SmFloat p0 = 0.0;
  ddMulExact(hi, b, p1, p0);
  if (!std::isfinite(p1)) {
    hi = p1;
    lo = 0.0;
    return;
  }
  SmFloat q1 = 0.0;
  SmFloat q0 = 0.0;
  ddMulExact(lo, b, q1, q0);
  SmFloat s1 = 0.0;
  SmFloat s0 = 0.0;
  ddTwoSum(p0, q1, s1, s0);
  SmFloat t1 = 0.0;
  SmFloat t0 = 0.0;
  ddTwoSum(s1, q0, t1, t0);
  s0 += t0;
  SmFloat r1 = 0.0;
  SmFloat r0 = 0.0;
  ddTwoSum(p1, t1, r1, r0);
  r0 += s0;
  ddTwoSum(r1, r0, hi, lo);
}

void ddDiv(SmFloat& hi, SmFloat& lo, SmFloat b) {
  const SmFloat qhi = hi / b;
  SmFloat p1 = 0.0;
  SmFloat p0 = 0.0;
  ddMulExact(qhi, b, p1, p0);
  const SmFloat rem = ((hi - p1) - p0) + lo;
  ddTwoSum(qhi, rem / b, hi, lo);
}

SmFloat nprProductFloat(SmFloat n, SmUInt r) {
  SmFloat hi = 1.0;
  SmFloat lo = 0.0;
  for (SmUInt i = 0; i < r; ++i) {
    ddScale(hi, lo, n - static_cast<SmFloat>(i));
    if (!std::isfinite(hi)) {
      break;
    }
  }
  return hi;
}

SmFloat ncrProductFloat(SmFloat n, SmUInt r) {
  SmFloat hi = 1.0;
  SmFloat lo = 0.0;
  const SmFloat rd = static_cast<SmFloat>(r);
  for (SmUInt i = 1; i <= r; ++i) {
    const SmFloat factor = n - rd + static_cast<SmFloat>(i);
    int eh = 0;
    int ef = 0;
    (void)std::frexp(hi, &eh);
    (void)std::frexp(factor, &ef);
    int scale = 0;
    if (eh + ef > smartmath::kFloatMaxExp2 - 44) {
      scale = eh + ef - (smartmath::kFloatMaxExp2 - 124);
      hi = std::ldexp(hi, -scale);
      lo = std::ldexp(lo, -scale);
    }
    ddScale(hi, lo, factor);
    if (!std::isfinite(hi)) {
      break;
    }
    ddDiv(hi, lo, static_cast<SmFloat>(i));
    if (scale != 0) {
      hi = std::ldexp(hi, scale);
      lo = std::ldexp(lo, scale);
    }
    if (!std::isfinite(hi)) {
      break;
    }
  }
  return hi;
}

bool tryApplyNcrNprScalar(
    const EvalValue::ScalarValue& nS,
    const EvalValue::ScalarValue& rS,
    bool doPerm,
    EvalValue& outS,
    int& status) {
  bool nNeg = false;
  bool rNeg = false;
  SmUInt nU = 0;
  SmUInt rU = 0;
  if (!tryGetExactSignMagnitudeFromScalar(nS, nNeg, nU) || !tryGetExactSignMagnitudeFromScalar(rS, rNeg, rU)) {
    status = 1;
    return false;
  }
  if (rNeg) {
    status = 2;
    return false;
  }
  if (rU == 0u) {
    status = 0;
    outS = makeScalarInt(1);
    return true;
  }
  if (!nNeg && rU > nU) {
    status = 0;
    outS = makeScalarInt(0);
    return true;
  }
  if (!nNeg) {
    SmUInt exact = 0;
    if (doPerm ? tryComputeNprUInt(nU, rU, exact) : tryComputeNcrUInt(nU, rU, exact)) {
      status = 0;
      outS = makeScalarUIntPreferInt(exact);
      return true;
    }
    status = 0;
    const SmFloat nD = static_cast<SmFloat>(nU);
    if (doPerm) {
      outS = makeScalar(nprProductFloat(nD, rU));
    } else {
      SmUInt rEff = rU;
      if ((nU - rU) < rEff) {
        rEff = nU - rU;
      }
      outS = makeScalar(ncrProductFloat(nD, rEff));
    }
    return true;
  }
  if (doPerm) {
    if (nU <= static_cast<SmUInt>(kIntMax)) {
      SmInt acc = 1;
      SmInt term = -static_cast<SmInt>(nU);
      bool exact = true;
      for (SmUInt i = 0; i < rU; ++i) {
        if (!tryMulSIntChecked(acc, term, acc)) {
          exact = false;
          break;
        }
        if (i + 1u == rU) {
          break;
        }
        if (term == kIntMin) {
          exact = false;
          break;
        }
        --term;
      }
      if (exact) {
        status = 0;
        outS = makeScalarInt(acc);
        return true;
      }
    }
    status = 0;
    outS = makeScalar(nprProductFloat(-static_cast<SmFloat>(nU), rU));
    return true;
  }
  if (rU <= kUIntMax - nU) {
    const SmUInt top = nU + rU - 1u;
    SmUInt comb = 0;
    if (tryComputeNcrUInt(top, rU, comb)) {
      status = 0;
      if ((rU & 1u) == 0u) {
        outS = makeScalarUIntPreferInt(comb);
      } else if (comb <= static_cast<SmUInt>(kIntMax)) {
        outS = makeScalarInt(-static_cast<SmInt>(comb));
      } else {
        outS = makeScalar(-static_cast<SmFloat>(comb));
      }
      return true;
    }
  }
  status = 0;
  outS = makeScalar(ncrProductFloat(-static_cast<SmFloat>(nU), rU));
  return true;
}

EvalValue applyNcrNprEvalValues(
    const EvalValue& n,
    const EvalValue& r,
    bool doPerm,
    int& status) {
  bool ok = true;
  const auto applyPair = [&](const EvalValue::ScalarValue& nS, const EvalValue::ScalarValue& rS,
                             EvalValue& outS) -> bool {
    return tryApplyNcrNprScalar(nS, rS, doPerm, outS, status);
  };
  EvalValue result = mapBinaryBroadcast(n, r, applyPair, ok);
  if (!ok && status == 0) {
    status = 3;
  }
  return result;
}

bool tryApplyModScalars(
    EvalContext& ctx,
    const std::string& fnName,
    const EvalValue::ScalarValue& aS,
    const EvalValue::ScalarValue& bS,
    EvalValue& outS) {
  if (aS.hasExactUInt() && bS.hasExactUInt() && (!aS.hasExactSInt() || !bS.hasExactSInt())) {
    if (bS.exactUInt == 0u) {
      ctx.failCall(CallErrorId::Numeric, fnName);
      return false;
    }
    outS = makeScalarUInt(aS.exactUInt % bS.exactUInt);
    return true;
  }
  bool aNeg = false;
  bool bNeg = false;
  SmUInt aMag = 0;
  SmUInt bMag = 0;
  if (!tryGetExactSignMagnitudeFromScalar(aS, aNeg, aMag) || !tryGetExactSignMagnitudeFromScalar(bS, bNeg, bMag)) {
    ctx.failCall(CallErrorId::IntegerValues, fnName);
    return false;
  }
  if (bMag == 0u) {
    ctx.failCall(CallErrorId::Numeric, fnName);
    return false;
  }
  outS = makeScalarFromSignMagnitude(aNeg, aMag % bMag);
  return true;
}

EvalValue mapBinaryBuiltinMathFunction(
    const EvalValue& left,
    const EvalValue& right,
    BuiltinFunctionId id,
    bool& ok,
    bool supportComplexNumbers) {
  ok = true;
  const bool isLog = (id == BuiltinFunctionId::Log);
  SmFloat (*binaryFn)(SmFloat, SmFloat) = nullptr;
  if (!isLog) {
    if (id == BuiltinFunctionId::Atan2) {
      binaryFn = std::atan2;
    } else if (id == BuiltinFunctionId::Hypot) {
      binaryFn = calcHypotBasic;
    } else {
      ok = false;
      return makeScalar(0);
    }
  }

  auto applyScalarValue = [&](const EvalValue::ScalarValue& l, const EvalValue::ScalarValue& r, EvalValue& outS) -> bool {
#if SMARTMATH_COMPLEX_NUMBERS
    if (supportComplexNumbers && id == BuiltinFunctionId::Hypot &&
        (scalarHasNonzeroImaginaryPart(l) || scalarHasNonzeroImaginaryPart(r))) {
      SmFloat ar = 0.0;
      SmFloat ai = 0.0;
      SmFloat br = 0.0;
      SmFloat bi = 0.0;
      scalarLoadCartesian(l, ar, ai);
      scalarLoadCartesian(r, br, bi);
      if (!std::isfinite(ar) || !std::isfinite(ai) || !std::isfinite(br) || !std::isfinite(bi)) {
        outS = makeScalar(std::numeric_limits<SmFloat>::quiet_NaN());
        return true;
      }
      const SmFloat ml = std::hypot(ar, ai);
      const SmFloat mr = std::hypot(br, bi);
      outS = makeScalarMaybeExact(std::hypot(ml, mr));
      return true;
    }
#endif
    SmFloat outD = 0.0;
    if (isLog) {
#if SMARTMATH_COMPLEX_NUMBERS
      if (!supportComplexNumbers) {
#endif
        if (l.scalar <= 0.0 || r.scalar <= 0.0 || r.scalar == 1.0) {
          return false;
        }
        outD = std::log(l.scalar) / std::log(r.scalar);
        outS = makeScalarMaybeExact(outD);
        return true;
#if SMARTMATH_COMPLEX_NUMBERS
      }
      SmFloat ar = 0.0;
      SmFloat ai = 0.0;
      SmFloat br = 0.0;
      SmFloat bi = 0.0;
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
        outS = makeScalarComplexFromFloats(
            std::numeric_limits<SmFloat>::quiet_NaN(), std::numeric_limits<SmFloat>::quiet_NaN());
        return true;
      }
      // Same domain errors as real-only log: log of 0, base 0, base 1.
      if ((ar == 0.0 && ai == 0.0) || (br == 0.0 && bi == 0.0) || (br == 1.0 && bi == 0.0)) {
        return false;
      }
      SmFloat lnr = 0.0;
      SmFloat lni = 0.0;
      SmFloat rnr = 0.0;
      SmFloat rni = 0.0;
      scalarPrincipalLnCartesian(ar, ai, lnr, lni);
      scalarPrincipalLnCartesian(br, bi, rnr, rni);
      const SmFloat den = rnr * rnr + rni * rni;
      if (den == 0.0) {
        outS = makeScalarComplexFromFloats(
            std::numeric_limits<SmFloat>::quiet_NaN(), std::numeric_limits<SmFloat>::quiet_NaN());
        return true;
      }
      SmFloat outRe = (lnr * rnr + lni * rni) / den;
      SmFloat outIm = (lni * rnr - lnr * rni) / den;
      snapComplexNearZeroAxis(outRe, outIm);
      outS = makeScalarComplexFromFloats(outRe, outIm);
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

EvalValue builtinApplyClamp(
    EvalContext& ctx,
    const EvalValue& valueV,
    const EvalValue& minV,
    const EvalValue& maxV,
    bool supportComplexNumbers,
    bool supportTimeValues) {
  if (minV.kind != ValueKind::Scalar || maxV.kind != ValueKind::Scalar) {
    ctx.failCall(CallErrorId::ScalarMinMax, getFunctionName(BuiltinFunctionId::Clamp));
    return makeScalar(0);
  }
#if SMARTMATH_TIME_VALUES
  auto scalarIsTimeKind = [](const EvalValue::ScalarValue& s) {
    return s.scalarKind == ScalarKind::Time;
  };
  bool timeHit = false;
  if (supportTimeValues) {
    if (valueV.kind != ValueKind::Scalar) {
      for (const auto& item : valueV.arr) {
        if (scalarIsTimeKind(item)) {
          timeHit = true;
          break;
        }
      }
    } else {
      timeHit = scalarIsTimeKind(valueV.scalarValue);
    }
    timeHit = timeHit || scalarIsTimeKind(minV.scalarValue) || scalarIsTimeKind(maxV.scalarValue);
  }
  if (timeHit) {
    ctx.fail(ParseErrorId::IncompatibleOperands);
    return makeScalar(0);
  }
#endif
#if SMARTMATH_COMPLEX_NUMBERS
  if (supportComplexNumbers &&
      (evalValueHasNonzeroImaginary(valueV) || evalValueHasNonzeroImaginary(minV) ||
          evalValueHasNonzeroImaginary(maxV))) {
    ctx.fail(ParseErrorId::IncompatibleOperands);
    return makeScalar(0);
  }
#endif
  if (valueV.kind != ValueKind::Scalar && valueV.kind != ValueKind::Array) {
    ctx.failCall(CallErrorId::Numeric, getFunctionName(BuiltinFunctionId::Clamp));
    return makeScalar(0);
  }
  const EvalValue::ScalarValue& minSv = minV.scalarValue;
  const EvalValue::ScalarValue& maxSv = maxV.scalarValue;
  return mapUnaryEvalValue(valueV, [&minSv, &maxSv](const EvalValue::ScalarValue& s) -> EvalValue {
    return scalarFromScalarValue(clampScalarValue(s, minSv, maxSv));
  });
}

// Builtins are lowercase a-z except names with digits (atan2, log10). Skip hash lookup for other idents with digit/_ (e.g. myconst1).
bool identMayBeBareBuiltinName(const char* begin, const char* end) {
  if (!identHasDigitOrUnderscore(begin, end)) {
    return true;
  }
  std::string low;
  assignLowerIdentFromRange(low, begin, end);
  BuiltinFunctionId bid = BuiltinFunctionId::Count;
  return tryGetBuiltinFunctionId(low, bid);
}

bool identMayBeBareBuiltinName(const std::string& ident) {
  return identMayBeBareBuiltinName(ident.data(), ident.data() + ident.size());
}

bool tryPowVerifyRootExactValue(SmInt valueInt, SmInt rootCand, SmInt rootDeg) {
  if (rootDeg < 2) {
    return false;
  }
  if (valueInt < 0 && (rootDeg % 2) == 0) {
    return false;
  }
  if (rootCand >= 0 && valueInt >= 0) {
    SmUInt reconU = 0;
    if (!tryPowUIntChecked(static_cast<SmUInt>(rootCand), static_cast<SmUInt>(rootDeg), reconU)) {
      return false;
    }
    return static_cast<SmInt>(reconU) == valueInt;
  }
  SmInt reconI = 0;
  if (!tryPowSIntChecked(rootCand, rootDeg, reconI)) {
    return false;
  }
  return reconI == valueInt;
}

// CPython-style length of range(start, stop, step) using unsigned wraparound arithmetic.
SmUInt estimateRangeItemCount(SmInt startV, SmInt stopV, SmInt stepV) {
  SmUInt uLo = 0;
  SmUInt uHi = 0;
  SmUInt uStep = 0;
  if (stepV > 0) {
    if (startV >= stopV) {
      return 0;
    }
    uLo = static_cast<SmUInt>(startV);
    uHi = static_cast<SmUInt>(stopV);
    uStep = static_cast<SmUInt>(stepV);
  } else {
    if (startV <= stopV) {
      return 0;
    }
    uLo = static_cast<SmUInt>(stopV);
    uHi = static_cast<SmUInt>(startV);
    uStep = 0ull - static_cast<SmUInt>(stepV);
  }
  return (uHi - uLo - 1ull) / uStep + 1ull;
}

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

SmFloat MedianFromFloatFlatInPlace(std::vector<SmFloat>& flat) {
  const std::size_t n = flat.size();
  const std::size_t mid = n / 2U;
  std::nth_element(flat.begin(), flat.begin() + static_cast<std::ptrdiff_t>(mid), flat.end());
  const SmFloat upper = flat[mid];
  if (n % 2U == 1U) {
    return upper;
  }
  const SmFloat lower = *std::max_element(flat.begin(), flat.begin() + static_cast<std::ptrdiff_t>(mid));
  return (lower + upper) / 2.0;
}

inline SmFloat ratioApproxAbsNumerator(SmFloat v, SmInt p, SmUInt q) {
  return std::fabs(v * static_cast<SmFloat>(q) - static_cast<SmFloat>(p));
}

inline bool ratioApproxErrLess(SmFloat n1, SmUInt q1, SmFloat n2, SmUInt q2) {
  return n1 * static_cast<SmFloat>(q2) < n2 * static_cast<SmFloat>(q1);
}

struct RatioRationalSearch {
  const SmFloat v;
  SmInt bestP = 0;
  SmUInt bestQ = 1;
  SmFloat bestAbsNum = kHugeSentinel;

  explicit RatioRationalSearch(SmFloat vIn) : v(vIn) {}

  bool consider(SmInt p, SmUInt q) {
    if (q == 0 || q > static_cast<SmUInt>(RATIO_MAX_DENOMINATOR)) {
      return false;
    }
    const SmFloat absNum = ratioApproxAbsNumerator(v, p, q);
    if (ratioApproxErrLess(absNum, q, bestAbsNum, bestQ)) {
      bestAbsNum = absNum;
      bestP = p;
      bestQ = q;
      return absNum == 0.0;
    }
    return false;
  }

  bool semiconvergentAbsNumLess(SmInt p1, SmInt q1, SmInt p0, SmInt q0, SmInt h1,
                                SmInt h2) const {
    const SmInt qh1 = q1 + h1 * q0;
    const SmInt qh2 = q1 + h2 * q0;
    if (qh1 <= 0 || qh2 <= 0) {
      return false;
    }
    const SmInt ph1 = p1 + h1 * p0;
    const SmInt ph2 = p1 + h2 * p0;
    const SmFloat n1 = ratioApproxAbsNumerator(v, ph1, static_cast<SmUInt>(qh1));
    const SmFloat n2 = ratioApproxAbsNumerator(v, ph2, static_cast<SmUInt>(qh2));
    return ratioApproxErrLess(n1, static_cast<SmUInt>(qh1), n2, static_cast<SmUInt>(qh2));
  }

  void scanSemiconvergentRange(SmInt p1, SmInt q1, SmInt p0, SmInt q0, SmInt hMax) {
    if (hMax < 0) {
      return;
    }
    if (hMax <= RATIO_SEMICONV_LINEAR_THRESH) {
      for (SmInt h = 0; h <= hMax; ++h) {
        if (consider(p1 + h * p0, static_cast<SmUInt>(q1 + h * q0))) {
          return;
        }
      }
      return;
    }
    SmInt lo = 0;
    SmInt hi = hMax;
    if (consider(p1, static_cast<SmUInt>(q1))) {
      return;
    }
    while (hi - lo > 2) {
      const SmInt span = hi - lo;
      const SmInt m1 = lo + span / 3;
      const SmInt m2 = hi - span / 3;
      if (semiconvergentAbsNumLess(p1, q1, p0, q0, m1, m2)) {
        hi = m2;
      } else {
        lo = m1;
      }
    }
    for (SmInt h = lo; h <= hi; ++h) {
      if (consider(p1 + h * p0, static_cast<SmUInt>(q1 + h * q0))) {
        return;
      }
    }
  }
};

bool tryExactPower10Rational(SmFloat v, RatioRationalSearch& search) {
  for (int k = 1; k <= RATIO_MAX_POWER10_EXP; ++k) {
    const SmUInt denPow = pow10UInt(k);
    const SmFloat scaled = v * static_cast<SmFloat>(denPow);
    const SmInt n = smartmath::roundFloatToSInt(scaled);
    if (n == 0) {
      continue;
    }
    SmFloat scaleErr = RATIO_APPROX_EPS * std::fabs(scaled);
    if (scaleErr < RATIO_APPROX_EPS) {
      scaleErr = RATIO_APPROX_EPS;
    }
    if (std::fabs(scaled - static_cast<SmFloat>(n)) > scaleErr) {
      continue;
    }
    const SmInt denPowLong = static_cast<SmInt>(denPow);
    const SmInt g = GcdSInt(n, denPowLong);
    const SmInt candNum = n / g;
    const SmUInt candDen = static_cast<SmUInt>(denPowLong / g);
    if (candDen == 0) {
      continue;
    }
    const SmFloat absNum = ratioApproxAbsNumerator(v, candNum, candDen);
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

bool tryApproximateRational(SmFloat x, SmInt& num, SmUInt& den) {
  if (std::isnan(x) || std::isinf(x)) {
    return false;
  }
  const bool neg = x < 0.0;
  const SmFloat v = std::fabs(x);
  SmFloat scale = v;
  if (scale < 1.0) {
    scale = 1.0;
  }
  const SmFloat tol = RATIO_APPROX_EPS * scale;
  if (v < tol) {
    num = 0;
    den = 1;
    return true;
  }
  RatioRationalSearch search(v);
  SmInt p0 = 0;
  SmInt p1 = 1;
  SmInt q0 = 1;
  SmInt q1 = 0;
  SmFloat frac = v;
  for (int i = 0; i < RATIO_MAX_CF_TERMS; ++i) {
    SmFloat ipart = 0.0;
    const SmFloat t = std::modf(frac, &ipart);
    const SmInt a = static_cast<SmInt>(ipart);
    const SmInt p = a * p1 + p0;
    const SmInt q = a * q1 + q0;
    if (q > RATIO_MAX_DENOMINATOR) {
      if (q1 > 0 && q0 > 0) {
        const SmInt hMax = (RATIO_MAX_DENOMINATOR - q1) / q0;
        search.scanSemiconvergentRange(p1, q1, p0, q0, hMax);
      }
      break;
    }
    if (search.consider(p, static_cast<SmUInt>(q))) {
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
  if (v <= 1.0 / static_cast<SmFloat>(RATIO_MAX_DENOMINATOR)) {
    tryExactPower10Rational(v, search);
  }
  if (search.bestP == 0) {
    num = 0;
    den = 1;
    return true;
  }
  const SmInt g = GcdSInt(search.bestP, static_cast<SmInt>(search.bestQ));
  num = search.bestP / g;
  den = static_cast<SmUInt>(static_cast<SmInt>(search.bestQ) / g);
  if (neg) {
    num = -num;
  }
  return true;
}

void setBinaryBuiltinBroadcastFailure(
    EvalContext& ctx,
    const std::string& fnName,
    const EvalValue& left,
    const EvalValue& right,
    int pairStatus) {
  if (pairStatus == 1) {
    ctx.failCall(CallErrorId::IntegerValues, fnName);
    return;
  }
  const bool mismatchedLengths = left.kind == ValueKind::Array &&
                                 right.kind == ValueKind::Array &&
                                 left.arr.size() != right.arr.size();
  if (pairStatus == 3 || mismatchedLengths) {
    ctx.fail(ParseErrorId::IncompatibleOperands);
    return;
  }
  if (!ctx.parseError && ctx.errorText.empty()) {
    ctx.failCall(CallErrorId::Numeric, fnName);
  }
}

bool isReservedBuiltinConstantName(const std::string& nameText, bool supportTimeValues) {
  if (nameText == STR_PI || nameText == STR_E || nameText == STR_INF || nameText == STR_NAN) {
    return true;
  }
  return supportTimeValues &&
         (nameText == STR_MILLISECOND || nameText == STR_SECOND || nameText == STR_MINUTE ||
          nameText == STR_HOUR || nameText == STR_DAY);
}

bool isReservedBuiltinVariableNameForUserFunctionDefinition(const std::string& fnName) {
  std::string lowered = fnName;
  std::transform(lowered.begin(), lowered.end(), lowered.begin(), [](unsigned char c) {
    return static_cast<char>(std::tolower(c));
  });
  return lowered == STR_ANS;
}

bool isPercentPostfixContext(const char* p) {
  if (*p != '%') {
    return false;
  }
  const char* q = p + 1;
  while (*q && std::isspace(static_cast<unsigned char>(*q))) {
    ++q;
  }
  if (*q == '\0') return true;
  if (*q == ')' || *q == ']' || *q == '}' || *q == ',' || *q == ';') return true;
  if (*q == '+' || *q == '-' || *q == '*' || *q == '/' || *q == '^') return true;
  if (*q == '<' || *q == '>' || *q == '=' || *q == '!') return true;
  if (*q == '&' || *q == '|') return true;
  return false;
}

bool percentTailStartsPower(const char* percentPtr) {
  if (!isPercentPostfixContext(percentPtr)) {
    return false;
  }
  const char* q = percentPtr + 1;
  while (*q && std::isspace(static_cast<unsigned char>(*q))) {
    ++q;
  }
  return q[0] == '*' && q[1] == '*';
}

}  // namespace

namespace smartmath_internal {

#if SMARTMATH_COMPLEX_NUMBERS
void lnOfTwice(SmFloat ar, SmFloat ai, SmFloat& outR, SmFloat& outI) {
  scalarPrincipalLnCartesian(ar, ai, outR, outI);
  outR += std::log(2.0);
}

bool magnitudeAtLeast(SmFloat ar, SmFloat ai, SmFloat limit) {
  return std::fabs(ar) >= limit || std::fabs(ai) >= limit;
}

bool complexUnaryTrigCartesian(BuiltinFunctionId id, SmFloat ar, SmFloat ai, SmFloat& outR, SmFloat& outI) {
  const SmFloat nanv = std::numeric_limits<SmFloat>::quiet_NaN();
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
      SmFloat numR = 0.0;
      SmFloat numI = 0.0;
      SmFloat denR = 0.0;
      SmFloat denI = 0.0;
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
      const SmFloat z2r = ar * ar - ai * ai;
      const SmFloat z2i = 2.0 * ar * ai;
      SmFloat sqrR = 0.0;
      SmFloat sqrI = 0.0;
      complexPrincipalSqrt(z2r + 1.0, z2i, sqrR, sqrI);
      scalarPrincipalLnCartesian(ar + sqrR, ai + sqrI, outR, outI);
      break;
    }
    case BuiltinFunctionId::Acosh: {
      if (ai == 0.0 && std::isfinite(ar)) {
        if (ar <= -1.0) {
          outR = std::acosh(-ar);
          outI = kPi;
          break;
        }
        if (ar >= 1.0) {
          outR = std::acosh(ar);
          outI = 0.0;
          break;
        }
      }
      if (std::isfinite(ar) && std::isfinite(ai) && magnitudeAtLeast(ar, ai, 1e8)) {
        lnOfTwice(ar, ai, outR, outI);
        break;
      }
      SmFloat s1r = 0.0;
      SmFloat s1i = 0.0;
      SmFloat s2r = 0.0;
      SmFloat s2i = 0.0;
      complexPrincipalSqrt(ar - 1.0, ai, s1r, s1i);
      complexPrincipalSqrt(ar + 1.0, ai, s2r, s2i);
      const SmFloat wr = s1r * s2r - s1i * s2i;
      const SmFloat wi = s1r * s2i + s1i * s2r;
      scalarPrincipalLnCartesian(ar + wr, ai + wi, outR, outI);
      break;
    }
    case BuiltinFunctionId::Atanh: {
      SmFloat quotR = 0.0;
      SmFloat quotI = 0.0;
      complexDivide(1.0 + ar, ai, 1.0 - ar, -ai, quotR, quotI);
      SmFloat lnR = 0.0;
      SmFloat lnI = 0.0;
      scalarPrincipalLnCartesian(quotR, quotI, lnR, lnI);
      outR = lnR * 0.5;
      outI = lnI * 0.5;
      break;
    }
    case BuiltinFunctionId::Asin: {
      if (ai == 0.0 && std::isfinite(ar) && std::fabs(ar) > 1.0) {
        const SmFloat ah = std::acosh(std::fabs(ar));
        const SmFloat sgn = (ar < 0.0) ? -1.0 : 1.0;
        outR = sgn * (kPi * 0.5);
        outI = -sgn * ah;
        break;
      }
      if (std::isfinite(ar) && std::isfinite(ai) && magnitudeAtLeast(ar, ai, kHypotDirectMax)) {
        const bool neg = ar < 0.0;
        const SmFloat br = neg ? -ar : ar;
        const SmFloat bi = neg ? -ai : ai;
        SmFloat lr = 0.0;
        SmFloat li = 0.0;
        lnOfTwice(br, bi, lr, li);
        outR = kPi * 0.5 + li;
        outI = -lr;
        if (neg) {
          outR = -outR;
          outI = -outI;
        }
        break;
      }
      const SmFloat izR = -ai;
      const SmFloat izI = ar;
      const SmFloat z2r = ar * ar - ai * ai;
      const SmFloat z2i = 2.0 * ar * ai;
      SmFloat sqrR = 0.0;
      SmFloat sqrI = 0.0;
      complexPrincipalSqrt(1.0 - z2r, -z2i, sqrR, sqrI);
      SmFloat lnR = 0.0;
      SmFloat lnI = 0.0;
      scalarPrincipalLnCartesian(izR + sqrR, izI + sqrI, lnR, lnI);
      outR = lnI;
      outI = -lnR;
      break;
    }
    case BuiltinFunctionId::Acos: {
      SmFloat asR = 0.0;
      SmFloat asI = 0.0;
      if (!complexUnaryTrigCartesian(BuiltinFunctionId::Asin, ar, ai, asR, asI)) {
        return false;
      }
      outR = kPi / 2.0 - asR;
      outI = -asI;
      break;
    }
    case BuiltinFunctionId::Atan: {
      const SmFloat izR = -ai;
      const SmFloat izI = ar;
      SmFloat ln1mR = 0.0;
      SmFloat ln1mI = 0.0;
      SmFloat ln1pR = 0.0;
      SmFloat ln1pI = 0.0;
      scalarPrincipalLnCartesian(1.0 - izR, -izI, ln1mR, ln1mI);
      scalarPrincipalLnCartesian(1.0 + izR, izI, ln1pR, ln1pI);
      const SmFloat dR = ln1mR - ln1pR;
      const SmFloat dI = ln1mI - ln1pI;
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

bool nearlyInt(SmFloat v, SmInt& out) {
  if (!std::isfinite(v)) {
    return false;
  }
  if (std::fabs(v) > kMaxExactIntFromFloat) {
    return false;
  }
  const SmFloat r = std::round(v);
  if (std::fabs(v - r) > kTolLost3) {
    return false;
  }
  if (!isFloatInSIntRange(r)) {
    return false;
  }
  out = static_cast<SmInt>(r);
  return true;
}

bool tryGetExactSignedIntFromScalarPolicy(
    const EvalValue::ScalarValue& sIn,
    SmInt& outI,
    ExactSignedIntPolicy policy) {
  if (sIn.scalarKind == ScalarKind::Time) {
    return false;
  }
#if SMARTMATH_COMPLEX_NUMBERS
  if (scalarHasNonzeroImaginaryPart(sIn)) {
    return false;
  }
#endif
  EvalValue::ScalarValue s = sIn;
  if (policy == ExactSignedIntPolicy::NoUIntWrapWithRepairAndStorageKind) {
    scalarRepairExactMetadata(s);
  }
  if (s.hasExactSInt()) {
    outI = s.exactSInt;
    return true;
  }
  if (s.hasExactUInt()) {
    if (policy == ExactSignedIntPolicy::AllowUIntBitReinterpret) {
      static_assert(sizeof(SmInt) == sizeof(SmUInt), "SmInt/SmUInt size mismatch");
      SmUInt bits = s.exactUInt;
      std::memcpy(&outI, &bits, sizeof(bits));
      return true;
    }
    if (s.exactUInt <= static_cast<SmUInt>(kIntMax)) {
      outI = static_cast<SmInt>(s.exactUInt);
      return true;
    }
    return false;
  }
  if (policy == ExactSignedIntPolicy::NoUIntWrapWithRepairAndStorageKind) {
    if (s.scalarKind == ScalarKind::SInt) {
      outI = s.exactSInt;
      return true;
    }
    if (s.scalarKind == ScalarKind::UInt &&
        s.exactUInt <= static_cast<SmUInt>(kIntMax)) {
      outI = static_cast<SmInt>(s.exactUInt);
      return true;
    }
  }
  return false;
}

#if SMARTMATH_COMPLEX_NUMBERS
bool tryGetExactImagSIntStrict(const EvalValue::ScalarValue& s, SmInt& out) {
  if (s.hasImagExactSInt()) {
    out = s.imagExactSInt;
    return true;
  }
  return tryExtractExactSIntFromFloatStrict(s.imag, out);
}
#endif

bool scalarHasExactIntegerPayload(const EvalValue::ScalarValue& sv) {
  if (sv.scalarKind == ScalarKind::Time) {
    return false;
  }
  return sv.hasExactSInt() || sv.hasExactUInt() || sv.scalarKind == ScalarKind::SInt ||
         sv.scalarKind == ScalarKind::UInt;
}

bool TryMulExactSIntSquare(SmInt i, SmInt& outSq) {
  SmUInt u = 0;
  if (i >= 0) {
    u = static_cast<SmUInt>(i);
  } else if (i == kIntMin) {
    return false;
  } else {
    u = static_cast<SmUInt>(-i);
  }
  SmUInt sqU = 0;
  if (!tryMulUIntChecked(u, u, sqU) ||
      sqU > static_cast<SmUInt>(kIntMax)) {
    return false;
  }
  outSq = static_cast<SmInt>(sqU);
  return true;
}

bool TryGetExactSignedIntNoUIntWrapScalarStrict(const EvalValue::ScalarValue& s, SmInt& outI) {
  return tryGetExactSignedIntFromScalarPolicy(
      s, outI, ExactSignedIntPolicy::NoUIntWrapWithRepairAndStorageKind);
}

ArrayIndexIntKind classifyArrayIndexInt(const EvalValue::ScalarValue& sIn, SmInt& outI) {
#if SMARTMATH_COMPLEX_NUMBERS
  if (scalarHasNonzeroImaginaryPart(sIn)) {
    return ArrayIndexIntKind::NotInteger;
  }
#endif
  EvalValue::ScalarValue s = sIn;
  scalarRepairExactMetadata(s);
  if (s.scalarKind == ScalarKind::SInt) {
    outI = s.exactSInt;
    return ArrayIndexIntKind::Ok;
  }
  if (s.scalarKind == ScalarKind::UInt) {
    if (s.exactUInt > static_cast<SmUInt>(kIntMax)) {
      return ArrayIndexIntKind::OutOfSignedRange;
    }
    outI = static_cast<SmInt>(s.exactUInt);
    return ArrayIndexIntKind::Ok;
  }
  return ArrayIndexIntKind::NotInteger;
}

bool tryGetArrayIndexIntFromScalarValue(const EvalValue::ScalarValue& sIn, SmInt& outI) {
  return classifyArrayIndexInt(sIn, outI) == ArrayIndexIntKind::Ok;
}

void applySqrtScalarValue(const EvalValue::ScalarValue& sv, EvalValue& outV) {
  EvalValue::ScalarValue s = sv;
  scalarRepairExactMetadata(s);
  const SmFloat r = std::sqrt(scalarNumericReal(s));
  if (!scalarHasExactIntegerPayload(s)) {
    outV = makeScalar(r);
    return;
  }
  if (!std::isfinite(r)) {
    outV = makeScalar(r);
    return;
  }
  SmUInt inpU = 0;
  if (!tryGetExactNonNegativeUIntFromScalar(s, inpU)) {
    outV = makeScalar(r);
    return;
  }
  const SmUInt n = static_cast<SmUInt>(std::round(r));
  SmUInt sq = 0;
  if (tryMulUIntChecked(n, n, sq) && sq == inpU) {
    outV = makeScalarUIntPreferInt(n);
  } else {
    outV = makeScalar(r);
  }
}

bool tryApplySqrExactScalar(const EvalValue::ScalarValue& sv, EvalValue& outV) {
  EvalValue::ScalarValue s = sv;
  scalarRepairExactMetadata(s);
  if (!scalarHasExactIntegerPayload(s)) {
    return false;
  }
  SmInt i = 0;
  if (TryGetExactSignedIntNoUIntWrapScalarStrict(s, i)) {
    SmInt sq = 0;
    if (TryMulExactSIntSquare(i, sq)) {
      outV = makeScalarInt(sq);
      return true;
    }
  } else {
    SmUInt u = 0;
    if (tryGetExactNonNegativeUIntFromScalar(s, u)) {
      SmUInt sqU = 0;
      if (tryMulUIntChecked(u, u, sqU)) {
        outV = makeScalarUIntPreferInt(sqU);
        return true;
      }
    }
  }
  return false;
}

bool tryApplyHypotExactScalars(
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
  SmInt la = 0;
  SmInt lb = 0;
  if (TryGetExactSignedIntNoUIntWrapScalarStrict(l, la) &&
      TryGetExactSignedIntNoUIntWrapScalarStrict(r, lb)) {
    SmInt aa = 0;
    SmInt bb = 0;
    SmInt sumSq = 0;
    if (TryMulExactSIntSquare(la, aa) && TryMulExactSIntSquare(lb, bb) && checkedAddLL(aa, bb, sumSq)) {
      outV = makeScalar(std::sqrt(static_cast<SmFloat>(sumSq)));
      return true;
    }
    return false;
  }
  SmUInt ua = 0;
  SmUInt ub = 0;
  if (tryGetExactNonNegativeUIntFromScalar(l, ua) && tryGetExactNonNegativeUIntFromScalar(r, ub)) {
    SmUInt aaU = 0;
    SmUInt bbU = 0;
    SmUInt sumSqU = 0;
    if (tryMulUIntChecked(ua, ua, aaU) && tryMulUIntChecked(ub, ub, bbU) &&
        tryAddUIntChecked(aaU, bbU, sumSqU)) {
      outV = makeScalar(std::sqrt(static_cast<SmFloat>(sumSqU)));
      return true;
    }
  }
  return false;
}

bool tryApplyPowExactScalarsByKind(
    ScalarExactKind kind,
    SmInt valueSigned,
    SmUInt valueUnsigned,
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
  const SmFloat p = rightS.scalar;
  if (!std::isfinite(p)) {
    return false;
  }
  if (p == 0.0) {
    outV = makeScalarInt(1);
    return true;
  }
  SmFloat r = 0.0;
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
    SmInt nRoot = 0;
    SmInt rInt = 0;
    SmUInt rU = 0;
    if (fractionalPowerIsOddUnitRoot(p, nRoot)) {
      if (isSigned) {
        if (!tryRoundFloatToSInt(r, rInt)) {
          return false;
        }
        if (rInt >= 0 && valueSigned >= 0) {
          SmUInt reconU = 0;
          if (tryPowUIntChecked(static_cast<SmUInt>(rInt), static_cast<SmUInt>(nRoot), reconU) &&
              reconU == static_cast<SmUInt>(valueSigned)) {
            outV = makeScalarInt(rInt);
            return true;
          }
        } else {
          SmInt reconI = 0;
          if (tryPowSIntChecked(rInt, nRoot, reconI) && reconI == valueSigned) {
            outV = makeScalarInt(rInt);
            return true;
          }
        }
      } else {
        SmUInt sq = 0;
        if (tryRoundFloatToUInt(r, rU) &&
            tryPowUIntChecked(rU, static_cast<SmUInt>(nRoot), sq) && sq == valueUnsigned) {
          outV = makeScalarUIntPreferInt(rU);
          return true;
        }
      }
    }
    return false;
  }
  SmInt nExp = 0;
  if (p >= 1.0 && tryGetSmallIntegerExponent(p, nExp)) {
    if (isSigned) {
      SmInt rInt = 0;
      SmInt recon = 0;
      if (tryRoundFloatToSInt(r, rInt) && tryPowSIntChecked(valueSigned, nExp, recon) && recon == rInt) {
        outV = makeScalarInt(rInt);
        return true;
      }
    } else {
      SmUInt powResult = 0;
      SmUInt recon = 0;
      if (tryRoundFloatToUInt(r, powResult) &&
          tryPowUIntChecked(valueUnsigned, static_cast<SmUInt>(nExp), recon) && recon == powResult) {
        outV = makeScalarUIntPreferInt(powResult);
        return true;
      }
    }
  }
  return false;
}

bool tryRecoverLargeSignedOddRoot(SmInt valueInt, SmFloat p, EvalValue& outV, bool& decided) {
  decided = false;
  if (!std::isfinite(p) || !(p > 0.0 && p < 1.0)) {
    return false;
  }
  SmInt nRoot = 0;
  if (!fractionalPowerIsOddUnitRoot(p, nRoot)) {
    return false;
  }
  const bool isMin = valueInt == kIntMin;
  const bool neg = valueInt < 0;
  const SmUInt mag = isMin ? kIntMinMag : static_cast<SmUInt>(neg ? -valueInt : valueInt);
  if (!isMin && smartmath::isMagBelowExactFloatLimit(mag)) {
    return false;
  }
  decided = true;
  const SmFloat magD = isMin ? kFloatSIntBound : static_cast<SmFloat>(mag);
  const SmFloat rootMag = std::pow(magD, p);
  SmUInt rU = 0;
  SmUInt recon = 0;
  if (!std::isfinite(rootMag) || rootMag < 0.0 || !tryRoundFloatToUInt(rootMag, rU) ||
      !tryPowUIntChecked(rU, static_cast<SmUInt>(nRoot), recon) || recon != mag) {
    return false;
  }
  if (neg) {
    if (rU <= static_cast<SmUInt>(kIntMax)) {
      outV = makeScalarInt(-static_cast<SmInt>(rU));
    } else {
      outV = makeScalar(-static_cast<SmFloat>(rU));
    }
  } else {
    outV = makeScalarUIntPreferInt(rU);
  }
  return true;
}

bool tryApplyPowExactScalars(
    const EvalValue::ScalarValue& leftS,
    const EvalValue::ScalarValue& rightS,
    EvalValue& outV) {
#if SMARTMATH_COMPLEX_NUMBERS
  if (scalarHasNonzeroImaginaryPart(leftS) || scalarHasNonzeroImaginaryPart(rightS)) {
    return false;
  }
#endif
  if (leftS.scalarKind == ScalarKind::Time) {
    return false;
  }
  SmInt valueInt = 0;
  if (TryGetExactSignedIntNoUIntWrapScalarStrict(leftS, valueInt)) {
    bool oddRootDecided = false;
    if (tryRecoverLargeSignedOddRoot(valueInt, rightS.scalar, outV, oddRootDecided)) {
      return true;
    }
    if (oddRootDecided) {
      return false;
    }
    if (valueInt == kIntMin) {
      return false;
    }
    const SmUInt mag = (valueInt >= 0) ? static_cast<SmUInt>(valueInt)
                                              : static_cast<SmUInt>(-valueInt);
    if (!smartmath::isMagBelowExactFloatLimit(mag)) {
      return false;
    }
    return tryApplyPowExactScalarsByKind(
        ScalarExactKind::Signed, valueInt, 0u, leftS, rightS, outV);
  }
  SmUInt inpU = 0;
  if (tryGetExactNonNegativeUIntFromScalar(leftS, inpU)) {
    return tryApplyPowExactScalarsByKind(
        ScalarExactKind::Unsigned, 0LL, inpU, leftS, rightS, outV);
  }
  return false;
}

bool tryApplyRealScalarPowNegFractional(
    const EvalValue::ScalarValue& leftS,
    SmFloat p,
    EvalValue& outV) {
  SmInt nRoot = 0;
  if (!fractionalPowerIsOddUnitRoot(p, nRoot)) {
    return false;
  }
  SmFloat ar = 0.0;
  SmFloat ai = 0.0;
  scalarLoadCartesian(leftS, ar, ai);
  if (ai != 0.0 || ar >= 0.0) {
    return false;
  }
  SmInt valueInt = 0;
  if (TryGetExactSignedIntNoUIntWrapScalarStrict(leftS, valueInt)) {
    if (valueInt != kIntMin) {
      SmFloat realR = 0.0;
      if (tryPowFloatGuess(valueInt, leftS.scalar, p, realR)) {
        outV = makeScalar(realR);
        return true;
      }
    }
  }
  const SmFloat rootMag = std::pow(-ar, p);
  if (!std::isfinite(rootMag)) {
    return false;
  }
  outV = makeScalar(-rootMag);
  return true;
}

bool tryApplyScalarPowSpecialPaths(
    const EvalValue::ScalarValue& leftS,
    const EvalValue::ScalarValue& rightS,
    EvalValue& outV) {
  if (tryApplyPowExactScalars(leftS, rightS, outV)) {
    return true;
  }
  return tryApplyRealScalarPowNegFractional(leftS, rightS.scalar, outV);
}

#if SMARTMATH_COMPLEX_NUMBERS
bool tryRefinePowPrincipalToExactScalarResult(
    const EvalValue::ScalarValue& leftS,
    const EvalValue::ScalarValue& rightS,
    SmFloat powR,
    SmFloat powI,
    EvalValue& outV) {
  constexpr SmFloat kEps = kTolLost3;
  SmFloat ar = 0.0;
  SmFloat ai = 0.0;
  SmFloat br = 0.0;
  SmFloat bi = 0.0;
  scalarLoadCartesian(leftS, ar, ai);
  scalarLoadCartesian(rightS, br, bi);
  if (std::fabs(ai) > kEps || std::fabs(bi) > kEps) {
    return false;
  }
  const SmFloat p = br;
  if (!std::isfinite(p) || !std::isfinite(powR) || !std::isfinite(powI)) {
    return false;
  }
  SmInt valueInt = 0;
  const bool hasSigned = TryGetExactSignedIntNoUIntWrapScalarStrict(leftS, valueInt);
  SmUInt inpU = 0;
  const bool hasUInt = tryGetExactNonNegativeUIntFromScalar(leftS, inpU);

  if (p > 0.0 && p < 1.0) {
    SmInt nRoot = 0;
    if (!fractionalPowerResolveRootDegree(p, nRoot)) {
      return false;
    }
    if (hasSigned && valueInt < 0) {
      if (std::fabs(powR) > kEps) {
        return false;
      }
      const SmInt rootCand = static_cast<SmInt>(std::round(powI));
      SmInt verifyBase = valueInt;
      if (nRoot == 2 && valueInt != kIntMin) {
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
      const SmInt rootCand = static_cast<SmInt>(std::round(powR));
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
      const SmUInt rootCand = static_cast<SmUInt>(std::round(powR));
      SmUInt recon = 0;
      if (!tryPowUIntChecked(rootCand, static_cast<SmUInt>(nRoot), recon) || recon != inpU) {
        return false;
      }
      outV = makeScalarUIntPreferInt(rootCand);
      return true;
    }
    return false;
  }
  if (p >= 1.0) {
    SmInt nExp = 0;
    if (!tryGetSmallIntegerExponent(p, nExp)) {
      return false;
    }
    if (std::fabs(powI) > kEps) {
      return false;
    }
    SmInt powResult = 0;
    if (!tryRoundFloatToSInt(powR, powResult)) {
      return false;
    }
    if (hasSigned) {
      SmInt recon = 0;
      if (!tryPowSIntChecked(valueInt, nExp, recon) || recon != powResult) {
        return false;
      }
      outV = makeScalarInt(powResult);
      return true;
    }
    if (hasUInt) {
      SmUInt recon = 0;
      if (!tryPowUIntChecked(inpU, static_cast<SmUInt>(nExp), recon) ||
          recon != static_cast<SmUInt>(powResult)) {
        return false;
      }
      if (static_cast<SmUInt>(powResult) <=
          static_cast<SmUInt>(kIntMax)) {
        outV = makeScalarInt(powResult);
      } else {
        outV = makeScalarUInt(static_cast<SmUInt>(powResult));
      }
      return true;
    }
  }
  return false;
}

bool tryVerifyComplexCartesianSquareExact(
    SmInt rootR, SmInt rootI, SmInt expR, SmInt expI) {
  SmInt rr = 0;
  SmInt ii = 0;
  SmInt ri = 0;
  SmInt sqR = 0;
  SmInt sqI = 0;
  if (!TryMulExactSIntSquare(rootR, rr) || !TryMulExactSIntSquare(rootI, ii)) {
    return false;
  }
  if (!checkedSubLL(rr, ii, sqR)) {
    return false;
  }
  if (!tryMulSIntChecked(rootR, rootI, ri)) {
    return false;
  }
  if (!tryAddSIntChecked(ri, ri, sqI)) {
    return false;
  }
  return sqR == expR && sqI == expI;
}

bool tryRefineSqrtPrincipalToExactComplex(
    const EvalValue::ScalarValue& inS,
    SmFloat sqrtR,
    SmFloat sqrtI,
    EvalValue& outV) {
  if (!std::isfinite(sqrtR) || !std::isfinite(sqrtI)) {
    return false;
  }
  ExactCartesianComponent reC{};
  ExactCartesianComponent imC{};
  if (!tryExtractExactRealComponent(inS, reC) || !tryExtractExactImagComponent(inS, imC)) {
    return false;
  }
  SmInt arI = 0;
  SmInt aiI = 0;
  if (!TryExactCartesianComponentToSInt(reC, arI) || !TryExactCartesianComponentToSInt(imC, aiI)) {
    return false;
  }
  const SmInt rootRi = static_cast<SmInt>(std::round(sqrtR));
  const SmInt rootIi = static_cast<SmInt>(std::round(sqrtI));
  if (!tryVerifyComplexCartesianSquareExact(rootRi, rootIi, arI, aiI)) {
    return false;
  }
  outV = setScalarComplexFromEvalRealImagParts(makeScalarInt(rootRi), makeScalarInt(rootIi));
  return true;
}

EvalValue applySqrtComplexPrincipalUnary(const EvalValue::ScalarValue& inS) {
  SmFloat cr = 0.0;
  SmFloat ai = 0.0;
  scalarLoadCartesian(inS, cr, ai);
  if (!std::isfinite(cr) || !std::isfinite(ai)) {
    return makeScalarComplexFromFloats(
        std::numeric_limits<SmFloat>::quiet_NaN(), std::numeric_limits<SmFloat>::quiet_NaN());
  }
  const SmFloat mag = std::hypot(cr, ai);
  if (mag == 0.0) {
    return makeScalarComplexFromFloats(0.0, 0.0);
  }
  SmFloat sqrtR = 0.0;
  SmFloat sqrtI = 0.0;
  complexPrincipalSqrt(cr, ai, sqrtR, sqrtI);
  EvalValue out{};
  if (tryRefineSqrtPrincipalToExactComplex(inS, sqrtR, sqrtI, out)) {
    return out;
  }
  return makeScalarComplexFromFloats(sqrtR, sqrtI);
}
#endif

bool tryGetExactSignedIntFromScalar(const EvalValue::ScalarValue& s, SmInt& outI) {
  return tryGetExactSignedIntFromScalarPolicy(
      s, outI, ExactSignedIntPolicy::AllowUIntBitReinterpret);
}

bool tryGetExactSignedIntNoUIntWrapFromScalar(const EvalValue::ScalarValue& s, SmInt& outI) {
  return tryGetExactSignedIntFromScalarPolicy(s, outI, ExactSignedIntPolicy::NoUIntWrap);
}

bool tryGetExactNonNegativeUIntFromScalar(const EvalValue::ScalarValue& s, SmUInt& outU) {
  if (s.scalarKind == ScalarKind::Time) {
    return false;
  }
#if SMARTMATH_COMPLEX_NUMBERS
  if (scalarHasNonzeroImaginaryPart(s)) {
    return false;
  }
#endif
  EvalValue::ScalarValue sv = s;
  scalarRepairExactMetadata(sv);
  if (sv.hasExactUInt() || sv.scalarKind == ScalarKind::UInt) {
    outU = sv.exactUInt;
    return true;
  }
  if (sv.hasExactSInt() && sv.exactSInt >= 0) {
    outU = static_cast<SmUInt>(sv.exactSInt);
    return true;
  }
  if (sv.scalarKind == ScalarKind::SInt && sv.exactSInt >= 0) {
    outU = static_cast<SmUInt>(sv.exactSInt);
    return true;
  }
  return false;
}

bool tryGetBothExactSignedIntNoUIntWrapFromScalars(
    const EvalValue::ScalarValue& a,
    const EvalValue::ScalarValue& b,
    SmInt& outA,
    SmInt& outB) {
  return tryGetExactSignedIntNoUIntWrapFromScalar(a, outA) &&
         tryGetExactSignedIntNoUIntWrapFromScalar(b, outB);
}

bool tryGetBothExactNonNegativeUIntFromScalars(
    const EvalValue::ScalarValue& a,
    const EvalValue::ScalarValue& b,
    SmUInt& outA,
    SmUInt& outB) {
  return tryGetExactNonNegativeUIntFromScalar(a, outA) &&
         tryGetExactNonNegativeUIntFromScalar(b, outB);
}

bool tryShiftLeftU64ExactOrMaybe(
    SmUInt aU,
    SmUInt bU,
    EvalValue& outV) {
  if (bU > static_cast<SmUInt>(kMaxShift)) return false;
  if (bU > 0u && aU > (kUIntMax >> bU)) {
    outV = makeScalarMaybeExact(std::ldexp(static_cast<SmFloat>(aU), static_cast<int>(bU)));
    return true;
  }
  outV = makeScalarUInt(aU << bU);
  return true;
}

bool tryGetSignedIntFromScalar(const EvalValue::ScalarValue& s, SmInt& outI) {
  if (s.scalarKind == ScalarKind::Time) {
    return false;
  }
  if (tryGetExactSignedIntFromScalar(s, outI)) {
    return true;
  }
  return tryExtractExactSIntFromFloatStrict(s.scalar, outI);
}

bool tryGetExactSignMagnitudeFromScalar(const EvalValue::ScalarValue& s, bool& isNeg, SmUInt& mag) {
  if (tryGetExactNonNegativeUIntFromScalar(s, mag)) {
    isNeg = false;
    return true;
  }
  SmInt i = 0;
  if (!tryGetSignedIntFromScalar(s, i)) {
    return false;
  }
  isNeg = i < 0;
  mag = uintMagnitudeOfSInt(i);
  return true;
}

EvalValue makeScalarFromSignMagnitude(bool isNeg, SmUInt mag) {
  if (isNeg && mag != 0u) {
    return makeScalarInt(sintFromNegatedMagnitude(mag));
  }
  return makeScalarUIntPreferInt(mag);
}

bool parseUIntFromFloat(SmFloat v, SmUInt& out) {
  if (!isWithinExactIntFromFloatRange(v) || v < 0.0) {
    return false;
  }
  if (!isFloatInUIntRange(v)) {
    return false;
  }
  const SmUInt u = static_cast<SmUInt>(v);
  if (v != static_cast<SmFloat>(u)) {
    return false;
  }
  out = u;
  return true;
}

EvalValue makeScalar(SmFloat v) {
  EvalValue out;
  out.kind = ValueKind::Scalar;
  out.scalarValue.scalarKind = ScalarKind::FloatingPoint;
  out.scalarValue.scalar = v;
  out.scalarValue.setExactSIntValid(false);
  out.scalarValue.exactSInt = 0;
  out.scalarValue.setExactUIntValid(false);
  out.scalarValue.exactUInt = 0;
  out.scalarValue.setDecScientificPow63High(false);
  out.scalarValue.imag = 0.0;
  out.scalarValue.imagExactSInt = 0;
  out.scalarValue.imagExactUInt = 0;
  out.scalarValue.setImagExactSIntValid(false);
  out.scalarValue.setImagExactUIntValid(false);
  return out;
}

#if SMARTMATH_COMPLEX_NUMBERS
void scalarClearCartesianRenderExact(EvalValue::ScalarValue& s) {
  s.setRenderRational(false);
  s.setImagRenderRational(false);
  s.setExactUIntValid(false);
  s.setImagExactUIntValid(false);
}
#endif

void scalarClearImaginary(EvalValue::ScalarValue& s) {
  s.imag = 0.0;
  s.imagExactSInt = 0;
  s.imagExactUInt = 0;
  s.setImagExactSIntValid(false);
  s.setImagExactUIntValid(false);
}

#if SMARTMATH_COMPLEX_NUMBERS
void ScalarApplyExactSIntPart(EvalValue::ScalarValue& s, bool imagPart, SmInt n) {
  if (imagPart) {
    s.imagExactSInt = n;
    s.setImagExactSIntValid(true);
    s.setImagExactUIntValid(false);
    s.setImagRenderRational(false);
  } else {
    s.exactSInt = n;
    s.setExactSIntValid(true);
    s.setExactUIntValid(false);
    s.setRenderRational(false);
  }
}
#endif

void scalarApplyReducedRationalPart(EvalValue::ScalarValue& s, bool imagPart, SmInt num,
                                                SmUInt den) {
  SmInt n = num;
  SmUInt d = den;
  const SmInt g = GcdSInt(n, static_cast<SmInt>(d));
  if (g > 0) {
    n /= g;
    d = static_cast<SmUInt>(static_cast<SmInt>(d) / g);
  }
#if SMARTMATH_COMPLEX_NUMBERS
  if (imagPart) {
    s.imagExactSInt = n;
    s.setImagExactSIntValid(true);
    if (d == 1u) {
      s.setImagExactUIntValid(false);
      s.setImagRenderRational(false);
    } else {
      s.imagExactUInt = d;
      s.setImagExactUIntValid(true);
      s.setImagRenderRational(true);
    }
    return;
  }
#else
  (void)imagPart;
#endif
  s.exactSInt = n;
  s.setExactSIntValid(true);
  if (d == 1u) {
    s.setExactUIntValid(false);
    s.setRenderRational(false);
  } else {
    s.exactUInt = d;
    s.setExactUIntValid(true);
    s.setRenderRational(true);
  }
}

#if SMARTMATH_COMPLEX_NUMBERS
EvalValue makeImaginaryUnit() {
  EvalValue out = makeScalarComplexFromFloats(0.0, 1.0);
  out.scalarValue.scalarKind = ScalarKind::SInt;
  out.scalarValue.setExactSIntValid(true);
  out.scalarValue.exactSInt = 0;
  out.scalarValue.setExactUIntValid(true);
  out.scalarValue.exactUInt = 0;
  out.scalarValue.setImagExactSIntValid(true);
  out.scalarValue.imagExactSInt = 1;
  out.scalarValue.setImagExactUIntValid(true);
  out.scalarValue.imagExactUInt = 1u;
  out.scalarValue.scalar = 0.0;
  out.scalarValue.imag = 1.0;
  return out;
}

EvalValue makeScalarComplexFromFloats(SmFloat re, SmFloat im) {
  if (std::isnan(re) && std::isnan(im)) {
    return makeScalar(std::numeric_limits<SmFloat>::quiet_NaN());
  }
  EvalValue out = makeScalar(re);
  out.scalarValue.setRenderRational(false);
  out.scalarValue.setImagRenderRational(false);
#if SMARTMATH_FACTORINT
  out.scalarValue.setRenderIntPower(false);
#endif
  out.scalarValue.imag = im;
  out.scalarValue.setImagExactSIntValid(false);
  out.scalarValue.setImagExactUIntValid(false);
  out.scalarValue.imagExactSInt = 0;
  out.scalarValue.imagExactUInt = 0;
  SmInt tri = 0;
  SmInt tii = 0;
  if (tryExtractExactSIntFromFloatStrict(re, tri)) {
    out.scalarValue.scalarKind = ScalarKind::SInt;
    out.scalarValue.setExactSIntValid(true);
    out.scalarValue.exactSInt = tri;
    if (tri >= 0) {
      out.scalarValue.setExactUIntValid(true);
      out.scalarValue.exactUInt = static_cast<SmUInt>(tri);
    } else {
      out.scalarValue.setExactUIntValid(false);
      out.scalarValue.exactUInt = 0;
    }
  }
  if (tryExtractExactSIntFromFloatStrict(im, tii)) {
    out.scalarValue.setImagExactSIntValid(true);
    out.scalarValue.imagExactSInt = tii;
    if (tii >= 0) {
      out.scalarValue.setImagExactUIntValid(true);
      out.scalarValue.imagExactUInt = static_cast<SmUInt>(tii);
    } else {
      out.scalarValue.setImagExactUIntValid(false);
      out.scalarValue.imagExactUInt = 0;
    }
  }
  if (!scalarHasNonzeroImaginaryPart(out.scalarValue)) {
    scalarClearImaginary(out.scalarValue);
  }
  return out;
}
#endif

void exactCartesianComponentClear(ExactCartesianComponent& c) {
  c.hasInt = false;
  c.intV = 0;
  c.hasUInt = false;
  c.uintV = 0;
}

void ExactCartesianComponentAssignFromSignedInt(ExactCartesianComponent& c, SmInt n) {
  exactCartesianComponentClear(c);
  c.hasInt = true;
  c.intV = n;
  c.hasUInt = true;
  c.uintV = static_cast<SmUInt>(n);
}

void ExactCartesianComponentAssignFromUInt(ExactCartesianComponent& c, SmUInt u) {
  exactCartesianComponentClear(c);
  c.hasUInt = true;
  c.uintV = u;
  if (u <= static_cast<SmUInt>(kIntMax)) {
    c.hasInt = true;
    c.intV = static_cast<SmInt>(u);
  }
}

bool tryExtractExactRealComponent(const EvalValue::ScalarValue& sv, ExactCartesianComponent& c) {
  if (sv.hasRenderRational()) {
    return false;
  }
  if (sv.hasExactSInt()) {
    ExactCartesianComponentAssignFromSignedInt(c, sv.exactSInt);
    if (sv.exactSInt < 0 && sv.hasExactUInt()) {
      c.hasUInt = true;
      c.uintV = sv.exactUInt;
    }
    return true;
  }
  if (sv.hasExactUInt()) {
    ExactCartesianComponentAssignFromUInt(c, sv.exactUInt);
    return true;
  }
  SmInt t = 0;
  if (tryExtractExactSIntFromFloatStrict(sv.scalar, t)) {
    ExactCartesianComponentAssignFromSignedInt(c, t);
    return true;
  }
  exactCartesianComponentClear(c);
  return false;
}

#if SMARTMATH_COMPLEX_NUMBERS
bool complexNeedsPrincipalNegRealPow(SmFloat ar, SmFloat ai, SmFloat br, SmFloat bi) {
  if (ai != 0.0 || bi != 0.0) {
    return false;
  }
  if (ar >= 0.0) {
    return false;
  }
  if (!std::isfinite(ar) || !std::isfinite(br)) {
    return false;
  }
  if (std::fabs(br - std::trunc(br)) < kTolLost3) {
    return false;
  }
  return true;
}

void complexCartesianBinary(SmFloat ar, SmFloat ai, SmFloat br, SmFloat bi, char op, SmFloat& outR,
                                          SmFloat& outI) {
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
      outR = std::numeric_limits<SmFloat>::quiet_NaN();
      outI = std::numeric_limits<SmFloat>::quiet_NaN();
      break;
  }
}

bool TryExactCartesianComponentToSInt(const ExactCartesianComponent& c, SmInt& outI) {
  if (c.hasInt) {
    outI = c.intV;
    return true;
  }
  if (c.hasUInt && c.uintV <= static_cast<SmUInt>(kIntMax)) {
    outI = static_cast<SmInt>(c.uintV);
    return true;
  }
  if (!c.hasInt && !c.hasUInt) {
    outI = 0;
    return true;
  }
  return false;
}

bool tryExtractExactCartesianComponent(bool isImag, const EvalValue::ScalarValue& sv,
                                                   ExactCartesianComponent& c) {
  if (!isImag) {
    return tryExtractExactRealComponent(sv, c);
  }
  if (sv.hasImagRenderRational()) {
    return false;
  }
  if (sv.hasImagExactSInt()) {
    if (sv.imagExactSInt == 0 && sv.imag != 0.0) {
      return false;
    }
    ExactCartesianComponentAssignFromSignedInt(c, sv.imagExactSInt);
    if (sv.imagExactSInt < 0 && sv.hasImagExactUInt()) {
      c.hasUInt = true;
      c.uintV = sv.imagExactUInt;
    }
    return true;
  }
  if (sv.hasImagExactUInt()) {
    if (sv.imagExactUInt == 0u && sv.imag != 0.0) {
      return false;
    }
    ExactCartesianComponentAssignFromUInt(c, sv.imagExactUInt);
    return true;
  }
  if (imagExactMetadataMatchesFloat(sv)) {
    if (sv.hasImagExactSIntMetadata()) {
      ExactCartesianComponentAssignFromSignedInt(c, sv.imagExactSInt);
      if (sv.imagExactSInt < 0 && sv.hasImagExactUIntMetadata()) {
        c.hasUInt = true;
        c.uintV = sv.imagExactUInt;
      }
      return true;
    }
    if (sv.hasImagExactUIntMetadata()) {
      ExactCartesianComponentAssignFromUInt(c, sv.imagExactUInt);
      return true;
    }
  }
  if (!scalarHasNonzeroImaginaryPart(sv)) {
    exactCartesianComponentClear(c);
    return true;
  }
  SmInt t = 0;
  if (tryExtractExactSIntFromFloatStrict(sv.imag, t)) {
    if (t == 0 && sv.imag != 0.0) {
      return false;
    }
    ExactCartesianComponentAssignFromSignedInt(c, t);
    return true;
  }
  exactCartesianComponentClear(c);
  return false;
}

bool tryExtractExactImagComponent(const EvalValue::ScalarValue& sv, ExactCartesianComponent& c) {
  return tryExtractExactCartesianComponent(true, sv, c);
}
#endif

void setScalarFromExactCartesianComponent(EvalValue& v, const ExactCartesianComponent& c) {
  v.kind = ValueKind::Scalar;
  if (c.hasInt) {
    v.scalarValue.scalarKind = ScalarKind::SInt;
    v.scalarValue.scalar = static_cast<SmFloat>(c.intV);
    v.scalarValue.setExactSIntValid(true);
    v.scalarValue.exactSInt = c.intV;
    if (c.intV >= 0) {
      v.scalarValue.setExactUIntValid(true);
      v.scalarValue.exactUInt = static_cast<SmUInt>(c.intV);
    } else if (c.hasUInt) {
      v.scalarValue.setExactUIntValid(true);
      v.scalarValue.exactUInt = c.uintV;
    } else {
      v.scalarValue.setExactUIntValid(false);
      v.scalarValue.exactUInt = 0;
    }
  } else if (c.hasUInt) {
    v.scalarValue.scalarKind = ScalarKind::UInt;
    v.scalarValue.scalar = static_cast<SmFloat>(c.uintV);
    v.scalarValue.setExactUIntValid(true);
    v.scalarValue.exactUInt = c.uintV;
    if (c.uintV <= static_cast<SmUInt>(kIntMax)) {
      v.scalarValue.setExactSIntValid(true);
      v.scalarValue.exactSInt = static_cast<SmInt>(c.uintV);
    } else {
      v.scalarValue.setExactSIntValid(false);
      v.scalarValue.exactSInt = 0;
    }
  } else {
    v.scalarValue.scalarKind = ScalarKind::FloatingPoint;
    v.scalarValue.scalar = 0.0;
    v.scalarValue.setExactSIntValid(false);
    v.scalarValue.setExactUIntValid(false);
  }
  v.scalarValue.setDecScientificPow63High(false);
  scalarClearImaginary(v.scalarValue);
}

#if SMARTMATH_COMPLEX_NUMBERS
void scalarApplyExactImagFromCartesianComponent(EvalValue::ScalarValue& sv,
                                                          const ExactCartesianComponent& c) {
  sv.imag = 0.0;
  sv.imagExactSInt = 0;
  sv.imagExactUInt = 0;
  sv.setImagExactSIntValid(false);
  sv.setImagExactUIntValid(false);
  if (!c.hasInt && !c.hasUInt) {
    return;
  }
  if (c.hasInt) {
    sv.setImagExactSIntValid(true);
    sv.imagExactSInt = c.intV;
    sv.imag = static_cast<SmFloat>(c.intV);
    sv.setImagExactUIntValid(true);
    sv.imagExactUInt = c.hasUInt ? c.uintV : static_cast<SmUInt>(c.intV);
  } else {
    sv.setImagExactUIntValid(true);
    sv.imagExactUInt = c.uintV;
    sv.imag = static_cast<SmFloat>(c.uintV);
    if (c.uintV <= static_cast<SmUInt>(kIntMax)) {
      sv.setImagExactSIntValid(true);
      sv.imagExactSInt = static_cast<SmInt>(c.uintV);
    }
  }
}

void setScalarComplexFromExactCartesian(EvalValue& v, const ExactCartesianComponent& re,
                                                    const ExactCartesianComponent& im) {
  setScalarFromExactCartesianComponent(v, re);
  v.scalarValue.imag = 0.0;
  scalarApplyExactImagFromCartesianComponent(v.scalarValue, im);
  if (!scalarHasNonzeroImaginaryPart(v.scalarValue)) {
    scalarClearImaginary(v.scalarValue);
  }
}

bool tryApplyExactCartesianBinaryOp(bool isAdd, const ExactCartesianComponent& a,
                                                const ExactCartesianComponent& b,
                                                ExactCartesianComponent& out) {
  exactCartesianComponentClear(out);
  SmInt ai = 0;
  SmInt bi = 0;
  SmInt oi = 0;
  if (TryExactCartesianComponentToSInt(a, ai) && TryExactCartesianComponentToSInt(b, bi)) {
    if (isAdd) {
      if (!checkedAddLL(ai, bi, oi)) {
        return false;
      }
    } else if (!checkedSubLL(ai, bi, oi)) {
      return false;
    }
    ExactCartesianComponentAssignFromSignedInt(out, oi);
    return true;
  }
  if (isAdd && a.hasUInt && b.hasUInt) {
    SmUInt ou = 0;
    if (tryAddUIntChecked(a.uintV, b.uintV, ou)) {
      ExactCartesianComponentAssignFromUInt(out, ou);
      return true;
    }
  }
  return false;
}
#endif

bool TryQuotExactSInt(SmInt num, SmInt den, SmInt& quo) {
  if (den == 0 || (den == -1 && num == kIntMin)) {
    return false;
  }
  quo = num / den;
  return quo * den == num;
}

bool tryMulUIntChecked(SmUInt a, SmUInt b, SmUInt& out) {
  if (b != 0u && a > (kUIntMax / b)) {
    return false;
  }
  out = a * b;
  return true;
}

bool tryPowUIntChecked(SmUInt base, SmUInt exp, SmUInt& out) {
  SmUInt r = 1;
  SmUInt b = base;
  SmUInt e = exp;
  while (e > 0u) {
    if ((e & 1u) != 0u && !tryMulUIntChecked(r, b, r)) {
      return false;
    }
    e >>= 1;
    if (e > 0u && !tryMulUIntChecked(b, b, b)) {
      return false;
    }
  }
  out = r;
  return true;
}

bool passExactAbsFloatFactorGates(SmFloat f, bool strictAbsFAboveOne) {
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
  SmInt tmp = 0;
  return !nearlyInt(f, tmp);
}

bool getExactRoundedIntFromFloatResult(SmFloat r, bool boundAbsResultTo2_53, SmInt& outI) {
  const SmFloat rounded = std::round(r);
  if (!isWithinExactIntFromFloatRange(rounded)) {
    return false;
  }
  const SmInt roundedI = static_cast<SmInt>(rounded);
  if (static_cast<SmFloat>(roundedI) != r || roundedI == 0) {
    return false;
  }
  if (boundAbsResultTo2_53 && std::fabs(static_cast<SmFloat>(roundedI)) > kMaxExactIntFromFloat) {
    return false;
  }
  outI = roundedI;
  return true;
}

bool verifyExactIntFloatOpResidue(
    ScalarExactKind kind,
    SmInt intSigned,
    SmUInt intUnsigned,
    SmFloat f,
    SmInt resultSigned,
    SmUInt resultUnsigned,
    bool isMultiply) {
  if (kind == ScalarExactKind::Signed) {
    if (isMultiply) {
      return relCloseExactIntFloat(
          static_cast<SmFloat>(resultSigned) - static_cast<SmFloat>(intSigned) * f, 0.0);
    }
    return relCloseExactIntFloat(
        static_cast<SmFloat>(intSigned) - static_cast<SmFloat>(resultSigned) * f, 0.0);
  }
  if (isMultiply) {
    return relCloseExactIntFloat(
        static_cast<SmFloat>(resultUnsigned) - static_cast<SmFloat>(intUnsigned) * f, 0.0);
  }
  return relCloseExactIntFloat(
      static_cast<SmFloat>(intUnsigned) - static_cast<SmFloat>(resultUnsigned) * f, 0.0);
}

bool verifyExactIntFloatOpCrossMultiply(
    ScalarExactKind kind,
    SmInt intSigned,
    SmUInt intUnsigned,
    SmFloat floatV,
    SmInt resultSigned,
    SmUInt resultUnsigned,
    SmInt ratA,
    SmInt ratB,
    bool isMultiply) {
  if (kind == ScalarExactKind::Signed) {
    const SmInt signF = floatV < 0.0 ? -1 : 1;
    SmInt resultSignedScaled = 0;
    SmInt lhs = 0;
    SmInt rhs = 0;
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
  SmUInt lhs = 0;
  SmUInt rhs = 0;
  if (isMultiply) {
    if (!tryMulUIntChecked(resultUnsigned, static_cast<SmUInt>(ratB), lhs) ||
        !tryMulUIntChecked(intUnsigned, static_cast<SmUInt>(ratA), rhs)) {
      return false;
    }
  } else if (!tryMulUIntChecked(resultUnsigned, static_cast<SmUInt>(ratA), lhs) ||
             !tryMulUIntChecked(intUnsigned, static_cast<SmUInt>(ratB), rhs)) {
    return false;
  }
  return lhs == rhs;
}

bool tryPromoteExactIntFromFloatOp(
    ScalarExactKind kind,
    SmInt intSigned,
    SmUInt intUnsigned,
    SmFloat f,
    SmFloat r,
    bool isMultiply,
    EvalValue& outV) {
  if (!passExactAbsFloatFactorGates(f, isMultiply)) {
    return false;
  }
  if (kind == ScalarExactKind::Signed) {
    if (std::fabs(static_cast<SmFloat>(intSigned)) > kMaxExactIntFromFloat) {
      return false;
    }
  } else if (f < 0.0 || static_cast<SmFloat>(intUnsigned) > kMaxExactIntFromFloat) {
    return false;
  }
  SmInt resultSigned = 0;
  if (!getExactRoundedIntFromFloatResult(r, isMultiply, resultSigned)) {
    return false;
  }
  SmUInt resultUnsigned = 0;
  if (kind == ScalarExactKind::Unsigned) {
    if (resultSigned < 0) {
      return false;
    }
    resultUnsigned = static_cast<SmUInt>(resultSigned);
    if (isMultiply && static_cast<SmFloat>(resultUnsigned) > kMaxExactIntFromFloat) {
      return false;
    }
  }
  if (!verifyExactIntFloatOpResidue(
          kind, intSigned, intUnsigned, f, resultSigned, resultUnsigned, isMultiply)) {
    return false;
  }
  SmInt ratA = 0;
  SmInt ratB = 0;
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

bool tryPromoteExactDivisionByFloatDivisorSigned(
    SmInt intI, SmFloat f, SmFloat r, EvalValue& outV) {
  return tryPromoteExactIntFromFloatOp(ScalarExactKind::Signed, intI, 0u, f, r, false, outV);
}

bool tryPromoteExactDivisionByFloatDivisorUnsigned(
    SmUInt intU, SmFloat f, SmFloat r, EvalValue& outV) {
  return tryPromoteExactIntFromFloatOp(ScalarExactKind::Unsigned, 0LL, intU, f, r, false, outV);
}

bool tryPromoteExactMultiplicationByFloatFactorSigned(
    SmInt intI, SmFloat f, SmFloat r, EvalValue& outV) {
  return tryPromoteExactIntFromFloatOp(ScalarExactKind::Signed, intI, 0u, f, r, true, outV);
}

bool tryPromoteExactMultiplicationByFloatFactorUnsigned(
    SmUInt intU, SmFloat f, SmFloat r, EvalValue& outV) {
  return tryPromoteExactIntFromFloatOp(ScalarExactKind::Unsigned, 0LL, intU, f, r, true, outV);
}

bool tryApplyExactIntFloatOpFromFloatSide(
    ScalarExactKind kind,
    bool isMultiply,
    SmInt intSigned,
    SmUInt intUnsigned,
    SmFloat f,
    SmFloat r,
    EvalValue& outV) {
  if (!std::isfinite(f) || f == 0.0) {
    return false;
  }
  constexpr SmFloat kEps = kTolLost3;
  SmInt tmp = 0;

  if (kind == ScalarExactKind::Unsigned) {
    if (isMultiply) {
      if (f <= 0.0) {
        return false;
      }
      if (f < 1.0 && !nearlyInt(f, tmp)) {
        const SmFloat inv = 1.0 / f;
        const SmFloat invRound = std::round(inv);
        if (invRound > 1.0 &&
            invRound < kFloatUIntBound &&
            std::fabs(inv - invRound) <= kEps * std::max(static_cast<SmFloat>(1), std::fabs(inv))) {
          const SmUInt nU = static_cast<SmUInt>(invRound);
          if (nU != 0u && (intUnsigned % nU) == 0u) {
            outV = makeScalarUInt(intUnsigned / nU);
            return true;
          }
        }
      }
      return tryPromoteExactMultiplicationByFloatFactorUnsigned(intUnsigned, f, r, outV);
    }
    if (std::fabs(f) < 1.0 && !nearlyInt(f, tmp)) {
      const SmFloat inv = 1.0 / f;
      const SmFloat invRound = std::round(inv);
      if (std::fabs(invRound) > 1.0 &&
          std::fabs(invRound) < kFloatUIntBound &&
          std::fabs(inv - invRound) <= kEps * std::max(static_cast<SmFloat>(1), std::fabs(inv))) {
        if (invRound > 0.0) {
          const SmUInt nU = static_cast<SmUInt>(invRound);
          SmUInt qU = 0;
          if (tryMulUIntChecked(intUnsigned, nU, qU)) {
            outV = makeScalarUInt(qU);
            return true;
          }
        } else if (invRound < 0.0) {
          const SmUInt nU = static_cast<SmUInt>(-invRound);
          SmUInt qU = 0;
          if (tryMulUIntChecked(intUnsigned, nU, qU) &&
              qU <= static_cast<SmUInt>(kIntMax)) {
            outV = makeScalarInt(-static_cast<SmInt>(qU));
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
      const SmFloat inv = 1.0 / f;
      const SmFloat invAbs = std::fabs(inv);
      const SmFloat invRound = std::round(invAbs);
      if (invRound <= 1.0 ||
          invRound >= kFloatSIntBound ||
          std::fabs(invAbs - invRound) > kEps * std::max(static_cast<SmFloat>(1), invAbs)) {
        return false;
      }
      const SmInt nI = static_cast<SmInt>(invRound);
      if (nI <= 1 || (intSigned % nI) != 0) {
        return false;
      }
      SmInt qI = intSigned / nI;
      if (f < 0.0) {
        if (qI == kIntMin) {
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
    const SmFloat inv = 1.0 / f;
    const SmFloat invRound = std::round(inv);
    if (std::fabs(invRound) > 1.0 &&
        std::fabs(invRound) < kFloatSIntBound &&
        std::fabs(inv - invRound) <= kEps * std::max(static_cast<SmFloat>(1), std::fabs(inv))) {
      const SmInt nI = static_cast<SmInt>(invRound);
      SmInt qI = 0;
      if (checkedMulLL(intSigned, nI, qI)) {
        outV = makeScalarInt(qI);
        return true;
      }
      if (intSigned >= 0 && nI > 0) {
        const SmUInt nU = static_cast<SmUInt>(nI);
        SmUInt qU = 0;
        if (tryMulUIntChecked(static_cast<SmUInt>(intSigned), nU, qU)) {
          outV = makeScalarUInt(qU);
          return true;
        }
      }
    }
  }
  return tryPromoteExactDivisionByFloatDivisorSigned(intSigned, f, r, outV);
}

#if SMARTMATH_COMPLEX_NUMBERS
bool tryApplyExactComplexCartesianBinary(const EvalValue::ScalarValue& leftS,
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
    SmInt lar = 0;
    SmInt lai = 0;
    SmInt lbr = 0;
    SmInt lbi = 0;
    SmInt p1 = 0;
    SmInt p2 = 0;
    SmInt p3 = 0;
    SmInt p4 = 0;
    SmInt oreI = 0;
    SmInt oimI = 0;
    if (!TryExactCartesianComponentToSInt(lRe, lar) || !TryExactCartesianComponentToSInt(lIm, lai) ||
        !TryExactCartesianComponentToSInt(rRe, lbr) || !TryExactCartesianComponentToSInt(rIm, lbi) ||
        !checkedMulLL(lar, lbr, p1) || !checkedMulLL(lai, lbi, p2) || !checkedMulLL(lar, lbi, p3) ||
        !checkedMulLL(lai, lbr, p4) || !checkedSubLL(p1, p2, oreI) || !checkedAddLL(p3, p4, oimI)) {
      return false;
    }
    ExactCartesianComponentAssignFromSignedInt(oRe, oreI);
    ExactCartesianComponentAssignFromSignedInt(oIm, oimI);
    setScalarComplexFromExactCartesian(outV, oRe, oIm);
    return true;
  }
  if (op == '/') {
    SmInt lar = 0;
    SmInt lai = 0;
    SmInt lbr = 0;
    SmInt lbi = 0;
    SmInt p1 = 0;
    SmInt p2 = 0;
    SmInt p3 = 0;
    SmInt p4 = 0;
    SmInt p5 = 0;
    SmInt p6 = 0;
    SmInt numRe = 0;
    SmInt numIm = 0;
    SmInt denom = 0;
    SmInt qRe = 0;
    SmInt qIm = 0;
    if (!TryExactCartesianComponentToSInt(lRe, lar) || !TryExactCartesianComponentToSInt(lIm, lai) ||
        !TryExactCartesianComponentToSInt(rRe, lbr) || !TryExactCartesianComponentToSInt(rIm, lbi) ||
        !checkedMulLL(lar, lbr, p1) || !checkedMulLL(lai, lbi, p2) || !checkedAddLL(p1, p2, numRe) ||
        !checkedMulLL(lbr, lbr, p3) || !checkedMulLL(lbi, lbi, p4) || !checkedAddLL(p3, p4, denom) ||
        denom == 0 || !checkedMulLL(lai, lbr, p5) || !checkedMulLL(lar, lbi, p6) ||
        !checkedSubLL(p5, p6, numIm) || !TryQuotExactSInt(numRe, denom, qRe) ||
        !TryQuotExactSInt(numIm, denom, qIm)) {
      return false;
    }
    ExactCartesianComponentAssignFromSignedInt(oRe, qRe);
    ExactCartesianComponentAssignFromSignedInt(oIm, qIm);
    setScalarComplexFromExactCartesian(outV, oRe, oIm);
    return true;
  }
  return false;
}

EvalValue setScalarComplexFromEvalRealImagParts(const EvalValue& rePart,
                                                                        const EvalValue& imPart) {
  ExactCartesianComponent reC{};
  ExactCartesianComponent imC{};
  if (tryExtractExactRealComponent(rePart.scalarValue, reC) &&
      tryExtractExactRealComponent(imPart.scalarValue, imC)) {
    EvalValue out{};
    setScalarComplexFromExactCartesian(out, reC, imC);
    return out;
  }
  return makeScalarComplexFromFloats(rePart.scalarValue.scalar, imPart.scalarValue.scalar);
}

void setPureImaginaryFromMagnitudeScalar(EvalValue& outV, const EvalValue::ScalarValue& magSv) {
  // Copy first: callers may pass outV.scalarValue (e.g. tight-imag suffix on a literal).
  const EvalValue::ScalarValue mag = magSv;
  outV = makeScalarInt(0);
  if (scalarHasExactIntegerPayload(mag)) {
    if (mag.hasExactSInt()) {
      outV.scalarValue.setImagExactSIntValid(true);
      outV.scalarValue.imagExactSInt = mag.exactSInt;
      if (mag.exactSInt >= 0) {
        outV.scalarValue.setImagExactUIntValid(true);
        outV.scalarValue.imagExactUInt = static_cast<SmUInt>(mag.exactSInt);
      } else {
        outV.scalarValue.setImagExactUIntValid(false);
        outV.scalarValue.imagExactUInt = 0;
      }
      outV.scalarValue.imag = static_cast<SmFloat>(mag.exactSInt);
    } else if (mag.hasExactUInt()) {
      outV.scalarValue.setImagExactUIntValid(true);
      outV.scalarValue.imagExactUInt = mag.exactUInt;
      outV.scalarValue.imag = static_cast<SmFloat>(mag.exactUInt);
      if (mag.exactUInt <= static_cast<SmUInt>(kIntMax)) {
        outV.scalarValue.setImagExactSIntValid(true);
        outV.scalarValue.imagExactSInt = static_cast<SmInt>(mag.exactUInt);
      } else {
        outV.scalarValue.setImagExactSIntValid(false);
        outV.scalarValue.imagExactSInt = 0;
      }
    } else {
      outV.scalarValue.imag = mag.scalar;
      outV.scalarValue.setImagExactSIntValid(false);
      outV.scalarValue.setImagExactUIntValid(false);
    }
  } else {
    outV.scalarValue.imag = mag.scalar;
    outV.scalarValue.setImagExactSIntValid(false);
    outV.scalarValue.setImagExactUIntValid(false);
    outV.scalarValue.imagExactSInt = 0;
    outV.scalarValue.imagExactUInt = 0;
  }
  scalarRepairExactMetadata(outV.scalarValue);
}

bool tryNegateExactCartesianComponent(const ExactCartesianComponent& c, ExactCartesianComponent& outC) {
  SmInt i = 0;
  if (TryExactCartesianComponentToSInt(c, i)) {
    if (i == kIntMin) {
      return false;
    }
    ExactCartesianComponentAssignFromSignedInt(outC, -i);
    return true;
  }
  if (c.hasUInt) {
    return false;
  }
  ExactCartesianComponentAssignFromSignedInt(outC, 0);
  return true;
}

bool tryNegateExactComplexScalar(const EvalValue::ScalarValue& sv, EvalValue& out) {
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

bool tryFoldExactComplexCartesian(const std::vector<EvalValue>& args, char op, EvalValue& out) {
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

bool tryAvgExactComplexFromSum(const EvalValue& sumV, std::size_t itemCount, EvalValue& out) {
  if (itemCount == 0) {
    return false;
  }
  const SmInt n = static_cast<SmInt>(itemCount);
  ExactCartesianComponent lRe{};
  ExactCartesianComponent lIm{};
  ExactCartesianComponent oRe{};
  ExactCartesianComponent oIm{};
  SmInt reI = 0;
  SmInt imI = 0;
  SmInt qRe = 0;
  SmInt qIm = 0;
  if (!tryExtractExactRealComponent(sumV.scalarValue, lRe) || !tryExtractExactImagComponent(sumV.scalarValue, lIm) ||
      !TryExactCartesianComponentToSInt(lRe, reI) || !TryExactCartesianComponentToSInt(lIm, imI) ||
      !TryQuotExactSInt(reI, n, qRe) || !TryQuotExactSInt(imI, n, qIm)) {
    return false;
  }
  ExactCartesianComponentAssignFromSignedInt(oRe, qRe);
  ExactCartesianComponentAssignFromSignedInt(oIm, qIm);
  setScalarComplexFromExactCartesian(out, oRe, oIm);
  return true;
}
#endif

EvalValue makeScalarMaybeExact(SmFloat v) {
  EvalValue out = makeScalar(v);
  SmInt asInt = 0;
  if (tryExtractExactSIntFromFloatStrict(v, asInt)) {
    out.scalarValue.scalarKind = ScalarKind::SInt;
    out.scalarValue.setExactSIntValid(true);
    out.scalarValue.exactSInt = asInt;
    if (asInt >= 0) {
      out.scalarValue.setExactUIntValid(true);
      out.scalarValue.exactUInt = static_cast<SmUInt>(asInt);
    }
  }
  return out;
}

EvalValue makeScalarInt(SmInt v) {
  EvalValue out;
  out.kind = ValueKind::Scalar;
  out.scalarValue.scalarKind = ScalarKind::SInt;
  out.scalarValue.scalar = static_cast<SmFloat>(v);
  out.scalarValue.setExactSIntValid(true);
  out.scalarValue.exactSInt = v;
  out.scalarValue.setExactUIntValid(v >= 0);
  out.scalarValue.exactUInt = (v >= 0) ? static_cast<SmUInt>(v) : 0;
  out.scalarValue.setDecScientificPow63High(false);
  out.scalarValue.imag = 0.0;
  out.scalarValue.imagExactSInt = 0;
  out.scalarValue.imagExactUInt = 0;
  out.scalarValue.setImagExactSIntValid(false);
  out.scalarValue.setImagExactUIntValid(false);
  return out;
}

EvalValue makeScalarUInt(SmUInt v) {
  EvalValue out;
  out.kind = ValueKind::Scalar;
  out.scalarValue.scalarKind = ScalarKind::UInt;
  out.scalarValue.scalar = static_cast<SmFloat>(v);
  out.scalarValue.setExactUIntValid(true);
  out.scalarValue.exactUInt = v;
  if (v <= static_cast<SmUInt>(kIntMax)) {
    out.scalarValue.setExactSIntValid(true);
    out.scalarValue.exactSInt = static_cast<SmInt>(v);
  } else {
    out.scalarValue.setExactSIntValid(false);
    out.scalarValue.exactSInt = 0;
  }
  out.scalarValue.setDecScientificPow63High(false);
  out.scalarValue.imag = 0.0;
  out.scalarValue.imagExactSInt = 0;
  out.scalarValue.imagExactUInt = 0;
  out.scalarValue.setImagExactSIntValid(false);
  out.scalarValue.setImagExactUIntValid(false);
  return out;
}

EvalValue makeScalarUIntPreferInt(SmUInt v) {
  if (v <= static_cast<SmUInt>(kIntMax)) {
    return makeScalarInt(static_cast<SmInt>(v));
  }
  return makeScalarUInt(v);
}

#if SMARTMATH_TIME_VALUES
EvalValue makeScalarTimeMs(SmInt totalMs) {
  EvalValue out;
  out.kind = ValueKind::Scalar;
  out.scalarValue.scalarKind = ScalarKind::Time;
  out.scalarValue.scalar = static_cast<SmFloat>(totalMs) / 1000.0;
  out.scalarValue.exactSInt = totalMs;
  out.scalarValue.setExactSIntValid(false);
  out.scalarValue.setExactUIntValid(false);
  out.scalarValue.setDecScientificPow63High(false);
  out.scalarValue.imag = 0.0;
  out.scalarValue.imagExactSInt = 0;
  out.scalarValue.imagExactUInt = 0;
  out.scalarValue.setImagExactSIntValid(false);
  out.scalarValue.setImagExactUIntValid(false);
  return out;
}

SmInt timeTotalMsFromScalarValue(const EvalValue::ScalarValue& s) {
  return s.exactSInt;
}
#endif

EvalValue makeArrayFromScalars(const std::vector<EvalValue>& v) {
  EvalValue out;
  out.kind = ValueKind::Array;
  out.scalarValue.scalarKind = ScalarKind::FloatingPoint;
  out.arr.resize(v.size());
  for (std::size_t i = 0; i < v.size(); ++i) {
    out.arr[i] = v[i].scalarValue;
  }
  return out;
}

int expandUnpackedArgs(const std::vector<EvalValue>& in, std::vector<EvalValue>& out) {
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

EvalValue negateEvalValue(const EvalValue& v) {
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
    const SmInt ms = timeTotalMsFromScalarValue(v.scalarValue);
    if (ms == kIntMin) {
      return makeScalarTimeMs(kIntMax);
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
    SmFloat re = 0.0;
    SmFloat im = 0.0;
    scalarLoadCartesian(v.scalarValue, re, im);
    return makeScalarComplexFromFloats(-re, -im);
  }
#endif
  if (v.scalarValue.hasDecScientificPow63High()) {
    const SmFloat p63 = kFloatSIntBound;
    if (v.scalarValue.scalar == p63) {
      return makeScalarInt(kIntMin);
    }
    if (v.scalarValue.scalar == -p63) {
      return makeScalarUInt(kIntMinMag);
    }
  }
  if (v.scalarValue.hasExactSInt()) {
    const SmInt x = v.scalarValue.exactSInt;
    if (x == kIntMin) {
      return makeScalarUInt(kIntMinMag);
    }
    return makeScalarInt(-x);
  }
  if (v.scalarValue.hasExactUInt()) {
    const SmUInt u = v.scalarValue.exactUInt;
    if (u == 0) {
      return makeScalarInt(0);
    }
    if (u == kIntMinMag) {
      return makeScalarInt(kIntMin);
    }
    if (u <= static_cast<SmUInt>(kIntMax)) {
      return makeScalarInt(-static_cast<SmInt>(u));
    }
    return makeScalar(-static_cast<SmFloat>(u));
  }
  return makeScalar(-v.scalarValue.scalar);
}

#if SMARTMATH_LAMBDA_FUNCTIONS
EvalValue makeInlineLambdaValue(std::vector<std::string> params, std::string body) {
  EvalValue v;
  v.kind = ValueKind::InlineLambda;
  v.lambdaParams = std::move(params);
  v.lambdaBody = std::move(body);
  return v;
}
#endif

EvalValue makeRationalReduced(SmInt num, SmUInt den) {
  EvalValue out = makeScalarMaybeExact(static_cast<SmFloat>(num) / static_cast<SmFloat>(den));
  scalarApplyReducedRationalPart(out.scalarValue, false, num, den);
  return out;
}

EvalValue applyAbsScalarValue(const EvalValue::ScalarValue& s) {
  const SmFloat x = s.scalar;
  if (std::isnan(x)) {
    return makeScalar(std::numeric_limits<SmFloat>::quiet_NaN());
  }
  if (s.hasExactUInt() && !s.hasExactSInt()) {
    return makeScalarUInt(s.exactUInt);
  }
  if (s.hasExactSInt()) {
    if (s.exactSInt >= 0) {
      return makeScalarInt(s.exactSInt);
    }
    if (s.exactSInt == kIntMin) {
      return makeScalarUInt(kIntMinMag);
    }
    return makeScalarInt(-s.exactSInt);
  }
  if (s.hasExactUInt()) {
    return makeScalarUInt(s.exactUInt);
  }
  return makeScalarMaybeExact(std::fabs(x));
}

EvalValue calcRoundingFn(BuiltinFunctionId id, const EvalValue::ScalarValue& s)
{
  const SmFloat x = s.scalar;
  if (!std::isfinite(x))
    return makeScalar(x);

  if (s.hasExactUInt())
    return makeScalarUInt(s.exactUInt);

  if (s.hasExactSInt())
    return makeScalarInt(s.exactSInt);

  if (x > kMaxExactIntFromFloat || x < -kMaxExactIntFromFloat)
  {
    const SmFloat p63 = kFloatSIntBound;
    // [-2^63, 2^63): kIntMax is not a double, so do not convert it (clang -Wimplicit-const-int-float-conversion).
    if (x < p63 && x >= -p63)
      return makeScalarInt(static_cast<SmInt>(x));

    const SmFloat p64 = kFloatUIntBound;
    if (x >= 0.0 && x < p64 && x == std::trunc(x)) {
      const SmUInt u = static_cast<SmUInt>(x);
      if (static_cast<SmFloat>(u) == x)
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

  SmInt rounded = 0;
  switch (id) {
    case BuiltinFunctionId::Floor:
      rounded = static_cast<SmInt>(std::floor(x));
      break;
    case BuiltinFunctionId::Ceil:
      rounded = static_cast<SmInt>(std::ceil(x));
      break;
    case BuiltinFunctionId::Trunc:
    case BuiltinFunctionId::Int:
      rounded = static_cast<SmInt>(std::trunc(x));
      break;
    case BuiltinFunctionId::Round:
      rounded = static_cast<SmInt>(std::round(x));
      break;
    default:
      break;
  }
  return makeScalarInt(rounded);
}

#if SMARTMATH_COMPLEX_NUMBERS
void scalarSplitCartesianParts(const EvalValue::ScalarValue& sv, SmFloat ar, SmFloat ai,
                               EvalValue::ScalarValue& re, EvalValue::ScalarValue& im) {
  re = sv;
  re.scalar = ar;
  scalarClearImaginary(re);
  im = EvalValue::ScalarValue{};
  im.scalar = ai;
  scalarClearImaginary(im);
  if (sv.hasImagExactSInt()) {
    im.setExactSIntValid(true);
    im.exactSInt = sv.imagExactSInt;
    if (sv.hasImagExactUInt()) {
      im.setExactUIntValid(true);
      im.exactUInt = sv.imagExactUInt;
    }
  }
}

EvalValue calcRoundingFnCartesian(
    BuiltinFunctionId rid,
    const EvalValue::ScalarValue& sv,
    SmFloat ar,
    SmFloat ai) {
  EvalValue::ScalarValue svRe;
  EvalValue::ScalarValue svIm;
  scalarSplitCartesianParts(sv, ar, ai, svRe, svIm);
  const EvalValue outRe = calcRoundingFn(rid, svRe);
  const EvalValue outIm = calcRoundingFn(rid, svIm);
  return setScalarComplexFromEvalRealImagParts(outRe, outIm);
}
#endif

std::string formatRationalParts(SmInt num, SmUInt den) {
  if (num == 0LL) {
    return "0";
  }
  if (den == 1u) {
    return smartmath::intToString(num);
  }
  return smartmath::intToString(num) + "/" + smartmath::uintToString(den);
}

bool tryFormatRationalScalar(const EvalValue::ScalarValue& sv, std::string& outText) {
  if (!sv.hasRenderRational()) {
    return false;
  }
  const SmInt num = sv.exactSInt;
  const SmUInt den = sv.exactUInt;
  if (den == 0u) {
    return false;
  }
  outText = formatRationalParts(num, den);
  return true;
}

#if SMARTMATH_FACTORINT
std::string formatIntPowerParts(SmInt baseV, SmUInt expV) {
  if (expV <= 1u) {
    return smartmath::intToString(baseV);
  }
  return smartmath::intToString(baseV) + "**" + smartmath::uintToString(expV);
}

bool tryFormatIntPowerScalar(const EvalValue::ScalarValue& sv, std::string& outText) {
  if (!sv.hasRenderIntPower() || sv.imagExactUInt <= 1u) {
    return false;
  }
  outText = formatIntPowerParts(sv.imagExactSInt, sv.imagExactUInt);
  return true;
}
#endif

#if SMARTMATH_COMPLEX_NUMBERS
bool tryFormatComplexRationalScalar(const EvalValue::ScalarValue& sv, std::string& outText) {
  if (std::isnan(sv.scalar) || std::isnan(sv.imag)) {
    return false;
  }
  if (!sv.hasRenderRational() && !sv.hasImagRenderRational() && !sv.hasImagExactSInt()) {
    return false;
  }
  std::string rePart;
  if (sv.hasRenderRational()) {
    rePart = formatRationalParts(sv.exactSInt, sv.exactUInt);
  } else if (sv.hasExactSInt()) {
    if (sv.exactSInt != 0LL) {
      rePart = smartmath::intToString(sv.exactSInt);
    }
  } else if (sv.scalar != 0.0) {
    return false;
  }
  std::string imTail;
  if (sv.hasImagRenderRational()) {
    const std::string rp = formatRationalParts(sv.imagExactSInt, sv.imagExactUInt);
    if (sv.imagExactUInt > 1u) {
      imTail = rp + "*i";
    } else if (sv.imagExactSInt == 1LL) {
      imTail = STR_I;
    } else if (sv.imagExactSInt == -1LL) {
      imTail = std::string("-") + STR_I;
    } else {
      imTail = rp + STR_I;
    }
  } else if (sv.hasImagExactSInt()) {
    const SmInt ni = sv.imagExactSInt;
    if (ni == 1LL) {
      imTail = STR_I;
    } else if (ni == -1LL) {
      imTail = std::string("-") + STR_I;
    } else if (ni != 0LL) {
      imTail = smartmath::intToString(ni) + STR_I;
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

std::string formatScalar(const EvalValue& v, RenderBase base, bool supportComplexNumbers) {
  const int baseCode = static_cast<int>(base);
  const SmFloat dval = v.scalarValue.scalar;
#if SMARTMATH_TIME_VALUES
  if (v.scalarValue.scalarKind == ScalarKind::Time) {
    return formatTimeCanonicalFromMs(timeTotalMsFromScalarValue(v.scalarValue));
  }
#endif
#if SMARTMATH_COMPLEX_NUMBERS
  if (supportComplexNumbers && scalarHasNonzeroImaginaryPart(v.scalarValue)) {
    if (base == RenderBase::Dec) {
      std::string ratCx;
      if (tryFormatComplexRationalScalar(v.scalarValue, ratCx)) {
        return ratCx;
      }
    }
    if (base != RenderBase::Dec) {
      return formatComplexScalarWithRenderBase(v.scalarValue, base, v.hasRenderUnsigned(), supportComplexNumbers);
    }
    return formatComplexScalarValue(v.scalarValue, supportComplexNumbers);
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
    if (v.scalarValue.hasExactSInt()) {
      return smartmath::intToString(v.scalarValue.exactSInt);
    }
    if (v.scalarValue.hasExactUInt()) {
      return smartmath::uintToString(v.scalarValue.exactUInt);
    }
    const SmFloat p63 = kFloatSIntBound;
    if (v.scalarValue.hasDecScientificPow63High()) {
      if (dval == p63) {
        return posPow63DecText();
      }
      if (dval == -p63) {
        return negPow63DecText();
      }
    }
    return formatFloatFast(dval);
  }

  SmUInt u = 0;
  if (v.scalarValue.hasExactUInt()) {
    u = v.scalarValue.exactUInt;
  } else if (v.scalarValue.hasExactSInt() && v.scalarValue.exactSInt >= 0) {
    u = static_cast<SmUInt>(v.scalarValue.exactSInt);
  } else if (v.scalarValue.hasExactSInt() && v.scalarValue.exactSInt < 0) {
    const SmInt iv = v.scalarValue.exactSInt;
    return formatSignedMagnitudeForRenderBase(iv, baseCode, v.hasRenderUnsigned());
  } else if (!parseUIntFromFloat(dval, u)) {
    return formatScalar(v, RenderBase::Dec, supportComplexNumbers);
  }

  return formatUnsignedForRenderBase(u, baseCode);
}

std::string valueToString(const EvalValue& v, RenderBase forcedBase, bool supportComplexNumbers) {
  if (v.kind == ValueKind::DefinedFunction) {
    return v.funcRefName;
  }
  if (v.kind == ValueKind::Scalar) {
    return formatScalar(v, forcedBase, supportComplexNumbers);
  }
  if (v.kind == ValueKind::InlineLambda) {
    return "";
  }
  std::string out = "(";
  for (std::size_t i = 0; i < v.arr.size(); ++i) {
    if (i) out += STR_COMMA;
    EvalValue e = scalarFromScalarValue(v.arr[i]);
    e.setRenderUnsigned(v.hasRenderUnsigned());
    out += formatScalar(e, forcedBase, supportComplexNumbers);
  }
  out += ")";
  return out;
}

std::string valueToString(const EvalValue& v, bool supportComplexNumbers) {
  return valueToString(v, v.getRenderBase(), supportComplexNumbers);
}

MathParser::RawResult::CartesianScalar rawCartesianAssignReducedRational(SmInt num,
                                                                                    SmUInt den) {
  MathParser::RawResult::CartesianScalar out;
  SmInt n = num;
  SmUInt d = den;
  const SmInt g = GcdSInt(n, static_cast<SmInt>(d));
  if (g > 0) {
    n /= g;
    d = static_cast<SmUInt>(static_cast<SmInt>(d) / g);
  }
  if (d == 1u) {
    out.kind = MathParser::RawResult::ScalarKind::SInt;
    out.intValue = n;
  } else {
    out.kind = MathParser::RawResult::ScalarKind::Rational;
    out.rational.numerator = n;
    out.rational.denominator = d;
  }
  return out;
}

#if SMARTMATH_COMPLEX_NUMBERS
std::string formatComplexScalarValue(const EvalValue::ScalarValue& sv, bool supportComplexNumbers) {
  std::string ratCx;
  if (tryFormatComplexRationalScalar(sv, ratCx)) {
    return ratCx;
  }
  SmFloat ar = 0.0;
  SmFloat ai = 0.0;
  scalarLoadCartesian(sv, ar, ai);
  if (std::isnan(ar) || std::isnan(ai)) {
    return STR_NAN;
  }
  EvalValue tmp = makeScalar(ar);
  tmp.scalarValue.scalarKind = sv.scalarKind;
  tmp.scalarValue.setExactSIntValid(sv.hasExactSInt());
  tmp.scalarValue.exactSInt = sv.exactSInt;
  tmp.scalarValue.setExactUIntValid(sv.hasExactUInt());
  tmp.scalarValue.exactUInt = sv.exactUInt;
  tmp.scalarValue.setDecScientificPow63High(sv.hasDecScientificPow63High());
  const std::string rePart = formatScalar(tmp, RenderBase::Dec, supportComplexNumbers);
  const SmFloat coeffAbs = std::fabs(ai);
  bool negUnit = false;
  std::string tail;
  if (std::isnan(ai)) {
    tail = std::string("nan*") + STR_I;
  } else if (std::isinf(ai)) {
    tail = (ai < 0.0) ? (std::string("-") + STR_INF + "*" + STR_I)
                        : (std::string(STR_INF) + "*" + STR_I);
  } else if (sv.hasImagExactSInt()) {
    const SmInt ii = sv.imagExactSInt;
    const SmInt ac = ii >= 0 ? ii : -ii;
    if (ac == 1LL) {
      tail = STR_I;
      negUnit = ii < 0;
    } else if (ii < 0) {
      tail = std::string("-") + smartmath::intToString(ac) + STR_I;
    } else {
      tail = smartmath::intToString(ac) + STR_I;
    }
  } else if (coeffAbs == 1.0) {
    tail = STR_I;
    negUnit = ai < 0.0;
  } else if (ai < 0.0) {
    tail = std::string("-") + formatFloatFast(coeffAbs) + STR_I;
  } else {
    tail = formatFloatFast(coeffAbs) + STR_I;
  }
  const bool reZero = (sv.hasExactSInt() && sv.exactSInt == 0) || (!sv.hasExactSInt() && ar == 0.0);
  return assembleComplexDecimalText(rePart, tail, negUnit, reZero);
}

std::string assembleComplexDecimalText(
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

std::string formatComplexScalarWithRenderBase(
    const EvalValue::ScalarValue& sv, RenderBase base, bool asUnsigned, bool supportComplexNumbers) {
  if (!scalarHasNonzeroImaginaryPart(sv) || base == RenderBase::Dec) {
    return formatComplexScalarValue(sv, supportComplexNumbers);
  }
  SmInt im = 0;
  if (!tryGetExactImagSIntStrict(sv, im)) {
    return formatComplexScalarValue(sv, supportComplexNumbers);
  }
  if (im == kIntMin) {
    return formatComplexScalarValue(sv, supportComplexNumbers);
  }

  EvalValue reTmp = makeScalar(0);
  SmFloat ar = 0.0;
  SmFloat ai = 0.0;
  scalarLoadCartesian(sv, ar, ai);
  (void)ai;
  reTmp.scalarValue.scalar = ar;
  reTmp.scalarValue.scalarKind = sv.scalarKind;
  reTmp.scalarValue.setExactSIntValid(sv.hasExactSInt());
  reTmp.scalarValue.exactSInt = sv.exactSInt;
  // A negative real part renders as signed magnitude, not via a two's-complement SmUInt mirror.
  reTmp.scalarValue.setExactUIntValid(sv.hasExactUInt() && !(sv.hasExactSInt() && sv.exactSInt < 0));
  reTmp.scalarValue.exactUInt = sv.exactUInt;
  reTmp.scalarValue.setDecScientificPow63High(sv.hasDecScientificPow63High());
  scalarClearImaginary(reTmp.scalarValue);
  reTmp.setRenderUnsigned(asUnsigned);
  std::string rePart = formatScalar(reTmp, base, supportComplexNumbers);

  bool negIm = false;
  SmInt absIm = 0;
  if (im >= 0) {
    absIm = im;
  } else {
    negIm = true;
    absIm = static_cast<SmInt>(-static_cast<SmInt>(im));
  }
  std::string tail;
  if (absIm == 1LL) {
    tail = STR_I;
  } else {
    EvalValue imTmp = makeScalarInt(absIm);
    // Imaginary coefficient stays signed magnitude; unsigned applies to the real part only.
    imTmp.setRenderUnsigned(false);
    tail = formatScalar(imTmp, base, supportComplexNumbers) + STR_I;
  }

  const bool reZero =
      (sv.hasExactSInt() && sv.exactSInt == 0) || (!sv.hasExactSInt() && ar == 0.0);
  if (reZero) {
    return negIm ? (std::string("-") + tail) : tail;
  }
  return negIm ? (rePart + std::string("-") + tail) : (rePart + std::string("+") + tail);
}
#endif

MathParser::RawResult::CartesianScalar toRawCartesianScalar(const EvalValue::ScalarValue& v, bool imagPart) {
  MathParser::RawResult::CartesianScalar out;
  if (!imagPart) {
#if SMARTMATH_TIME_VALUES
    if (v.scalarKind == ScalarKind::Time) {
      out.kind = MathParser::RawResult::ScalarKind::Time;
      out.intValue = timeTotalMsFromScalarValue(v);
      return out;
    }
#endif
    if (v.hasRenderRational() && (v.hasExactUIntMetadata() || v.exactUInt != 0)) {
      return rawCartesianAssignReducedRational(v.exactSInt, v.exactUInt);
    }
#if SMARTMATH_FACTORINT
    if (v.hasRenderIntPower() && v.imagExactUInt > 1u) {
      out.kind = MathParser::RawResult::ScalarKind::IntPower;
      out.rational.numerator = v.imagExactSInt;
      out.rational.denominator = v.imagExactUInt;
      return out;
    }
#endif
    if (v.hasExactSInt()) {
      out.kind = MathParser::RawResult::ScalarKind::SInt;
      out.intValue = v.exactSInt;
      return out;
    }
    if (v.hasExactUInt()) {
      out.kind = MathParser::RawResult::ScalarKind::UInt;
      out.uintValue = v.exactUInt;
      return out;
    }
    if (v.scalarKind == ScalarKind::SInt) {
      out.kind = MathParser::RawResult::ScalarKind::SInt;
      out.intValue = v.exactSInt;
      return out;
    }
    if (v.scalarKind == ScalarKind::UInt) {
      out.kind = MathParser::RawResult::ScalarKind::UInt;
      out.uintValue = v.exactUInt;
      return out;
    }
    out.kind = MathParser::RawResult::ScalarKind::FloatingPoint;
    out.floatingPoint = v.scalar;
    return out;
  }

#if SMARTMATH_COMPLEX_NUMBERS
  if (v.hasImagRenderRational() && (v.hasImagExactUIntMetadata() || v.imagExactUInt != 0)) {
    return rawCartesianAssignReducedRational(v.imagExactSInt, v.imagExactUInt);
  }
  if (v.hasImagExactSInt()) {
    out.kind = MathParser::RawResult::ScalarKind::SInt;
    out.intValue = v.imagExactSInt;
    return out;
  }
  if (v.hasImagExactUInt()) {
    if (v.imagExactUInt <= static_cast<SmUInt>(kIntMax)) {
      out.kind = MathParser::RawResult::ScalarKind::SInt;
      out.intValue = static_cast<SmInt>(v.imagExactUInt);
    } else {
      out.kind = MathParser::RawResult::ScalarKind::UInt;
      out.uintValue = v.imagExactUInt;
    }
    return out;
  }
  out.kind = MathParser::RawResult::ScalarKind::FloatingPoint;
  out.floatingPoint = v.imag;
#endif
  return out;
}

MathParser::RawResult::Scalar toRawScalar(const EvalValue::ScalarValue& v) {
  MathParser::RawResult::Scalar out;
#if SMARTMATH_COMPLEX_NUMBERS
  if (scalarHasNonzeroImaginaryPart(v)) {
    out.kind = MathParser::RawResult::ScalarKind::Complex;
    out.real = toRawCartesianScalar(v, false);
    out.imag = toRawCartesianScalar(v, true);
    return out;
  }
#endif
  out.real = toRawCartesianScalar(v, false);
  out.kind = out.real.kind;
  out.imag.clear();
  return out;
}

MathParser::RawResult toRawResult(const EvalValue& v) {
  MathParser::RawResult out;
  if (v.kind == ValueKind::DefinedFunction) {
    return out;
  }
  if (v.kind == ValueKind::Scalar) {
    out.kind = MathParser::RawResult::Kind::Scalar;
    out.scalar = toRawScalar(v.scalarValue);
    return out;
  }

  out.kind = MathParser::RawResult::Kind::Array;
  out.array.reserve(v.arr.size());
  for (std::size_t i = 0; i < v.arr.size(); ++i) {
    out.array.push_back(toRawScalar(v.arr[i]));
  }
  return out;
}

void EvalContext::fail(const std::string& msg) {
  if (parseError) {
    return;
  }
  parseError = true;
  errorText = msg;
}

void EvalContext::fail(ParseErrorId id) {
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
      STR_INTERNAL_ERROR_IN_AGGREGATE_BUILTIN,
      STR_INTERNAL_ERROR_IN_SCALAR_BINARY_BUILTIN,
      STR_INTERNAL_ERROR_IN_UNARY_MATH_BUILTIN,
      STR_UNEXPECTED_TOKEN_AFTER_EXPRESSION,
      STR_SCALAR_ONLY_EXPRESSION_ENCOUNTERED_NON,
      STR_PARSE_FAILED,
      STR_UNEXPECTED_TOKEN,
      STR_UNEXPECTED_INPUT,
      STR_CANNOT_ASSIGN_TO_AN_ARRAY_SLICE,
  };
  fail(texts[static_cast<std::size_t>(id)]);
}

void EvalContext::failCall(CallErrorId id, const std::string& fnName, std::size_t n) {
  switch (id) {
    case CallErrorId::Numeric:
      fail(STR_NUMERIC_ERROR_IN + fnName + STR_PAR);
      return;
    case CallErrorId::AtLeastN:
      fail(fnName + STR_PAR_EXPECTS_AT_LEAST + std::to_string(n) + (n == 1 ? STR_ARGUMENT : STR_ARGUMENTS));
      return;
    case CallErrorId::ExactCount:
      fail(fnName + STR_PAR_EXPECTS + std::to_string(n) + STR_ARGUMENT_PAR_S);
      return;
    case CallErrorId::ScalarValues:
      fail(fnName + STR_PAR_EXPECTS_SCALAR_VALUES);
      return;
    case CallErrorId::IntegerValues:
      fail(fnName + STR_PAR_EXPECTS_INTEGER_VALUES);
      return;
    case CallErrorId::ScalarMinMax:
      fail(fnName + STR_PAR_EXPECTS_SCALAR_MIN_SLASH);
      return;
    case CallErrorId::NonNegativeInteger:
      fail(fnName + STR_PAR_EXPECTS_A_NON_DASH);
      return;
    case CallErrorId::PositiveInteger:
      fail(fnName + STR_PAR_EXPECTS_A_POSITIVE_INTEGER);
      return;
    case CallErrorId::RangeStepZero:
      fail(fnName + STR_PAR_STEP_MUST_NOT_BE_ZERO);
      return;
    case CallErrorId::TooManyValues:
      fail(fnName + STR_PAR_PRODUCES_TOO_MANY_VALUES);
      return;
  }
}

void EvalContext::failHint(const std::string& hintText) {
  fail(STR_HINT_PREFIX + hintText);
}

void EvalContext::failPrefixedLiteral(char prefixChar) {
  switch (prefixChar) {
    case 'x':
    case 'X':
      fail(ParseErrorId::InvalidHexLiteral);
      return;
    case 'b':
    case 'B':
      fail(ParseErrorId::InvalidBinaryLiteral);
      return;
    case 'o':
    case 'O':
      fail(ParseErrorId::InvalidOctalLiteral);
      return;
    default:
      fail(ParseErrorId::InvalidNumericLiteral);
      return;
  }
}

void EvalContext::failRecursiveUserFunction(const std::string& fnName) {
  fail(STR_RECURSIVE_USER_FUNCTION_CALL_COLON + fnName);
}

void EvalContext::failMissingClosingParenLike() {
  const unsigned char c = static_cast<unsigned char>(*p);
  if (localIsIdentStart(*p) || std::isdigit(c) || *p == '.') {
    fail(ParseErrorId::UnexpectedToken);
  } else {
    fail(ParseErrorId::MissingClosingParenthesis);
  }
}

void EvalContext::failMissingClosingBracketAfterIndex() {
  if (*p == ')') {
    fail(ParseErrorId::MismatchedClosingParenthesis);
  } else if (*p == '}') {
    fail(ParseErrorId::MismatchedClosingBrace);
  } else {
    fail(ParseErrorId::MissingClosingBracket);
  }
}

}  // namespace smartmath_internal

MathParser::MathParser() {
  assert(functionNames().size() == static_cast<std::size_t>(BuiltinFunctionId::Count));
  assert(operatorNames().size() == static_cast<std::size_t>(OperatorNameId::Count));
  addConst(STR_PI, kPi);
  addConst(STR_E, std::exp(static_cast<SmFloat>(1)));
  addConst(STR_INF, std::numeric_limits<SmFloat>::infinity());
  addConst(STR_NAN, std::numeric_limits<SmFloat>::quiet_NaN());
#if SMARTMATH_TIME_VALUES
  setVariable(STR_MILLISECOND, makeScalarTimeMs(1LL));
  setVariable(STR_SECOND, makeScalarTimeMs(1000LL));
  setVariable(STR_MINUTE, makeScalarTimeMs(60000LL));
  setVariable(STR_HOUR, makeScalarTimeMs(3600000LL));
  setVariable(STR_DAY, makeScalarTimeMs(86400000LL));
#endif
  setVariable(STR_ANS, makeScalarInt(0));
}

bool MathParser::isReservedConstantIdent(const std::string& ident) const {
#if SMARTMATH_COMPLEX_NUMBERS
  if (getSupportComplexNumbers()) {
    std::string low = ident;
    std::transform(low.begin(), low.end(), low.begin(), [](unsigned char c) { return static_cast<char>(std::tolower(c)); });
    if (low == STR_I) {
      return true;
    }
  }
#endif
  if (isReservedBuiltinConstantName(ident, getSupportTimeValues())) {
    return true;
  }
  return false;
}

std::string MathParser::getReservedIdentifierError(const std::string& ident) const {
  if (isReservedFunctionName(ident)) {
    return std::string(STR_RESERVED_FUNCTION_NAME_COLON) + ident;
  }
  if (isReservedConstantIdent(ident)) {
    return std::string(STR_RESERVED_CONSTANT_NAME_COLON) + ident;
  }
  return "";
}

bool MathParser::isBareFunctionNameAtExpressionTail(EvalContext& ctx) const {
  if (!ctx.p) {
    return false;
  }
  skipSpaces(ctx.p);
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
      ctx.fail(ParseErrorId::MismatchedClosingParenthesis);
      return true;
    case ']':
      ctx.fail(ParseErrorId::MismatchedClosingBracket);
      return true;
    case '}':
      ctx.fail(ParseErrorId::MismatchedClosingBrace);
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
  if (begin >= end || !localIsIdentStart(*begin)) {
    return false;
  }
  const char* identEnd = begin + 1;
  while (identEnd < end && localIsIdentChar(*identEnd)) {
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

bool MathParser::tryHandleBareFunctionNameCommonTail(
    EvalContext& ctx,
    const std::string& ident,
    const char* identStartForCloser) const {
  skipSpaces(ctx.p);
  if (*ctx.p == ';') {
    ctx.fail(ParseErrorId::UnexpectedToken);
    return true;
  }
  if (!isBareFunctionNameAtExpressionTail(ctx)) {
    if (identStartForCloser != nullptr
        && trySetBareFunctionImmediateCloserError(ctx, identStartForCloser)) {
      return true;
    }
    ctx.fail(buildUnknownVariableErrorText(ident));
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
    ctx.fail(buildUnknownVariableErrorText(ident));
    return true;
  }
  std::string hint = getBuiltinFunctionMissingCallHint(bid);
  if (!hint.empty()) ctx.failHint(hint);
  else ctx.failHint(ident + STR_PAR_DOTDOTDOT);
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
  if (identMayBeBareBuiltinName(ident) && trySetMissingFunctionCallError(ctx, ident)) {
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
    ctx.fail(buildUnknownVariableErrorText(ident));
    return true;
  }
  if (uf) {
    ctx.fail(std::string(STR_USER_DEFINED_FUNCTION_COLON) + formatUserFunctionSignature(*uf));
  } else {
    UserFunction stub;
    stub.name = ident;
    stub.params = *compilingParams;
    ctx.fail(std::string(STR_USER_DEFINED_FUNCTION_COLON) + formatUserFunctionSignature(stub));
  }
  return true;
}

std::string MathParser::formatUserFunctionSignature(const UserFunction& uf) {
  std::string sig = uf.name + "(";
  for (std::size_t i = 0; i < uf.params.size(); ++i) {
    if (i > 0) {
      sig += STR_COMMA;
    }
    sig += uf.params[i];
  }
  sig += ")";
  return sig;
}

bool MathParser::finishSubEvalFromCompiledProgram(EvalContext& parent, EvalContext& sub) const {
  if (sub.referencedUnresolvedName || !sub.unknownVarsText.empty() || !sub.unknownFuncsText.empty()) {
    parent.referencedUnresolvedName = true;
  }
  if (!sub.unknownVarsText.empty()) {
    parent.fail(buildUnknownVariableErrorText(sub.unknownVarsText));
    return false;
  }
  if (sub.parseError) {
    parent.fail(sub.errorText);
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
    const std::vector<std::string>* boundBodyParamNames,
    std::vector<AstStatement>& outProgram) {
  EvalContext parseCtx;
  parseCtx.p = expr.c_str();
  parseCtx.sourceExpr = expr;
  parseCtx.userFunctionCallStack = callStack;
  parseCtx.validatingUserFunctionName = validatingUserFunctionName;
  parseCtx.compilingUserFunctionParams = parentCtx.compilingUserFunctionParams;
  parseCtx.boundBodyParamNames = boundBodyParamNames;
  parseCtx.suppressBareFunctionTailHint = suppressBareFunctionTailHint;
  // Bodies are expressions, same as definition checking. ``name=expr`` stays a comparison.
  skipSpaces(parseCtx.p);
  parseCtx.start = parseCtx.p;
  auto ex = parseExpression(parseCtx);
  if (parseCtx.parseError || !ex) {
    parentCtx.fail(parseCtx.errorText.empty() ? STR_FAILED_TO_PARSE_USER_FUNCTION_BODY : parseCtx.errorText);
    return false;
  }
  skipSpaces(parseCtx.p);
  if (*parseCtx.p != '\0') {
    parentCtx.fail(ParseErrorId::UnexpectedInput);
    return false;
  }
  outProgram.clear();
  AstStatement st;
  st.kind = AstStatement::Kind::Expr;
  st.expr = std::move(ex);
  outProgram.emplace_back(std::move(st));
  return true;
}

bool MathParser::handleUnknownIdentifier(EvalContext& ctx, const std::string& ident, std::string& unknownList) const {
  if (isLogicalBinaryOperatorKeyword(ident)) {
    ctx.fail(ParseErrorId::UnexpectedToken);
    return true;
  }
  if (!ident.empty()) {
    if (identMayBeBareBuiltinName(ident) && trySetMissingFunctionCallError(ctx, ident)) {
      return true;
    }
    if ((!userFunctionIndex_.empty() || ctx.compilingUserFunctionParams != nullptr)
        && trySetBareUserFunctionNameError(ctx, ident, nullptr)) {
      return true;
    }
  }
  appendUniqueName(unknownList, ident);
  ctx.referencedUnresolvedName = true;
  return false;
}

bool MathParser::tryResolveVariableValue(
    const Expr& e,
    const std::unordered_map<std::string, EvalValue>* scopedVars,
    EvalValue& out) const {
  if (scopedVars) {
    auto it = scopedVars->find(e.name);
    if (it != scopedVars->end()) {
      out = it->second;
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

std::string MathParser::validateParameterNames(const std::vector<std::string>& fnParams) const {
  std::unordered_map<std::string, bool> seen;
  for (const auto& p : fnParams) {
    if (seen.find(p) != seen.end()) {
      return STR_DUPLICATE_PARAMETER_NAME;
    }
    if (p == STR_ANS) {
      return std::string(STR_RESERVED_BUILTIN_VARIABLE_NAME_COLON) + p;
    }
    if (isReservedConstantIdent(p)) {
      return std::string(STR_RESERVED_CONSTANT_NAME_COLON) + p;
    }
    if (isReservedFunctionName(p)) {
      return std::string(STR_RESERVED_FUNCTION_PARAM_NAME_COLON) + p;
    }
    seen[p] = true;
  }
  return "";
}

std::string MathParser::validateUserFunctionDefinitionNames(
    const std::string& fnName,
    const std::vector<std::string>& fnParams) const {
  if (isReservedBuiltinVariableNameForUserFunctionDefinition(fnName)) {
    return std::string(STR_RESERVED_BUILTIN_VARIABLE_NAME_COLON) + fnName;
  }
  const std::string reservedErr = getReservedIdentifierError(fnName);
  if (!reservedErr.empty()) {
    return reservedErr;
  }
  return validateParameterNames(fnParams);
}

std::string MathParser::getUserFunctionDefinitionErrorText(
    const std::string& fnName,
    const std::vector<std::string>& fnParams,
    const std::string& fnExpr,
    const bool evaluateBody) {
  const std::string udfNameErr = validateUserFunctionDefinitionNames(fnName, fnParams);
  if (!udfNameErr.empty()) {
    return udfNameErr;
  }
  if (bodyIsEffectivelyEmpty(fnExpr)) {
    return STR_FUNCTION_BODY_IS_EMPTY;
  }
  if (udfBodyCallsDefinedFunction(fnExpr, fnName)) {
    return STR_RECURSIVE_USER_FUNCTION_CALL_COLON + fnName;
  }
  EvalContext bodyCtx;
  std::unordered_map<std::string, std::vector<std::string>> validatingUserFunctionParams;
  validatingUserFunctionParams[fnName] = fnParams;
  bodyCtx.compilingUserFunctionParams = &validatingUserFunctionParams;
  bodyCtx.validatingUserFunctionName = fnName.c_str();
  bodyCtx.boundBodyParamNames = &fnParams;
  std::unique_ptr<Expr> ex;
  if (!tryParseSyntaxCheckedExpression(bodyCtx, fnExpr, evaluateBody, ex)) {
    if (!bodyCtx.errorText.empty()) {
      return bodyCtx.errorText;
    }
    return STR_FAILED_TO_PARSE_USER_FUNCTION_BODY;
  }
  // Unknown names stay non-fatal, matching Basic: a later statement may define them.
  return "";
}

bool MathParser::tryParseSyntaxCheckedExpression(
    EvalContext& bodyCtx,
    const std::string& source,
    const bool checkCallShapes,
    std::unique_ptr<Expr>& out) {
  bodyCtx.sourceExpr = source;
  bodyCtx.p = bodyCtx.sourceExpr.c_str();
  skipSpaces(bodyCtx.p);
  bodyCtx.start = bodyCtx.p;
  out = parseExpression(bodyCtx);
  if (bodyCtx.parseError || !out) {
    if (bodyCtx.errorText.empty()) {
      bodyCtx.fail(STR_FAILED_TO_PARSE_USER_FUNCTION_BODY);
    }
    return false;
  }
  skipSpaces(bodyCtx.p);
  if (*bodyCtx.p != '\0') {
    bodyCtx.fail(ParseErrorId::UnexpectedInput);
    return false;
  }
  if (!checkCallShapes) {
    return true;
  }
  bodyCtx.parseError = false;
  bodyCtx.errorText.clear();
  bodyCtx.unknownVarsText.clear();
  bodyCtx.unknownFuncsText.clear();
  if (!checkUnevaluatedCallShapes(bodyCtx, *out)) {
    if (bodyCtx.errorText.empty()) {
      bodyCtx.fail(STR_FAILED_TO_PARSE_USER_FUNCTION_BODY);
    }
    return false;
  }
  return true;
}

bool MathParser::trySetUserFunctionDefinitionError(
    EvalContext& ctx,
    const std::string& fnName,
    const std::vector<std::string>& fnParams,
    const std::string& fnExpr) {
  // Parse-only here. Call arity is checked when the definition runs, after earlier
  // statements have been applied. Basic checks the body as soon as it is parsed.
  const std::string err = getUserFunctionDefinitionErrorText(fnName, fnParams, fnExpr, false);
  if (err.empty()) {
    return false;
  }
  ctx.fail(err.c_str());
  return true;
}

std::string MathParser::validateAssignmentTargetName(const std::string& ident) const {
  return getReservedIdentifierError(ident);
}

bool MathParser::trySetUnknownNameError(const EvalContext& ctx) {
  lastError_ = buildUnknownNameErrorText(ctx.unknownVarsText, ctx.unknownFuncsText);
  return !lastError_.empty();
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
    std::string ident = consumeLowerIdentToken(ctx.p);
    skipSpaces(ctx.p);
    if (*ctx.p == '=' && ctx.p[1] != '=') {
      const std::string assignNameErr = validateAssignmentTargetName(ident);
      if (!assignNameErr.empty()) {
        ctx.fail(assignNameErr);
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
        ctx.fail(ParseErrorId::UnexpectedToken);
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
  if (peekIdentFollowedByIndexThenAssignEquals(assignSave)) {
    ctx.p = assignSave;
    std::string ident = consumeLowerIdentToken(ctx.p);
    const std::string assignNameErr = validateAssignmentTargetName(ident);
    if (!assignNameErr.empty()) {
      ctx.fail(assignNameErr);
      return false;
    }
    skipSpaces(ctx.p);
    auto bracket = parseOneBracket(ctx, nullptr);
    if (ctx.parseError || !bracket) {
      return false;
    }
    skipSpaces(ctx.p);
    if (*ctx.p != '=' || ctx.p[1] == '=') {
      ctx.fail(ParseErrorId::UnexpectedToken);
      return false;
    }
    ++ctx.p;
    auto ex = parseExpression(ctx);
    if (ctx.parseError || !ex) {
      return false;
    }
    AstStatement st;
    st.kind = AstStatement::Kind::IndexAssign;
    st.assignName = std::move(ident);
    st.indexExpr = std::move(bracket);
    st.expr = std::move(ex);
    out.emplace_back(std::move(st));
    return true;
  }
  return tryAppendParsedExpressionStatement(ctx, out);
}

bool MathParser::consumeProgramStatementSeparator(EvalContext& ctx) {
  skipSpaces(ctx.p);
  if (*ctx.p == ';') {
    ++ctx.p;
    return true;
  }
  if (*ctx.p == '\0') {
    return false;
  }
  ctx.fail(ParseErrorId::UnexpectedTokenAfterExpression);
  return false;
}

bool MathParser::hasExprParseFailure(const EvalContext& ctx, const std::unique_ptr<Expr>& node) {
  return ctx.parseError || !node;
}

#if SMARTMATH_TIME_VALUES
bool MathParser::rejectBinaryBuiltinTimeOperands(
    EvalContext& ctx,
    const EvalValue& left,
    const EvalValue& right) const {
  if (evalValueInvolvesTime(left) || evalValueInvolvesTime(right)) {
    ctx.fail(ParseErrorId::IncompatibleOperands);
    return true;
  }
  return false;
}
#endif

bool MathParser::RejectSIntBinaryOperands(
    EvalContext& ctx,
    const EvalValue& left,
    const EvalValue& right,
    const bool isModulo) const {
#if SMARTMATH_TIME_VALUES
  if (getSupportTimeValues() && (evalValueInvolvesTime(left) || evalValueInvolvesTime(right))) {
    if (isModulo) {
      ctx.fail(ParseErrorId::ModuloIntegerOperands);
    } else {
      ctx.fail(ParseErrorId::IncompatibleOperands);
    }
    return true;
  }
#endif
#if SMARTMATH_COMPLEX_NUMBERS
  if (getSupportComplexNumbers() &&
      (evalValueHasNonzeroImaginary(left) || evalValueHasNonzeroImaginary(right))) {
    ctx.fail(ParseErrorId::IncompatibleOperands);
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
    ctx.fail(ParseErrorId::IncompatibleOperands);
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
    ctx.fail(ParseErrorId::IncompatibleOperands);
    return makeScalar(0);
  }
#endif
  bool ok = false;
  EvalValue out = mapBinaryBuiltinMathFunction(args[0], args[1], id, ok, getSupportComplexNumbers());
  if (!ok) {
    if (!ctx.parseError) {
      setBinaryBuiltinBroadcastFailure(ctx, fnName, args[0], args[1], numericErrorOnMapFailure ? 2 : 3);
    }
    return makeScalar(0);
  }
  return out;
}

bool MathParser::acceptProducedItemCount(EvalContext& ctx, const std::string& fnName, SmUInt count) const {
  if (count > static_cast<SmUInt>(BUILTIN_MAX_PRODUCED_ITEMS)) {
    ctx.failCall(CallErrorId::TooManyValues, fnName);
    return false;
  }
  return true;
}

bool MathParser::tryConsumeCommaArgSeparator(EvalContext& ctx, bool& hasComma) const {
  hasComma = false;
  if (*ctx.p != ',') {
    return true;
  }
  hasComma = true;
  ++ctx.p;
  skipSpaces(ctx.p);
  if (*ctx.p == ')' || *ctx.p == ',') {
    ctx.fail(ParseErrorId::UnexpectedComma);
    return false;
  }
  return true;
}

bool MathParser::tryConsumeLogicalBinaryOperator(EvalContext& ctx, OperatorNameId keywordId, char symbol) const {
  skipSpaces(ctx.p);
  if (ctx.p[0] == symbol && ctx.p[1] == symbol) {
    ctx.p += 2;
    return true;
  }
  return consumeKeyword(ctx.p, opName(keywordId).c_str());
}

bool MathParser::parseParenthesizedExprList(
    EvalContext& ctx,
    std::vector<std::unique_ptr<Expr>>& outValues) {
  skipSpaces(ctx.p);
  if (*ctx.p == ')') {
    return true;
  }
  while (true) {
    outValues.emplace_back(parseExpression(ctx));
    if (ctx.parseError) {
      return false;
    }
    skipSpaces(ctx.p);
    bool hasComma = false;
    if (!tryConsumeCommaArgSeparator(ctx, hasComma)) {
      return false;
    }
    if (!hasComma) {
      return true;
    }
  }
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
  skipSpaces(ctx.p);
  if (*ctx.p != '&' || ctx.p[1] == '&') {
    return false;
  }
  ++ctx.p;
  outOp = Expr::BinaryOp::BitAnd;
  return true;
}

bool MathParser::tryConsumeBitXorOp(EvalContext& ctx, Expr::BinaryOp& outOp) {
  skipSpaces(ctx.p);
  if (*ctx.p != '^') {
    return false;
  }
  ++ctx.p;
  outOp = Expr::BinaryOp::BitXor;
  return true;
}

bool MathParser::tryConsumeBitOrOp(EvalContext& ctx, Expr::BinaryOp& outOp) {
  skipSpaces(ctx.p);
  if (*ctx.p != '|' || ctx.p[1] == '|') {
    return false;
  }
  ++ctx.p;
  outOp = Expr::BinaryOp::BitOr;
  return true;
}

bool MathParser::tryConsumeShiftOp(EvalContext& ctx, Expr::BinaryOp& outOp) {
  skipSpaces(ctx.p);
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
  skipSpaces(ctx.p);
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
  skipSpaces(ctx.p);
  if (*ctx.p == '*' && ctx.p[1] != '*') {
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
  skipSpaces(ctx.p);
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


#if SMARTMATH_COMPLEX_NUMBERS
void MathParser::applyComplexCaretPrincipalEval(
    const EvalValue::ScalarValue& lv,
    const EvalValue::ScalarValue& rv,
    EvalValue& outS) const {
  SmFloat ar = 0.0;
  SmFloat ai = 0.0;
  SmFloat br = 0.0;
  SmFloat bi = 0.0;
  scalarLoadCartesian(lv, ar, ai);
  scalarLoadCartesian(rv, br, bi);
  SmFloat powR = 0.0;
  SmFloat powI = 0.0;
  complexPowPrincipal(ar, ai, br, bi, powR, powI);
  if (tryRefinePowPrincipalToExactScalarResult(lv, rv, powR, powI, outS)) {
    return;
  }
  outS = makeScalarComplexFromFloats(powR, powI);
}
#endif

MathParser::EvalValue MathParser::applyUnarySqrtEval(const EvalValue::ScalarValue& s) const {
#if SMARTMATH_COMPLEX_NUMBERS
  if (getSupportComplexNumbers()) {
    if (scalarHasNonzeroImaginaryPart(s)) {
      return applySqrtComplexPrincipalUnary(s);
    }
    const SmFloat x = s.scalar;
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

std::string MathParser::getError() const {
  return lastError_;
}

std::string MathParser::getResult() const {
  if (!lastError_.empty() || !hasResult_) {
    return "";
  }
  return valueToString(lastResult_, getSupportComplexNumbers());
}

std::string MathParser::getResultWithBase(RenderBase base) const {
  if (!lastError_.empty() || !hasResult_) {
    return "";
  }
  return valueToString(lastResult_, base, getSupportComplexNumbers());
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

void MathParser::addConst(const std::string& constName, SmInt intValue) {
  EvalValue v = makeScalarInt(intValue);
  setVariable(toLower(constName), v);
}

void MathParser::addConst(const std::string& constName, SmFloat dblValue) {
  EvalValue v = makeScalar(dblValue);
  setVariable(toLower(constName), v);
}

#if SMARTMATH_COMPLEX_NUMBERS
bool MathParser::tryApplyComplexBinaryScalars(
    const EvalValue::ScalarValue& lv, const EvalValue::ScalarValue& rv, char op, EvalValue& outS) const {
  if (!getSupportComplexNumbers()) {
    return false;
  }
  if (op != '+' && op != '-' && op != '*' && op != '/' && op != '^') {
    return false;
  }
  SmFloat ar = 0.0;
  SmFloat ai = 0.0;
  SmFloat br = 0.0;
  SmFloat bi = 0.0;
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
  SmFloat floatR = 0.0;
  SmFloat floatI = 0.0;
  complexCartesianBinary(ar, ai, br, bi, op, floatR, floatI);
  outS = makeScalarComplexFromFloats(floatR, floatI);
  return true;
}
#endif

#if SMARTMATH_TIME_VALUES
bool MathParser::scalarValueIsTime(const EvalValue::ScalarValue& s) const {
  if (!getSupportTimeValues()) {
    return false;
  }
  return s.scalarKind == ScalarKind::Time;
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

#if SMARTMATH_TIME_VALUES
bool MathParser::scalarMsForCompare(const EvalValue::ScalarValue& sv, SmInt& outMs) const {
  if (scalarValueIsTime(sv)) {
    outMs = timeTotalMsFromScalarValue(sv);
    return true;
  }
  if (std::isnan(sv.scalar)) {
    return false;
  }
  const SmFloat ms = sv.scalar * 1000.0;
  if (tryRoundHalfUpFloatToSInt(ms, outMs)) {
    return true;
  }
  // Saturate so ordering against time values (always in range) stays correct.
  outMs = (ms > 0.0) ? kIntMax : kIntMin;
  return true;
}
#endif

#if SMARTMATH_COMPLEX_NUMBERS
bool MathParser::rejectBuiltinArgsWithComplexImaginary(
    EvalContext& ctx,
    const std::vector<EvalValue>& args) const {
  if (!getSupportComplexNumbers()) {
    return false;
  }
  for (const auto& v : args) {
    if (evalValueHasNonzeroImaginary(v)) {
      ctx.fail(ParseErrorId::IncompatibleOperands);
      return true;
    }
  }
  return false;
}
#endif

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
#if SMARTMATH_COMPLEX_NUMBERS
  const bool ha = getSupportComplexNumbers() && scalarHasNonzeroImaginaryPart(sa);
  const bool hb = getSupportComplexNumbers() && scalarHasNonzeroImaginaryPart(sb);
  if ((ha || hb) && (ta || tb)) {
    if (policy == CmpScalarIncompatiblePolicy::SetError && ctx != nullptr) {
      ctx->fail(ParseErrorId::IncompatibleOperands);
    }
    cmpOut = (policy == CmpScalarIncompatiblePolicy::SortUniqueReturnOne) ? 1 : 0;
    return policy == CmpScalarIncompatiblePolicy::SortUniqueReturnOne;
  }
#endif
#if SMARTMATH_TIME_VALUES
  if (ta || tb) {
    SmInt ams = 0;
    SmInt bms = 0;
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
    SmFloat ar = 0.0;
    SmFloat ai = 0.0;
    SmFloat br = 0.0;
    SmFloat bi = 0.0;
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

#if SMARTMATH_TIME_VALUES
MathParser::EvalValue MathParser::evalValueFromTimeMs(BuiltinFunctionId id, SmInt ms) const {
  if (id == BuiltinFunctionId::Milliseconds) {
    return makeScalarInt(ms);
  }
  if (id == BuiltinFunctionId::Seconds) {
    return makeScalar(static_cast<SmFloat>(ms) / 1000.0);
  }
  if (id == BuiltinFunctionId::Minutes) {
    return makeScalar(static_cast<SmFloat>(ms) / 60000.0);
  }
  if (id == BuiltinFunctionId::Hours) {
    return makeScalar(static_cast<SmFloat>(ms) / 3600000.0);
  }
  return makeScalar(static_cast<SmFloat>(ms) / 86400000.0);
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
      ctx.fail(fnName + STR_TIME_EXPECTS_TIME_ARG);
      return makeScalar(0);
    }
    outs.push_back(evalValueFromTimeMs(id, timeTotalMsFromScalarValue(item)));
  }
  return makeArrayFromScalars(outs);
}
#endif

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
      ctx.fail(ParseErrorId::IncompatibleOperands);
      return true;
    }
  }
#endif
  if (op != '+' && op != '-' && op != '*' && op != '/') {
    ctx.fail(ParseErrorId::IncompatibleOperands);
    return true;
  }
  if ((!lt && !std::isfinite(lv.scalar)) || (!rt && !std::isfinite(rv.scalar))) {
    ctx.fail(STR_TIME_NON_FINITE);
    return true;
  }
  SmInt lms = 0;
  SmInt rms = 0;
  if (lt) {
    lms = timeTotalMsFromScalarValue(lv);
  } else if (!tryRoundHalfUpFloatToSInt(lv.scalar * 1000.0, lms)) {
    ctx.fail(ParseErrorId::IncompatibleOperands);
    return true;
  }
  if (rt) {
    rms = timeTotalMsFromScalarValue(rv);
  } else if (!tryRoundHalfUpFloatToSInt(rv.scalar * 1000.0, rms)) {
    ctx.fail(ParseErrorId::IncompatibleOperands);
    return true;
  }
  if (op == '+') {
    SmInt o = 0;
    if (!tryAddTimeMsChecked(lms, rms, o)) {
      ctx.fail(ParseErrorId::IncompatibleOperands);
      return true;
    }
    outS = makeScalarTimeMs(o);
    return true;
  }
  if (op == '-') {
    SmInt o = 0;
    if (!trySubTimeMsChecked(lms, rms, o)) {
      ctx.fail(ParseErrorId::IncompatibleOperands);
      return true;
    }
    outS = makeScalarTimeMs(o);
    return true;
  }
  if (op == '*') {
    if (lt && rt) {
      ctx.fail(ParseErrorId::IncompatibleOperands);
      return true;
    }
    const SmInt baseMs = lt ? lms : rms;
    const SmFloat mult = lt ? rv.scalar : lv.scalar;
    if (!std::isfinite(mult)) {
      ctx.fail(ParseErrorId::IncompatibleOperands);
      return true;
    }
    SmInt o = 0;
    if (!tryRoundHalfUpFloatToSInt(static_cast<SmFloat>(baseMs) * mult, o)) {
      ctx.fail(ParseErrorId::IncompatibleOperands);
      return true;
    }
    outS = makeScalarTimeMs(o);
    return true;
  }
  if (lt && rt) {
    if (rms == 0) {
      ctx.fail(ParseErrorId::IncompatibleOperands);
      return true;
    }
    const SmFloat ratio = (static_cast<SmFloat>(lms) / 1000.0) / (static_cast<SmFloat>(rms) / 1000.0);
    if (!std::isfinite(ratio)) {
      ctx.fail(ParseErrorId::IncompatibleOperands);
      return true;
    }
    outS = makeScalarMaybeExact(ratio);
    return true;
  }
  if (lt && !rt) {
    if (rv.scalar == 0.0) {
      ctx.fail(ParseErrorId::IncompatibleOperands);
      return true;
    }
    const SmFloat td = (static_cast<SmFloat>(lms) / 1000.0) / rv.scalar;
    SmInt o = 0;
    if (!tryRoundHalfUpFloatToSInt(td * 1000.0, o)) {
      ctx.fail(ParseErrorId::IncompatibleOperands);
      return true;
    }
    outS = makeScalarTimeMs(o);
    return true;
  }
  ctx.fail(ParseErrorId::IncompatibleOperands);
  return true;
}
#endif

inline bool tryApplyCheckedAddSubMulOrSignedUIntPromote(
    char op,
    SmInt li,
    SmInt ri,
    EvalValue& outS) {
  if (op != '+' && op != '-' && op != '*') {
    return false;
  }
  SmInt outI = 0;
  const bool okInt =
      (op == '*') ? checkedMulLL(li, ri, outI) :
      (op == '+') ? checkedAddLL(li, ri, outI) :
                    checkedSubLL(li, ri, outI);
  if (okInt) {
    outS = makeScalarInt(outI);
    return true;
  }
  if (li >= 0 && ri >= 0) {
    const SmUInt lu = static_cast<SmUInt>(li);
    const SmUInt ru = static_cast<SmUInt>(ri);
    if (op == '+') {
      SmUInt outU = 0;
      if (tryAddUIntChecked(lu, ru, outU)) {
        outS = makeScalarUInt(outU);
        return true;
      }
    } else if (op == '*') {
      SmUInt outU = 0;
      if (tryMulUIntChecked(lu, ru, outU)) {
        outS = makeScalarUInt(outU);
        return true;
      }
    }
  }
  return false;
}

inline bool tryApplyUIntAddOrMul(char op, SmUInt lu, SmUInt ru, EvalValue& outS) {
  if (op != '+' && op != '*') {
    return false;
  }
  SmUInt outU = 0;
  const bool ok = (op == '+') ? tryAddUIntChecked(lu, ru, outU) : tryMulUIntChecked(lu, ru, outU);
  if (!ok) {
    return false;
  }
  outS = makeScalarUInt(outU);
  return true;
}

inline bool tryApplyExactNonNegUIntAddOrMul(
    char op,
    const EvalValue::ScalarValue& lv,
    const EvalValue::ScalarValue& rv,
    EvalValue& outS) {
  SmUInt lu = 0;
  SmUInt ru = 0;
  if (!tryGetBothExactNonNegativeUIntFromScalars(lv, rv, lu, ru)) {
    return false;
  }
  return tryApplyUIntAddOrMul(op, lu, ru, outS);
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
      SmFloat outD = 0.0;
      if (!applyBinary(lv.scalar, rv.scalar, op, outD)) {
        return false;
      }
      outS = makeScalarMaybeExact(outD);
      return true;
    }

    if (lv.hasExactUInt() && rv.hasExactUInt()
        && (!lv.hasExactSInt() || !rv.hasExactSInt())) {
      if (TryApplyExactUIntBinaryScalarOp(op, lv.exactUInt, rv.exactUInt, outS)) {
        return true;
      }
    }

    SmInt li = 0;
    SmInt ri = 0;
    const bool leftExactSInt = tryGetExactSignedIntNoUIntWrapFromScalar(lv, li);
    const bool rightExactSInt = tryGetExactSignedIntNoUIntWrapFromScalar(rv, ri);
    SmUInt lu = 0;
    SmUInt ru = 0;
    if (op == '^' && tryGetBothExactNonNegativeUIntFromScalars(lv, rv, lu, ru)) {
      SmUInt outU = 0;
      if (tryPowUIntChecked(lu, ru, outU)) {
        outS = makeScalarUInt(outU);
        return true;
      }
    }
    if (op == '^' && leftExactSInt && li < 0 && tryGetExactNonNegativeUIntFromScalar(rv, ru)) {
      const SmUInt baseMag = (li == kIntMin)
        ? kIntMinMag
        : static_cast<SmUInt>(-li);
      SmUInt powMag = 0;
      if (tryPowUIntChecked(baseMag, ru, powMag)) {
        if ((ru & 1u) == 0u) {
          outS = makeScalarUInt(powMag);
          return true;
        }
        if (powMag <= kIntMinMag) {
          outS = (powMag == kIntMinMag)
            ? makeScalarInt(kIntMin)
            : makeScalarInt(-static_cast<SmInt>(powMag));
          return true;
        }
      }
    }
    if (leftExactSInt && rightExactSInt) {
      if (op == '-' && li == kIntMin && ri == 1LL) {
        outS = makeScalar(-kFloatSIntBound); outS.scalarValue.setDecScientificPow63High(true);
        return true;
      }
      if (tryApplyCheckedAddSubMulOrSignedUIntPromote(op, li, ri, outS)) {
        return true;
      }
      SmInt outI = 0;
      if (op == '^' && tryPowSIntChecked(li, ri, outI)) {
        outS = makeScalarInt(outI);
        return true;
      }
      if (op == '/' && TryQuotExactSInt(li, ri, outI)) {
        outS = makeScalarInt(outI);
        return true;
      }
      if (op == '/' && ri == -1 && li == kIntMin) {
        outS = makeScalarUInt(kIntMinMag);
        return true;
      }
    }

    if (tryApplyExactNonNegUIntAddOrMul(op, lv, rv, outS)) {
      return true;
    }
    if (op == '/') {
      lu = 0;
      ru = 0;
      if (tryGetBothExactNonNegativeUIntFromScalars(lv, rv, lu, ru) &&
          TryApplyExactUIntBinaryScalarOp(op, lu, ru, outS)) {
        return true;
      }
    }
    if (op == '+' && lv.hasExactSInt() && lv.exactSInt == kIntMax && !rv.hasExactSInt() &&
        !rv.hasExactUInt() && std::fabs(rv.scalar - 0.5) < kTolLost3) {
      outS = makeScalar(kFloatSIntBound); outS.scalarValue.setDecScientificPow63High(true);
      return true;
    }
    if (op == '+' && rv.hasExactSInt() && rv.exactSInt == kIntMax && !lv.hasExactSInt() &&
        !lv.hasExactUInt() && std::fabs(lv.scalar - 0.5) < kTolLost3) {
      outS = makeScalar(kFloatSIntBound); outS.scalarValue.setDecScientificPow63High(true);
      return true;
    }
#if SMARTMATH_COMPLEX_NUMBERS
    if (getSupportComplexNumbers() && (scalarHasNonzeroImaginaryPart(lv) || scalarHasNonzeroImaginaryPart(rv))) {
      return false;
    }
#endif
    SmFloat outD = 0;
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

bool tryConsumeParenthesizedIdentList(
    smartmath_internal::EvalContext& ctx, std::vector<std::string>& outParams, bool quiet) {
  outParams.clear();
  if (*ctx.p != '(') {
    return false;
  }
  ++ctx.p;
  skipSpaces(ctx.p);
  if (*ctx.p == ')') {
    ++ctx.p;
    return true;
  }
  while (true) {
    if (!localIsIdentStart(*ctx.p)) {
      if (!quiet) {
        ctx.fail(ParseErrorId::UnexpectedToken);
      }
      return false;
    }
    outParams.emplace_back(consumeLowerIdentToken(ctx.p));
    skipSpaces(ctx.p);
    if (*ctx.p == ',') {
      ++ctx.p;
      skipSpaces(ctx.p);
      continue;
    }
    if (*ctx.p == ')') {
      ++ctx.p;
      return true;
    }
    if (!quiet) {
      ctx.fail(ParseErrorId::UnexpectedToken);
    }
    return false;
  }
}

bool MathParser::parseFunctionDefinition(
    EvalContext& ctx,
    std::string& outName,
    std::vector<std::string>& outParams,
    std::string& outExpr) const {
  const char* save = ctx.p;
  skipSpaces(ctx.p);

  if (!localIsIdentStart(*ctx.p)) {
    ctx.p = save;
    return false;
  }
  std::string fnName = consumeLowerIdentToken(ctx.p);
  skipSpaces(ctx.p);
  if (*ctx.p != '(') {
    ctx.p = save;
    return false;
  }
  std::vector<std::string> params;
  if (!tryConsumeParenthesizedIdentList(ctx, params, true)) {
    ctx.p = save;
    return false;
  }
  skipSpaces(ctx.p);
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
  skipSpaces(ctx.p);

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

#if SMARTMATH_LAMBDA_FUNCTIONS
bool MathParser::lambdaBodyConsume(
    EvalContext& ctx,
    std::string& outBody,
    LambdaBodyStop stop,
    const char* lambdaArgStart) {
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
        case LambdaBodyStop::ArgDelim:
          if ((c == ',' || c == ')') && lambdaArgStart != nullptr &&
              netRoundParenDepthBetween(lambdaArgStart, ctx.p) == 0) {
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
  return stop == LambdaBodyStop::TopLevelSemicolonOrEof || stop == LambdaBodyStop::ArgDelim;
}

bool MathParser::tryConsumeLambdaParameterList(EvalContext& ctx, std::vector<std::string>& outParams, bool quiet) const {
  outParams.clear();
  skipSpaces(ctx.p);
  if (*ctx.p == '(') {
    return tryConsumeParenthesizedIdentList(ctx, outParams, quiet);
  }
  if (!localIsIdentStart(*ctx.p)) {
    return false;
  }
  outParams.emplace_back(consumeLowerIdentToken(ctx.p));
  skipSpaces(ctx.p);
  while (*ctx.p == ',') {
    ++ctx.p;
    skipSpaces(ctx.p);
    if (!localIsIdentStart(*ctx.p)) {
      if (!quiet) {
        ctx.fail(ParseErrorId::UnexpectedToken);
      }
      return false;
    }
    outParams.emplace_back(consumeLowerIdentToken(ctx.p));
    skipSpaces(ctx.p);
  }
  return true;
}

bool MathParser::tryParseLambdaInnerUnwrappedSuffix(
    EvalContext& ctx,
    std::vector<std::string>& outParams,
    std::string& outBody,
    LambdaBodyStop bodyStop,
    const char* lambdaArgStart,
    bool quiet) const {
  if (!tryConsumeLambdaParameterList(ctx, outParams, quiet)) {
    return false;
  }
  skipSpaces(ctx.p);
  if (*ctx.p != ':') {
    return false;
  }
  ++ctx.p;
  skipSpaces(ctx.p);
  return lambdaBodyConsume(ctx, outBody, bodyStop, lambdaArgStart);
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
  skipSpaces(ctx.p);
  if (*ctx.p != ')') {
    ctx.p = saveOuter;
    return false;
  }
  ++ctx.p;
  skipSpaces(ctx.p);
  return true;
}

bool MathParser::tryParseLambdaRhsAfterEquals(
    EvalContext& ctx, std::vector<std::string>& outParams, std::string& outExpr) const {
  const char* const saveOuter = ctx.p;
  if (!peekRhsMayBeLambdaSyntaxAt(saveOuter)) {
    return false;
  }
  skipSpaces(ctx.p);
  if (tryParseWrappedParenLambdaSuffix(
          ctx, outParams, outExpr, LambdaBodyStop::WrappedParenClose, true)) {
    if (*ctx.p == '\0' || *ctx.p == ';') {
      return true;
    }
    ctx.p = saveOuter;
    outParams.clear();
    outExpr.clear();
  }
  skipSpaces(ctx.p);
  if (!tryParseLambdaInnerUnwrappedSuffix(
          ctx, outParams, outExpr, LambdaBodyStop::TopLevelSemicolonOrEof, nullptr, true)) {
    return false;
  }
  skipSpaces(ctx.p);
  return *ctx.p == '\0' || *ctx.p == ';';
}

std::vector<std::string> enclosingParamNamesThen(
    const std::vector<std::string>* enclosing,
    const std::vector<std::string>& own) {
  std::vector<std::string> names;
  if (enclosing != nullptr) {
    names = *enclosing;
  }
  names.insert(names.end(), own.begin(), own.end());
  return names;
}

std::unique_ptr<MathParser::Expr> MathParser::makeUnaryInlineLambdaExpr(
    EvalContext& ctx,
    std::vector<std::string>&& params,
    std::string&& body,
    const char* expectsUnaryFunction,
    const bool syntaxCheckBody) {
  const std::string nameErr = validateParameterNames(params);
  if (!nameErr.empty()) {
    ctx.fail(nameErr);
    return nullptr;
  }
  if (params.size() != 1U) {
    ctx.fail(expectsUnaryFunction);
    return nullptr;
  }
  if (bodyIsEffectivelyEmpty(body)) {
    ctx.fail(STR_FUNCTION_BODY_IS_EMPTY);
    return nullptr;
  }
  if (syntaxCheckBody) {
    const std::vector<std::string> boundNames = enclosingParamNamesThen(ctx.boundBodyParamNames, params);
    EvalContext bodyCtx;
    bodyCtx.suppressBareFunctionTailHint = true;
    bodyCtx.compilingUserFunctionParams = ctx.compilingUserFunctionParams;
    bodyCtx.validatingUserFunctionName = ctx.validatingUserFunctionName;
    bodyCtx.boundBodyParamNames = &boundNames;
    std::unique_ptr<Expr> parsed;
    if (!tryParseSyntaxCheckedExpression(bodyCtx, body, true, parsed)) {
      ctx.fail(bodyCtx.errorText.empty() ? std::string(STR_FAILED_TO_PARSE_USER_FUNCTION_BODY)
                                         : bodyCtx.errorText);
      return nullptr;
    }
  }
  auto lit = std::make_unique<Expr>();
  lit->tag = Expr::Tag::Literal;
  lit->literalValue = makeInlineLambdaValue(std::move(params), std::move(body));
  return lit;
}

std::unique_ptr<MathParser::Expr> MathParser::tryParseUnaryLambdaOrFunctionRef(
    EvalContext& ctx,
    const char* expectsOneFunction,
    const char* expectsUnaryFunction) {
  ctx.parseError = false;
  ctx.errorText.clear();
  const char* const save = ctx.p;
  skipSpaces(ctx.p);

  if (*ctx.p == '(' && peekRhsMayBeLambdaSyntaxAt(ctx.p + 1)) {
    std::vector<std::string> params;
    std::string body;
    if (tryParseWrappedParenLambdaSuffix(ctx, params, body, LambdaBodyStop::WrappedParenClose, true)) {
      if (*ctx.p == ',' || *ctx.p == ')' || *ctx.p == '\0') {
        return makeUnaryInlineLambdaExpr(
            ctx, std::move(params), std::move(body), expectsUnaryFunction, *ctx.p != '\0');
      }
    }
  }

  ctx.p = save;
  skipSpaces(ctx.p);
  if (!peekRhsMayBeLambdaSyntaxAt(ctx.p)) {
    return parseUnaryFunctionRef(ctx, expectsOneFunction, expectsUnaryFunction);
  }

  std::vector<std::string> params;
  std::string body;
  if (!tryParseLambdaInnerUnwrappedSuffix(
          ctx, params, body, LambdaBodyStop::ArgDelim, save, true)) {
    ctx.p = save;
    return parseUnaryFunctionRef(ctx, expectsOneFunction, expectsUnaryFunction);
  }
  skipSpaces(ctx.p);
  if (*ctx.p == ',' || *ctx.p == ')' || *ctx.p == '\0') {
    return makeUnaryInlineLambdaExpr(
        ctx, std::move(params), std::move(body), expectsUnaryFunction, *ctx.p != '\0');
  }
  ctx.p = save;
  return parseUnaryFunctionRef(ctx, expectsOneFunction, expectsUnaryFunction);
}

MathParser::EvalValue MathParser::evalInlineLambdaCall(
    EvalContext& ctx,
    const std::vector<std::string>& paramNames,
    const std::string& bodyExpr,
    std::vector<EvalValue>&& args,
    const std::unordered_map<std::string, EvalValue>* parentScopedVars) {
  static const std::string kLambda = "lambda";
  if (args.size() != paramNames.size()) {
    ctx.failCall(CallErrorId::ExactCount, kLambda, paramNames.size());
    return makeScalar(0);
  }
  if (ctx.evalDepth >= kMaxEvalDepth) {
    ctx.fail(ParseErrorId::MaxEvaluationDepthReached);
    return makeScalar(0);
  }

  std::unordered_map<std::string, EvalValue> localVars;
  if (parentScopedVars != nullptr) {
    localVars = *parentScopedVars;
  }
  for (std::size_t i = 0; i < paramNames.size(); ++i) {
    localVars[paramNames[i]] = std::move(args[i]);
  }

  const std::string trimmedBody = trim(bodyExpr);
  if (bodyIsEffectivelyEmpty(trimmedBody)) {
    ctx.fail(STR_FUNCTION_BODY_IS_EMPTY);
    return makeScalar(0);
  }

  std::vector<std::string> enclosingNames;
  if (parentScopedVars != nullptr) {
    enclosingNames.reserve(parentScopedVars->size());
    for (const auto& kv : *parentScopedVars) {
      enclosingNames.push_back(kv.first);
    }
  }
  const std::vector<std::string> boundNames = enclosingParamNamesThen(
      parentScopedVars != nullptr ? &enclosingNames : nullptr, paramNames);
  std::vector<AstStatement> compiledBody;
  if (!tryCompileExprToProgram(trimmedBody, ctx, nullptr, nullptr, true, &boundNames, compiledBody)) {
    return makeScalar(0);
  }
  return runBoundProgram(ctx, compiledBody, &localVars, true);
}
#endif

std::unique_ptr<MathParser::Expr> MathParser::parseUnaryFunctionArg(
    EvalContext& ctx,
    const char* expectsOneFunction,
    const char* expectsUnaryFunction) {
#if SMARTMATH_LAMBDA_FUNCTIONS
  if (getSupportLambdaFunctions()) {
    return tryParseUnaryLambdaOrFunctionRef(ctx, expectsOneFunction, expectsUnaryFunction);
  }
#endif
  return parseUnaryFunctionRefOnly(ctx, expectsOneFunction, expectsUnaryFunction);
}

std::unique_ptr<MathParser::Expr> MathParser::parseUnaryFunctionRefOnly(
    EvalContext& ctx,
    const char* expectsOneFunction,
    const char* expectsUnaryFunction) {
  ctx.parseError = false;
  ctx.errorText.clear();
  skipSpaces(ctx.p);
#if SMARTMATH_LAMBDA_FUNCTIONS
  if (peekRhsMayBeLambdaSyntaxAt(ctx.p)) {
    ctx.fail(ParseErrorId::UnexpectedToken);
    return nullptr;
  }
#endif
  return parseUnaryFunctionRef(ctx, expectsOneFunction, expectsUnaryFunction);
}

#if SMARTMATH_COMPLEX_NUMBERS
std::unique_ptr<MathParser::Expr> MathParser::wrapExprWithTightImagSuffixIfPresent(
    EvalContext& ctx,
    std::unique_ptr<Expr> expr) {
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
}
#endif

std::unique_ptr<MathParser::Expr> MathParser::parsePrimary(EvalContext& ctx) {
  skipSpaces(ctx.p);
  std::unique_ptr<Expr> prim;
  bool allowTightImag = true;
  if (*ctx.p == '(') {
    prim = parsePrimaryParenthesized(ctx);
  } else if (isNumericLiteralStart(*ctx.p)) {
    prim = parsePrimaryNumericLiteral(ctx, allowTightImag);
  } else if (localIsIdentStart(*ctx.p)) {
    prim = parsePrimaryIdentifierOrCall(ctx);
  } else {
    ctx.fail(ParseErrorId::UnexpectedToken);
    return nullptr;
  }
  if (ctx.parseError || !prim) {
    return nullptr;
  }
  prim = parseTrailingIndexes(ctx, std::move(prim));
  if (ctx.parseError || !prim) {
    return nullptr;
  }
#if SMARTMATH_COMPLEX_NUMBERS
  if (allowTightImag) {
    return wrapExprWithTightImagSuffixIfPresent(ctx, std::move(prim));
  }
#else
  (void)allowTightImag;
#endif
  return prim;
}

std::unique_ptr<MathParser::Expr> MathParser::parsePrimaryParenthesized(EvalContext& ctx) {
  ++ctx.p;
  skipSpaces(ctx.p);
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
    ctx.failMissingClosingParenLike();
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
  SmInt ms = 0;
  if (!parseTimeLiteralStringToMs(p0, q, ms, timeErr)) {
    if (timeErr != nullptr) {
      ctx.fail(timeErr);
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
  SmInt totalMs = 0;
  int lastUnitRank = -1;
  int comps = 0;
  while (true) {
    if (*ctx.p < '0' || *ctx.p > '9') {
      break;
    }
    SmUInt uv = 0;
    while (*ctx.p >= '0' && *ctx.p <= '9') {
      const int dig = *ctx.p - '0';
      if (uv > (kUIntMax - static_cast<SmUInt>(dig)) / 10u) {
        ctx.p = pSave;
        if (comps > 0) {
          ctx.fail(STR_TIME_LITERAL_INVALID_SEGMENT);
        }
        return false;
      }
      uv = uv * 10u + static_cast<SmUInt>(dig);
      ++ctx.p;
    }
    skipSpaces(ctx.p);
    int ur = -1;
    SmInt fac = 0;
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
      ctx.fail(STR_TIME_COMPACT_EXPECTED_UNIT);
      return false;
    }
    if (ur <= lastUnitRank) {
      ctx.p = pSave;
      ctx.fail(STR_TIME_COMPACT_UNIT_ORDER);
      return false;
    }
    lastUnitRank = ur;
    if (uv > static_cast<SmUInt>(kIntMax)) {
      ctx.p = pSave;
      ctx.fail(STR_TIME_LITERAL_INVALID_SEGMENT);
      return false;
    }
    const SmInt cf = static_cast<SmInt>(uv);
    SmInt dlt = 0;
    if (!tryProductCoeffUnitMs(cf, fac, dlt)) {
      ctx.p = pSave;
      ctx.fail(STR_TIME_LITERAL_INVALID_SEGMENT);
      return false;
    }
    if (!tryAddTimeMsChecked(totalMs, dlt, totalMs)) {
      ctx.p = pSave;
      ctx.fail(STR_TIME_LITERAL_INVALID_SEGMENT);
      return false;
    }
    ++comps;
    skipSpaces(ctx.p);
  }
  if (comps <= 0) {
    ctx.p = pSave;
    return false;
  }
  skipSpaces(ctx.p);
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
      ctx.fail(STR_TIME_COMPACT_INVALID_SUFFIX);
      return false;
    }
  }
  out = makeScalarTimeMs(totalMs);
  return true;
}
#endif

std::unique_ptr<MathParser::Expr> MathParser::parsePrimaryNumericLiteral(EvalContext& ctx, bool& allowTightImag) {
  allowTightImag = true;
  auto emitNumericLiteral = [&](EvalValue litV) -> std::unique_ptr<Expr> {
    auto lit = std::make_unique<Expr>();
    lit->tag = Expr::Tag::Literal;
    lit->literalValue = std::move(litV);
    return lit;
  };
  auto finishNumericLiteral = [&](std::unique_ptr<Expr> lit) -> std::unique_ptr<Expr> {
    return lit;
  };

  const char* parsedEnd = nullptr;
  SmUInt parsedUInt = 0;
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
        ctx.failPrefixedLiteral(px);
        return nullptr;
      }
      ctx.p = parsedEnd;
      if (parsedUInt > static_cast<SmUInt>(kIntMax)) {
        return finishNumericLiteral(emitNumericLiteral(makeScalarUInt(parsedUInt)));
      }
      return finishNumericLiteral(emitNumericLiteral(makeScalarInt(static_cast<SmInt>(parsedUInt))));
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
          allowTightImag = false;
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
          allowTightImag = false;
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
      ctx.fail(ParseErrorId::UnexpectedToken);
    } else {
      ctx.fail(ParseErrorId::InvalidNumericLiteral);
    }
    return nullptr;
  }

  if (tryParseInputNumberAsInteger(numStart, numEnd, parsedUInt)) {
    ctx.p = numEnd;
    return finishNumericLiteral(emitNumericLiteral(makeScalarUIntPreferInt(parsedUInt)));
  }

  char* end = nullptr;
  const SmFloat d = smartmath::parseFloatPrefix(numStart, &end);
  if (end == numStart) {
    if (*numStart == '.') {
      ctx.fail(ParseErrorId::UnexpectedToken);
    } else {
      ctx.fail(ParseErrorId::InvalidNumericLiteral);
    }
    return nullptr;
  }
  if ((*end == 'e' || *end == 'E')) {
    const char* pExp = end + 1;
    if (*pExp == '+' || *pExp == '-') {
      ++pExp;
    }
    if (!std::isdigit(static_cast<unsigned char>(*pExp))) {
      ctx.fail(ParseErrorId::UnexpectedToken);
      return nullptr;
    }
  }
  ctx.p = end;
  return finishNumericLiteral(emitNumericLiteral(makeScalar(d)));
}

namespace {

bool isInstanceTypeName(const std::string& typeName) {
  return typeName == STR_INT || typeName == STR_FLOAT || typeName == STR_TIME || typeName == STR_COMPLEX ||
         typeName == STR_ARRAY || typeName == STR_LIST;
}

bool valueIsInstanceOfType(const smartmath_internal::EvalValue& v, const std::string& typeName, bool supportComplex,
                           bool supportTime) {
  if (v.kind == smartmath_internal::ValueKind::Array) {
    return typeName == STR_ARRAY || typeName == STR_LIST;
  }
  if (v.kind != smartmath_internal::ValueKind::Scalar) {
    return false;
  }
#if SMARTMATH_COMPLEX_NUMBERS
  if (supportComplex && scalarHasNonzeroImaginaryPart(v.scalarValue)) {
    return typeName == STR_COMPLEX;
  }
#else
  (void)supportComplex;
#endif
  if (supportTime && v.scalarValue.scalarKind == smartmath_internal::ScalarKind::Time) {
    return typeName == STR_TIME;
  }
  if (scalarHasExactIntegerPayload(v.scalarValue)) {
    return typeName == STR_INT;
  }
  return typeName == STR_FLOAT;
}

}  // namespace

std::unique_ptr<MathParser::Expr> MathParser::parseIsInstanceCall(EvalContext& ctx) {
  const std::string fnName = getFunctionName(BuiltinFunctionId::IsInstance);
  skipSpaces(ctx.p);
  if (*ctx.p == ')') {
    ++ctx.p;
    ctx.failCall(CallErrorId::ExactCount, fnName, 2);
    return nullptr;
  }
  if (*ctx.p == ',') {
    ctx.fail(ParseErrorId::UnexpectedComma);
    return nullptr;
  }
  auto value = parseExpression(ctx);
  if (ctx.parseError || !value) {
    return nullptr;
  }
  skipSpaces(ctx.p);
  bool hasComma = false;
  if (!tryConsumeCommaArgSeparator(ctx, hasComma)) {
    return nullptr;
  }
  if (!hasComma) {
    if (*ctx.p != ')') {
      ctx.failMissingClosingParenLike();
      return nullptr;
    }
    ++ctx.p;
    ctx.failCall(CallErrorId::ExactCount, fnName, 2);
    return nullptr;
  }
  skipSpaces(ctx.p);
  if (!localIsIdentStart(*ctx.p)) {
    ctx.fail(fnName + STR_PAR_EXPECTS_A_TYPE_NAME);
    return nullptr;
  }
  const char* identStart = ctx.p;
  const char* identEnd = ctx.p + 1;
  while (localIsIdentChar(*identEnd)) {
    ++identEnd;
  }
  ctx.p = identEnd;
  std::string typeName;
  assignLowerIdentFromRange(typeName, identStart, identEnd);
  skipSpaces(ctx.p);
  if (*ctx.p != ')' && *ctx.p != ',') {
    ctx.fail(fnName + STR_PAR_EXPECTS_A_TYPE_NAME);
    return nullptr;
  }
  if (!isInstanceTypeName(typeName)) {
    ctx.fail(std::string(STR_UNKNOWN_TYPE_NAME_COLON) + typeName);
    return nullptr;
  }
  std::size_t argc = 2;
  while (true) {
    skipSpaces(ctx.p);
    hasComma = false;
    if (!tryConsumeCommaArgSeparator(ctx, hasComma)) {
      return nullptr;
    }
    if (!hasComma) {
      break;
    }
    auto extra = parseExpression(ctx);
    if (ctx.parseError || !extra) {
      return nullptr;
    }
    ++argc;
  }
  if (*ctx.p != ')') {
    ctx.failMissingClosingParenLike();
    return nullptr;
  }
  ++ctx.p;
  if (argc != 2) {
    ctx.failCall(CallErrorId::ExactCount, fnName, 2);
    return nullptr;
  }
  auto typeExpr = std::make_unique<Expr>();
  typeExpr->tag = Expr::Tag::TypeName;
  typeExpr->name = std::move(typeName);
  auto call = std::make_unique<Expr>();
  call->tag = Expr::Tag::Call;
  call->name = fnName;
  call->builtinFunctionId = BuiltinFunctionId::IsInstance;
  call->elements.push_back(std::move(value));
  call->elements.push_back(std::move(typeExpr));
  return call;
}

bool identIsBoundBodyParam(const smartmath_internal::EvalContext& ctx, const std::string& ident) {
  const std::vector<std::string>* names = ctx.boundBodyParamNames;
  if (names == nullptr) {
    return false;
  }
  for (const std::string& name : *names) {
    if (name == ident) {
      return true;
    }
  }
  return false;
}

std::unique_ptr<MathParser::Expr> MathParser::parsePrimaryIdentifierOrCall(EvalContext& ctx) {
  const char* identStart = ctx.p;
  if (!localIsIdentStart(*ctx.p)) {
    ctx.fail(ParseErrorId::UnexpectedToken);
    return nullptr;
  }
  const char* identEnd = ctx.p + 1;
  while (localIsIdentChar(*identEnd)) {
    ++identEnd;
  }
  ctx.p = identEnd;
  skipSpaces(ctx.p);
  if (*ctx.p != '(') {
    if (!identMayBeBareBuiltinName(identStart, identEnd) && userFunctionIndex_.empty()
        && ctx.compilingUserFunctionParams == nullptr) {
      auto v = std::make_unique<Expr>();
      v->tag = Expr::Tag::Variable;
      assignLowerIdentFromRange(v->name, identStart, identEnd);
      return v;
    }
    std::string ident;
    assignLowerIdentFromRange(ident, identStart, identEnd);
    if (identIsBoundBodyParam(ctx, ident)) {
      auto v = std::make_unique<Expr>();
      v->tag = Expr::Tag::Variable;
      v->name = std::move(ident);
      return v;
    }
    if (identMayBeBareBuiltinName(ident) && trySetMissingFunctionCallError(ctx, ident, identStart)) {
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
  skipSpaces(ctx.p);
  if (trySetIncompleteOpenedFunctionCallHint(ctx, ident)) {
    return nullptr;
  }
  if (ident == STR_ISINSTANCE) {
    return parseIsInstanceCall(ctx);
  }
  std::vector<std::unique_ptr<Expr>> args;
  const bool isSortbyCall = (ident == STR_SORTBY);
  const bool isFilterCall = (ident == STR_FILTER);
  if (isSortbyCall || isFilterCall) {
    const char* expectsOne = isFilterCall ? STR_FILTER_EXPECTS_ONE_FUNCTION : STR_SORTBY_EXPECTS_ONE_FUNCTION;
    const char* expectsUnary = isFilterCall ? STR_FILTER_EXPECTS_UNARY_FUNCTION : STR_SORTBY_EXPECTS_UNARY_FUNCTION;
    if (!parseUnaryFunctionCallArguments(ctx, args, expectsOne, expectsUnary)) {
      return nullptr;
    }
  } else if (!parseParenthesizedExprList(ctx, args)) {
    return nullptr;
  }
  if (*ctx.p != ')') {
    ctx.failMissingClosingParenLike();
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
  skipSpaces(ctx.p);
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
  if (consumeKeyword(ctx.p, opName(OperatorNameId::Not).c_str())) {
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
  return parsePower(ctx);
}

std::unique_ptr<MathParser::Expr> MathParser::parseOneBracket(EvalContext& ctx, std::unique_ptr<Expr> base) {
  ++ctx.p;
  ++ctx.arrayIndexBracketDepth;
  skipSpaces(ctx.p);
  if (*ctx.p == ']') {
    --ctx.arrayIndexBracketDepth;
    ctx.fail(ParseErrorId::MissingIndex);
    return nullptr;
  }

  bool hasStart = false;
  bool hasStop = false;
  bool hasStep = false;
  std::unique_ptr<Expr> startExpr;
  std::unique_ptr<Expr> stopExpr;
  std::unique_ptr<Expr> stepExpr;
  // Called after the second slice ':'; leaves ctx.p on the closing ']'.
  const auto parseSliceStepAfterColon = [&]() -> bool {
    skipSpaces(ctx.p);
    if (*ctx.p == ']') {
      hasStep = false;
      return true;
    }
    stepExpr = parseExpression(ctx);
    if (ctx.parseError || !stepExpr) {
      --ctx.arrayIndexBracketDepth;
      return false;
    }
    hasStep = true;
    skipSpaces(ctx.p);
    if (*ctx.p != ']') {
      --ctx.arrayIndexBracketDepth;
      ctx.failMissingClosingBracketAfterIndex();
      return false;
    }
    return true;
  };

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
    skipSpaces(ctx.p);
    if (*ctx.p == ':') {
      ++ctx.p;
    } else if (*ctx.p == ']') {
      ++ctx.p;
      --ctx.arrayIndexBracketDepth;
      auto w = std::make_unique<Expr>();
      w->tag = Expr::Tag::Index;
      w->left = std::move(base);
      w->right = std::move(startExpr);
      return w;
    } else {
      --ctx.arrayIndexBracketDepth;
      ctx.failMissingClosingBracketAfterIndex();
      return nullptr;
    }
  }

  // Slice form: [start:stop:step] with any bound omitted.
  skipSpaces(ctx.p);
  if (*ctx.p == ']') {
    hasStop = false;
    hasStep = false;
  } else if (*ctx.p == ':') {
    hasStop = false;
    ++ctx.p;
    if (!parseSliceStepAfterColon()) {
      return nullptr;
    }
  } else {
    stopExpr = parseExpression(ctx);
    if (ctx.parseError || !stopExpr) {
      --ctx.arrayIndexBracketDepth;
      return nullptr;
    }
    hasStop = true;
    skipSpaces(ctx.p);
    if (*ctx.p == ':') {
      ++ctx.p;
      if (!parseSliceStepAfterColon()) {
        return nullptr;
      }
    } else if (*ctx.p == ']') {
      hasStep = false;
    } else {
      --ctx.arrayIndexBracketDepth;
      ctx.failMissingClosingBracketAfterIndex();
      return nullptr;
    }
  }

  ++ctx.p;
  --ctx.arrayIndexBracketDepth;
  auto w = std::make_unique<Expr>();
  w->tag = Expr::Tag::Slice;
  w->left = std::move(base);
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
  return w;
}

std::unique_ptr<MathParser::Expr> MathParser::parseTrailingIndexes(EvalContext& ctx, std::unique_ptr<Expr> prim) {
  while (true) {
    const char* const saved = ctx.p;
    skipSpaces(ctx.p);
    if (*ctx.p != '[') {
      ctx.p = saved;
      break;
    }
    prim = parseOneBracket(ctx, std::move(prim));
    if (!prim) {
      return nullptr;
    }
  }
  return prim;
}

std::unique_ptr<MathParser::Expr> MathParser::parsePower(EvalContext& ctx) {
  auto left = parsePrimary(ctx);
  if (ctx.parseError || !left) {
    return nullptr;
  }
  skipSpaces(ctx.p);
  // `%` immediately before `**` is the base: `2%**3` is `(2%)**3`.
  // A trailing `%` stays for parsePostfixPercent: `2**3%` is `(2**3)%`.
  if (percentTailStartsPower(ctx.p)) {
    ++ctx.p;
    left = makePostfixPercentExpr(std::move(left));
    skipSpaces(ctx.p);
  }
  if (!(ctx.p[0] == '*' && ctx.p[1] == '*')) {
    return left;
  }
  ctx.p += 2;
  auto right = parseUnary(ctx);
  if (ctx.parseError || !right) {
    return nullptr;
  }
  auto out = std::make_unique<Expr>();
  out->tag = Expr::Tag::Binary;
  out->binaryOp = Expr::BinaryOp::Pow;
  out->left = std::move(left);
  out->right = std::move(right);
  return out;
}

std::unique_ptr<MathParser::Expr> MathParser::parsePostfixPercent(EvalContext& ctx) {
  auto prim = parseUnary(ctx);
  if (hasExprParseFailure(ctx, prim)) {
    return nullptr;
  }
  skipSpaces(ctx.p);
  if (!isPercentPostfixContext(ctx.p)) {
    return prim;
  }
  ++ctx.p;
  return makePostfixPercentExpr(std::move(prim));
}

std::unique_ptr<MathParser::Expr> MathParser::parseMulDivMod(EvalContext& ctx) {
  return parseLeftAssocBinary(ctx, &MathParser::parsePostfixPercent, &MathParser::tryConsumeMulDivModOp);
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
  skipSpaces(ctx.p);
  if (consumeKeyword(ctx.p, opName(OperatorNameId::Not).c_str())) {
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
  skipSpaces(ctx.p);
  // Statement start is fixed by the statement parser after skipSpaces.
  // A nested expression must not move it, or a bare name before ')' inside an
  // argument is reported as a mismatched closer.
  return parseOr(ctx);
}

bool MathParser::tryAppendTrailingFormatterSugarStatement(
    EvalContext& ctx,
    std::vector<AstStatement>& out,
    bool& handled) {
  handled = false;
  if (out.empty() || !localIsIdentStart(*ctx.p)) {
    return true;
  }
  const char* sugarStart = ctx.p;
  const char* q = ctx.p + 1;
  while (localIsIdentChar(*q)) {
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
  skipSpaces(ctx.p);
  if (*ctx.p == '\0') {
    return false;
  }
  if (*ctx.p == ';') {
    ctx.fail(STR_EMPTY_STATEMENT);
    return false;
  }
  if (std::strchr(ctx.p, ';') != nullptr) {
    return false;
  }
  ctx.start = ctx.p;
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
  skipSpaces(ctx.p);
  if (!ctx.parseError && *ctx.p != '\0') {
    ctx.fail(ParseErrorId::UnexpectedInput);
    return false;
  }
  return !ctx.parseError && !out.empty();
}

bool MathParser::parseProgram(EvalContext& ctx, std::vector<AstStatement>& out) {
  out.clear();
  std::unique_ptr<std::unordered_map<std::string, std::vector<std::string>>> compilingUserFunctionParams;
  ctx.compilingUserFunctionParams = nullptr;
  while (true) {
    skipSpaces(ctx.p);
    if (*ctx.p == '\0') {
      break;
    }
    if (*ctx.p == ';') {
      ctx.fail(STR_EMPTY_STATEMENT);
      ctx.compilingUserFunctionParams = nullptr;
      return false;
    }
    ctx.start = ctx.p;
    const std::size_t stmtBefore = out.size();
    const char* save = ctx.p;
    if (peekIdentFollowedByChar(save, '(') && tryAppendFunctionDefinitionStatement(ctx, out)) {
      // handled above
    } else {
      ctx.p = save;
      skipSpaces(ctx.p);
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
    if (st.kind == AstStatement::Kind::FunDef || st.kind == AstStatement::Kind::IndexAssign) {
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
    if (st.indexExpr) {
      bindExprVariableRefs(*st.indexExpr, state.hasAssignments, variables_);
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
    if (getSupportComplexNumbers() && evalValueHasNonzeroImaginary(v)) {
      ctx.fail(ParseErrorId::IncompatibleOperands);
      return makeScalar(0);
    }
#endif
    return evalSIntBinaryOp(ctx, v, makeScalarInt(-1), Expr::BinaryOp::BitXor);
  }

  ctx.fail(ParseErrorId::InternalUnaryOp);
  return makeScalar(0);
}

bool MathParser::tryScaleRhsByDirectPostfixPercent(
    EvalContext& ctx,
    Expr::BinaryOp op,
    bool rhsIsDirectPostfixPercent,
    const EvalValue& left,
    EvalValue& right) const {
  if (!rhsIsDirectPostfixPercent ||
      (op != Expr::BinaryOp::Add && op != Expr::BinaryOp::Sub)) {
    return true;
  }
  EvalValue scaled = evalMappedBinaryOp(ctx, Expr::BinaryOp::Mul, left, right);
  if (ctx.parseError) {
    return false;
  }
  right = std::move(scaled);
  return true;
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

  if (left.kind == ValueKind::Array && right.kind == ValueKind::Array && left.arr.size() != right.arr.size()) {
    ctx.fail(ParseErrorId::IncompatibleOperands);
  } else if (op == Expr::BinaryOp::Pow) {
    ctx.fail(ParseErrorId::NumericErrorInPowerOperation);
  } else {
    ctx.fail(ParseErrorId::NumericErrorInExpression);
  }
  return makeScalar(0);
}

enum class ExactUIntBitwiseGate {
  OnlyWhenNotBothExactSInt,
  WheneverBothExactUInt
};

inline bool applyIntBitwiseOrModulo(
    EvalContext& ctx,
    Expr::BinaryOp op,
    const EvalValue::ScalarValue& lv,
    const EvalValue::ScalarValue& rv,
    ExactUIntBitwiseGate gate,
    bool reportIntegerOperandError,
    EvalValue& outS) {
  const bool isModulo = (op == Expr::BinaryOp::Modulo);
  if (op == Expr::BinaryOp::ShiftLeft) {
    SmUInt aU = 0;
    SmUInt bU = 0;
    if (tryGetBothExactNonNegativeUIntFromScalars(lv, rv, aU, bU)) {
      if (!tryShiftLeftU64ExactOrMaybe(aU, bU, outS)) {
        ctx.fail(ParseErrorId::IncompatibleOperands);
        return false;
      }
      return true;
    }
  }

  const bool uintGateOpen =
      (gate == ExactUIntBitwiseGate::WheneverBothExactUInt) ||
      (!lv.hasExactSInt() || !rv.hasExactSInt());
  if (lv.hasExactUInt() && rv.hasExactUInt() && uintGateOpen) {
    const SmUInt aU = lv.exactUInt;
    const SmUInt bU = rv.exactUInt;
    if (isModulo) {
      if (bU == 0) {
        ctx.fail(ParseErrorId::IncompatibleOperands);
        return false;
      }
      outS = makeScalarUInt(aU % bU);
      return true;
    }
    if ((op == Expr::BinaryOp::ShiftLeft || op == Expr::BinaryOp::ShiftRight) && bU > static_cast<SmUInt>(kMaxShift)) {
      ctx.fail(ParseErrorId::IncompatibleOperands);
      return false;
    }
    if (op == Expr::BinaryOp::ShiftLeft) {
      if (!tryShiftLeftU64ExactOrMaybe(aU, bU, outS)) {
        ctx.fail(ParseErrorId::IncompatibleOperands);
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

  if (isModulo) {
    bool aNeg = false, bNeg = false;
    SmUInt aMag = 0, bMag = 0;
    if (!tryGetExactSignMagnitudeFromScalar(lv, aNeg, aMag) || !tryGetExactSignMagnitudeFromScalar(rv, bNeg, bMag)) {
      if (reportIntegerOperandError) {
        ctx.fail(ParseErrorId::ModuloIntegerOperands);
      }
      return false;
    }
    if (bMag == 0u) {
      ctx.fail(ParseErrorId::IncompatibleOperands);
      return false;
    }
    outS = makeScalarFromSignMagnitude(aNeg, aMag % bMag);
    return true;
  }
  SmInt a = 0, b = 0;
  if (!tryGetSignedIntFromScalar(lv, a) || !tryGetSignedIntFromScalar(rv, b)) {
    if (reportIntegerOperandError) {
      ctx.fail(ParseErrorId::BitwiseIntegerOperands);
    }
    return false;
  }
  if ((op == Expr::BinaryOp::ShiftLeft || op == Expr::BinaryOp::ShiftRight) && (b < 0 || b > kMaxShift)) {
    ctx.fail(ParseErrorId::IncompatibleOperands);
    return false;
  }
  if (op == Expr::BinaryOp::ShiftLeft) {
    outS = makeScalarShiftLeftSInt(a, static_cast<unsigned int>(b));
    return true;
  }
  const SmInt out = (op == Expr::BinaryOp::BitAnd) ? (a & b) :
                        (op == Expr::BinaryOp::BitOr) ? (a | b) :
                        (op == Expr::BinaryOp::BitXor) ? (a ^ b) :
                                                         bitwiseShiftRightDefined(a, static_cast<unsigned int>(b));
  outS = makeScalarInt(out);
  return true;
}

MathParser::EvalValue MathParser::evalSIntBinaryOp(
    EvalContext& ctx,
    const EvalValue& left,
    const EvalValue& right,
    Expr::BinaryOp op) const {
  const bool isModulo = (op == Expr::BinaryOp::Modulo);
  if (RejectSIntBinaryOperands(ctx, left, right, isModulo)) {
    return makeScalar(0);
  }
  const auto setIntegerOperandError = [&]() {
    if (isModulo) ctx.fail(ParseErrorId::ModuloIntegerOperands);
    else ctx.fail(ParseErrorId::BitwiseIntegerOperands);
  };
  const auto applyOp = [&](const EvalValue::ScalarValue& lv, const EvalValue::ScalarValue& rv, EvalValue& outS) -> bool {
    return applyIntBitwiseOrModulo(
        ctx, op, lv, rv, ExactUIntBitwiseGate::WheneverBothExactUInt, false, outS);
  };
  if (left.kind == ValueKind::Array && right.kind == ValueKind::Array && left.arr.size() != right.arr.size()) {
    ctx.fail(ParseErrorId::IncompatibleOperands);
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

namespace {

void scalarMarkInexactFloat(EvalValue::ScalarValue& sv) {
  sv.setExactSIntValid(false);
  sv.setExactUIntValid(false);
  sv.scalarKind = ScalarKind::FloatingPoint;
}

bool tryApplyPostfixPercent(EvalContext& ctx, EvalValue& v, bool supportTimeValues, bool supportComplexNumbers) {
#if SMARTMATH_TIME_VALUES
  if (supportTimeValues && v.scalarValue.scalarKind == ScalarKind::Time) {
    ctx.fail(ParseErrorId::IncompatibleOperands);
    return false;
  }
#else
  (void)supportTimeValues;
#endif
#if SMARTMATH_COMPLEX_NUMBERS
  if (supportComplexNumbers && scalarHasNonzeroImaginaryPart(v.scalarValue)) {
    ctx.fail(ParseErrorId::IncompatibleOperands);
    return false;
  }
#else
  (void)supportComplexNumbers;
#endif
  v.scalarValue.scalar /= 100.0;
  scalarMarkInexactFloat(v.scalarValue);
  return true;
}

bool timeComparedWithNonFinite(bool lTime, SmFloat l, bool rTime, SmFloat r) {
  return (lTime || rTime) && ((!lTime && !std::isfinite(l)) || (!rTime && !std::isfinite(r)));
}

// IEEE unordered NaN for comparison ops (not sort). Time values skip this check.
bool scalarPairHasUnorderedNan(const EvalValue::ScalarValue& a, const EvalValue::ScalarValue& b,
                               bool supportComplex) {
#if SMARTMATH_TIME_VALUES
  // Session time flag is enforced by callers; a Time kind still skips NaN unordered.
  if (a.scalarKind == ScalarKind::Time || b.scalarKind == ScalarKind::Time) {
    return false;
  }
#endif
#if SMARTMATH_COMPLEX_NUMBERS
  if (supportComplex) {
    SmFloat ar = 0.0;
    SmFloat ai = 0.0;
    SmFloat br = 0.0;
    SmFloat bi = 0.0;
    scalarLoadCartesian(a, ar, ai);
    scalarLoadCartesian(b, br, bi);
    return std::isnan(ar) || std::isnan(ai) || std::isnan(br) || std::isnan(bi);
  }
#else
  (void)supportComplex;
#endif
  return std::isnan(a.scalar) || std::isnan(b.scalar);
}

}  // namespace

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
        (void)handleUnknownIdentifier(ctx, e.name, ctx.unknownVarsText);
        return makeScalar(0);
      }
      if (out.kind != ValueKind::Scalar) {
        ctx.fail(ParseErrorId::ScalarOnlyExpressionEncounteredNon);
        return makeScalar(0);
      }
      return out;
    }
    case Expr::Tag::Unary:
      return evalUnaryExpr(e, ctx, scopedVars, true);
    case Expr::Tag::PostfixPercent: {
      EvalValue v = evalExprScalar(*e.child, ctx, scopedVars);
      if (ctx.parseError) return v;
      if (!tryApplyPostfixPercent(ctx, v, getSupportTimeValues(), getSupportComplexNumbers())) {
        return makeScalar(0);
      }
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
        case Expr::BinaryOp::Modulo:
        case Expr::BinaryOp::BitAnd:
        case Expr::BinaryOp::BitOr:
        case Expr::BinaryOp::BitXor:
        case Expr::BinaryOp::ShiftLeft:
        case Expr::BinaryOp::ShiftRight: {
#if SMARTMATH_COMPLEX_NUMBERS
          if (getSupportComplexNumbers() &&
              (evalValueHasNonzeroImaginary(l) || evalValueHasNonzeroImaginary(r))) {
            ctx.fail(ParseErrorId::IncompatibleOperands);
            return makeScalar(0);
          }
#endif
          EvalValue outS;
          if (!applyIntBitwiseOrModulo(
                  ctx, e.binaryOp, l.scalarValue, r.scalarValue,
                  ExactUIntBitwiseGate::OnlyWhenNotBothExactSInt, true, outS)) {
            return makeScalar(0);
          }
          return outS;
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
#if SMARTMATH_COMPLEX_NUMBERS
          const bool lIm = getSupportComplexNumbers() && scalarHasNonzeroImaginaryPart(l.scalarValue);
          const bool rIm = getSupportComplexNumbers() && scalarHasNonzeroImaginaryPart(r.scalarValue);
#else
          const bool lIm = false;
          const bool rIm = false;
#endif
          const bool orderingOp = e.binaryOp == Expr::BinaryOp::CmpLt || e.binaryOp == Expr::BinaryOp::CmpLe ||
                                  e.binaryOp == Expr::BinaryOp::CmpGt || e.binaryOp == Expr::BinaryOp::CmpGe;
          if (orderingOp && (lIm || rIm)) {
            ctx.fail(ParseErrorId::IncompatibleOperands);
            return makeScalar(0);
          }
          if (((lIm || rIm) && (lTime || rTime)) ||
              timeComparedWithNonFinite(lTime, l.scalarValue.scalar, rTime, r.scalarValue.scalar)) {
            ctx.fail(ParseErrorId::IncompatibleOperands);
            return makeScalar(0);
          }
          EvalValue cmpOut;
          if (!tryEvalComparisonOp(ctx, e.binaryOp, l, r, cmpOut)) {
            return makeScalar(0);
          }
          return cmpOut;
        }
        case Expr::BinaryOp::Pow:
          return evalMappedBinaryOp(ctx, e.binaryOp, l, r);
        case Expr::BinaryOp::Mul:
        case Expr::BinaryOp::Div:
        case Expr::BinaryOp::Add:
        case Expr::BinaryOp::Sub: {
          if (!tryScaleRhsByDirectPostfixPercent(
                  ctx, e.binaryOp, e.rhsIsDirectPostfixPercent, l, r)) {
            return makeScalar(0);
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
            const char opCh = (e.binaryOp == Expr::BinaryOp::Mul) ? '*' :
                              (e.binaryOp == Expr::BinaryOp::Add) ? '+' : '-';
            SmInt li = 0;
            SmInt ri = 0;
            if (tryGetBothExactSignedIntNoUIntWrapFromScalars(l.scalarValue, r.scalarValue, li, ri)) {
              EvalValue exactOut;
              if (tryApplyCheckedAddSubMulOrSignedUIntPromote(opCh, li, ri, exactOut)) {
                return exactOut;
              }
            }
            SmUInt lu = 0;
            SmUInt ru = 0;
            if (tryGetBothExactNonNegativeUIntFromScalars(l.scalarValue, r.scalarValue, lu, ru)) {
              EvalValue exactOut;
              if (tryApplyUIntAddOrMul(opCh, lu, ru, exactOut)) {
                return exactOut;
              }
            }
          }

          return evalMappedBinaryOp(ctx, e.binaryOp, l, r);
        }
        default:
          ctx.fail(ParseErrorId::InternalBinaryOp);
          return makeScalar(0);
      }
    }
    default:
      ctx.fail(ParseErrorId::InternalEvalError);
      return makeScalar(0);
  }
}

bool MathParser::checkUnevaluatedCallShapes(EvalContext& ctx, const Expr& e) const {
  if (e.child && !checkUnevaluatedCallShapes(ctx, *e.child)) {
    return false;
  }
  if (e.left && !checkUnevaluatedCallShapes(ctx, *e.left)) {
    return false;
  }
  if (e.right && !checkUnevaluatedCallShapes(ctx, *e.right)) {
    return false;
  }
  bool directUnpack = false;
  for (const auto& ch : e.elements) {
    if (!ch) {
      continue;
    }
    if (e.tag == Expr::Tag::Call && ch->tag == Expr::Tag::Call &&
        ch->builtinFunctionId == BuiltinFunctionId::Unpack) {
      directUnpack = true;
    }
    if (!checkUnevaluatedCallShapes(ctx, *ch)) {
      return false;
    }
  }
  if (ctx.parseError) {
    return false;
  }
  if (e.tag != Expr::Tag::Call || directUnpack) {
    return true;
  }
  if (e.builtinFunctionId != BuiltinFunctionId::Count) {
    uint8_t minArgs = 0;
    uint8_t maxArgs = 0;
    if (!getBuiltinArity(e.builtinFunctionId, minArgs, maxArgs)) {
      return true;
    }
    return validateCallArity(
        ctx, getFunctionName(e.builtinFunctionId), minArgs, maxArgs, e.elements.size());
  }
  const UserFunction* uf = findUserFunction(e.name);
  if (!uf) {
    return true;
  }
  if (uf->params.size() != e.elements.size()) {
    ctx.failCall(CallErrorId::ExactCount, e.name, uf->params.size());
    return false;
  }
  return true;
}

MathParser::EvalValue MathParser::evalIsInstance(
    EvalContext& ctx,
    const Expr& e,
    const std::unordered_map<std::string, EvalValue>* scopedVars) {
  if (e.elements.size() != 2 || !e.elements[0] || !e.elements[1] ||
      e.elements[1]->tag != Expr::Tag::TypeName) {
    ctx.fail(ParseErrorId::InternalEvalError);
    return makeScalar(0);
  }
  EvalValue value = evalExpr(*e.elements[0], ctx, scopedVars);
  if (ctx.parseError) {
    return value;
  }
  if (value.hasExpandArgs()) {
    const std::size_t nItems = (value.kind == ValueKind::Array) ? value.arr.size() : 1u;
    if (nItems != 1u) {
      ctx.failCall(CallErrorId::ExactCount, getFunctionName(BuiltinFunctionId::IsInstance), 2);
      return makeScalar(0);
    }
    if (value.kind == ValueKind::Array) {
      value = scalarFromArrayAt(value, 0);
    }
  }
  const bool match = valueIsInstanceOfType(
      value, e.elements[1]->name, getSupportComplexNumbers(), getSupportTimeValues());
  return makeScalarInt(match ? 1LL : 0LL);
}

MathParser::EvalValue MathParser::evalIif(
    EvalContext& ctx,
    const Expr& e,
    const std::unordered_map<std::string, EvalValue>* scopedVars) {
  if (e.elements.size() != 3 || !e.elements[0] || !e.elements[1] || !e.elements[2]) {
    ctx.failCall(CallErrorId::ExactCount, getFunctionName(BuiltinFunctionId::Iif), 3);
    return makeScalar(0);
  }
  EvalValue cond = evalExpr(*e.elements[0], ctx, scopedVars);
  if (ctx.parseError) {
    return cond;
  }
  const bool takeTrue = isTruthy(cond);
  const Expr& other = takeTrue ? *e.elements[2] : *e.elements[1];
  const Expr& selected = takeTrue ? *e.elements[1] : *e.elements[2];
  if (!checkUnevaluatedCallShapes(ctx, other)) {
    return makeScalar(0);
  }
  return evalExpr(selected, ctx, scopedVars);
}

MathParser::EvalValue MathParser::evalExpr(
    const Expr& e,
    EvalContext& ctx,
    const std::unordered_map<std::string, EvalValue>* scopedVars) {
  switch (e.tag) {
    case Expr::Tag::Literal:
      return e.literalValue;
    case Expr::Tag::Variable: {
      EvalValue out = makeScalar(0);
      if (tryResolveVariableValue(e, scopedVars, out)) {
        return out;
      }
      (void)handleUnknownIdentifier(ctx, e.name, ctx.unknownVarsText);
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
        ctx.fail(ParseErrorId::PercentageRequiresScalarValue);
        return makeScalar(0);
      }
      if (!tryApplyPostfixPercent(ctx, v, getSupportTimeValues(), getSupportComplexNumbers())) {
        return makeScalar(0);
      }
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
        ctx.fail(constructTimeArrayMixedError("array literal", false));
        return makeScalar(0);
      }
#endif
      return makeArrayFromScalars(flatVals);
    }
    case Expr::Tag::Call: {
      if (e.builtinFunctionId == BuiltinFunctionId::Iif) {
        return evalIif(ctx, e, scopedVars);
      }
      if (e.builtinFunctionId == BuiltinFunctionId::IsInstance) {
        return evalIsInstance(ctx, e, scopedVars);
      }
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
        ctx.fail(ParseErrorId::ArrayIndexMustBeScalar);
        return makeScalar(0);
      }
      SmInt idx = 0;
      const ArrayIndexIntKind idxKind = classifyArrayIndexInt(idxv.scalarValue, idx);
      if (idxKind != ArrayIndexIntKind::Ok) {
        ctx.fail(idxKind == ArrayIndexIntKind::OutOfSignedRange
                     ? ParseErrorId::ArrayIndexOutOfRange
                     : ParseErrorId::ArrayIndexMustBeInteger);
        return makeScalar(0);
      }
      if (base.kind != ValueKind::Array) {
        ctx.fail(ParseErrorId::IndexingRequiresArray);
        return makeScalar(0);
      }
      SmInt realIdx = 0;
      if (!tryResolveArrayIndex(static_cast<SmInt>(base.arr.size()), idx, realIdx)) {
        ctx.fail(ParseErrorId::ArrayIndexOutOfRange);
        return makeScalar(0);
      }
      return scalarFromArrayAt(base, static_cast<std::size_t>(realIdx));
    }
    case Expr::Tag::Slice: {
      EvalValue base = evalExpr(*e.left, ctx, scopedVars);
      if (ctx.parseError) return base;

      SliceBoundValues bounds;
      if (!tryEvalSliceBounds(ctx, e, scopedVars, bounds)) {
        return makeScalar(0);
      }
      if (base.kind != ValueKind::Array) {
        ctx.fail(ParseErrorId::IndexingRequiresArray);
        return makeScalar(0);
      }

      const SmInt length = static_cast<SmInt>(base.arr.size());
      SmInt normStart = 0;
      SmInt normStop = 0;
      normalizePythonSliceBounds(
          length, bounds.hasStart, bounds.startIdx, bounds.hasStop, bounds.stopIdx, bounds.stepIdx,
          normStart, normStop);
      const SmInt outCount = estimateSliceItemCount(normStart, normStop, bounds.stepIdx);
      if (outCount <= 0) {
        return makeArrayFromScalars({});
      }
      EvalValue sliced;
      sliced.kind = ValueKind::Array;
      sliced.scalarValue.scalarKind = ScalarKind::FloatingPoint;
      sliced.arr.resize(static_cast<std::size_t>(outCount));
      SmInt cur = normStart;
      for (SmInt i = 0; i < outCount; ++i) {
        sliced.arr[static_cast<std::size_t>(i)] = base.arr[static_cast<std::size_t>(cur)];
        if (i + 1 < outCount) {
          cur += bounds.stepIdx;
        }
      }
      return sliced;
    }
    case Expr::Tag::Binary: {
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
        return evalSIntBinaryOp(ctx, l, r, e.binaryOp);
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
            ctx.fail(ParseErrorId::IncompatibleOperands);
            return makeScalar(0);
          }
#if SMARTMATH_TIME_VALUES
          if ((valueHasComplex(l) || valueHasComplex(r)) &&
              (evalValueInvolvesTime(l) || evalValueInvolvesTime(r))) {
            ctx.fail(ParseErrorId::IncompatibleOperands);
            return makeScalar(0);
          }
#endif
        }
#endif
        if (l.kind == ValueKind::Scalar && r.kind == ValueKind::Scalar) {
#if SMARTMATH_TIME_VALUES
          const bool lTime = scalarValueIsTime(l.scalarValue);
          const bool rTime = scalarValueIsTime(r.scalarValue);
          if (timeComparedWithNonFinite(lTime, l.scalarValue.scalar, rTime, r.scalarValue.scalar)) {
            ctx.fail(ParseErrorId::IncompatibleOperands);
            return makeScalar(0);
          }
#endif
        }
        EvalValue cmpOut;
        if (!tryEvalComparisonOp(ctx, e.binaryOp, l, r, cmpOut)) {
          return makeScalar(0);
        }
        return cmpOut;
      }
      EvalValue l = evalExpr(*e.left, ctx, scopedVars);
      if (ctx.parseError) {
        return l;
      }
      EvalValue r = evalExpr(*e.right, ctx, scopedVars);
      if (ctx.parseError) {
        return r;
      }
      if (!tryScaleRhsByDirectPostfixPercent(
              ctx, e.binaryOp, e.rhsIsDirectPostfixPercent, l, r)) {
        return makeScalar(0);
      }
      if (e.binaryOp == Expr::BinaryOp::Pow) {
        return evalMappedBinaryOp(ctx, e.binaryOp, l, r);
      }
      if (e.binaryOp == Expr::BinaryOp::Mul || e.binaryOp == Expr::BinaryOp::Div ||
          e.binaryOp == Expr::BinaryOp::Add || e.binaryOp == Expr::BinaryOp::Sub) {
        return evalMappedBinaryOp(ctx, e.binaryOp, l, r);
      }
      ctx.fail(ParseErrorId::InternalBinaryOp);
      return makeScalar(0);
    }
    default:
      break;
  }
  ctx.fail(ParseErrorId::InternalEvalError);
  return makeScalar(0);
}

bool MathParser::tryEvalSliceBounds(
    EvalContext& ctx,
    const Expr& slice,
    const std::unordered_map<std::string, EvalValue>* scopedVars,
    SliceBoundValues& bounds) {
  const auto evalOptionalBound = [&](std::size_t slot, bool& hasBound, SmInt& outBound) -> bool {
    hasBound = (slot < slice.elements.size() && slice.elements[slot] != nullptr);
    if (!hasBound) {
      outBound = 0;
      return true;
    }
    const EvalValue bv = evalExpr(*slice.elements[slot], ctx, scopedVars);
    if (ctx.parseError) {
      return false;
    }
    if (bv.kind != ValueKind::Scalar) {
      ctx.fail(ParseErrorId::ArrayIndexMustBeScalar);
      return false;
    }
    const ArrayIndexIntKind boundKind = classifyArrayIndexInt(bv.scalarValue, outBound);
    if (boundKind == ArrayIndexIntKind::OutOfSignedRange) {
      ctx.fail(ParseErrorId::ArrayIndexOutOfRange);
      return false;
    }
    if (boundKind != ArrayIndexIntKind::Ok) {
      ctx.fail(ParseErrorId::ArrayIndexMustBeInteger);
      return false;
    }
    return true;
  };

  bounds = SliceBoundValues{};
  if (!evalOptionalBound(0, bounds.hasStart, bounds.startIdx) || ctx.parseError) {
    return false;
  }
  if (!evalOptionalBound(1, bounds.hasStop, bounds.stopIdx) || ctx.parseError) {
    return false;
  }
  if (!evalOptionalBound(2, bounds.hasStep, bounds.stepIdx) || ctx.parseError) {
    return false;
  }
  if (!bounds.hasStep) {
    bounds.stepIdx = 1;
  }
  if (bounds.hasStep && bounds.stepIdx == 0) {
    ctx.fail(STR_SLICE_STEP_MUST_NOT_BE_ZERO);
    return false;
  }
  return true;
}

bool MathParser::tryEvalElementAssign(EvalContext& ctx, const AstStatement& st, EvalValue& out) {
  const auto unresolved = [&ctx]() {
    return ctx.parseError || !ctx.unknownVarsText.empty() || !ctx.unknownFuncsText.empty();
  };
  auto found = variables_.find(st.assignName);
  if (found == variables_.end()) {
    appendUniqueName(ctx.unknownVarsText, st.assignName);
    ctx.referencedUnresolvedName = true;
    return false;
  }
  EvalValue& slot = found->second;
  if (slot.kind != ValueKind::Array) {
    ctx.fail(ParseErrorId::IndexingRequiresArray);
    return false;
  }
  if (!st.indexExpr) {
    ctx.fail(ParseErrorId::InternalEvalError);
    return false;
  }

  SmInt resolved = 0;
  if (st.indexExpr->tag == Expr::Tag::Slice) {
    SliceBoundValues bounds;
    if (!tryEvalSliceBounds(ctx, *st.indexExpr, nullptr, bounds)) {
      return false;
    }
    if (unresolved()) {
      return false;
    }
    ctx.fail(ParseErrorId::CannotAssignToArraySlice);
    return false;
  }

  if (st.indexExpr->tag != Expr::Tag::Index || !st.indexExpr->right) {
    ctx.fail(ParseErrorId::InternalEvalError);
    return false;
  }
  const EvalValue idxv = evalExpr(*st.indexExpr->right, ctx, nullptr);
  if (unresolved()) {
    return false;
  }
  if (idxv.kind != ValueKind::Scalar) {
    ctx.fail(ParseErrorId::ArrayIndexMustBeScalar);
    return false;
  }
  SmInt idx = 0;
  const ArrayIndexIntKind idxKind = classifyArrayIndexInt(idxv.scalarValue, idx);
  if (idxKind == ArrayIndexIntKind::OutOfSignedRange) {
    ctx.fail(ParseErrorId::ArrayIndexOutOfRange);
    return false;
  }
  if (idxKind != ArrayIndexIntKind::Ok) {
    ctx.fail(ParseErrorId::ArrayIndexMustBeInteger);
    return false;
  }
  if (!tryResolveArrayIndex(static_cast<SmInt>(slot.arr.size()), idx, resolved)) {
    ctx.fail(ParseErrorId::ArrayIndexOutOfRange);
    return false;
  }

  if (!st.expr) {
    ctx.fail(ParseErrorId::InternalEvalError);
    return false;
  }
  const EvalValue rhs = evalExpr(*st.expr, ctx, nullptr);
  if (unresolved()) {
    return false;
  }
  if (rhs.kind != ValueKind::Scalar && rhs.kind != ValueKind::Array) {
    ctx.fail(ParseErrorId::IncompatibleOperands);
    return false;
  }

#if SMARTMATH_TIME_VALUES
  if (!(rhs.kind == ValueKind::Array && rhs.arr.empty())) {
    bool baseAllTime = !slot.arr.empty();
    bool baseAllNonTime = !slot.arr.empty();
    for (const auto& item : slot.arr) {
      if (scalarValueIsTime(item)) {
        baseAllNonTime = false;
      } else {
        baseAllTime = false;
      }
    }
    const auto rejects = [&](bool rhsTime) {
      return (baseAllTime && !rhsTime) || (baseAllNonTime && rhsTime);
    };
    bool mixed = false;
    if (rhs.kind == ValueKind::Scalar) {
      mixed = rejects(scalarValueIsTime(rhs.scalarValue));
    } else {
      for (const auto& item : rhs.arr) {
        if (rejects(scalarValueIsTime(item))) {
          mixed = true;
          break;
        }
      }
    }
    if (mixed) {
      ctx.fail(STR_TIME_ARRAY_MIXED);
      return false;
    }
  }
#endif

  const std::size_t at = static_cast<std::size_t>(resolved);
  if (rhs.kind == ValueKind::Scalar) {
    slot.arr[at] = rhs.scalarValue;
  } else {
    const std::size_t oldLen = slot.arr.size();
    EvalValue updated;
    updated.kind = ValueKind::Array;
    updated.scalarValue.scalarKind = ScalarKind::FloatingPoint;
    updated.arr.reserve(oldLen - 1 + rhs.arr.size());
    updated.arr.insert(updated.arr.end(), slot.arr.begin(), slot.arr.begin() + static_cast<std::ptrdiff_t>(at));
    updated.arr.insert(updated.arr.end(), rhs.arr.begin(), rhs.arr.end());
    updated.arr.insert(updated.arr.end(), slot.arr.begin() + static_cast<std::ptrdiff_t>(at + 1), slot.arr.end());
    slot = std::move(updated);
  }
  out = slot;
  if (!ctx.referencedUnresolvedName) {
    setVariable(STR_ANS, out);
  }
  return true;
}

MathParser::EvalValue MathParser::runCompiledProgram(
    EvalContext& ctx,
    const std::vector<AstStatement>& program,
    const std::unordered_map<std::string, EvalValue>* scopedVars,
    bool scalarOnlyMode) {
  EvalValue out = makeScalarInt(0);
  const auto evalIntoOut = [&](const Expr& ex) {
    if (!scalarOnlyMode) {
      out = evalExpr(ex, ctx, scopedVars);
      return;
    }
    out = evalExprScalar(ex, ctx, scopedVars);
    // Compile time cannot know that a variable holds an array. Retry on the general path.
    if (ctx.parseError && ctx.errorText == STR_SCALAR_ONLY_EXPRESSION_ENCOUNTERED_NON) {
      ctx.parseError = false;
      ctx.errorText.clear();
      out = evalExpr(ex, ctx, scopedVars);
    }
  };
  for (const auto& st : program) {
    ctx.referencedUnresolvedName = false;
    if (st.kind == AstStatement::Kind::FunDef) {
      const std::string udfErr =
          getUserFunctionDefinitionErrorText(st.fun.name, st.fun.params, st.fun.expr, true);
      if (!udfErr.empty()) {
        ctx.fail(udfErr);
        return out;
      }
      upsertUserFunction(UserFunction{st.fun.name, st.fun.params, st.fun.expr});
      EvalValue defined;
      defined.kind = ValueKind::DefinedFunction;
      defined.funcRefName = std::string(STR_DEFINED) + formatUserFunctionSignature(st.fun);
      out = std::move(defined);
    } else if (st.kind == AstStatement::Kind::Assign) {
      evalIntoOut(*st.expr);
      if (ctx.parseError || !ctx.unknownVarsText.empty() || !ctx.unknownFuncsText.empty()) {
        return out;
      }
      removeUserFunctionByName(st.assignName);
      setVariable(st.assignName, out);
      if (!ctx.referencedUnresolvedName) {
        setVariable(STR_ANS, out);
      }
    } else if (st.kind == AstStatement::Kind::IndexAssign) {
      if (!tryEvalElementAssign(ctx, st, out)) {
        return out;
      }
    } else {
      evalIntoOut(*st.expr);
      if (ctx.parseError) {
        return out;
      }
      if (!ctx.referencedUnresolvedName) {
        setVariable(STR_ANS, out);
      }
    }
  }
  return out;
}

bool MathParser::validateCallArity(
    EvalContext& ctx,
    const std::string& fnName,
    uint8_t minArgs,
    uint8_t maxArgs,
    const std::size_t argc) const {
  if (minArgs == maxArgs) {
    if (argc != minArgs) {
      ctx.failCall(CallErrorId::ExactCount, fnName, minArgs);
      return false;
    }
    return true;
  }
  if (argc < minArgs) {
    ctx.failCall(CallErrorId::AtLeastN, fnName, minArgs);
    return false;
  }
  if (maxArgs != kBuiltinArityUnbounded && argc > maxArgs) {
    ctx.failCall(CallErrorId::ExactCount, fnName, maxArgs);
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
    if (!(s.hasExactSInt() || s.hasExactUInt())) {
      SmInt signedV = 0;
      if (!tryGetSignedIntFromScalar(s, signedV)) {
        ctx.failCall(CallErrorId::IntegerValues, fnName);
        return false;
      }
    }
#if SMARTMATH_COMPLEX_NUMBERS
    if (getSupportComplexNumbers() && scalarHasNonzeroImaginaryPart(s)) {
      SmFloat ai = 0.0;
      SmFloat ar = 0.0;
      scalarLoadCartesian(s, ar, ai);
      (void)ar;
      if (!std::isfinite(ai)) {
        if (!allowNonFiniteForFormat) {
          ctx.failCall(CallErrorId::IntegerValues, fnName);
          return false;
        }
        return true;
      }
      if (s.hasImagExactSInt()) {
        return true;
      }
      SmInt tmp = 0;
      if (!tryExtractExactSIntFromFloatStrict(ai, tmp)) {
        ctx.failCall(CallErrorId::IntegerValues, fnName);
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
    ctx.failCall(CallErrorId::IntegerValues, fnName);
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
#if SMARTMATH_TIME_VALUES
    for (const EvalValue& arg : args) {
      if (evalValueInvolvesTime(arg)) {
        ctx.failCall(CallErrorId::IntegerValues, fnName);
        return false;
      }
    }
#endif
    return validateIntegerRepresentableArgs(ctx, fnName, args, true);
  }
  if (hasBuiltinFlag(id, BuiltinFlags::IntegerOnly)) {
    return validateIntegerRepresentableArgs(ctx, fnName, args, false);
  }
  if (hasBuiltinFlag(id, BuiltinFlags::NonCalculating)) {
    return true;
  }
  if (hasBuiltinFlag(id, BuiltinFlags::FiniteRequired) && argsContainNonFinite(args)) {
    ctx.failCall(CallErrorId::Numeric, fnName);
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
    if (!args.empty()) {
      return markExpanded(makeArrayFromScalars({}));
    }
    ctx.failCall(CallErrorId::AtLeastN, getFunctionName(BuiltinFunctionId::Unpack), 1);
    return makeScalar(0);
  }
  return markExpanded(makeArrayFromScalars(elems));
}

MathParser::EvalValue MathParser::builtinLen(
    EvalContext& /*ctx*/,
    const std::vector<EvalValue>& args) const {
  return makeScalarInt(static_cast<SmInt>(countAggregateArgScalarItems(args)));
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
  const bool emptyArray = vals.empty() && args[0].kind == ValueKind::Array;
  if (vals.empty() && !emptyArray) {
    ctx.failCall(CallErrorId::AtLeastN, fnName, 1);
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
    ctx.failCall(CallErrorId::IntegerValues, fnName);
    return makeScalar(0);
  }

  SmInt n = 0;
  if (!tryGetSignedIntFromScalar(countArg.scalarValue, n)) {
    ctx.failCall(CallErrorId::IntegerValues, fnName);
    return makeScalar(0);
  }
  if (n <= 0) {
    ctx.failCall(CallErrorId::PositiveInteger, fnName);
    return makeScalar(0);
  }
  if (emptyArray) {
    return makeArrayFromScalars({});
  }

  const std::size_t blockSize = vals.size();
  SmUInt estimated = 0;
  if (!tryMulUIntChecked(static_cast<SmUInt>(blockSize), static_cast<SmUInt>(n), estimated)) {
    ctx.failCall(CallErrorId::TooManyValues, fnName);
    return makeScalar(0);
  }
  if (!acceptProducedItemCount(ctx, fnName, estimated)) {
    return makeScalar(0);
  }

  const std::size_t totalSize = static_cast<std::size_t>(estimated);
  std::vector<EvalValue> out;
  out.reserve(totalSize);
  for (SmInt rep = 0; rep < n; ++rep) {
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
  const auto tryGetRangeIntegerScalarArg = [this](const EvalValue& v, SmInt& outN) -> bool {
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
    return tryGetSignedIntFromScalar(v.scalarValue, outN);
  };

  SmInt startV = 0;
  SmInt stopV = 0;
  SmInt stepV = 1;
  if (!tryGetRangeIntegerScalarArg(args[0], startV) || !tryGetRangeIntegerScalarArg(args[1], stopV)) {
    ctx.failCall(CallErrorId::IntegerValues, fnName);
    return makeScalar(0);
  }
  if (args.size() >= 3U) {
    if (!tryGetRangeIntegerScalarArg(args[2], stepV)) {
      ctx.failCall(CallErrorId::IntegerValues, fnName);
      return makeScalar(0);
    }
  }

  if (stepV == 0) {
    ctx.failCall(CallErrorId::RangeStepZero, fnName);
    return makeScalar(0);
  }
  if ((stepV > 0 && startV >= stopV) || (stepV < 0 && startV <= stopV)) {
    return makeArrayFromScalars({});
  }

  const SmUInt estimated = estimateRangeItemCount(startV, stopV, stepV);
  if (estimated == 0ull) {
    return makeArrayFromScalars({});
  }
  if (!acceptProducedItemCount(ctx, fnName, estimated)) {
    return makeScalar(0);
  }

  const std::size_t count = static_cast<std::size_t>(estimated);
  std::vector<EvalValue> out;
  out.reserve(count);
  SmInt cur = startV;
  for (std::size_t i = 0; i < count; ++i) {
    out.push_back(makeScalarInt(cur));
    if (i + 1U < count) {
      cur += stepV;
    }
  }
  return makeArrayFromScalars(out);
}

std::size_t MathParser::countAggregateArgScalarItems(const std::vector<EvalValue>& args) const {
  std::size_t itemCount = 0;
  forEachCallArgScalarValues(args, [&](const EvalValue::ScalarValue&) { ++itemCount; });
  return itemCount;
}

template <typename IntT, typename ReadFn, typename AddFn, typename MakeFn>
bool tryFoldExactSumMinMax(
    BuiltinFunctionId id,
    const std::vector<EvalValue>& args,
    ReadFn read,
    AddFn add,
    MakeFn make,
    EvalValue& out) {
  bool ok = true;
  IntT acc = 0;
  bool hasMm = false;
  IntT mm = 0;
  const bool wantMin = (id == BuiltinFunctionId::Min);
  const bool isSum = (id == BuiltinFunctionId::Sum);
  forEachCallArgScalarValues(args, [&](const EvalValue::ScalarValue& s) {
    IntT v = 0;
    (void)read(s, v);
    if (isSum) {
      IntT next = 0;
      if (!add(acc, v, next)) {
        ok = false;
      }
      acc = next;
    } else if (!hasMm || (wantMin ? (v < mm) : (v > mm))) {
      mm = v;
      hasMm = true;
    }
  });
  if (!ok) {
    return false;
  }
  out = isSum ? make(acc) : make(mm);
  return true;
}

MathParser::EvalValue MathParser::aggregateFoldByMode(
    EvalContext& ctx,
    const std::string& fnName,
    BuiltinFunctionId id,
    const std::vector<EvalValue>& args,
    AggregateFoldMode mode) const {
  if (args.empty()) {
    ctx.failCall(CallErrorId::AtLeastN, fnName, 1);
    return makeScalar(0);
  }
  if (countAggregateArgScalarItems(args) == 0) {
    return makeArrayFromScalars({});
  }
  switch (mode) {
#if SMARTMATH_COMPLEX_NUMBERS
  case AggregateFoldMode::Complex: {
  const auto forEachArgScalarValue = [&](const auto& fn) -> std::size_t {
    return forEachCallArgScalarValues(args, fn);
  };
  if (args.empty()) {
    ctx.failCall(CallErrorId::AtLeastN, fnName, 1);
    return makeScalar(0);
  }
  if (args.size() == 1 && args[0].kind == ValueKind::Scalar) {
    return args[0];
  }
  const std::size_t itemCount = countAggregateArgScalarItems(args);
  if (itemCount == 0) {
    ctx.failCall(CallErrorId::AtLeastN, fnName, 1);
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
      SmFloat ar = 0.0;
      SmFloat ai = 0.0;
      scalarLoadCartesian(folded.scalarValue, ar, ai);
      ar /= static_cast<SmFloat>(itemCount);
      ai /= static_cast<SmFloat>(itemCount);
      return makeScalarComplexFromFloats(ar, ai);
    }
    return folded;
  }
  SmFloat ar = 0.0;
  SmFloat ai = 0.0;
  if (id == BuiltinFunctionId::Product) {
    ar = 1.0;
    ai = 0.0;
  }
  SmFloat br = 0.0;
  SmFloat bi = 0.0;
  forEachArgScalarValue([&](const EvalValue::ScalarValue& s) {
    scalarLoadCartesian(s, br, bi);
    if (id == BuiltinFunctionId::Product) {
      // A later 0 or sign change still applies after Inf (Inf*0 is NaN).
      complexMultiply(ar, ai, br, bi, ar, ai);
    } else {
      ar += br;
      ai += bi;
    }
  });
  if (id == BuiltinFunctionId::Avg || id == BuiltinFunctionId::Mean) {
    ar /= static_cast<SmFloat>(itemCount);
    ai /= static_cast<SmFloat>(itemCount);
  }
  return makeScalarComplexFromFloats(ar, ai);
  }
#endif
#if SMARTMATH_TIME_VALUES
  case AggregateFoldMode::Time: {
  const auto forEachArgScalarValue = [&](const auto& fn) -> std::size_t {
    return forEachCallArgScalarValues(args, fn);
  };
  if (id == BuiltinFunctionId::Median) {
    std::vector<SmFloat> flatMs;
    bool okM = true;
    forEachArgScalarValue([&](const EvalValue::ScalarValue& s) {
      if (!okM) {
        return;
      }
      if (!scalarValueIsTime(s)) {
        ctx.fail(constructTimeArrayMixedError(fnName));
        okM = false;
        return;
      }
      flatMs.push_back(static_cast<SmFloat>(timeTotalMsFromScalarValue(s)));
    });
    if (!okM) {
      return makeScalar(0);
    }
    if (flatMs.empty()) {
      ctx.failCall(CallErrorId::AtLeastN, fnName, 1);
      return makeScalar(0);
    }
    return makeScalarTimeMs(roundHalfUpFloatToSInt(MedianFromFloatFlatInPlace(flatMs)));
  }

  SmInt accMs = 0;
  bool accInit = false;
  std::size_t itemCount = 0;
  bool okAgg = true;
  forEachArgScalarValue([&](const EvalValue::ScalarValue& s) {
    if (!okAgg) {
      return;
    }
    if (!scalarValueIsTime(s)) {
      ctx.fail(constructTimeArrayMixedError(fnName));
      okAgg = false;
      return;
    }
    const SmInt curMs = timeTotalMsFromScalarValue(s);
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
      SmInt nacc = 0;
      if (!tryAddTimeMsChecked(accMs, curMs, nacc)) {
        ctx.fail(ParseErrorId::IncompatibleOperands);
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
    ctx.failCall(CallErrorId::AtLeastN, fnName, 1);
    return makeScalar(0);
  }
  if (id == BuiltinFunctionId::Avg || id == BuiltinFunctionId::Mean) {
    const SmFloat avgD = static_cast<SmFloat>(accMs) / static_cast<SmFloat>(itemCount);
    return makeScalarTimeMs(roundHalfUpFloatToSInt(avgD));
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
        SmUInt u = 0;
        SmInt i = 0;
        const bool nnU = tryGetExactNonNegativeUIntFromScalar(s, u);
        const bool swI = tryGetExactSignedIntNoUIntWrapFromScalar(s, i);
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
        ctx.failCall(CallErrorId::AtLeastN, fnName, 1);
        return makeScalar(0);
      }
      if (!anyNonInteger && id == BuiltinFunctionId::Product) {
        bool ok = true;
        SmUInt productMag = 1;
        bool isNegativeProduct = false;
        forEachArgScalarValue([&](const EvalValue::ScalarValue& s) {
          if (!ok) {
            return;
          }
          SmUInt termMag = 0;
          if (!tryGetExactNonNegativeUIntFromScalar(s, termMag)) {
            SmInt termI = 0;
            if (!tryGetExactSignedIntNoUIntWrapFromScalar(s, termI)) {
              ok = false;
              return;
            }
            if (termI < 0) {
              isNegativeProduct = !isNegativeProduct;
              termMag = (termI == kIntMin)
                ? kIntMinMag
                : static_cast<SmUInt>(-termI);
            } else {
              termMag = static_cast<SmUInt>(termI);
            }
          }
          SmUInt next = 0;
          if (!tryMulUIntChecked(productMag, termMag, next)) {
            ok = false;
            return;
          }
          productMag = next;
        });
        if (ok) {
          if (isNegativeProduct) {
            if (productMag <= kIntMinMag) {
              return (productMag == kIntMinMag)
                ? makeScalarInt(kIntMin)
                : makeScalarInt(-static_cast<SmInt>(productMag));
            }
          } else {
            return makeScalarUInt(productMag);
          }
        }
      } else if (!anyNonInteger && allNnU) {
        EvalValue folded;
        if (tryFoldExactSumMinMax<SmUInt>(
                id, args,
                [](const EvalValue::ScalarValue& s, SmUInt& u) {
                  return tryGetExactNonNegativeUIntFromScalar(s, u);
                },
                [](SmUInt a, SmUInt b, SmUInt& o) {
                  return tryAddUIntChecked(a, b, o);
                },
                [](SmUInt u) { return makeScalarUInt(u); },
                folded)) {
          return folded;
        }
      } else if (!anyNonInteger && allSwI) {
        EvalValue folded;
        if (tryFoldExactSumMinMax<SmInt>(
                id, args,
                [](const EvalValue::ScalarValue& s, SmInt& i) {
                  return tryGetExactSignedIntNoUIntWrapFromScalar(s, i);
                },
                [](SmInt a, SmInt b, SmInt& o) {
                  return tryAddSIntChecked(a, b, o);
                },
                [](SmInt i) { return makeScalarInt(i); },
                folded)) {
          return folded;
        }
      }
    }
    if (id == BuiltinFunctionId::Min || id == BuiltinFunctionId::Max) {
      const bool wantMin = (id == BuiltinFunctionId::Min);
      const auto hasExactIntegerMetadata = [](const EvalValue::ScalarValue& sv) {
        if (sv.hasRenderRational()) {
          return false;
        }
        return sv.hasExactSInt() || sv.hasExactUInt();
      };
      const auto shouldReplace = [&](SmFloat v, SmFloat winnerV, bool hasWinner) {
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
      SmFloat winnerV = 0.0;
      bool hasWinner = false;
      bool hasNonNanWinner = false;
      std::size_t itemCount = 0;
      forEachArgScalarValue([&](const EvalValue::ScalarValue& s) {
        ++itemCount;
        const SmFloat v = scalarNumericReal(s);
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
        ctx.failCall(CallErrorId::AtLeastN, fnName, 1);
        return makeScalar(0);
      }
      return scalarFromScalarValue(winnerSv);
    }
    SmFloat acc = 0.0;
    std::size_t n = 0;
    if (id == BuiltinFunctionId::Product) {
      acc = 1.0;
      // A later 0 or sign change still applies after Inf (Inf*0 is NaN).
      n = forEachArgScalar([&](SmFloat v) { acc *= v; });
    } else {
      acc = 0.0;
      n = forEachArgScalar([&](SmFloat v) { acc += v; });
    }

    if (n == 0) {
      ctx.failCall(CallErrorId::AtLeastN, fnName, 1);
        return makeScalar(0);
    }
    if (id == BuiltinFunctionId::Avg || id == BuiltinFunctionId::Mean) {
      acc /= static_cast<SmFloat>(n);
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
      std::vector<SmFloat> flat;
      if (!flattenArgs(args, flat)) {
        ctx.failCall(CallErrorId::AtLeastN, fnName, 1);
        return makeScalar(0);
      }
      return makeScalar(MedianFromFloatFlatInPlace(flat));
    }
    case BuiltinFunctionId::Variance:
    case BuiltinFunctionId::Stddev: {
#if SMARTMATH_TIME_VALUES
      if (std::any_of(args.begin(), args.end(), [this](const EvalValue& v) { return evalValueInvolvesTime(v); })) {
        ctx.fail(ParseErrorId::IncompatibleOperands);
        return makeScalar(0);
      }
#endif
      // Welford single-pass accumulation for better stability and fewer passes.
      SmFloat mean = 0.0;
      SmFloat m2 = 0.0;
      std::size_t n = 0;
      forEachArgScalar([&](SmFloat v) {
        ++n;
        const SmFloat delta = v - mean;
        mean += delta / static_cast<SmFloat>(n);
        const SmFloat delta2 = v - mean;
        m2 += delta * delta2;
      });
      if (n == 0) {
        ctx.failCall(CallErrorId::AtLeastN, fnName, 1);
        return makeScalar(0);
      }
      SmFloat var = m2 / static_cast<SmFloat>(n);
      if (id == BuiltinFunctionId::Variance) {
        return makeScalar(var);
      }
      return makeScalar(std::sqrt(var));
    }
    default:
      break;
  }
  ctx.fail(ParseErrorId::InternalAggregateBuiltin);
  return makeScalar(0);
  }
  }
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
    ctx.fail(ParseErrorId::IncompatibleOperands);
    return makeScalar(0);
  }
#else
  const bool anyTimeAgg = false;
#endif
#if SMARTMATH_COMPLEX_NUMBERS
  if (getSupportComplexNumbers()) {
    const bool anyComplex = std::any_of(
        args.begin(), args.end(),
        [](const EvalValue& v) { return evalValueHasNonzeroImaginary(v); });
    if (anyComplex) {
#if SMARTMATH_TIME_VALUES
      if (anyTimeAgg) {
        ctx.fail(ParseErrorId::IncompatibleOperands);
        return makeScalar(0);
      }
#endif
      if (id == BuiltinFunctionId::Min || id == BuiltinFunctionId::Max || id == BuiltinFunctionId::Median ||
          id == BuiltinFunctionId::Variance || id == BuiltinFunctionId::Stddev) {
        ctx.fail(ParseErrorId::IncompatibleOperands);
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
      ctx.fail(ParseErrorId::IncompatibleOperands);
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

bool MathParser::isEligibleUnaryFunctionName(
    const std::string& funcName,
    std::string& outErr,
    const char* expectsUnaryFunction) const {
  const std::string lowFn = toLower(funcName);
  BuiltinFunctionId fnId = BuiltinFunctionId::Count;
  if (tryGetBuiltinFunctionId(lowFn, fnId)) {
    if (isIneligibleUnaryBuiltin(fnId)) {
      outErr = expectsUnaryFunction;
      return false;
    }
    return true;
  }
  const UserFunction* uf = findUserFunction(lowFn);
  if (!uf) {
    if (variables_.find(lowFn) != variables_.end()) {
      outErr = expectsUnaryFunction;
      return false;
    }
    outErr = std::string(STR_UNKNOWN_FUNCTION_COLON) + funcName;
    return false;
  }
  if (uf->params.size() != 1U) {
    outErr = expectsUnaryFunction;
    return false;
  }
  return true;
}

bool MathParser::tryLexicographicCompareEvalValues(
    EvalContext& ctx,
    const EvalValue& a,
    const EvalValue& b,
    int& cmpOut,
    CmpScalarIncompatiblePolicy policy) const {
#if SMARTMATH_COMPLEX_NUMBERS
  if (!getSupportComplexNumbers()) {
    if (evalValueHasNonzeroImaginary(a) || evalValueHasNonzeroImaginary(b)) {
      ctx.fail(ParseErrorId::IncompatibleOperands);
      cmpOut = 0;
      return false;
    }
  }
#endif
  const auto refAt = [](const EvalValue& v, std::size_t i) -> const EvalValue::ScalarValue& {
    return (v.kind == ValueKind::Scalar) ? v.scalarValue : v.arr[i];
  };
  const std::size_t na = (a.kind == ValueKind::Scalar) ? 1U : a.arr.size();
  const std::size_t nb = (b.kind == ValueKind::Scalar) ? 1U : b.arr.size();
  const std::size_t n = (na < nb) ? na : nb;
  for (std::size_t i = 0; i < n; ++i) {
    int c = 0;
    if (!cmpScalarValuesForCompare(&ctx, refAt(a, i), refAt(b, i), c, policy)) {
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

bool MathParser::tryEvalComparisonOp(
    EvalContext& ctx,
    Expr::BinaryOp op,
    const EvalValue& a,
    const EvalValue& b,
    EvalValue& out) const {
#if SMARTMATH_COMPLEX_NUMBERS
  if (!getSupportComplexNumbers()) {
    if (evalValueHasNonzeroImaginary(a) || evalValueHasNonzeroImaginary(b)) {
      ctx.fail(ParseErrorId::IncompatibleOperands);
      return false;
    }
  }
#endif
  const auto refAt = [](const EvalValue& v, std::size_t i) -> const EvalValue::ScalarValue& {
    return (v.kind == ValueKind::Scalar) ? v.scalarValue : v.arr[i];
  };
  const std::size_t na = (a.kind == ValueKind::Scalar) ? 1U : a.arr.size();
  const std::size_t nb = (b.kind == ValueKind::Scalar) ? 1U : b.arr.size();
  const std::size_t n = (na < nb) ? na : nb;
  const bool supportComplex = getSupportComplexNumbers();
  for (std::size_t i = 0; i < n; ++i) {
    const EvalValue::ScalarValue& lsv = refAt(a, i);
    const EvalValue::ScalarValue& rsv = refAt(b, i);
    if (scalarPairHasUnorderedNan(lsv, rsv, supportComplex)) {
      out = makeScalarInt(evalComparisonTruthWhenUnorderedNan(op) ? 1LL : 0LL);
      return true;
    }
    int c = 0;
    if (!cmpScalarValuesForCompare(
            &ctx, lsv, rsv, c, CmpScalarIncompatiblePolicy::SetError)) {
      return false;
    }
    if (ctx.parseError) {
      return false;
    }
    if (c != 0) {
      out = makeScalarInt(evalComparisonByOp(op, c) ? 1LL : 0LL);
      return true;
    }
  }
  int cmp = 0;
  if (na < nb) {
    cmp = -1;
  } else if (na > nb) {
    cmp = 1;
  } else {
    cmp = 0;
  }
  out = makeScalarInt(evalComparisonByOp(op, cmp) ? 1LL : 0LL);
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
              cmp,
              CmpScalarIncompatiblePolicy::SortLess)) {
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

bool MathParser::tryBuiltinRatioScalar(EvalContext& ctx, const EvalValue::ScalarValue& sv, EvalValue& outV) const {
#if SMARTMATH_TIME_VALUES
  if (scalarValueIsTime(sv)) {
    ctx.parseError = true;
    return false;
  }
#endif
  SmFloat ar = 0.0;
  SmFloat ai = 0.0;
  scalarLoadCartesian(sv, ar, ai);
  if (std::isnan(ar) || std::isnan(ai)) {
    if (std::isnan(ar)) {
      if (std::isnan(ai) || std::fabs(ai) < RATIO_APPROX_EPS) {
        outV = makeScalar(std::numeric_limits<SmFloat>::quiet_NaN());
        return true;
      }
    }
    ctx.parseError = true;
    return false;
  }
#if SMARTMATH_COMPLEX_NUMBERS
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
  if (getSupportComplexNumbers() && scalarHasNonzeroImaginaryPart(sv)) {
    SmInt numR = 0;
    SmUInt denR = 1;
    SmInt numI = 0;
    SmUInt denI = 1;
    ExactCartesianComponent reC;
    SmFloat aiScale = std::fabs(ai);
    if (aiScale < 1.0) {
      aiScale = 1.0;
    }
    SmFloat arScale = std::fabs(ar);
    if (arScale < 1.0) {
      arScale = 1.0;
    }
    outV = scalarFromScalarValue(sv);
    if (sv.hasRenderRational() || sv.hasImagRenderRational()) {
      return true;
    }
    scalarClearCartesianRenderExact(outV.scalarValue);
    if (sv.hasImagExactSInt() && imagExactMetadataMatchesFloat(sv)) {
      ScalarApplyExactSIntPart(outV.scalarValue, true, sv.imagExactSInt);
    } else if (std::fabs(ai) >= RATIO_APPROX_EPS * aiScale && std::fabs(ai) <= kMaxExactIntFromFloat) {
      if (!tryApproximateRational(ai, numI, denI)) {
        ctx.parseError = true;
        return false;
      }
      scalarApplyReducedRationalPart(outV.scalarValue, true, numI, denI);
    }
    if (tryExtractExactRealComponent(sv, reC) && TryExactCartesianComponentToSInt(reC, numR)) {
      ScalarApplyExactSIntPart(outV.scalarValue, false, numR);
    } else if (std::fabs(ar) >= RATIO_APPROX_EPS * arScale && std::fabs(ar) <= kMaxExactIntFromFloat) {
      if (!tryApproximateRational(ar, numR, denR)) {
        ctx.parseError = true;
        return false;
      }
      scalarApplyReducedRationalPart(outV.scalarValue, false, numR, denR);
    } else if (sv.hasExactSInt()) {
      ScalarApplyExactSIntPart(outV.scalarValue, false, sv.exactSInt);
    }
    return true;
  }
#else
  if (std::isinf(ar)) {
    outV = makeScalar(ar);
    return true;
  }
#endif
  if (sv.hasExactSInt()) {
    outV = makeScalarInt(sv.exactSInt);
    return true;
  }
  if (sv.hasExactUInt()) {
    outV = makeScalarUInt(sv.exactUInt);
    return true;
  }
  SmFloat arScale0 = std::fabs(ar);
  if (arScale0 < 1.0) {
    arScale0 = 1.0;
  }
  if (std::fabs(ar) < RATIO_APPROX_EPS * arScale0) {
    outV = makeScalar(0.0);
    return true;
  }
  SmInt nearInt = 0;
  if (tryExtractExactSIntFromFloatStrict(ar, nearInt)) {
    outV = makeScalarInt(nearInt);
    return true;
  }
  if (std::fabs(ar) > kMaxExactIntFromFloat) {
    outV = makeScalar(ar);
    return true;
  }
  SmInt num = 0;
  SmUInt den = 1;
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
    if (!args.empty() && args[0].kind == ValueKind::Array) {
      return makeArrayFromScalars({});
    }
    ctx.failCall(CallErrorId::AtLeastN, fnName, 1);
    return makeScalar(0);
  }
  if (flat.size() == 1U && args[0].kind == ValueKind::Scalar) {
    EvalValue out = makeScalar(0);
    if (!tryBuiltinRatioScalar(sub, flat[0].scalarValue, out)) {
      if (sub.parseError) {
        ctx.fail(ParseErrorId::IncompatibleOperands);
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
        ctx.fail(ParseErrorId::IncompatibleOperands);
      }
      return makeScalar(0);
    }
    wrapped.push_back(scalarFromScalarValue(tmp.scalarValue));
  }
  return makeArrayFromScalars(wrapped);
}

MathParser::EvalValue MathParser::invokeUnaryFunctionOnElement(
    EvalContext& ctx,
    const std::string& funcName,
    const EvalValue::ScalarValue& elem,
    const std::unordered_map<std::string, EvalValue>* scopedVars,
    const char* expectsUnaryFunction) {
  std::vector<EvalValue> args;
  args.push_back(scalarFromScalarValue(elem));
  BuiltinFunctionId fnId = BuiltinFunctionId::Count;
  if (!tryGetBuiltinFunctionId(toLower(funcName), fnId)) {
    return evalUserFunctionCall(ctx, funcName, args, scopedVars);
  }
  const BuiltinCategory cat = getBuiltinCategory(fnId);
  if (cat == BuiltinCategory::Sortby || cat == BuiltinCategory::Ratio || cat == BuiltinCategory::Filter) {
    ctx.fail(expectsUnaryFunction);
    return makeScalar(0);
  }
  EvalValue out = evalFunctionCall(ctx, funcName, std::move(args), fnId, scopedVars);
  if (ctx.parseError) {
    return out;
  }
  if (out.kind == ValueKind::Scalar || out.kind == ValueKind::Array) {
    return out;
  }
  ctx.fail(expectsUnaryFunction);
  return makeScalar(0);
}

bool MathParser::prepareUnaryCallableScalars(
    EvalContext& ctx,
    const std::string& fnName,
    const EvalValue& dataArg,
    const EvalValue& fnArg,
    const char* expectsOneFunction,
    const char* expectsUnaryFunction,
    std::vector<EvalValue::ScalarValue>& vals,
    bool& isLambda,
    bool& emptyArray) const {
  emptyArray = false;
  isLambda = (fnArg.kind == ValueKind::InlineLambda);
  const bool isRef = (fnArg.kind == ValueKind::FunctionRef);
  if (!isLambda && !isRef) {
    if (fnArg.kind == ValueKind::Array) {
      ctx.fail(expectsOneFunction);
    } else {
      ctx.fail(expectsUnaryFunction);
    }
    return false;
  }
  if (isRef) {
    std::string errText;
    if (!isEligibleUnaryFunctionName(fnArg.funcRefName, errText, expectsUnaryFunction)) {
      ctx.fail(errText.c_str());
      return false;
    }
  }
  std::vector<EvalValue> flat;
  if (!flattenArgsToScalars({dataArg}, flat)) {
    if (dataArg.kind == ValueKind::Array) {
      emptyArray = true;
      return false;
    }
    ctx.failCall(CallErrorId::AtLeastN, fnName, 1);
    return false;
  }
  vals.clear();
  vals.reserve(flat.size());
  for (const auto& item : flat) {
    vals.push_back(item.scalarValue);
  }
  return true;
}

MathParser::EvalValue MathParser::evalUnaryCallableOnScalar(
    EvalContext& ctx,
    const EvalValue& fnArg,
    bool isLambda,
    const EvalValue::ScalarValue& elem,
    const std::unordered_map<std::string, EvalValue>* scopedVars,
    const char* expectsUnaryFunction) {
#if SMARTMATH_LAMBDA_FUNCTIONS
  if (isLambda) {
    std::vector<EvalValue> lamArgs;
    lamArgs.push_back(scalarFromScalarValue(elem));
    return evalInlineLambdaCall(
        ctx, fnArg.lambdaParams, fnArg.lambdaBody, std::move(lamArgs), scopedVars);
  }
#else
  (void)isLambda;
#endif
  return invokeUnaryFunctionOnElement(ctx, fnArg.funcRefName, elem, scopedVars, expectsUnaryFunction);
}

MathParser::EvalValue MathParser::builtinSortby(
    EvalContext& ctx,
    const std::string& fnName,
    const std::vector<EvalValue>& args,
    const std::unordered_map<std::string, EvalValue>* scopedVars) {
  bool isLambda = false;
  bool emptyArray = false;
  std::vector<EvalValue::ScalarValue> vals;
  if (!prepareUnaryCallableScalars(
          ctx, fnName, args[0], args[1],
          STR_SORTBY_EXPECTS_ONE_FUNCTION, STR_SORTBY_EXPECTS_UNARY_FUNCTION,
          vals, isLambda, emptyArray)) {
    if (emptyArray) {
      return makeArrayFromScalars({});
    }
    return makeScalar(0);
  }
  std::vector<EvalValue> keys;
  keys.reserve(vals.size());
  for (std::size_t i = 0; i < vals.size(); ++i) {
    const EvalValue keyV = evalUnaryCallableOnScalar(
        ctx, args[1], isLambda, vals[i], scopedVars, STR_SORTBY_EXPECTS_UNARY_FUNCTION);
    if (ctx.parseError) {
      return makeScalar(0);
    }
    if (keyV.kind == ValueKind::FunctionRef || keyV.kind == ValueKind::InlineLambda ||
        (keyV.kind != ValueKind::Scalar && keyV.kind != ValueKind::Array)) {
      ctx.fail(STR_SORTBY_KEY_MUST_BE_SCALAR_OR_ARRAY);
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

std::unique_ptr<MathParser::Expr> MathParser::parseUnaryFunctionRef(
    EvalContext& ctx,
    const char* expectsOneFunction,
    const char* expectsUnaryFunction) {
  skipSpaces(ctx.p);
  if (*ctx.p == '(') {
    auto parsed = parseExpression(ctx);
    if (ctx.parseError || !parsed) {
      return nullptr;
    }
    if (parsed->tag == Expr::Tag::ArrayOrParens) {
      ctx.fail(expectsOneFunction);
      return nullptr;
    }
    ctx.fail(expectsUnaryFunction);
    return nullptr;
  }
  if (!localIsIdentStart(*ctx.p)) {
    ctx.fail(expectsUnaryFunction);
    return nullptr;
  }
  const char* start = ctx.p;
  ++ctx.p;
  while (localIsIdentChar(*ctx.p)) {
    ++ctx.p;
  }
  const std::string nameText(start, static_cast<std::size_t>(ctx.p - start));
  skipSpaces(ctx.p);
  if (*ctx.p == ':') {
    ctx.fail(ParseErrorId::UnexpectedToken);
    return nullptr;
  }
  if (*ctx.p == '(') {
    ctx.fail(expectsOneFunction);
    return nullptr;
  }
  auto ref = std::make_unique<Expr>();
  ref->tag = Expr::Tag::FunctionRef;
  ref->name = toLower(nameText);
  return ref;
}

bool MathParser::parseUnaryFunctionCallArguments(
    EvalContext& ctx,
    std::vector<std::unique_ptr<Expr>>& outArgs,
    const char* expectsOneFunction,
    const char* expectsUnaryFunction) {
  skipSpaces(ctx.p);
  if (*ctx.p == ',') {
    ctx.fail(ParseErrorId::UnexpectedComma);
    return false;
  }
  if (*ctx.p == ')') {
    ctx.fail(expectsUnaryFunction);
    return false;
  }
  auto first = parseExpression(ctx);
  if (ctx.parseError || !first) {
    return false;
  }
  outArgs.push_back(std::move(first));
  skipSpaces(ctx.p);
  bool hasComma = false;
  if (!tryConsumeCommaArgSeparator(ctx, hasComma)) {
    return false;
  }
  if (!hasComma) {
    ctx.fail(expectsUnaryFunction);
    return false;
  }
  auto keyArg = parseUnaryFunctionArg(ctx, expectsOneFunction, expectsUnaryFunction);
  if (ctx.parseError || !keyArg) {
    return false;
  }
  outArgs.push_back(std::move(keyArg));
  skipSpaces(ctx.p);
  if (!tryConsumeCommaArgSeparator(ctx, hasComma)) {
    return false;
  }
  if (hasComma) {
    ctx.fail(expectsOneFunction);
    return false;
  }
  return true;
}

MathParser::EvalValue MathParser::builtinFilter(
    EvalContext& ctx,
    const std::string& fnName,
    const std::vector<EvalValue>& args,
    const std::unordered_map<std::string, EvalValue>* scopedVars) {
  bool isLambda = false;
  bool emptyArray = false;
  std::vector<EvalValue::ScalarValue> vals;
  if (!prepareUnaryCallableScalars(
          ctx, fnName, args[0], args[1],
          STR_FILTER_EXPECTS_ONE_FUNCTION, STR_FILTER_EXPECTS_UNARY_FUNCTION,
          vals, isLambda, emptyArray)) {
    if (emptyArray) {
      return makeArrayFromScalars({});
    }
    return makeScalar(0);
  }
  std::vector<EvalValue> kept;
  kept.reserve(vals.size());
  for (const auto& elem : vals) {
    const EvalValue predV = evalUnaryCallableOnScalar(
        ctx, args[1], isLambda, elem, scopedVars, STR_FILTER_EXPECTS_UNARY_FUNCTION);
    if (ctx.parseError) {
      return makeScalar(0);
    }
    if (predV.kind != ValueKind::Scalar && predV.kind != ValueKind::Array) {
      ctx.fail(STR_FILTER_PREDICATE_MUST_BE_SCALAR_OR_ARRAY);
      return makeScalar(0);
    }
    if (isTruthy(predV)) {
      kept.push_back(scalarFromScalarValue(elem));
    }
  }
  return makeArrayFromScalars(kept);
}

MathParser::EvalValue MathParser::builtinSortFamily(
    EvalContext& ctx,
    const std::string& fnName,
    BuiltinFunctionId id,
    const std::vector<EvalValue>& args) const {
  if (id == BuiltinFunctionId::Flat && args.empty()) {
    return makeArrayFromScalars({});
  }
  auto dedupUniqueInPlace = [&](std::vector<EvalValue>& values) {
    bool seenNan = false;
    auto isNanBearing = [](const EvalValue& ev) {
      SmFloat re = 0.0;
      SmFloat im = 0.0;
      scalarLoadCartesian(ev.scalarValue, re, im);
      return std::isnan(re) || std::isnan(im);
    };
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
        if (isNanBearing(ev)) {
          if (!seenNan) {
            seenNan = true;
            out.push_back(ev);
          }
          continue;
        }
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
    std::unordered_set<SmFloat> seen;
    seen.reserve(values.size() * 2 + 1);
    std::size_t writePos = 0;
    for (std::size_t readPos = 0; readPos < values.size(); ++readPos) {
      const SmFloat v = values[readPos].scalarValue.scalar;
      if (isNanBearing(values[readPos])) {
        // unique keeps one NaN. nan != nan remains false.
        if (!seenNan) {
          seenNan = true;
          values[writePos++] = values[readPos];
        }
        continue;
      }
      // Canonicalize signed zero to preserve equality semantics of (0 == -0).
      const SmFloat key = (v == 0.0) ? 0.0 : v;
      if (seen.insert(key).second) {
        values[writePos++] = values[readPos];
      }
    }
    values.resize(writePos);
  };

#if SMARTMATH_COMPLEX_NUMBERS
  if (getSupportComplexNumbers() && id == BuiltinFunctionId::Sort &&
      std::any_of(args.begin(), args.end(), [](const EvalValue& v) { return evalValueHasNonzeroImaginary(v); })) {
    ctx.fail(ParseErrorId::IncompatibleOperands);
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
        ctx.fail(constructTimeArrayMixedError(fnName));
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
      if (a.kind == ValueKind::Array) {
        return makeArrayFromScalars({});
      }
      ctx.failCall(CallErrorId::AtLeastN, fnName, 1);
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
    if (id == BuiltinFunctionId::Flat) {
      return makeArrayFromScalars(out);
    }
  }

  std::vector<EvalValue> flat;
  if (!flattenArgsToScalars(args, flat)) {
    if (!args.empty()) {
      return makeArrayFromScalars({});
    }
    ctx.failCall(CallErrorId::AtLeastN, fnName, 1);
    return makeScalar(0);
  }
  if (id == BuiltinFunctionId::Sort) {
    sortScalarsInPlace(flat);
  } else if (id == BuiltinFunctionId::Reverse) {
    std::reverse(flat.begin(), flat.end());
  } else if (id == BuiltinFunctionId::Unique) {
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
    ctx.fail(ParseErrorId::IncompatibleOperands);
    return makeScalar(0);
  }
#if SMARTMATH_TIME_VALUES
  const EvalValue& a0 = args[0];
  if (a0.kind == ValueKind::Scalar) {
    if (!scalarValueIsTime(a0.scalarValue)) {
      ctx.fail(fnName + STR_TIME_EXPECTS_TIME_ARG);
      return makeScalar(0);
    }
    return evalValueFromTimeMs(id, timeTotalMsFromScalarValue(a0.scalarValue));
  }
  if (a0.kind == ValueKind::Array) {
    return mapTimeUnitOverArray(ctx, fnName, id, a0);
  }
  ctx.fail(fnName + STR_TIME_EXPECTS_TIME_ARG);
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
      out = makeArrayFromScalars({});
    } else {
      out = makeArrayFromScalars(elems);
    }
  }
  applyBuiltinFormatRenderMeta(id, out);
  clearBaseFormatIntPowerDisplay(out);
  return out;
}

MathParser::EvalValue MathParser::builtinPow(EvalContext& ctx, const std::vector<EvalValue>& args) const {
  const std::string fnName = getFunctionName(BuiltinFunctionId::Pow);
#if SMARTMATH_TIME_VALUES
  if (evalValueInvolvesTime(args[0]) || evalValueInvolvesTime(args[1])) {
    ctx.fail(ParseErrorId::IncompatibleOperands);
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
      ctx.failCall(CallErrorId::ScalarMinMax, fnName);
      return makeScalar(0);
    }
    return builtinApplyClamp(ctx, args[0], args[1], args[2], getSupportComplexNumbers(), getSupportTimeValues());
  }

  if (id == BuiltinFunctionId::Gcd || id == BuiltinFunctionId::Lcm || id == BuiltinFunctionId::Ncr ||
      id == BuiltinFunctionId::Npr) {
#if SMARTMATH_COMPLEX_NUMBERS
    if (rejectBuiltinArgsWithComplexImaginary(ctx, args)) {
      return makeScalar(0);
    }
#endif
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
    ctx.failCall(CallErrorId::ScalarValues, fnName);
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
  ctx.fail(ParseErrorId::InternalScalarBinaryBuiltin);
  return makeScalar(0);
}

MathParser::EvalValue MathParser::builtinRand(EvalContext& /*ctx*/, const std::vector<EvalValue>& /*args*/) const {
  return makeScalar(randomUnitScalar());
}

MathParser::EvalValue MathParser::builtinModCall(EvalContext& ctx, const std::vector<EvalValue>& args) const {
  const std::string fnName = getFunctionName(BuiltinFunctionId::Mod);
#if SMARTMATH_COMPLEX_NUMBERS
  if (rejectBuiltinArgsWithComplexImaginary(ctx, args)) {
    return makeScalar(0);
  }
#endif
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

bool tryApplyFactorialRealNonNegative(const EvalValue::ScalarValue& sv, EvalValue& out) {
  SmInt n = 0;
  if (!nearlyInt(sv.scalar, n) || n < 0) {
    return false;
  }
  if (n <= kMaxFactorialArg) {
    out = makeScalarInt(factorialInt(static_cast<int>(n)));
    return true;
  }
  out = makeScalar(factorialScalarFloatFromInt(n));
  return true;
}

MathParser::EvalValue MathParser::builtinFactorial(
    EvalContext& ctx,
    const std::string& fnName,
    const std::vector<EvalValue>& args) const {
#if SMARTMATH_COMPLEX_NUMBERS
  if (getSupportComplexNumbers()) {
    const auto isNegativeIntegerFact = [](const EvalValue::ScalarValue& s) {
      if (scalarHasNonzeroImaginaryPart(s)) {
        return false;
      }
      if (s.hasExactSInt() && s.exactSInt < 0) {
        return true;
      }
      SmInt n = 0;
      return nearlyInt(s.scalar, n) && n < 0;
    };
    bool negativeInteger = false;
    if (args[0].kind == ValueKind::Array) {
      for (const auto& elem : args[0].arr) {
        if (isNegativeIntegerFact(elem)) {
          negativeInteger = true;
          break;
        }
      }
    } else if (args[0].kind == ValueKind::Scalar) {
      negativeInteger = isNegativeIntegerFact(args[0].scalarValue);
    }
    if (negativeInteger) {
      ctx.failCall(CallErrorId::NonNegativeInteger, fnName);
      return makeScalar(0);
    }
    return builtinUnaryMath(ctx, fnName, BuiltinFunctionId::Fact, args);
  }
#endif
  if (args[0].kind == ValueKind::Array) {
    if (args[0].arr.empty()) {
      return makeArrayFromScalars({});
    }
    std::vector<EvalValue> outVals;
    outVals.reserve(args[0].arr.size());
    for (const auto& elem : args[0].arr) {
      EvalValue one;
      if (!tryApplyFactorialRealNonNegative(elem, one)) {
        ctx.failCall(CallErrorId::NonNegativeInteger, fnName);
        return makeScalar(0);
      }
      outVals.push_back(std::move(one));
    }
    return makeArrayFromScalars(std::move(outVals));
  }
  EvalValue out;
  if (args[0].kind != ValueKind::Scalar || !tryApplyFactorialRealNonNegative(args[0].scalarValue, out)) {
    ctx.failCall(CallErrorId::NonNegativeInteger, fnName);
    return makeScalar(0);
  }
  return out;
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
        ctx.fail(ParseErrorId::IncompatibleOperands);
        return makeScalar(0);
      }
    }
  }
#endif
  const bool toDeg = (id == BuiltinFunctionId::Deg);
  if (args.size() == 1) {
    return mapUnaryEvalValue(args[0], [&](const EvalValue::ScalarValue& s) -> EvalValue {
      const SmFloat x = s.scalar;
      return makeScalarMaybeExact(toDeg ? (x * 180.0 / kPi) : (x * kPi / 180.0));
    });
  }
  std::vector<EvalValue> inVals;
  if (!flattenArgsToScalars(args, inVals)) {
    return makeArrayFromScalars({});
  }
  std::vector<EvalValue> outVals;
  outVals.reserve(inVals.size());
  for (const auto& item : inVals) {
    const SmFloat x = item.scalarValue.scalar;
    outVals.emplace_back(makeScalarMaybeExact(toDeg ? (x * 180.0 / kPi) : (x * kPi / 180.0)));
  }
  return makeArrayFromScalars(outVals);
}

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
  SmFloat ar = 0.0;
  SmFloat ai = 0.0;
  scalarLoadCartesian(sv, ar, ai);
#if SMARTMATH_COMPLEX_NUMBERS
  if (!getSupportComplexNumbers() && scalarHasNonzeroImaginaryPart(sv)) {
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
        const SmInt num = sv.imagExactSInt;
        const SmUInt den = sv.imagExactUInt;
        if (den == 0u) {
          out = makeScalar(0.0);
          scalarClearImaginary(out.scalarValue);
          return true;
        }
        if (den == 1u) {
          ExactCartesianComponent cIm{};
          if (num < 0) {
            ExactCartesianComponentAssignFromSignedInt(cIm, num);
          } else {
            ExactCartesianComponentAssignFromSignedInt(cIm, num);
          }
          setScalarFromExactCartesianComponent(out, cIm);
          return true;
        }
        const SmFloat imOut = static_cast<SmFloat>(num) / static_cast<SmFloat>(den);
        out = makeScalar(imOut);
        scalarClearImaginary(out.scalarValue);
        return true;
      }

      // Prefer correct "exact imag" when it's nonzero; otherwise, use the
      // floating component loaded from the scalar (used by complex formatting).
      if (sv.hasImagExactSInt() && sv.imagExactSInt != 0LL) {
        ExactCartesianComponent cIm{};
        ExactCartesianComponentAssignFromSignedInt(cIm, sv.imagExactSInt);
        setScalarFromExactCartesianComponent(out, cIm);
        return true;
      }
      if (sv.hasImagExactUInt() && sv.imagExactUInt != 0u) {
        ExactCartesianComponent cIm{};
        ExactCartesianComponentAssignFromUInt(cIm, sv.imagExactUInt);
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
          out = makeScalarComplexFromFloats(ar, ai);
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
            out = makeScalarComplexFromFloats(ar, -ai);
          } else if (scalarHasNonzeroImaginaryPart(sv)) {
            setScalarComplexFromExactCartesian(out, cRe, cImConj);
          } else {
            setScalarComplexFromExactCartesian(out, cRe, cIm);
          }
        } else {
          out = makeScalarComplexFromFloats(ar, -ai);
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
      EvalValue::ScalarValue svFracRe;
      EvalValue::ScalarValue svFracIm;
      scalarSplitCartesianParts(sv, ar, ai, svFracRe, svFracIm);
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
        out = makeScalar(std::numeric_limits<SmFloat>::quiet_NaN());
      } else {
        out = makeScalarMaybeExact(calcHypotBasic(ar, ai));
      }
      return true;
    case BuiltinFunctionId::Sign: {
      const SmFloat mag = calcHypotBasic(ar, ai);
      if (mag == 0.0 || !std::isfinite(mag) || !std::isfinite(ar) || !std::isfinite(ai)) {
        out = makeScalarComplexFromFloats(0.0, 0.0);
      } else {
        out = makeScalarComplexFromFloats(ar / mag, ai / mag);
      }
      return true;
    }
    case BuiltinFunctionId::Fact: {
      if (ai == 0.0 && !scalarHasNonzeroImaginaryPart(sv)) {
        SmInt n = 0;
        const bool exactNeg = sv.hasExactSInt() && sv.exactSInt < 0;
        if (exactNeg || (nearlyInt(sv.scalar, n) && n < 0)) {
          return false;
        }
        if (sv.hasExactSInt()) {
          n = sv.exactSInt;
        }
        if ((sv.hasExactSInt() || nearlyInt(sv.scalar, n)) && n >= 0) {
          if (n <= kMaxFactorialArg) {
            out = makeScalarInt(factorialInt(static_cast<int>(n)));
            return true;
          }
          out = makeScalar(factorialScalarFloatFromInt(n));
          return true;
        }
      }
      SmFloat gr = 0.0;
      SmFloat gi = 0.0;
      complexGamma(ar + 1.0, ai, gr, gi);
      out = makeScalarComplexFromFloats(gr, gi);
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
  const auto cartFromPolarScalars = [&](SmFloat rMag, SmFloat rAng) -> EvalValue {
#if SMARTMATH_COMPLEX_NUMBERS
    if (!getSupportComplexNumbers() && std::fabs(rAng) > kTolLost1) {
      ctx.fail(ParseErrorId::IncompatibleOperands);
      return makeScalar(0);
    }
    return makeScalarComplexFromFloats(rMag * calcCos(rAng), rMag * calcSin(rAng));
#else
    if (std::fabs(rAng) > kTolLost1) {
      ctx.fail(ParseErrorId::IncompatibleOperands);
      return makeScalar(0);
    }
    return makeScalarMaybeExact(rMag);
#endif
  };

  if (id == BuiltinFunctionId::Polar) {
    EvalValue polarIn = makeScalar(0);
    if (!tryNormalizeUnaryScalarInput(args[0], polarIn)) {
      ctx.fail(ParseErrorId::IncompatibleOperands);
      return makeScalar(0);
    }
    EvalValue outS = makeScalar(0);
    if (!tryUnaryComplexBuiltinSupport(BuiltinFunctionId::Polar, polarIn.scalarValue, outS)) {
      ctx.failCall(CallErrorId::Numeric, fnName);
      return makeScalar(0);
    }
    return outS;
  }

  if (args.size() == 1U) {
    const EvalValue& cartIn = args[0];
    if (cartIn.kind == ValueKind::Scalar) {
      EvalValue outS = makeScalar(0);
      if (!tryUnaryComplexBuiltinSupport(BuiltinFunctionId::Cart, cartIn.scalarValue, outS)) {
        ctx.failCall(CallErrorId::Numeric, fnName);
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
    ctx.fail(ParseErrorId::IncompatibleOperands);
    return makeScalar(0);
  }
  if (args.size() == 2U) {
    if (args[0].kind != ValueKind::Scalar || args[1].kind != ValueKind::Scalar) {
      ctx.fail(ParseErrorId::IncompatibleOperands);
      return makeScalar(0);
    }
#if SMARTMATH_COMPLEX_NUMBERS
    if (scalarHasNonzeroImaginaryPart(args[0].scalarValue) ||
        scalarHasNonzeroImaginaryPart(args[1].scalarValue)) {
      ctx.fail(ParseErrorId::IncompatibleOperands);
      return makeScalar(0);
    }
#endif
    return cartFromPolarScalars(args[0].scalarValue.scalar, args[1].scalarValue.scalar);
  }
  ctx.fail(ParseErrorId::IncompatibleOperands);
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
  const SmFloat nanv = std::numeric_limits<SmFloat>::quiet_NaN();
  const SmFloat x = sv.scalar;
  if (scalarHasNonzeroImaginaryPart(sv)) {
    SmFloat arCx = 0.0;
    SmFloat aiCx = 0.0;
    scalarLoadCartesian(sv, arCx, aiCx);
    if (id == BuiltinFunctionId::Sqrt || id == BuiltinFunctionId::Sqr) {
      if (!std::isfinite(arCx) || !std::isfinite(aiCx)) {
        outV = makeScalarComplexFromFloats(nanv, nanv);
        return true;
      }
      if (id == BuiltinFunctionId::Sqr) {
        outV = makeScalarComplexFromFloats(arCx * arCx - aiCx * aiCx, 2.0 * arCx * aiCx);
        return true;
      }
      outV = applyUnarySqrtEval(sv);
      return true;
    }
    if (id == BuiltinFunctionId::Exp) {
      if (!std::isfinite(arCx) || !std::isfinite(aiCx)) {
        outV = makeScalarComplexFromFloats(nanv, nanv);
        return true;
      }
      if (!isTrigRadiansInRange(aiCx)) {
        trigOutOfRange = true;
        outV = makeScalar(0);
        return true;
      }
      const SmFloat ea = std::exp(arCx);
      SmFloat pr = ea * calcCos(aiCx);
      SmFloat pi = ea * calcSin(aiCx);
      snapComplexNearZeroAxis(pr, pi);
      outV = makeScalarComplexFromFloats(pr, pi);
      return true;
    }
    if (id == BuiltinFunctionId::Ln || id == BuiltinFunctionId::Log10) {
      SmFloat loR = 0.0;
      SmFloat loI = 0.0;
      scalarPrincipalLnCartesian(arCx, aiCx, loR, loI);
      if (id == BuiltinFunctionId::Ln) {
        outV = makeScalarComplexFromFloats(loR, loI);
      } else {
        outV = makeScalarComplexFromFloats(loR * kInvLn10, loI * kInvLn10);
      }
      return true;
    }
    if (isComplexUnaryTrigBuiltin(id)) {
      SmFloat or_ = 0.0;
      SmFloat oi = 0.0;
      if (!complexUnaryTrigCartesian(id, arCx, aiCx, or_, oi)) {
        trigOutOfRange = true;
        outV = makeScalar(0);
        return true;
      }
      snapComplexNearZeroAxis(or_, oi);
      outV = makeScalarComplexFromFloats(or_, oi);
      return true;
    }
    realUnaryOk = false;
    outV = makeScalar(0);
    return true;
  }
  if ((id == BuiltinFunctionId::Ln || id == BuiltinFunctionId::Log10) && std::isfinite(x)) {
    if (x < 0.0) {
      if (id == BuiltinFunctionId::Ln) {
        SmFloat loR = std::log(-x);
        SmFloat loI = kPi;
        snapComplexNearZeroAxis(loR, loI);
        outV = makeScalarComplexFromFloats(loR, loI);
      } else {
        SmFloat loR = std::log10(-x);
        SmFloat loI = kPi * kInvLn10;
        snapComplexNearZeroAxis(loR, loI);
        outV = makeScalarComplexFromFloats(loR, loI);
      }
      return true;
    }
    if (x == 0.0) {
      SmFloat loR = 0.0;
      SmFloat loI = 0.0;
      scalarPrincipalLnCartesian(0.0, 0.0, loR, loI);
      if (id == BuiltinFunctionId::Ln) {
        outV = makeScalarComplexFromFloats(loR, loI);
      } else {
        outV = makeScalarComplexFromFloats(loR * kInvLn10, loI * kInvLn10);
      }
      return true;
    }
  }
  if (std::isfinite(x)) {
    const bool outOfRealDomain =
        ((id == BuiltinFunctionId::Asin || id == BuiltinFunctionId::Acos || id == BuiltinFunctionId::Atanh) &&
         std::fabs(x) > 1.0) ||
        (id == BuiltinFunctionId::Acosh && x < 1.0);
    if (outOfRealDomain) {
      SmFloat or_ = 0.0;
      SmFloat oi = 0.0;
      if (!complexUnaryTrigCartesian(id, x, 0.0, or_, oi)) {
        return false;
      }
      snapComplexNearZeroAxis(or_, oi);
      outV = makeScalarComplexFromFloats(or_, oi);
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

  const SmFloat x = sv.scalar;
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
        if (std::fabs(x) > kHypotDirectMax && std::isfinite(x)) {
          return makeScalar(std::copysign(std::log(std::fabs(x)) + std::log(2.0), x));
        }
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
      if (sv.hasExactSInt()) {
        return makeScalarInt((sv.exactSInt > 0) ? 1LL : ((sv.exactSInt < 0) ? -1LL : 0LL));
      }
      if (sv.hasExactUInt()) {
        return makeScalarInt((sv.exactUInt == 0u) ? 0LL : 1LL);
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
    ctx.fail(ParseErrorId::IncompatibleOperands);
    return makeScalar(0);
  }
#endif
  bool realUnaryOk = true;
  bool trigOutOfRange = false;
  const EvalValue result = mapUnaryEvalValue(args[0], [&](const EvalValue::ScalarValue& s) -> EvalValue {
    return applyUnaryBuiltin(id, s, trigOutOfRange, realUnaryOk);
  });
  if (trigOutOfRange) {
    ctx.failCall(CallErrorId::Numeric, fnName);
    return makeScalar(0);
  }
  if (!realUnaryOk) {
    ctx.fail(ParseErrorId::InternalUnaryMathBuiltin);
    return makeScalar(0);
  }
  return result;
}

MathParser::EvalValue MathParser::builtinLog(EvalContext& ctx, const std::vector<EvalValue>& args) const {
  return builtinMapBinaryTwoArg(
      ctx, getFunctionName(BuiltinFunctionId::Log), BuiltinFunctionId::Log, args, false, true);
}

MathParser::EvalValue MathParser::runBoundProgram(
    EvalContext& ctx,
    const std::vector<AstStatement>& program,
    const std::unordered_map<std::string, EvalValue>* localVars,
    bool suppressBareFunctionTailHint) {
  EvalContext sub;
  sub.evalDepth = ctx.evalDepth + 1;
  sub.userFunctionCallStack = ctx.userFunctionCallStack;
  sub.suppressBareFunctionTailHint = suppressBareFunctionTailHint;
  EvalValue v = runCompiledProgram(sub, program, localVars, false);
  if (!finishSubEvalFromCompiledProgram(ctx, sub)) {
    return makeScalar(0);
  }
  return v;
}

MathParser::EvalValue MathParser::evalUserFunctionCall(
    EvalContext& ctx,
    const std::string& fnName,
    const std::vector<EvalValue>& args,
    const std::unordered_map<std::string, EvalValue>* /*scopedVars*/) {
  UserFunction* uf = findUserFunction(fnName);
  if (!uf) {
    appendUniqueName(ctx.unknownFuncsText, fnName);
    ctx.referencedUnresolvedName = true;
    return makeScalar(0);
  }
  if (args.size() != uf->params.size()) {
    ctx.failCall(CallErrorId::ExactCount, fnName, uf->params.size());
    return makeScalar(0);
  }
  if (ctx.evalDepth >= kMaxEvalDepth) {
    ctx.fail(ParseErrorId::MaxEvaluationDepthReached);
    return makeScalar(0);
  }

  std::vector<std::string> localUserFunctionCallStack;
  std::vector<std::string>* const callStack =
      ctx.userFunctionCallStack ? ctx.userFunctionCallStack : &localUserFunctionCallStack;

  for (const std::string& active : *callStack) {
    if (active == fnName) {
      ctx.failRecursiveUserFunction(fnName);
      return makeScalar(0);
    }
  }
  if (callStack->size() >= static_cast<std::size_t>(kMaxEvalDepth)) {
    ctx.fail(ParseErrorId::UserFunctionCallStackOverflow);
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
    if (!tryCompileExprToProgram(uf->expr, ctx, callStack, fnName.c_str(), false, &uf->params, compiledBody)) {
      return makeScalar(0);
    }
    uf->compiledProgram = std::move(compiledBody);
    uf->compiledProgramReady = true;
  }

  return runBoundProgram(ctx, uf->compiledProgram, &localVars, false);
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
    case BuiltinCategory::Filter:
      return builtinFilter(ctx, fnName, args, scopedVars);
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
#if SMARTMATH_FACTORINT
    case BuiltinCategory::Factorint:
      return builtinFactorint(ctx, fnName, args);
#endif
    case BuiltinCategory::DegRad:
      return builtinDegRad(ctx, fnName, id, args);
    case BuiltinCategory::Log:
      return builtinLog(ctx, args);
    case BuiltinCategory::UnaryMath:
      return builtinUnaryMath(ctx, fnName, id, args);
    case BuiltinCategory::PolarCart:
      return builtinPolarCart(ctx, fnName, id, args);
    case BuiltinCategory::Iif:
    case BuiltinCategory::IsInstance:
      ctx.fail(ParseErrorId::InternalEvalError);
      return makeScalar(0);
    case BuiltinCategory::Unknown:
      break;
  }
  appendUniqueName(ctx.unknownFuncsText, fnName);
  ctx.referencedUnresolvedName = true;
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
      ctx.fail(ParseErrorId::ParseFailed);
    }
    lastError_ = ctx.errorText;
    return false;
  }
  skipSpaces(ctx.p);
  if (!ctx.parseError && *ctx.p != '\0') {
    ctx.fail(ParseErrorId::UnexpectedInput);
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
      [](const AstStatement& st) {
        return st.kind == AstStatement::Kind::Assign || st.kind == AstStatement::Kind::IndexAssign;
      });
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

  skipSpaces(ctx.p);
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
