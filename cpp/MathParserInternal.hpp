#pragma once

#include <cmath>
#include <cstdint>
#include <limits>
#include <string>
#include <memory>
#include <unordered_map>
#include <vector>

#include "MathParserTypes.hpp"

#ifndef SMARTMATH_COMPLEX_NUMBERS
#define SMARTMATH_COMPLEX_NUMBERS 1
#endif

#ifndef SMARTMATH_TIME_VALUES
#define SMARTMATH_TIME_VALUES 1
#endif

#ifndef SMARTMATH_FACTORINT
#define SMARTMATH_FACTORINT 1
#endif

#ifndef SMARTMATH_LAMBDA_FUNCTIONS
#define SMARTMATH_LAMBDA_FUNCTIONS 1
#endif

namespace smartmath_internal {

  using smartmath::SmFloat;
  using smartmath::SmInt;
  using smartmath::SmUInt;
  using smartmath::kIntMax;
  using smartmath::kIntMin;
  using smartmath::kIntMinMag;
  using smartmath::kUIntMax;

  enum class BuiltinFunctionId {
    Rand = 0,
    Random,
    Bin,
    Hex,
    Oct,
    Pow,
    Atan2,
    Sin,
    Cos,
    Tan,
    Asin,
    Acos,
    Atan,
    Sinh,
    Cosh,
    Tanh,
    Acosh,
    Asinh,
    Atanh,
    Exp,
    Log,
    Ln,
    Log10,
    Sqrt,
    Sqr,
    Int,
    Frac,
    Abs,
    Floor,
    Ceil,
    Trunc,
    Round,
    Sign,
    Deg,
    Rad,
    Sum,
    Median,
    Variance,
    Stddev,
    Sort,
    Sortby,
    Ratio,
    Range,
    Reverse,
    Unique,
    Unpack,
    Flat,
    Repeat,
    Len,
    Fact,
#if SMARTMATH_FACTORINT
    Factorint,
#endif
    Avg,
    Mean,
    Mod,
    Clamp,
    Hypot,
    Gcd,
    Lcm,
    Ncr,
    Npr,
    Product,
    Min,
    Max,
    Uhex,
    Uoct,
    Ubin,
    Milliseconds,
    Seconds,
    Minutes,
    Hours,
    Days,
    Real,
    Imag,
    Phase,
    Polar,
    Cart,
    Conj,
    Filter,
    Iif,
    IsInstance,
    Dec,
    Udec,
    Count
  };

  enum class ValueKind {
    Scalar,
    Array,
    FunctionRef,
    InlineLambda,
    DefinedFunction
  };

  enum class ScalarKind {
    FloatingPoint,
    SInt,
    UInt,
    Time
  };

  enum class RenderBase {
    PlainDec = 1, // dec/udec digits; Dec keeps rational and power display
    Dec = 10,
    Hex = 16,
    Oct = 8,
    Bin = 2
  };

  struct EvalValue {
    struct ScalarValue {
      enum : unsigned int {
        fExactSIntValid = 0x01u,
        fExactUIntValid = 0x02u,
        fDecScientificPow63High = 0x10u,
        fImagExactSIntValid = 0x20u,
        fImagExactUIntValid = 0x40u,
        fRenderRational = 0x80u,
        fImagRenderRational = 0x100u
      #if SMARTMATH_FACTORINT
        ,
        fRenderIntPower = 0x200u
      #endif
      };

      ScalarKind scalarKind = ScalarKind::FloatingPoint;
      /// Do not read for exact-int policy; use hasExactSInt() / hasRenderRational().
      unsigned int flags = 0;
      SmFloat scalar = 0.0;
      /// Numerator for render-rational and exact integers; prefer hasExactSInt() in eval paths.
      SmInt exactSInt = 0;
      /// Denominator for render-rational and exact uintegers; prefer hasExactUInt() in eval paths.
      SmUInt exactUInt = 0;
      SmFloat imag = 0.0;
      SmInt imagExactSInt = 0;
      SmUInt imagExactUInt = 0;

      bool hasExactSIntMetadata() const { return (flags & fExactSIntValid) != 0; }
      bool hasExactUIntMetadata() const { return (flags & fExactUIntValid) != 0; }
      bool hasImagExactSIntMetadata() const { return (flags & fImagExactSIntValid) != 0; }
      bool hasImagExactUIntMetadata() const { return (flags & fImagExactUIntValid) != 0; }
      /// Exact SmInt for arithmetic (not render-rational numerator storage).
      bool hasExactSInt() const { return hasExactSIntMetadata() && !hasRenderRational(); }
      void setExactSIntValid(bool v) { flags = v ? (flags | fExactSIntValid) : (flags & ~fExactSIntValid); }
      /// Exact SmUInt for arithmetic (not render-rational denominator storage).
      bool hasExactUInt() const { return hasExactUIntMetadata() && !hasRenderRational(); }
      void setExactUIntValid(bool v) { flags = v ? (flags | fExactUIntValid) : (flags & ~fExactUIntValid); }
      bool hasDecScientificPow63High() const { return (flags & fDecScientificPow63High) != 0; }
      void setDecScientificPow63High(bool v) {
        flags = v ? (flags | fDecScientificPow63High) : (flags & ~fDecScientificPow63High);
      }
      bool hasImagExactSInt() const { return hasImagExactSIntMetadata() && !hasImagRenderRational(); }
      void setImagExactSIntValid(bool v) { flags = v ? (flags | fImagExactSIntValid) : (flags & ~fImagExactSIntValid); }
      bool hasImagExactUInt() const { return hasImagExactUIntMetadata() && !hasImagRenderRational(); }
      void setImagExactUIntValid(bool v) {
        flags = v ? (flags | fImagExactUIntValid) : (flags & ~fImagExactUIntValid);
      }
      bool hasRenderRational() const { return (flags & fRenderRational) != 0; }
      void setRenderRational(bool v) { flags = v ? (flags | fRenderRational) : (flags & ~fRenderRational); }
      bool hasImagRenderRational() const { return (flags & fImagRenderRational) != 0; }
      void setImagRenderRational(bool v) {
        flags = v ? (flags | fImagRenderRational) : (flags & ~fImagRenderRational);
      }
    #if SMARTMATH_FACTORINT
      bool hasRenderIntPower() const { return (flags & fRenderIntPower) != 0; }
      void setRenderIntPower(bool v) { flags = v ? (flags | fRenderIntPower) : (flags & ~fRenderIntPower); }
    #endif
    };

    enum : unsigned int {
      fExpandArgs = 0x01u,   // Used by unpack(...): marks values for call-argument expansion.
      fRenderUnsigned = 0x02u, // For hex/oct/bin output: render negative integers as unsigned bit pattern.
      fRenderBaseMask = 0x0000FF00u, // Bit mask for packed RenderBase storage in flags.
      fRenderBaseShift = 8u // Bit shift for packed RenderBase storage in flags.
    };

    ValueKind kind = ValueKind::Scalar;
    unsigned int flags = 0;
    ScalarValue scalarValue{};
    std::vector<ScalarValue> arr;
    /// FunctionRef stores the referenced name. DefinedFunction stores the user-facing text, such as "defined f(x)".
    std::string funcRefName;
    /** Only when ``kind`` is InlineLambda (anonymous lambdas inside ``sortby``). */
    std::vector<std::string> lambdaParams;
    std::string lambdaBody;
    /// When true with hex/oct/bin/dec output, format negative integers as unsigned (two's complement) bit pattern.
    bool hasExpandArgs() const { return (flags & fExpandArgs) != 0; }
    void setExpandArgs(bool v) { flags = v ? (flags | fExpandArgs) : (flags & ~fExpandArgs); }
    bool hasRenderUnsigned() const { return (flags & fRenderUnsigned) != 0; }
    void setRenderUnsigned(bool v) { flags = v ? (flags | fRenderUnsigned) : (flags & ~fRenderUnsigned); }
    RenderBase getRenderBase() const {
      const unsigned int raw = (flags & fRenderBaseMask) >> fRenderBaseShift;
      return (raw == 0u) ? RenderBase::Dec : static_cast<RenderBase>(raw);
    }
    void setRenderBase(RenderBase v) {
      flags = (flags & ~fRenderBaseMask) |
              ((static_cast<unsigned int>(v) & 0xFFu) << fRenderBaseShift);
    }
  };

  EvalValue makeScalar(SmFloat v);
  EvalValue makeScalarMaybeExact(SmFloat v);
  EvalValue makeScalarInt(SmInt v);
  EvalValue makeScalarUInt(SmUInt v);
  EvalValue makeScalarUIntPreferInt(SmUInt v);

#if SMARTMATH_TIME_VALUES
  EvalValue makeScalarTimeMs(SmInt totalMs);
#endif

#if SMARTMATH_COMPLEX_NUMBERS
  EvalValue makeImaginaryUnit();
  EvalValue makeScalarComplexFromFloats(SmFloat re, SmFloat im);
#endif

  EvalValue makeArrayFromScalars(const std::vector<EvalValue>& v);
  EvalValue makeInlineLambdaValue(std::vector<std::string> params, std::string body);

  void scalarClearImaginary(EvalValue::ScalarValue& s);

  enum class ExactSignedIntPolicy {
    AllowUIntBitReinterpret,
    NoUIntWrap,
    NoUIntWrapWithRepairAndStorageKind,
  };

  enum class ScalarExactKind : std::uint8_t {
    Signed,
    Unsigned,
  };

  struct ExactCartesianComponent {
    bool hasInt = false;
    SmInt intV = 0;
    bool hasUInt = false;
    SmUInt uintV = 0;
  };

  
  inline bool isWithinExactIntFromFloatRange(SmFloat v) {
    return std::isfinite(v) && std::fabs(v) <= smartmath::kMaxExactIntFromFloat;
  }

  inline bool tryExtractExactSIntFromFloatStrict(SmFloat v, SmInt& out) {
    if (!isWithinExactIntFromFloatRange(v)) {
      return false;
    }
    if (!smartmath::isFloatInSIntRange(v)) {
      return false;
    }
    const SmInt asInt = static_cast<SmInt>(v);
    if (v != static_cast<SmFloat>(asInt)) {
      return false;
    }
    out = asInt;
    return true;
  }

  inline SmUInt gcdUInt(SmUInt a, SmUInt b) {
    while (b != 0u) {
      const SmUInt t = a % b;
      a = b;
      b = t;
    }
    return a;
  }

  inline SmUInt uintMagnitudeOfSInt(SmInt v) {
    return v < 0 ? (SmUInt{0} - static_cast<SmUInt>(v)) : static_cast<SmUInt>(v);
  }

  // mag must be <= 2^63.
  inline SmInt sintFromNegatedMagnitude(SmUInt mag) {
    return mag == smartmath::kIntMinMag ? kIntMin : -static_cast<SmInt>(mag);
  }

  // CPython slice.indices(length). ``length`` is an array size and is non-negative.
  inline void normalizePythonSliceBounds(
      SmInt length,
      bool hasStart,
      SmInt startIn,
      bool hasStop,
      SmInt stopIn,
      SmInt stepIn,
      SmInt& outStart,
      SmInt& outStop) {
    const bool stepNeg = (stepIn < 0);
    SmInt startIdx;
    SmInt stopIdx;
    if (!hasStart) {
      startIdx = stepNeg ? (length - 1) : 0;
    } else {
      startIdx = startIn;
      if (startIdx < 0) {
        startIdx += length;
      }
      if (startIdx < 0) {
        startIdx = stepNeg ? -1 : 0;
      } else if (startIdx >= length) {
        startIdx = stepNeg ? (length - 1) : length;
      }
    }
    if (!hasStop) {
      stopIdx = stepNeg ? -1 : length;
    } else {
      stopIdx = stopIn;
      if (stopIdx < 0) {
        stopIdx += length;
      }
      if (stopIdx < 0) {
        stopIdx = stepNeg ? -1 : 0;
      } else if (stopIdx >= length) {
        stopIdx = stepNeg ? (length - 1) : length;
      }
    }
    outStart = startIdx;
    outStop = stopIdx;
  }

  // ``stepV`` is not zero. Start and stop are already normalized into a range that fits in signed 64-bit.
  inline SmInt estimateSliceItemCount(SmInt startV, SmInt stopV, SmInt stepV) {
    if (stepV > 0) {
      if (startV >= stopV) {
        return 0;
      }
      const SmInt dist = stopV - startV;
      return (dist - 1) / stepV + 1;
    }
    if (startV <= stopV) {
      return 0;
    }
    const SmInt dist = startV - stopV;
    const SmUInt mag = uintMagnitudeOfSInt(stepV);
    return static_cast<SmInt>((static_cast<SmUInt>(dist - 1) / mag) + 1ull);
  }

  // Negative indexes wrap. An index that does not land in ``[0, length)`` returns false.
  // ``length + idx`` runs only when the sum is inside that range, so it cannot overflow.
  inline bool tryResolveArrayIndex(SmInt length, SmInt idx, SmInt& resolved) {
    if (idx >= 0) {
      if (idx >= length) {
        return false;
      }
      resolved = idx;
      return true;
    }
    if (length > 0 && idx >= -length) {
      resolved = length + idx;
      return true;
    }
    return false;
  }

  bool nearlyInt(SmFloat v, SmInt& out);
  bool parseUIntFromFloat(SmFloat v, SmUInt& out);
  bool tryGetExactSignedIntFromScalarPolicy(
    const EvalValue::ScalarValue& s,
    SmInt& outI,
    ExactSignedIntPolicy policy);
  bool tryGetExactSignedIntFromScalar(const EvalValue::ScalarValue& s, SmInt& outI);
  bool tryGetExactSignedIntNoUIntWrapFromScalar(const EvalValue::ScalarValue& s, SmInt& outI);
  bool scalarHasExactIntegerPayload(const EvalValue::ScalarValue& s);
  bool TryMulExactSIntSquare(SmInt i, SmInt& outSq);
  bool TryGetExactSignedIntNoUIntWrapScalarStrict(const EvalValue::ScalarValue& s, SmInt& outI);
  enum class ArrayIndexIntKind { Ok, NotInteger, OutOfSignedRange };
  ArrayIndexIntKind classifyArrayIndexInt(const EvalValue::ScalarValue& sIn, SmInt& outI);
  bool tryGetArrayIndexIntFromScalarValue(const EvalValue::ScalarValue& sIn, SmInt& outI);
  bool tryGetExactNonNegativeUIntFromScalar(const EvalValue::ScalarValue& s, SmUInt& outU);
  bool tryGetSignedIntFromScalar(const EvalValue::ScalarValue& s, SmInt& outI);
  bool tryGetExactSignMagnitudeFromScalar(const EvalValue::ScalarValue& s, bool& isNeg, SmUInt& mag);
  EvalValue makeScalarFromSignMagnitude(bool isNeg, SmUInt mag);
  bool tryGetBothExactSignedIntNoUIntWrapFromScalars(
    const EvalValue::ScalarValue& a,
    const EvalValue::ScalarValue& b,
    SmInt& outA,
    SmInt& outB);
  bool tryGetBothExactNonNegativeUIntFromScalars(
    const EvalValue::ScalarValue& a,
    const EvalValue::ScalarValue& b,
    SmUInt& outA,
    SmUInt& outB);
  bool tryShiftLeftU64ExactOrMaybe(
    SmUInt aU,
    SmUInt bU,
    EvalValue& outV);
  bool tryMulUIntChecked(SmUInt a, SmUInt b, SmUInt& out);
  bool tryPowUIntChecked(SmUInt base, SmUInt exp, SmUInt& out);
  bool TryQuotExactSInt(SmInt num, SmInt den, SmInt& quo);
  bool passExactAbsFloatFactorGates(SmFloat f, bool strictAbsFAboveOne);
  bool getExactRoundedIntFromFloatResult(SmFloat r, bool boundAbsResultTo2_53, SmInt& outI);
  bool verifyExactIntFloatOpResidue(
    ScalarExactKind kind,
    SmInt intSigned,
    SmUInt intUnsigned,
    SmFloat f,
    SmInt resultSigned,
    SmUInt resultUnsigned,
    bool isMultiply);
  bool verifyExactIntFloatOpCrossMultiply(
    ScalarExactKind kind,
    SmInt intSigned,
    SmUInt intUnsigned,
    SmFloat floatV,
    SmInt resultSigned,
    SmUInt resultUnsigned,
    SmInt ratA,
    SmInt ratB,
    bool isMultiply);
  bool tryPromoteExactIntFromFloatOp(
    ScalarExactKind kind,
    SmInt intSigned,
    SmUInt intUnsigned,
    SmFloat f,
    SmFloat r,
    bool isMultiply,
    EvalValue& outV);
  bool tryPromoteExactDivisionByFloatDivisorSigned(
    SmInt intI, SmFloat f, SmFloat r, EvalValue& outV);
  bool tryPromoteExactDivisionByFloatDivisorUnsigned(
    SmUInt intU, SmFloat f, SmFloat r, EvalValue& outV);
  bool tryPromoteExactMultiplicationByFloatFactorSigned(
    SmInt intI, SmFloat f, SmFloat r, EvalValue& outV);
  bool tryPromoteExactMultiplicationByFloatFactorUnsigned(
    SmUInt intU, SmFloat f, SmFloat r, EvalValue& outV);
  bool tryApplyExactIntFloatOpFromFloatSide(
    ScalarExactKind kind,
    bool isMultiply,
    SmInt intSigned,
    SmUInt intUnsigned,
    SmFloat f,
    SmFloat r,
    EvalValue& outV);
  void scalarApplyReducedRationalPart(EvalValue::ScalarValue& s, bool imagPart, SmInt num,
                                          SmUInt den);
  void applySqrtScalarValue(const EvalValue::ScalarValue& sv, EvalValue& outV);
  bool tryApplySqrExactScalar(const EvalValue::ScalarValue& sv, EvalValue& outV);
  bool tryApplyHypotExactScalars(
    const EvalValue::ScalarValue& leftS,
    const EvalValue::ScalarValue& rightS,
    EvalValue& outV);
  bool tryApplyPowExactScalars(
    const EvalValue::ScalarValue& leftS,
    const EvalValue::ScalarValue& rightS,
    EvalValue& outV);
  bool tryApplyPowExactScalarsByKind(
    ScalarExactKind kind,
    SmInt valueSigned,
    SmUInt valueUnsigned,
    const EvalValue::ScalarValue& leftS,
    const EvalValue::ScalarValue& rightS,
    EvalValue& outV);
  EvalValue makeRationalReduced(SmInt num, SmUInt den);
  EvalValue negateEvalValue(const EvalValue& v);
  EvalValue applyAbsScalarValue(const EvalValue::ScalarValue& s);
  void exactCartesianComponentClear(ExactCartesianComponent& c);
  void ExactCartesianComponentAssignFromSignedInt(ExactCartesianComponent& c, SmInt n);
  void ExactCartesianComponentAssignFromUInt(ExactCartesianComponent& c, SmUInt u);
  bool tryExtractExactRealComponent(const EvalValue::ScalarValue& sv, ExactCartesianComponent& c);
  void setScalarFromExactCartesianComponent(EvalValue& v, const ExactCartesianComponent& c);
  int expandUnpackedArgs(const std::vector<EvalValue>& in, std::vector<EvalValue>& out);

#if SMARTMATH_TIME_VALUES
  SmInt timeTotalMsFromScalarValue(const EvalValue::ScalarValue& s);
#endif

#if SMARTMATH_COMPLEX_NUMBERS
  void scalarClearCartesianRenderExact(EvalValue::ScalarValue& s);
  bool tryGetExactImagSIntStrict(const EvalValue::ScalarValue& s, SmInt& outI);
  void ScalarApplyExactSIntPart(EvalValue::ScalarValue& s, bool imagPart, SmInt n);
  bool TryExactCartesianComponentToSInt(const ExactCartesianComponent& c, SmInt& outI);
  bool tryExtractExactCartesianComponent(bool isImag, const EvalValue::ScalarValue& sv,
                                              ExactCartesianComponent& c);
  bool tryNegateExactCartesianComponent(const ExactCartesianComponent& c, ExactCartesianComponent& outC);
  void setPureImaginaryFromMagnitudeScalar(EvalValue& outV, const EvalValue::ScalarValue& magSv);
  void complexCartesianBinary(SmFloat ar, SmFloat ai, SmFloat br, SmFloat bi, char op, SmFloat& outR,
                                   SmFloat& outI);
  bool tryExtractExactImagComponent(const EvalValue::ScalarValue& sv, ExactCartesianComponent& c);
  bool tryApplyExactCartesianBinaryOp(bool isAdd, const ExactCartesianComponent& a,
                                           const ExactCartesianComponent& b, ExactCartesianComponent& out);
  void scalarApplyExactImagFromCartesianComponent(EvalValue::ScalarValue& sv,
                                                       const ExactCartesianComponent& c);
  void setScalarComplexFromExactCartesian(EvalValue& v, const ExactCartesianComponent& re,
                                               const ExactCartesianComponent& im);
  bool tryNegateExactComplexScalar(const EvalValue::ScalarValue& sv, EvalValue& out);
#endif

  bool tryApplyRealScalarPowNegFractional(
    const EvalValue::ScalarValue& leftS,
    SmFloat p,
    EvalValue& outV);
  bool tryApplyScalarPowSpecialPaths(
    const EvalValue::ScalarValue& leftS,
    const EvalValue::ScalarValue& rightS,
    EvalValue& outV);
  EvalValue calcRoundingFn(BuiltinFunctionId id, const EvalValue::ScalarValue& s);

#if SMARTMATH_COMPLEX_NUMBERS
  bool tryRefinePowPrincipalToExactScalarResult(
    const EvalValue::ScalarValue& leftS,
    const EvalValue::ScalarValue& rightS,
    SmFloat powR,
    SmFloat powI,
    EvalValue& outV);
  bool tryVerifyComplexCartesianSquareExact(
    SmInt rootR, SmInt rootI, SmInt expR, SmInt expI);
  bool tryRefineSqrtPrincipalToExactComplex(
    const EvalValue::ScalarValue& inS,
    SmFloat sqrtR,
    SmFloat sqrtI,
    EvalValue& outV);
  EvalValue applySqrtComplexPrincipalUnary(const EvalValue::ScalarValue& inS);
  bool complexNeedsPrincipalNegRealPow(SmFloat ar, SmFloat ai, SmFloat br, SmFloat bi);
  EvalValue setScalarComplexFromEvalRealImagParts(const EvalValue& rePart, const EvalValue& imPart);
  bool tryApplyExactComplexCartesianBinary(const EvalValue::ScalarValue& leftS,
                                                const EvalValue::ScalarValue& rightS, char op,
                                                EvalValue& outV);
  bool tryFoldExactComplexCartesian(const std::vector<EvalValue>& args, char op, EvalValue& out);
  bool tryAvgExactComplexFromSum(const EvalValue& sumV, std::size_t itemCount, EvalValue& out);
  bool complexUnaryTrigCartesian(BuiltinFunctionId id, SmFloat ar, SmFloat ai, SmFloat& outR, SmFloat& outI);
  void scalarSplitCartesianParts(const EvalValue::ScalarValue& sv, SmFloat ar, SmFloat ai,
                                 EvalValue::ScalarValue& re, EvalValue::ScalarValue& im);
  EvalValue calcRoundingFnCartesian(
    BuiltinFunctionId rid,
    const EvalValue::ScalarValue& sv,
    SmFloat ar,
    SmFloat ai);
#endif

#if SMARTMATH_FACTORINT
  void setFactorintTermValue(
    EvalValue::ScalarValue& sv,
    SmInt signedValueI,
    SmUInt valueU,
    bool hasUIntValue);
  void setFactorintPowerTerm(
    EvalValue::ScalarValue& sv,
    SmUInt baseU,
    int expV,
    SmInt signedValueI,
    SmUInt valueU,
    bool hasUIntValue);
#endif

  std::string formatRationalParts(SmInt num, SmUInt den);
  bool tryFormatRationalScalar(const EvalValue::ScalarValue& sv, std::string& outText);
  std::string formatScalar(const EvalValue& v, RenderBase base, bool supportComplexNumbers);
  std::string valueToString(const EvalValue& v, RenderBase forcedBase, bool supportComplexNumbers);
  std::string valueToString(const EvalValue& v, bool supportComplexNumbers);

#if SMARTMATH_FACTORINT
  bool tryFormatIntPowerScalar(const EvalValue::ScalarValue& sv, std::string& outText);
#endif

#if SMARTMATH_COMPLEX_NUMBERS
  bool tryFormatComplexRationalScalar(const EvalValue::ScalarValue& sv, std::string& outText);
  std::string assembleComplexDecimalText(
    const std::string& rePart,
    const std::string& imagTail,
    bool negUnitImag,
    bool reZero);
  std::string formatComplexScalarValue(const EvalValue::ScalarValue& sv, bool supportComplexNumbers);
  std::string formatComplexScalarWithRenderBase(
    const EvalValue::ScalarValue& sv,
    RenderBase base,
    bool asUnsigned,
    bool supportComplexNumbers);
#endif

  enum class AggregateFoldMode : std::uint8_t { Real, Complex, Time };

  enum class BuiltinCategory : std::uint8_t {
    Rand,
    ScalarBinary,
    BaseFormat,
    Pow,
    UnaryMath,
    DegRad,
    Aggregate,
    ArrayTransform,
    Repeat,
    Range,
    Sortby,
    Ratio,
    Factorial,
    Factorint,
    Mod,
    TimeUnit,
    PolarCart,
    Log,
    Filter,
    Iif,
    IsInstance,
    Unknown,
  };

  enum class UnaryScalarKind : std::uint8_t {
    None,
    PrimaryTrig,
    InverseTrig,
    HyperbolicTrig,
    InverseHyperbolicTrig,
    Exp,
    Ln,
    Log10,
    Sqrt,
    Sqr,
    Rounding,
    Frac,
    Abs,
    Sign,
    Deg,
    Rad,
  };

  enum class OperatorNameId {
    Not = 0,
    And,
    Or,
    Count
  };

  enum class BuiltinHintKind {
    None = 0,
    EmptyPar,
    MinMax,
    DotDotDot,
    ValuePower,
    YX,
    Angle,
    Value,
    ValueBase,
    N,
    ValueDivisor,
    ValueMinMax,
    XY,
    AB,
    XN,
    X,
    StartStopStep,
    ArrayFunc,
    Iif,
    IsInstance
  };

  enum class BuiltinFlags : unsigned {
    None = 0,
    Unary = 1u << 0,
    Format = 1u << 1,
    IntegerOnly = 1u << 2,
    NonCalculating = 1u << 3,
    FiniteRequired = 1u << 4,
    TrailingFormatter = 1u << 5,
  };

  constexpr BuiltinFlags operator|(BuiltinFlags a, BuiltinFlags b) noexcept {
    return static_cast<BuiltinFlags>(static_cast<unsigned>(a) | static_cast<unsigned>(b));
  }

  struct BuiltinMetaRow {
    BuiltinFlags flags;
    std::uint8_t minArgs;
    std::uint8_t maxArgs;
    BuiltinHintKind hintKind;
    BuiltinCategory category;
  };

  enum class ParseErrorId : std::uint8_t {
    UnexpectedComma,
    IndexingRequiresArray,
    MissingIndex,
    ArrayIndexMustBeScalar,
    ArrayIndexMustBeInteger,
    ArrayIndexOutOfRange,
    MissingClosingBracket,
    MissingClosingParenthesis,
    MismatchedClosingParenthesis,
    MismatchedClosingBracket,
    MismatchedClosingBrace,
    BitwiseIntegerOperands,
    ModuloIntegerOperands,
    IncompatibleOperands,
    InvalidHexLiteral,
    InvalidBinaryLiteral,
    InvalidOctalLiteral,
    InternalUnaryOp,
    InternalBinaryOp,
    InternalEvalError,
    NumericErrorInPowerOperation,
    NumericErrorInExpression,
    UserFunctionCallStackOverflow,
    MaxEvaluationDepthReached,
    InvalidNumericLiteral,
    PercentageRequiresScalarValue,
    InternalAggregateBuiltin,
    InternalScalarBinaryBuiltin,
    InternalUnaryMathBuiltin,
    UnexpectedTokenAfterExpression,
    ScalarOnlyExpressionEncounteredNon,
    ParseFailed,
    UnexpectedToken,
    UnexpectedInput,
    CannotAssignToArraySlice,
    Count,
  };

  enum class CallErrorId : std::uint8_t {
    Numeric,
    AtLeastN,
    ExactCount,
    ScalarValues,
    IntegerValues,
    ScalarMinMax,
    NonNegativeInteger,
    PositiveInteger,
    RangeStepZero,
    TooManyValues,
  };

  struct AstStatement;

  struct UserFunction {
    std::string name;
    std::vector<std::string> params;
    std::string expr;
    std::vector<AstStatement> compiledProgram;
    bool compiledProgramReady = false;
  };

  struct Expr {
    enum class Tag {
      Literal,
      Variable,
      Unary,
      Binary,
      Call,
      Index,
      Slice,
      ArrayOrParens,
      PostfixPercent,
      FunctionRef,
      TypeName
    };

    enum class BinaryOp {
      None = 0,
      LogicalOr,
      LogicalAnd,
      Modulo,
      BitAnd,
      BitOr,
      BitXor,
      ShiftLeft,
      ShiftRight,
      CmpLt,
      CmpGt,
      CmpLe,
      CmpGe,
      CmpEq,
      CmpNe,
      Add,
      Sub,
      Mul,
      Div,
      Pow
    };

    Tag tag = Tag::Literal;
    EvalValue literalValue{};
    std::string name;
    BuiltinFunctionId builtinFunctionId = BuiltinFunctionId::Count;
    const EvalValue* boundVariable = nullptr;
    BinaryOp binaryOp = BinaryOp::None;
    bool rhsIsDirectPostfixPercent = false;
    char unaryOp = 0;
    std::unique_ptr<Expr> child;
    std::unique_ptr<Expr> left;
    std::unique_ptr<Expr> right;
    std::vector<std::unique_ptr<Expr>> elements;
  };

  struct AstStatement {
    enum class Kind { FunDef, Assign, IndexAssign, Expr };
    Kind kind = Kind::Expr;
    UserFunction fun;
    std::string assignName;
    std::unique_ptr<Expr> expr;
    /// Bracket of an ``IndexAssign``: ``Index`` (right is the index) or ``Slice``.
    std::unique_ptr<Expr> indexExpr;
  };

  struct EvalContext {
    const char* p = nullptr;
    const char* start = nullptr;
    std::string sourceExpr;
    bool parseError = false;
    std::string errorText;
    int evalDepth = 0;
    /** Nesting depth of ``[...]`` while parsing; suppresses colon-time literals so ``a[1:30]`` is a slice. */
    int arrayIndexBracketDepth = 0;
    /** Collected like FreeBASIC (non-fatal until end of eval). */
    std::string unknownVarsText;
    std::string unknownFuncsText;
    /** This statement referenced an unknown variable or function. */
    bool referencedUnresolvedName = false;
    /** Params of UDFs seen earlier in the program currently being compiled (not yet in userFunctions_). */
    const std::unordered_map<std::string, std::vector<std::string>>* compilingUserFunctionParams = nullptr;
    /** Shared stack for the current top-level evaluate(); null outside evaluate(). */
    std::vector<std::string>* userFunctionCallStack = nullptr;
    /** Set only while parsing a UDF body for definition validation. */
    const char* validatingUserFunctionName = nullptr;
    /**
     * Parameter names in scope while syntax-checking or compiling this body.
     * A bare parameter is a variable even when the same spelling is a builtin or a user function.
     * A nested lambda binds its own parameters for that check, the same as Basic's override list.
     */
    const std::vector<std::string>* boundBodyParamNames = nullptr;
    /** Set while evaluating an inline lambda body (e.g. sortby key). */
    // Path C: set when trailing ';' was stripped. Path B: *p == ';' in trySetMissingFunctionCallError.
    bool suppressBareFunctionTailHint = false;

    /// First error wins. Later calls leave the existing text in place.
    void fail(const std::string& msg);
    void fail(ParseErrorId id);
    void failCall(CallErrorId id, const std::string& fnName, std::size_t n = 0);
    void failHint(const std::string& hintText);
    void failPrefixedLiteral(char prefixChar);
    void failRecursiveUserFunction(const std::string& fnName);
    void failMissingClosingParenLike();
    void failMissingClosingBracketAfterIndex();
  };

  struct CompiledProgramState {
    std::vector<AstStatement> program;
    bool ready = false;
    bool scalarOnly = false;
    bool hasAssignments = false;
    std::size_t boundVariablesVersion = static_cast<std::size_t>(-1);
  };
}  // namespace smartmath_internal
