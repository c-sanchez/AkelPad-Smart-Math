#pragma once

#include <cmath>
#include <cstdint>
#include <limits>
#include <string>
#include <memory>
#include <unordered_map>
#include <vector>

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
    Count
  };

  enum class ValueKind {
    Scalar,
    Array,
    FunctionRef,
    InlineLambda,
    UdfFormalValidationDummy,
    DefinedFunction
  };

  enum class ScalarKind {
    FloatingPoint,
    Int64,
    UInt64,
    Time
  };

  enum class RenderBase {
    Dec = 10,
    Hex = 16,
    Oct = 8,
    Bin = 2
  };

  struct EvalValue {
    struct ScalarValue {
      enum : unsigned int {
        fExactInt64Valid = 0x01u,
        fExactUInt64Valid = 0x02u,
        fDecScientificPow63High = 0x10u,
        fImagExactInt64Valid = 0x20u,
        fImagExactUInt64Valid = 0x40u,
        fRenderRational = 0x80u,
        fImagRenderRational = 0x100u
      #if SMARTMATH_FACTORINT
        ,
        fRenderIntPower = 0x200u
      #endif
      };

      ScalarKind scalarKind = ScalarKind::FloatingPoint;
      /// Do not read for exact-int policy; use hasExactInt64() / hasRenderRational().
      unsigned int flags = 0;
      double scalar = 0.0;
      /// Numerator for render-rational and exact integers; prefer hasExactInt64() in eval paths.
      long long exactInt64 = 0;
      /// Denominator for render-rational and exact uintegers; prefer hasExactUInt64() in eval paths.
      std::uint64_t exactUInt64 = 0;
      double imag = 0.0;
      long long imagExactInt64 = 0;
      std::uint64_t imagExactUInt64 = 0;

      bool hasExactInt64Metadata() const { return (flags & fExactInt64Valid) != 0; }
      bool hasExactUInt64Metadata() const { return (flags & fExactUInt64Valid) != 0; }
      bool hasImagExactInt64Metadata() const { return (flags & fImagExactInt64Valid) != 0; }
      bool hasImagExactUInt64Metadata() const { return (flags & fImagExactUInt64Valid) != 0; }
      /// Exact int64 for arithmetic (not render-rational numerator storage).
      bool hasExactInt64() const { return hasExactInt64Metadata() && !hasRenderRational(); }
      void setExactInt64Valid(bool v) { flags = v ? (flags | fExactInt64Valid) : (flags & ~fExactInt64Valid); }
      /// Exact uint64 for arithmetic (not render-rational denominator storage).
      bool hasExactUInt64() const { return hasExactUInt64Metadata() && !hasRenderRational(); }
      void setExactUInt64Valid(bool v) { flags = v ? (flags | fExactUInt64Valid) : (flags & ~fExactUInt64Valid); }
      bool hasDecScientificPow63High() const { return (flags & fDecScientificPow63High) != 0; }
      void setDecScientificPow63High(bool v) {
        flags = v ? (flags | fDecScientificPow63High) : (flags & ~fDecScientificPow63High);
      }
      bool hasImagExactInt64() const { return hasImagExactInt64Metadata() && !hasImagRenderRational(); }
      void setImagExactInt64Valid(bool v) { flags = v ? (flags | fImagExactInt64Valid) : (flags & ~fImagExactInt64Valid); }
      bool hasImagExactUInt64() const { return hasImagExactUInt64Metadata() && !hasImagRenderRational(); }
      void setImagExactUInt64Valid(bool v) {
        flags = v ? (flags | fImagExactUInt64Valid) : (flags & ~fImagExactUInt64Valid);
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
      fFormalProbeTaint = 0x08u, // UDF validation: value depends on a formal probe.
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
    /// When true with hex/oct/bin output, format negative integers as unsigned (two's complement) bit pattern.
    bool hasExpandArgs() const { return (flags & fExpandArgs) != 0; }
    void setExpandArgs(bool v) { flags = v ? (flags | fExpandArgs) : (flags & ~fExpandArgs); }
    bool hasRenderUnsigned() const { return (flags & fRenderUnsigned) != 0; }
    void setRenderUnsigned(bool v) { flags = v ? (flags | fRenderUnsigned) : (flags & ~fRenderUnsigned); }
    bool hasFormalProbeTaint() const { return (flags & fFormalProbeTaint) != 0; }
    void setFormalProbeTaint(bool v) { flags = v ? (flags | fFormalProbeTaint) : (flags & ~fFormalProbeTaint); }
    RenderBase getRenderBase() const {
      const unsigned int raw = (flags & fRenderBaseMask) >> fRenderBaseShift;
      return (raw == 0u) ? RenderBase::Dec : static_cast<RenderBase>(raw);
    }
    void setRenderBase(RenderBase v) {
      flags = (flags & ~fRenderBaseMask) |
              ((static_cast<unsigned int>(v) & 0xFFu) << fRenderBaseShift);
    }
  };

  EvalValue makeScalar(double v);
  EvalValue makeScalarMaybeExact(double v);
  EvalValue makeScalarInt(long long v);
  EvalValue makeScalarUInt(std::uint64_t v);
  EvalValue makeScalarUIntPreferInt(std::uint64_t v);

#if SMARTMATH_TIME_VALUES
  EvalValue makeScalarTimeMs(long long totalMs);
#endif

#if SMARTMATH_COMPLEX_NUMBERS
  EvalValue makeImaginaryUnit();
  EvalValue makeScalarComplexFromDoubles(double re, double im);
#endif

  EvalValue makeArrayFromScalars(const std::vector<EvalValue>& v);
  EvalValue makeInlineLambdaValue(std::vector<std::string> params, std::string body);
  EvalValue makeUdfFormalValidationDummy();

  void scalarClearImaginary(EvalValue::ScalarValue& s);

  enum class ExactSignedInt64Policy {
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
    long long intV = 0;
    bool hasUInt = false;
    std::uint64_t uintV = 0;
  };

  constexpr double kMaxExactIntFromDouble = 9007199254740992.0;  // 2^53

  inline bool isWithinExactIntFromDoubleRange(double v) {
    return std::isfinite(v) && std::fabs(v) <= kMaxExactIntFromDouble;
  }

  inline bool tryExtractExactInt64FromDoubleStrict(double v, long long& out) {
    if (!isWithinExactIntFromDoubleRange(v)) {
      return false;
    }
    const double kI64MinD = static_cast<double>((std::numeric_limits<long long>::min)());
    const double kI64MaxD = static_cast<double>((std::numeric_limits<long long>::max)());
    if (v < kI64MinD || v > kI64MaxD) {
      return false;
    }
    const long long asInt = static_cast<long long>(v);
    if (v != static_cast<double>(asInt)) {
      return false;
    }
    out = asInt;
    return true;
  }

  inline std::uint64_t gcdUInt64(std::uint64_t a, std::uint64_t b) {
    while (b != 0u) {
      const std::uint64_t t = a % b;
      a = b;
      b = t;
    }
    return a;
  }

  inline std::uint64_t uint64MagnitudeOfInt64(long long v) {
    return v < 0 ? (std::uint64_t{0} - static_cast<std::uint64_t>(v)) : static_cast<std::uint64_t>(v);
  }

  // mag must be <= 2^63.
  inline long long int64FromNegatedMagnitude(std::uint64_t mag) {
    return mag == (1ull << 63) ? (std::numeric_limits<long long>::min)() : -static_cast<long long>(mag);
  }

  bool nearlyInt(double v, long long& out);
  bool parseUInt64FromDouble(double v, std::uint64_t& out);
  bool tryGetExactSignedInt64FromScalarPolicy(
    const EvalValue::ScalarValue& s,
    long long& outI,
    ExactSignedInt64Policy policy);
  bool tryGetExactSignedInt64FromScalar(const EvalValue::ScalarValue& s, long long& outI);
  bool tryGetExactSignedInt64NoUIntWrapFromScalar(const EvalValue::ScalarValue& s, long long& outI);
  bool scalarHasExactIntegerPayload(const EvalValue::ScalarValue& s);
  bool tryMulExactInt64Square(long long i, long long& outSq);
  bool tryGetExactSignedInt64NoUIntWrapScalarStrict(const EvalValue::ScalarValue& s, long long& outI);
  bool tryGetArrayIndexIntFromScalarValue(const EvalValue::ScalarValue& sIn, long long& outI);
  bool tryGetExactNonNegativeUInt64FromScalar(const EvalValue::ScalarValue& s, std::uint64_t& outU);
  bool tryGetSignedInt64FromScalar(const EvalValue::ScalarValue& s, long long& outI);
  bool tryGetExactSignMagnitudeFromScalar(const EvalValue::ScalarValue& s, bool& isNeg, std::uint64_t& mag);
  EvalValue makeScalarFromSignMagnitude(bool isNeg, std::uint64_t mag);
  bool tryGetBothExactSignedInt64NoUIntWrapFromScalars(
    const EvalValue::ScalarValue& a,
    const EvalValue::ScalarValue& b,
    long long& outA,
    long long& outB);
  bool tryGetBothExactNonNegativeUInt64FromScalars(
    const EvalValue::ScalarValue& a,
    const EvalValue::ScalarValue& b,
    std::uint64_t& outA,
    std::uint64_t& outB);
  bool tryShiftLeftU64ExactOrMaybe(
    std::uint64_t aU,
    std::uint64_t bU,
    EvalValue& outV);
  bool tryMulUInt64Checked(std::uint64_t a, std::uint64_t b, std::uint64_t& out);
  bool tryPowUInt64Checked(std::uint64_t base, std::uint64_t exp, std::uint64_t& out);
  bool tryQuotExactInt64(long long num, long long den, long long& quo);
  bool passExactAbsFloatFactorGates(double f, bool strictAbsFAboveOne);
  bool getExactRoundedIntFromFloatResult(double r, bool boundAbsResultTo2_53, long long& outI);
  bool verifyExactIntFloatOpResidue(
    ScalarExactKind kind,
    long long intSigned,
    std::uint64_t intUnsigned,
    double f,
    long long resultSigned,
    std::uint64_t resultUnsigned,
    bool isMultiply);
  bool verifyExactIntFloatOpCrossMultiply(
    ScalarExactKind kind,
    long long intSigned,
    std::uint64_t intUnsigned,
    double floatV,
    long long resultSigned,
    std::uint64_t resultUnsigned,
    long long ratA,
    long long ratB,
    bool isMultiply);
  bool tryPromoteExactIntFromFloatOp(
    ScalarExactKind kind,
    long long intSigned,
    std::uint64_t intUnsigned,
    double f,
    double r,
    bool isMultiply,
    EvalValue& outV);
  bool tryPromoteExactDivisionByFloatDivisorSigned(
    long long intI, double f, double r, EvalValue& outV);
  bool tryPromoteExactDivisionByFloatDivisorUnsigned(
    std::uint64_t intU, double f, double r, EvalValue& outV);
  bool tryPromoteExactMultiplicationByFloatFactorSigned(
    long long intI, double f, double r, EvalValue& outV);
  bool tryPromoteExactMultiplicationByFloatFactorUnsigned(
    std::uint64_t intU, double f, double r, EvalValue& outV);
  bool tryApplyExactIntFloatOpFromFloatSide(
    ScalarExactKind kind,
    bool isMultiply,
    long long intSigned,
    std::uint64_t intUnsigned,
    double f,
    double r,
    EvalValue& outV);
  void scalarApplyReducedRationalPart(EvalValue::ScalarValue& s, bool imagPart, long long num,
                                          std::uint64_t den);
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
    long long valueSigned,
    std::uint64_t valueUnsigned,
    const EvalValue::ScalarValue& leftS,
    const EvalValue::ScalarValue& rightS,
    EvalValue& outV);
  EvalValue makeRationalReduced(long long num, std::uint64_t den);
  EvalValue negateEvalValue(const EvalValue& v);
  EvalValue applyAbsScalarValue(const EvalValue::ScalarValue& s);
  void exactCartesianComponentClear(ExactCartesianComponent& c);
  void exactCartesianComponentAssignFromSignedInt64(ExactCartesianComponent& c, long long n);
  void exactCartesianComponentAssignFromUInt64(ExactCartesianComponent& c, std::uint64_t u);
  bool tryExtractExactRealComponent(const EvalValue::ScalarValue& sv, ExactCartesianComponent& c);
  void setScalarFromExactCartesianComponent(EvalValue& v, const ExactCartesianComponent& c);
  int expandUnpackedArgs(const std::vector<EvalValue>& in, std::vector<EvalValue>& out);

#if SMARTMATH_TIME_VALUES
  long long timeTotalMsFromScalarValue(const EvalValue::ScalarValue& s);
#endif

#if SMARTMATH_COMPLEX_NUMBERS
  void scalarClearCartesianRenderExact(EvalValue::ScalarValue& s);
  bool tryGetExactImagInt64Strict(const EvalValue::ScalarValue& s, long long& outI);
  void scalarApplyExactInt64Part(EvalValue::ScalarValue& s, bool imagPart, long long n);
  bool tryExactCartesianComponentToInt64(const ExactCartesianComponent& c, long long& outI);
  bool tryExtractExactCartesianComponent(bool isImag, const EvalValue::ScalarValue& sv,
                                              ExactCartesianComponent& c);
  bool tryNegateExactCartesianComponent(const ExactCartesianComponent& c, ExactCartesianComponent& outC);
  void setPureImaginaryFromMagnitudeScalar(EvalValue& outV, const EvalValue::ScalarValue& magSv);
  void complexCartesianBinary(double ar, double ai, double br, double bi, char op, double& outR,
                                   double& outI);
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
    double p,
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
    double powR,
    double powI,
    EvalValue& outV);
  bool tryVerifyComplexCartesianSquareExact(
    long long rootR, long long rootI, long long expR, long long expI);
  bool tryRefineSqrtPrincipalToExactComplex(
    const EvalValue::ScalarValue& inS,
    double sqrtR,
    double sqrtI,
    EvalValue& outV);
  EvalValue applySqrtComplexPrincipalUnary(const EvalValue::ScalarValue& inS);
  bool complexNeedsPrincipalNegRealPow(double ar, double ai, double br, double bi);
  EvalValue setScalarComplexFromEvalRealImagParts(const EvalValue& rePart, const EvalValue& imPart);
  bool tryApplyExactComplexCartesianBinary(const EvalValue::ScalarValue& leftS,
                                                const EvalValue::ScalarValue& rightS, char op,
                                                EvalValue& outV);
  bool tryFoldExactComplexCartesian(const std::vector<EvalValue>& args, char op, EvalValue& out);
  bool tryAvgExactComplexFromSum(const EvalValue& sumV, std::size_t itemCount, EvalValue& out);
  bool complexUnaryTrigCartesian(BuiltinFunctionId id, double ar, double ai, double& outR, double& outI);
  void scalarSplitCartesianParts(const EvalValue::ScalarValue& sv, double ar, double ai,
                                 EvalValue::ScalarValue& re, EvalValue::ScalarValue& im);
  EvalValue calcRoundingFnCartesian(
    BuiltinFunctionId rid,
    const EvalValue::ScalarValue& sv,
    double ar,
    double ai);
#endif

#if SMARTMATH_FACTORINT
  void setFactorintTermValue(
    EvalValue::ScalarValue& sv,
    long long signedValueI,
    std::uint64_t valueU,
    bool hasUIntValue);
  void setFactorintPowerTerm(
    EvalValue::ScalarValue& sv,
    std::uint64_t baseU,
    int expV,
    long long signedValueI,
    std::uint64_t valueU,
    bool hasUIntValue);
#endif

  std::string formatRationalParts(long long num, std::uint64_t den);
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
    ArrayFunc
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
      FunctionRef
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
    enum class Kind { FunDef, Assign, Expr };
    Kind kind = Kind::Expr;
    UserFunction fun;
    std::string assignName;
    std::unique_ptr<Expr> expr;
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
  };

  struct CompiledProgramState {
    std::vector<AstStatement> program;
    bool ready = false;
    bool scalarOnly = false;
    bool hasAssignments = false;
    std::size_t boundVariablesVersion = static_cast<std::size_t>(-1);
  };
}  // namespace smartmath_internal
