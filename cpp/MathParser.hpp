#pragma once

#include <cstddef>
#include <cstdint>
#include <memory>
#include <string>
#include <unordered_map>
#include <vector>

#include "MathParserInternal.hpp"

class MathParser {
public:
  struct RawResult {
    enum class Kind { None, Scalar, Array };
    enum class ScalarKind { FloatingPoint, Int64, UInt64, Rational,
#if SMARTMATH_FACTORINT
                            IntPower,
#endif
                            Complex, Time };

    /// One real or imaginary component. Only the union member for ``kind`` is meaningful.
    struct CartesianScalar {
      ScalarKind kind = ScalarKind::FloatingPoint;
      union {
        double floatingPoint;    // FloatingPoint
        long long intValue;      // Int64 and Time (total milliseconds)
        std::uint64_t uintValue; // UInt64
        struct {
          long long numerator;       // Rational numerator; IntPower base
          std::uint64_t denominator; // Rational denominator; IntPower exponent
        } rational;
      };

      CartesianScalar() : kind(ScalarKind::FloatingPoint), floatingPoint(0.0) {}

      void clear();
      bool isZero() const;

      bool isFloatingPoint() const { return kind == ScalarKind::FloatingPoint; }
      bool isInt64() const { return kind == ScalarKind::Int64; }
      bool isUInt64() const { return kind == ScalarKind::UInt64; }
      bool isRational() const { return kind == ScalarKind::Rational; }
#if SMARTMATH_FACTORINT
      bool isIntPower() const { return kind == ScalarKind::IntPower; }
#endif
      bool isTime() const { return kind == ScalarKind::Time; }
    };

    /// Non-complex: payload in ``real`` only; ``imag`` cleared; ``kind`` mirrors ``real.kind``.
    /// Complex: ``kind == Complex``; each part uses ``real`` / ``imag``.
    /// Rational -> real.rational member.
    /// Time -> real.intValue ms.
    struct Scalar {
      ScalarKind kind = ScalarKind::FloatingPoint;
      CartesianScalar real{};
      CartesianScalar imag{};

      bool isComplex() const { return kind == ScalarKind::Complex; }
      bool isFloatingPoint() const { return !isComplex() && real.isFloatingPoint(); }
      bool isInt64() const { return !isComplex() && real.isInt64(); }
      bool isUInt64() const { return !isComplex() && real.isUInt64(); }
      bool isRational() const { return !isComplex() && real.isRational(); }
      bool isTime() const { return !isComplex() && real.isTime(); }
    };

    /// None: no value available (for example, after evaluation error).
    /// Scalar: scalar field is valid, array is empty.
    /// Array: array field is valid, scalar should be ignored.
    Kind kind = Kind::None;
    /// Meaningful only when kind == Scalar; otherwise ignore.
    Scalar scalar{};
    /// Meaningful only when kind == Array; otherwise empty.
    std::vector<Scalar> array;

    void clear();
    bool hasValue() const { return kind != Kind::None; }
    bool isScalar() const { return kind == Kind::Scalar; }
    bool isArray() const { return kind == Kind::Array; }
  };

  MathParser();

  /// Parse into an internal program (AST). O(expression length). Does not evaluate.
  /// On failure, getError() is non-empty and no program is retained.
  bool compile(const std::string& mathExpression);

  /// Run the last successful compile() against the current variables, constants, and user functions.
  /// Repeated calls re-evaluate without re-parsing. Fails if compile() was never successful.
  void evaluate();

  /// Equivalent to compile() followed by evaluate() when you only need a single shot.
  void parseAndEvaluate(const std::string& mathExpression);
  /// Single-shot parse+evaluate that returns raw value directly; returns empty RawResult on error.
  RawResult parseAndEvaluateRaw(const std::string& mathExpression);

  std::string getError() const;
  std::string getResult() const;
  std::string getResultAsHex() const;
  std::string getResultAsDec() const;
  std::string getResultAsOct() const;
  std::string getResultAsBin() const;
  RawResult getRawResult() const;

  /// When false (default), the parser keeps real-only scalar semantics for non-finite or non-real outcomes.
  /// When true, future releases may evaluate complex-valued builtins and accept complex literals/arrays.
  void setSupportComplexNumbers(bool enabled);
  bool getSupportComplexNumbers() const;
  void setSupportTimeValues(bool enabled);
  bool getSupportTimeValues() const;
  void setSupportLambdaFunctions(bool enabled);
  bool getSupportLambdaFunctions() const;

  void addConst(const std::string& constName, long long intValue);
  void addConst(const std::string& constName, double dblValue);
  std::string addUserFunction(const std::string& mathExpression);

private:
  using BuiltinFunctionId = smartmath_internal::BuiltinFunctionId;

  using AggregateFoldMode = smartmath_internal::AggregateFoldMode;
  using BuiltinCategory = smartmath_internal::BuiltinCategory;
  using UnaryScalarKind = smartmath_internal::UnaryScalarKind;
  using OperatorNameId = smartmath_internal::OperatorNameId;
  using BuiltinHintKind = smartmath_internal::BuiltinHintKind;
  using BuiltinFlags = smartmath_internal::BuiltinFlags;
  using BuiltinMetaRow = smartmath_internal::BuiltinMetaRow;

  using ValueKind = smartmath_internal::ValueKind;
  using ScalarKind = smartmath_internal::ScalarKind;
  using RenderBase = smartmath_internal::RenderBase;
  using EvalValue = smartmath_internal::EvalValue;
  using ExactSignedInt64Policy = smartmath_internal::ExactSignedInt64Policy;
  using ScalarExactKind = smartmath_internal::ScalarExactKind;
  using ExactCartesianComponent = smartmath_internal::ExactCartesianComponent;

  using UserFunction = smartmath_internal::UserFunction;
  using Expr = smartmath_internal::Expr;
  using AstStatement = smartmath_internal::AstStatement;
  using ParseErrorId = smartmath_internal::ParseErrorId;
  using CallErrorId = smartmath_internal::CallErrorId;
  using EvalContext = smartmath_internal::EvalContext;
  using CompiledProgramState = smartmath_internal::CompiledProgramState;

  EvalValue lastResult_;
  bool hasResult_ = false;
  std::string lastError_;
  std::unordered_map<std::string, EvalValue> variables_;
  std::vector<UserFunction> userFunctions_;
  std::unordered_map<std::string, std::size_t> userFunctionIndex_;
  CompiledProgramState compiled_;
  std::size_t variablesVersion_ = 0;
  bool supportComplexNumbers_ = false;
  bool supportTimeValues_ = true;
  bool supportLambdaFunctions_ = true;

  bool prepareCompileParseSource(
      const std::string& mathExpression,
      EvalContext& ctx,
      std::string& ownedParseBuffer,
      bool& strippedTrailingSemicolon);
  bool tryConsumeCommaArgSeparator(EvalContext& ctx, bool& hasComma) const;
  bool tryConsumeLogicalBinaryOperator(EvalContext& ctx, OperatorNameId keywordId, char symbol) const;
  bool parseParenthesizedExprList(
      EvalContext& ctx,
      std::vector<std::unique_ptr<Expr>>& outValues);
  std::unique_ptr<Expr> parsePrimaryParenthesized(EvalContext& ctx);
  std::unique_ptr<Expr> parsePrimaryNumericLiteral(EvalContext& ctx);
#if SMARTMATH_TIME_VALUES
  bool tryParseScalarTimeLiteral(EvalContext& ctx, EvalValue& out) const;
  bool tryParseCompactSuffixTimeLiteral(EvalContext& ctx, EvalValue& out) const;
#endif
  std::unique_ptr<Expr> parsePrimaryIdentifierOrCall(EvalContext& ctx);
  bool parseUnaryFunctionCallArguments(
      EvalContext& ctx,
      std::vector<std::unique_ptr<Expr>>& outArgs,
      const char* expectsOneFunction,
      const char* expectsUnaryFunction);
  std::unique_ptr<Expr> parseUnaryFunctionRef(
      EvalContext& ctx,
      const char* expectsOneFunction,
      const char* expectsUnaryFunction);
  enum class LambdaBodyStop {
    WrappedParenClose,
    ArgDelim,
    TopLevelSemicolonOrEof
  };
#if SMARTMATH_LAMBDA_FUNCTIONS
  bool tryConsumeLambdaParameterList(EvalContext& ctx, std::vector<std::string>& outParams, bool quiet) const;
  static bool lambdaBodyConsume(
      EvalContext& ctx,
      std::string& outBody,
      LambdaBodyStop stop,
      const char* lambdaArgStart);
  bool tryParseLambdaInnerUnwrappedSuffix(
      EvalContext& ctx,
      std::vector<std::string>& outParams,
      std::string& outBody,
      LambdaBodyStop bodyStop,
      const char* lambdaArgStart,
      bool quiet) const;
  bool tryParseWrappedParenLambdaSuffix(
      EvalContext& ctx,
      std::vector<std::string>& outParams,
      std::string& outBody,
      LambdaBodyStop bodyStop,
      bool quiet) const;
  bool tryParseLambdaRhsAfterEquals(EvalContext& ctx, std::vector<std::string>& outParams, std::string& outExpr) const;
  std::unique_ptr<Expr> tryParseUnaryLambdaOrFunctionRef(
      EvalContext& ctx,
      const char* expectsOneFunction,
      const char* expectsUnaryFunction);
  std::unique_ptr<Expr> makeUnaryInlineLambdaExpr(
      EvalContext& ctx,
      std::vector<std::string>&& params,
      std::string&& body,
      const char* expectsUnaryFunction);
  EvalValue evalInlineLambdaCall(
      EvalContext& ctx,
      const std::vector<std::string>& paramNames,
      const std::string& bodyExpr,
      std::vector<EvalValue>&& args);
#endif
  std::unique_ptr<Expr> parseUnaryFunctionArg(
      EvalContext& ctx,
      const char* expectsOneFunction,
      const char* expectsUnaryFunction);
  std::unique_ptr<Expr> parseUnaryFunctionRefOnly(
      EvalContext& ctx,
      const char* expectsOneFunction,
      const char* expectsUnaryFunction);
  static bool udfBodyIsEmptyTupleLiteral(const std::string& bodyExpr);

#if SMARTMATH_COMPLEX_NUMBERS
  void applyComplexCaretPrincipalEval(
      const EvalValue::ScalarValue& lv,
      const EvalValue::ScalarValue& rv,
      EvalValue& outS) const;
#endif
  EvalValue applyUnarySqrtEval(const EvalValue::ScalarValue& s) const;

  bool validateCallArity(
      EvalContext& ctx,
      const std::string& fnName,
      uint8_t minArgs,
      uint8_t maxArgs,
      std::size_t argc) const;
  bool validateBuiltinCallArity(
      EvalContext& ctx,
      const std::string& fnName,
      BuiltinFunctionId id,
      const std::vector<EvalValue>& args) const;
  bool validateIntegerRepresentableArgs(
      EvalContext& ctx,
      const std::string& fnName,
      const std::vector<EvalValue>& args,
      bool allowNonFiniteForFormat) const;
  bool validateBuiltinArgs(
      EvalContext& ctx,
      const std::string& fnName,
      BuiltinFunctionId id,
      const std::vector<EvalValue>& args) const;

  std::string getResultWithBase(RenderBase base) const;
  const char* getReservedIdentifierError(const std::string& ident) const;
  bool isBareFunctionNameAtExpressionTail(EvalContext& ctx) const;
  bool trySetBareFunctionImmediateCloserError(EvalContext& ctx, const char* identStart) const;

  bool identIsBareFunctionOrUdfName(const std::string& ident, const EvalContext& ctx) const;
  bool trimmedStmtIsBareFunctionOrUdfName(const std::string& stmt) const;
  bool trimmedStmtIsBareFunctionOrUdfName(const char* begin, const char* end) const;
  bool stripTrailingSemicolonsForTopLevelInput(std::string& s) const;
  bool trySetMissingFunctionCallError(
      EvalContext& ctx,
      const std::string& ident,
      const char* identStartForCloser = nullptr) const;
  bool trySetIncompleteOpenedFunctionCallHint(EvalContext& ctx, const std::string& ident) const;
  bool trySetBareUserFunctionNameError(
      EvalContext& ctx,
      const std::string& ident,
      const char* identStart) const;
  bool tryHandleBareFunctionNameCommonTail(
      EvalContext& ctx,
      const std::string& ident,
      const char* identStartForCloser) const;
  bool finishSubEvalFromCompiledProgram(EvalContext& parent, EvalContext& sub) const;
  bool tryCompileExprToProgram(
      const std::string& expr,
      EvalContext& parentCtx,
      std::vector<std::string>* callStack,
      const char* validatingUserFunctionName,
      bool suppressBareFunctionTailHint,
      std::vector<AstStatement>& outProgram);
  static std::string formatUserFunctionSignature(const UserFunction& uf);
  bool handleUnknownIdentifier(EvalContext& ctx, const std::string& ident, std::string& unknownList) const;
  bool tryResolveVariableValue(
      const Expr& e,
      const std::unordered_map<std::string, EvalValue>* scopedVars,
      EvalValue& out) const;
  const char* validateUserFunctionDefinitionNames(
      const std::string& fnName,
      const std::vector<std::string>& fnParams) const;
  std::string getUserFunctionDefinitionErrorText(
      const std::string& fnName,
      const std::vector<std::string>& fnParams,
      const std::string& fnExpr,
      bool evaluateBody);
  bool trySetUserFunctionDefinitionError(
      EvalContext& ctx,
      const std::string& fnName,
      const std::vector<std::string>& fnParams,
      const std::string& fnExpr);
  const char* validateAssignmentTargetName(const std::string& ident) const;
  bool trySetUnknownNameError(const EvalContext& ctx);
  bool tryAppendParsedExpressionStatement(
      EvalContext& ctx,
      std::vector<AstStatement>& out);
  bool tryAppendFunctionDefinitionStatement(
      EvalContext& ctx,
      std::vector<AstStatement>& out);
  bool tryAppendAssignOrExpressionStatement(
      EvalContext& ctx,
      std::vector<AstStatement>& out);
  bool consumeProgramStatementSeparator(EvalContext& ctx);
  bool tryAppendTrailingFormatterSugarStatement(
      EvalContext& ctx,
      std::vector<AstStatement>& out,
      bool& handled);
  static bool hasExprParseFailure(const EvalContext& ctx, const std::unique_ptr<Expr>& node);
#if SMARTMATH_TIME_VALUES
  bool rejectBinaryBuiltinTimeOperands(EvalContext& ctx, const EvalValue& left, const EvalValue& right) const;
  bool rejectNumericBinaryPowWithTime(
      EvalContext& ctx,
      const EvalValue& left,
      const EvalValue& right,
      char op) const;
#endif
  bool rejectInt64BinaryOperands(
      EvalContext& ctx,
      const EvalValue& left,
      const EvalValue& right,
      bool isModulo) const;
  EvalValue builtinMapBinaryTwoArg(
      EvalContext& ctx,
      const std::string& fnName,
      BuiltinFunctionId id,
      const std::vector<EvalValue>& args,
      bool rejectComplexOperands,
      bool numericErrorOnMapFailure) const;
  bool acceptProducedItemCount(EvalContext& ctx, const std::string& fnName, std::uint64_t count) const;

  bool tryCompileSingleExpressionProgram(EvalContext& ctx, std::vector<AstStatement>& out);
  bool parseProgram(EvalContext& ctx, std::vector<AstStatement>& out);
  std::unique_ptr<Expr> parseExpression(EvalContext& ctx);
  std::unique_ptr<Expr> parseOr(EvalContext& ctx);
  std::unique_ptr<Expr> parseAnd(EvalContext& ctx);
  std::unique_ptr<Expr> parseLogicalNot(EvalContext& ctx);
  std::unique_ptr<Expr> parseCompare(EvalContext& ctx);
  std::unique_ptr<Expr> parseLeftAssocBinary(
      EvalContext& ctx,
      std::unique_ptr<Expr> (MathParser::*parseOperand)(EvalContext&),
      bool (MathParser::*tryConsumeOp)(EvalContext&, Expr::BinaryOp&),
      bool setPercentFlag = false);
  bool tryConsumeBitAndOp(EvalContext& ctx, Expr::BinaryOp& outOp);
  bool tryConsumeBitXorOp(EvalContext& ctx, Expr::BinaryOp& outOp);
  bool tryConsumeBitOrOp(EvalContext& ctx, Expr::BinaryOp& outOp);
  bool tryConsumeShiftOp(EvalContext& ctx, Expr::BinaryOp& outOp);
  bool tryConsumeAddSubOp(EvalContext& ctx, Expr::BinaryOp& outOp);
  bool tryConsumeMulDivModOp(EvalContext& ctx, Expr::BinaryOp& outOp);
  bool tryConsumeCompareOp(EvalContext& ctx, Expr::BinaryOp& outOp);
  bool tryConsumeLogicalAndOp(EvalContext& ctx, Expr::BinaryOp& outOp);
  bool tryConsumeLogicalOrOp(EvalContext& ctx, Expr::BinaryOp& outOp);
  std::unique_ptr<Expr> parseBitOr(EvalContext& ctx);
  std::unique_ptr<Expr> parseBitXor(EvalContext& ctx);
  std::unique_ptr<Expr> parseBitAnd(EvalContext& ctx);
  std::unique_ptr<Expr> parseShift(EvalContext& ctx);
  std::unique_ptr<Expr> parseAddSub(EvalContext& ctx);
  std::unique_ptr<Expr> parseMulDivMod(EvalContext& ctx);
  std::unique_ptr<Expr> parsePower(EvalContext& ctx);
  std::unique_ptr<Expr> parseUnary(EvalContext& ctx);
  std::unique_ptr<Expr> parsePrimary(EvalContext& ctx);
#if SMARTMATH_COMPLEX_NUMBERS
  std::unique_ptr<Expr> wrapExprWithTightImagSuffixIfPresent(EvalContext& ctx, std::unique_ptr<Expr> expr);
#endif

  EvalValue evalExpr(const Expr& e, EvalContext& ctx, const std::unordered_map<std::string, EvalValue>* scopedVars);
  EvalValue evalExprScalar(const Expr& e, EvalContext& ctx, const std::unordered_map<std::string, EvalValue>* scopedVars);
  EvalValue evalUnaryExpr(
      const Expr& e,
      EvalContext& ctx,
      const std::unordered_map<std::string, EvalValue>* scopedVars,
      bool scalarOnlyMode);
  EvalValue evalMappedBinaryOp(
      EvalContext& ctx,
      Expr::BinaryOp op,
      const EvalValue& left,
      const EvalValue& right) const;
  EvalValue evalInt64BinaryOp(
      EvalContext& ctx,
      const EvalValue& left,
      const EvalValue& right,
      Expr::BinaryOp op) const;
  bool exprIsScalarOnly(const Expr& e) const;
  bool programIsScalarOnly(const std::vector<AstStatement>& program) const;
  EvalValue runCompiledProgram(
      EvalContext& ctx,
      const std::vector<AstStatement>& program,
      const std::unordered_map<std::string, EvalValue>* scopedVars,
      bool scalarOnlyMode = false);

  EvalValue evalFunctionCall(
      EvalContext& ctx,
      const std::string& fnName,
      std::vector<EvalValue>&& args,
      BuiltinFunctionId preboundId,
      const std::unordered_map<std::string, EvalValue>* scopedVars);

  EvalValue builtinUnpack(EvalContext& ctx, const std::vector<EvalValue>& args) const;
  EvalValue builtinLen(EvalContext& ctx, const std::vector<EvalValue>& args) const;
  EvalValue builtinRepeat(EvalContext& ctx, const std::string& fnName, const std::vector<EvalValue>& args) const;
  EvalValue builtinRange(EvalContext& ctx, const std::string& fnName, const std::vector<EvalValue>& args) const;
  EvalValue builtinAggregateFamily(
      EvalContext& ctx,
      const std::string& fnName,
      BuiltinFunctionId id,
      const std::vector<EvalValue>& args) const;
  std::size_t countAggregateArgScalarItems(const std::vector<EvalValue>& args) const;
  EvalValue aggregateFoldByMode(
      EvalContext& ctx,
      const std::string& fnName,
      BuiltinFunctionId id,
      const std::vector<EvalValue>& args,
      AggregateFoldMode mode) const;
  EvalValue builtinSortFamily(
      EvalContext& ctx,
      const std::string& fnName,
      BuiltinFunctionId id,
      const std::vector<EvalValue>& args) const;
  bool prepareUnaryCallableScalars(
      EvalContext& ctx,
      const std::string& fnName,
      const EvalValue& dataArg,
      const EvalValue& fnArg,
      const char* expectsOneFunction,
      const char* expectsUnaryFunction,
      std::vector<EvalValue::ScalarValue>& vals,
      bool& isLambda,
      bool& emptyArray) const;
  EvalValue evalUnaryCallableOnScalar(
      EvalContext& ctx,
      const EvalValue& fnArg,
      bool isLambda,
      const EvalValue::ScalarValue& elem,
      const std::unordered_map<std::string, EvalValue>* scopedVars,
      const char* expectsUnaryFunction);
  EvalValue builtinSortby(
      EvalContext& ctx,
      const std::string& fnName,
      const std::vector<EvalValue>& args,
      const std::unordered_map<std::string, EvalValue>* scopedVars);
  EvalValue builtinRatio(EvalContext& ctx, const std::string& fnName, const std::vector<EvalValue>& args) const;
  EvalValue invokeUnaryFunctionOnElement(
      EvalContext& ctx,
      const std::string& funcName,
      const EvalValue::ScalarValue& elem,
      const std::unordered_map<std::string, EvalValue>* scopedVars,
      const char* expectsUnaryFunction);
  bool isEligibleUnaryFunctionName(
      const std::string& funcName,
      std::string& outErr,
      const char* expectsUnaryFunction) const;
  EvalValue builtinFilter(
      EvalContext& ctx,
      const std::string& fnName,
      const std::vector<EvalValue>& args,
      const std::unordered_map<std::string, EvalValue>* scopedVars);
  bool tryLexicographicCompareEvalValues(
      EvalContext& ctx,
      const EvalValue& a,
      const EvalValue& b,
      int& cmpOut) const;
  bool sortbyStableSortIndicesFromKeys(
      EvalContext& ctx,
      const std::vector<EvalValue>& sortKeys,
      std::vector<int>& orderIdx) const;

  bool tryBuiltinRatioScalar(EvalContext& ctx, const EvalValue::ScalarValue& sv, EvalValue& outV) const;
  EvalValue builtinTimeUnit(
      EvalContext& ctx,
      const std::string& fnName,
      BuiltinFunctionId id,
      const std::vector<EvalValue>& args) const;
  EvalValue builtinBaseFormat(
      EvalContext& ctx,
      const std::string& fnName,
      BuiltinFunctionId id,
      const std::vector<EvalValue>& args) const;
  EvalValue builtinPow(EvalContext& ctx, const std::vector<EvalValue>& args) const;
  EvalValue builtinScalarBinaryFamily(
      EvalContext& ctx,
      const std::string& fnName,
      BuiltinFunctionId id,
      const std::vector<EvalValue>& args) const;
  EvalValue builtinRand(EvalContext& ctx, const std::vector<EvalValue>& args) const;
  EvalValue builtinModCall(EvalContext& ctx, const std::vector<EvalValue>& args) const;
  EvalValue builtinFactorial(
      EvalContext& ctx,
      const std::string& fnName,
      const std::vector<EvalValue>& args) const;
#if SMARTMATH_FACTORINT
  EvalValue builtinFactorint(
      EvalContext& ctx,
      const std::string& fnName,
      const std::vector<EvalValue>& args) const;
  bool tryGetFactorintInput(const EvalValue& v, bool& isNegative, std::uint64_t& absU) const;
#endif
  EvalValue builtinDegRad(
      EvalContext& ctx,
      const std::string& fnName,
      BuiltinFunctionId id,
      const std::vector<EvalValue>& args) const;
  EvalValue builtinPolarCart(
      EvalContext& ctx,
      const std::string& fnName,
      BuiltinFunctionId id,
      const std::vector<EvalValue>& args) const;
#if SMARTMATH_COMPLEX_NUMBERS
  bool applyUnaryComplexMathOverlay(
      BuiltinFunctionId id,
      const EvalValue::ScalarValue& sv,
      bool& trigOutOfRange,
      bool& realUnaryOk,
      EvalValue& outV) const;
#endif
  EvalValue applyUnaryBuiltin(
      BuiltinFunctionId id,
      const EvalValue::ScalarValue& sv,
      bool& trigOutOfRange,
      bool& realUnaryOk) const;
  EvalValue builtinUnaryMath(
      EvalContext& ctx,
      const std::string& fnName,
      BuiltinFunctionId id,
      const std::vector<EvalValue>& args) const;
  EvalValue builtinLog(EvalContext& ctx, const std::vector<EvalValue>& args) const;
  EvalValue evalUserFunctionCall(
      EvalContext& ctx,
      const std::string& fnName,
      const std::vector<EvalValue>& args,
      const std::unordered_map<std::string, EvalValue>* scopedVars);
  void resetCompileState();
  void resetEvaluateState();
  bool prepareEvaluate(EvalContext& ctx);
  bool finalizeEvaluate(EvalContext& ctx, EvalValue&& out);

  bool parseFunctionDefinition(
      EvalContext& ctx,
      std::string& outName,
      std::vector<std::string>& outParams,
      std::string& outExpr) const;

  void setVariable(const std::string& name, const EvalValue& value);
  void removeVariableByName(const std::string& name);
  void removeUserFunctionByName(const std::string& name);
  void normalizeCallArgs(std::vector<EvalValue>& args);
  static void bindExprVariableRefs(Expr& e, bool hasAssignments, const std::unordered_map<std::string, EvalValue>& variables);
  void bindCompiledVariableRefs(CompiledProgramState& state);

#if SMARTMATH_COMPLEX_NUMBERS
  bool tryApplyComplexBinaryScalars(const EvalValue::ScalarValue& lv, const EvalValue::ScalarValue& rv, char op, EvalValue& outS) const;
#endif
#if SMARTMATH_TIME_VALUES
  bool scalarValueIsTime(const EvalValue::ScalarValue& s) const;
  bool evalValueInvolvesTime(const EvalValue& v) const;
#endif
  enum class CmpScalarIncompatiblePolicy { SetError, SortUniqueReturnOne, SortLess, SortGreater };
  bool cmpScalarValuesForCompare(
      EvalContext* ctx,
      const EvalValue::ScalarValue& sa,
      const EvalValue::ScalarValue& sb,
      int& cmpOut,
      CmpScalarIncompatiblePolicy policy) const;
#if SMARTMATH_COMPLEX_NUMBERS
  bool rejectBuiltinArgsWithComplexImaginary(EvalContext& ctx, const std::vector<EvalValue>& args) const;
#endif
#if SMARTMATH_TIME_VALUES
  bool scalarMsForCompare(const EvalValue::ScalarValue& sv, long long& outMs) const;
  EvalValue mapTimeUnitOverArray(
      EvalContext& ctx,
      const std::string& fnName,
      BuiltinFunctionId id,
      const EvalValue& inV) const;
  EvalValue evalValueFromTimeMs(BuiltinFunctionId id, long long ms) const;
#endif
  bool tryUnaryComplexBuiltinSupport(BuiltinFunctionId id, const EvalValue::ScalarValue& sv, EvalValue& out) const;
  /** Unary minus with exact int/uint preservation; LLONG_MIN -> double. */

#if SMARTMATH_TIME_VALUES
  bool tryApplyTimeBinaryScalars(
      EvalContext& ctx,
      const EvalValue::ScalarValue& lv,
      const EvalValue::ScalarValue& rv,
      char op,
      EvalValue& outS) const;
#endif
  bool tryCombineBinaryScalars(
      EvalContext& ctx,
      char op,
      const EvalValue::ScalarValue& lv,
      const EvalValue::ScalarValue& rv,
      EvalValue& outS) const;
  EvalValue mapBinary(EvalContext& ctx, const EvalValue& a, const EvalValue& b, char op, bool& ok) const;
  UserFunction* findUserFunction(const std::string& fnName);
  const UserFunction* findUserFunction(const std::string& fnName) const;
  void upsertUserFunction(UserFunction uf);

};

namespace smartmath_internal {

MathParser::RawResult::CartesianScalar rawCartesianAssignReducedRational(long long num, std::uint64_t den);
MathParser::RawResult::CartesianScalar toRawCartesianScalar(const EvalValue::ScalarValue& v, bool imagPart);
MathParser::RawResult::Scalar toRawScalar(const EvalValue::ScalarValue& v);
MathParser::RawResult toRawResult(const EvalValue& v);

}  // namespace smartmath_internal
