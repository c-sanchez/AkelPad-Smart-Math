#include "MathParser.hpp"

#include <cctype>
#include <cmath>
#include <cstdlib>
#include <algorithm>
#include <limits>
#include <functional>
#include <iostream>
#include <iomanip>
#include <sstream>
#include <string>
#include <vector>
#include <unordered_map>
#include <unordered_set>

namespace {

struct TestState {
  int total = 0;
  int passed = 0;
  int failed = 0;
};

struct TestCase {
  std::string name;
  std::function<bool(std::string&)> run;
};

// Dedup support for expression-evaluation smoke cases.
// - runSuite() sets g_seen to the suite-local "seen" set.
// - expectEval()/expectEvalErrorContains() compute a signature and flip g_skipDuplicate on repeats.
static std::unordered_set<std::string>* g_seen = nullptr;
static bool g_skipDuplicate = false;
static std::size_t g_duplicateCount = 0;
static std::unordered_map<const MathParser*, std::vector<std::string>> g_parserConstDedupState;

std::string formatTrackedDouble(const double value) {
  if (std::isnan(value)) {
    return "NaN";
  }
  if (std::isinf(value)) {
    return (value < 0.0) ? "-Inf" : "+Inf";
  }
  std::ostringstream oss;
  oss << std::setprecision(17) << value;
  return oss.str();
}

void trackConstForDedup(const MathParser& p, const std::string& name, const std::string& valueSig) {
  g_parserConstDedupState[&p].push_back(name + "=" + valueSig);
}

void addConstTracked(MathParser& p, const std::string& name, const long long value) {
  p.addConst(name, value);
  trackConstForDedup(p, name, "i:" + std::to_string(value));
}

void addConstTracked(MathParser& p, const std::string& name, const double value) {
  p.addConst(name, value);
  trackConstForDedup(p, name, "d:" + formatTrackedDouble(value));
}

std::string getParserConstDedupSignature(const MathParser& p) {
  const auto it = g_parserConstDedupState.find(&p);
  if (it == g_parserConstDedupState.end() || it->second.empty()) {
    return "";
  }
  std::vector<std::string> items = it->second;
  std::sort(items.begin(), items.end());
  std::ostringstream oss;
  for (std::size_t i = 0; i < items.size(); ++i) {
    if (i != 0U) {
      oss << ";";
    }
    oss << items[i];
  }
  return oss.str();
}

void clearParserConstDedupState() {
  g_parserConstDedupState.clear();
}

bool parseQuotedBasicString(const std::string& s, std::size_t quotePos, std::string& out, std::size_t& endPos) {
  if (quotePos >= s.size() || s[quotePos] != '"') {
    return false;
  }
  out.clear();
  std::size_t i = quotePos + 1;
  while (i < s.size()) {
    const char c = s[i];
    if (c == '"') {
      if ((i + 1) < s.size() && s[i + 1] == '"') {
        out.push_back('"');
        i += 2;
        continue;
      }
      endPos = i + 1;
      return true;
    }
    out.push_back(c);
    ++i;
  }
  return false;
}

std::string trimSmokeToken(std::string s) {
  while (!s.empty() && std::isspace(static_cast<unsigned char>(s.back()))) {
    s.pop_back();
  }
  std::size_t i = 0;
  while (i < s.size() && std::isspace(static_cast<unsigned char>(s[i]))) {
    ++i;
  }
  return s.substr(i);
}

bool tryParseFullDouble(const std::string& s, double& out) {
  const char* p = s.c_str();
  char* e = nullptr;
  out = std::strtod(p, &e);
  if (e == p) {
    return false;
  }
  while (*e != '\0' && std::isspace(static_cast<unsigned char>(*e))) {
    ++e;
  }
  return *e == '\0';
}

namespace {

std::string smokeLowerAscii(std::string s) {
  for (char& c : s) {
    c = static_cast<char>(std::tolower(static_cast<unsigned char>(c)));
  }
  return s;
}

int smokeInfSignFromLiteral(const std::string& s) {
  const std::string t = smokeLowerAscii(trimSmokeToken(s));
  if (t == "inf" || t == "+inf") {
    return 1;
  }
  if (t == "-inf") {
    return -1;
  }
  return 0;
}

bool smokeIsExactZeroLiteral(const std::string& s) {
  const std::string t = trimSmokeToken(s);
  return t == "0" || t == "0.0" || t == "-0" || t == "-0.0" || t == "+0" || t == "+0.0";
}

bool smokeIsExactOneLiteral(const std::string& s) {
  const std::string t = trimSmokeToken(s);
  return t == "1" || t == "-1" || t == "1.0" || t == "-1.0" || t == "+1" || t == "+1.0";
}

int smokeOneSignFromLiteral(const std::string& s) {
  const std::string t = trimSmokeToken(s);
  if (t == "1" || t == "1.0" || t == "+1" || t == "+1.0") {
    return 1;
  }
  if (t == "-1" || t == "-1.0") {
    return -1;
  }
  return 0;
}

bool smokeExpectedRequiresStrictScalarMatch(const std::string& expected) {
  return smokeIsExactZeroLiteral(expected) || smokeIsExactOneLiteral(expected) ||
         smokeInfSignFromLiteral(expected) != 0;
}

bool smokeStrictScalarMatch(const std::string& actual, const std::string& expected) {
  if (smokeIsExactZeroLiteral(expected)) {
    if (!smokeIsExactZeroLiteral(actual)) {
      return false;
    }
    double da = 0.0;
    double de = 0.0;
    if (!tryParseFullDouble(actual, da) || !tryParseFullDouble(expected, de)) {
      return false;
    }
    return da == 0.0 && de == 0.0;
  }
  if (smokeIsExactOneLiteral(expected)) {
    if (!smokeIsExactOneLiteral(actual)) {
      return false;
    }
    if (smokeOneSignFromLiteral(actual) != smokeOneSignFromLiteral(expected)) {
      return false;
    }
    double da = 0.0;
    double de = 0.0;
    if (!tryParseFullDouble(actual, da) || !tryParseFullDouble(expected, de)) {
      return false;
    }
    if (smokeOneSignFromLiteral(expected) > 0) {
      return da == 1.0 && de == 1.0;
    }
    return da == -1.0 && de == -1.0;
  }
  const int eInf = smokeInfSignFromLiteral(expected);
  if (eInf != 0) {
    return smokeInfSignFromLiteral(actual) == eInf;
  }
  return false;
}

}  // namespace

std::vector<std::string> splitTopLevelCsvInParens(const std::string& s) {
  std::vector<std::string> parts;
  if (s.size() < 2U || s.front() != '(' || s.back() != ')') {
    return parts;
  }
  std::size_t start = 1;
  for (std::size_t i = 1; i + 1U < s.size(); ++i) {
    if (s[i] == ',') {
      parts.push_back(s.substr(start, i - start));
      start = i + 1;
    }
  }
  parts.push_back(s.substr(start, s.size() - 1U - start));
  return parts;
}

bool smokeScalarCloseEnough(const std::string& actual, const std::string& expected) {
  const std::string ea = trimSmokeToken(actual);
  const std::string ee = trimSmokeToken(expected);
  if (ea == ee) {
    return true;
  }
  if (smokeExpectedRequiresStrictScalarMatch(ee)) {
    return smokeStrictScalarMatch(ea, ee);
  }
  double da = 0.0;
  double de = 0.0;
  if (!tryParseFullDouble(ea, da) || !tryParseFullDouble(ee, de)) {
    return false;
  }
  const double p63 = std::ldexp(1.0, 63);
  if (da == p63 || da == -p63 || de == p63 || de == -p63) {
    return false;
  }
  if (da == de) {
    return true;
  }
  if (std::nextafter(da, de) == de || std::nextafter(de, da) == da) {
    return true;
  }
  const double scale = std::max(1.0, std::max(std::fabs(da), std::fabs(de)));
  const double tol = 16.0 * std::numeric_limits<double>::epsilon() * scale;
  return std::fabs(da - de) <= tol;
}

namespace {

bool splitComplexText(const std::string& s, std::string& rePart, std::string& imPart) {
  std::string t = trimSmokeToken(s);
  if (t.empty()) {
    rePart = "0";
    imPart = "0";
    return true;
  }
  const char last = static_cast<char>(std::tolower(static_cast<unsigned char>(t.back())));
  if (last != 'i') {
    return false;
  }
  std::size_t splitAt = 0;
  for (std::size_t i = t.size(); i >= 2; --i) {
    const char ch = t[i - 1];
    if (ch == '+' || ch == '-') {
      if (i > 1) {
        const char prevCh = static_cast<char>(std::tolower(static_cast<unsigned char>(t[i - 2])));
        if (prevCh != 'e') {
          splitAt = i - 1;
          break;
        }
      } else if (ch == '-') {
        splitAt = i - 1;
        break;
      }
    }
  }
  if (splitAt == 0) {
    rePart = "0";
    imPart = t.substr(0, t.size() - 1);
    return true;
  }
  rePart = trimSmokeToken(t.substr(0, splitAt));
  if (rePart.empty()) {
    rePart = "0";
  }
  imPart = trimSmokeToken(t.substr(splitAt));
  if (!imPart.empty() && imPart.front() == '+') {
    imPart.erase(imPart.begin());
  }
  if (!imPart.empty()) {
    const char imLast = static_cast<char>(std::tolower(static_cast<unsigned char>(imPart.back())));
    if (imLast == 'i') {
      imPart.pop_back();
    }
  }
  return true;
}

bool complexResultCloseEnough(const std::string& actual, const std::string& expected) {
  std::string aRe;
  std::string aIm;
  std::string eRe;
  std::string eIm;
  if (!splitComplexText(actual, aRe, aIm) || !splitComplexText(expected, eRe, eIm)) {
    return false;
  }
  return smokeScalarCloseEnough(aRe, eRe) && smokeScalarCloseEnough(aIm, eIm);
}

}  // namespace

bool smokeResultCloseEnough(const std::string& actual, const std::string& expected) {
  if (actual == expected) {
    return true;
  }
  if (actual.size() >= 2U && expected.size() >= 2U && actual.front() == '(' && expected.front() == '(' &&
      actual.back() == ')' && expected.back() == ')') {
    const std::vector<std::string> pa = splitTopLevelCsvInParens(actual);
    const std::vector<std::string> pe = splitTopLevelCsvInParens(expected);
    if (pa.size() != pe.size()) {
      return false;
    }
    for (std::size_t i = 0; i < pa.size(); ++i) {
      if (!smokeScalarCloseEnough(trimSmokeToken(pa[i]), trimSmokeToken(pe[i]))) {
        return false;
      }
    }
    return true;
  }
  const auto looksLikeComplexLiteral = [](const std::string& s) {
    const std::string t = trimSmokeToken(s);
    if (t.empty()) {
      return false;
    }
    const char last = static_cast<char>(std::tolower(static_cast<unsigned char>(t.back())));
    return last == 'i';
  };
  if (looksLikeComplexLiteral(actual) || looksLikeComplexLiteral(expected)) {
    return complexResultCloseEnough(actual, expected);
  }
  return smokeScalarCloseEnough(actual, expected);
}

bool expectEval(
    MathParser& p,
    const std::string& expr,
    const std::string& expected,
    std::string& why) {
  if (g_seen) {
    const std::string sig = "expected|" + expr + "|" + expected + "|consts|" + getParserConstDedupSignature(p);
    if (g_seen->find(sig) != g_seen->end()) {
      g_skipDuplicate = true;
      ++g_duplicateCount;
      return true;
    }
    g_seen->insert(sig);
  }
  p.parseAndEvaluate(expr);
  if (!p.getError().empty()) {
    why = "unexpected error: " + p.getError();
    return false;
  }
  const std::string actual = p.getResult();
  if (!smokeResultCloseEnough(actual, expected)) {
    why = "expected \"" + expected + "\", got \"" + actual + "\"";
    return false;
  }
  return true;
}

bool errorContainsExpectedPart(const std::string& err, const std::string& part) {
  if (part == "unexpected token") {
    return err.find("unexpected token") != std::string::npos ||
           err.find("unexpected characters") != std::string::npos;
  }
  return err.find(part) != std::string::npos;
}

bool expectEvalErrorContains(
    MathParser& p,
    const std::string& expr,
    const std::string& expectedErrPart,
    std::string& why) {
  if (g_seen) {
    const std::string sig =
        "errorContains|" + expr + "|" + expectedErrPart + "|consts|" + getParserConstDedupSignature(p);
    if (g_seen->find(sig) != g_seen->end()) {
      g_skipDuplicate = true;
      ++g_duplicateCount;
      return true;
    }
    g_seen->insert(sig);
  }
  p.parseAndEvaluate(expr);
  const std::string err = p.getError();
  if (err.empty()) {
    why = "expected error containing \"" + expectedErrPart + "\", got success";
    return false;
  }
  if (!errorContainsExpectedPart(err, expectedErrPart)) {
    why = "expected error containing \"" + expectedErrPart + "\", got \"" + err + "\"";
    return false;
  }
  return true;
}

std::vector<TestCase> buildUnitCases() {
  std::vector<TestCase> t;

  // C++ API unit tests.
  t.push_back({"unit/addConst int overload", [](std::string& why) {
                 MathParser p;
                 addConstTracked(p, "A", 42LL);
                 return expectEval(p, "a+1", "43", why);
               }});
  t.push_back({"unit/addConst double overload", [](std::string& why) {
                 MathParser p;
                 addConstTracked(p, "rate", 1.5);
                 return expectEval(p, "rate*8", "12", why);
               }});
  t.push_back({"unit/addUserFunction api", [](std::string& why) {
                 MathParser p;
                 const std::string err = p.addUserFunction("mul3(x)=x*3");
                 if (!err.empty()) {
                   why = "unexpected addUserFunction error: " + err;
                   return false;
                 }
                 return expectEval(p, "mul3(7)", "21", why);
               }});
  t.push_back({"unit/addUserFunction invalid", [](std::string& why) {
                 MathParser p;
                 const std::string err = p.addUserFunction("bad(x)");
                 if (err.empty()) {
                   why = "expected addUserFunction error";
                   return false;
                 }
                 return true;
               }});
  t.push_back({"unit/addUserFunction rejects reserved function name", [](std::string& why) {
                 MathParser p;
                 const std::string err = p.addUserFunction("hex(x)=x");
                 if (err != "reserved function name") {
                   why = "expected reserved function name, got: " + err;
                   return false;
                 }
                 return true;
               }});
  t.push_back({"unit/addUserFunction rejects reserved constant name", [](std::string& why) {
                 MathParser p;
                 const std::string err = p.addUserFunction("e(x)=x");
                 if (err != "reserved constant name") {
                   why = "expected reserved constant name, got: " + err;
                   return false;
                 }
                 return true;
               }});
  t.push_back({"unit/addUserFunction rejects ans as function name", [](std::string& why) {
                 MathParser p;
                 const std::string err = p.addUserFunction("ans(x)=x*x");
                 if (err != "reserved built-in variable name") {
                   why = "expected reserved built-in variable name, got: " + err;
                   return false;
                 }
                 return true;
               }});
  t.push_back({"unit/addUserFunction rejects formal probe as function name", [](std::string& why) {
                 MathParser p;
                 const std::string err = p.addUserFunction("_(x)=x*x");
                 if (err != "reserved built-in variable name") {
                   why = "expected reserved built-in variable name, got: " + err;
                   return false;
                 }
                 return true;
               }});
  t.push_back({"unit/addUserFunction rejects trailing junk in body", [](std::string& why) {
                 MathParser p;
                 const std::string err = p.addUserFunction("bad1(x)=x)");
                 if (err != "unexpected characters") {
                   why = "expected unexpected characters, got: " + err;
                   return false;
                 }
                 return true;
               }});
  t.push_back({"unit/addUserFunction rejects invalid body expression", [](std::string& why) {
                 MathParser p;
                 const std::string err = p.addUserFunction("bad2(x)=x$$5");
                 if (err != "unexpected characters") {
                   why = "expected unexpected characters, got: " + err;
                   return false;
                 }
                 return true;
               }});
  t.push_back({"unit/format getResultAsHexOctBinDec", [](std::string& why) {
                 MathParser p;
                 p.parseAndEvaluate("255");
                 if (!p.getError().empty()) {
                   why = p.getError();
                   return false;
                 }
                 if (p.getResultAsHex() != "0xFF") {
                   why = "hex mismatch: " + p.getResultAsHex();
                   return false;
                 }
                 if (p.getResultAsOct() != "0o377") {
                   why = "oct mismatch: " + p.getResultAsOct();
                   return false;
                 }
                 if (p.getResultAsBin() != "0b11111111") {
                   why = "bin mismatch: " + p.getResultAsBin();
                   return false;
                 }
                 if (p.getResultAsDec() != "255") {
                   why = "dec mismatch: " + p.getResultAsDec();
                   return false;
                 }
                 return true;
               }});
  t.push_back({"unit/array formatting in bases", [](std::string& why) {
                 MathParser p;
                 p.parseAndEvaluate("(1,2,5)");
                 if (!p.getError().empty()) {
                   why = p.getError();
                   return false;
                 }
                 if (p.getResultAsHex() != "(0x1,0x2,0x5)") {
                   why = "hex array mismatch: " + p.getResultAsHex();
                   return false;
                 }
                 if (p.getResultAsBin() != "(0b1,0b10,0b101)") {
                   why = "bin array mismatch: " + p.getResultAsBin();
                   return false;
                 }
                 return true;
               }});
  t.push_back({"unit/error clears result text", [](std::string& why) {
                 MathParser p;
                 p.parseAndEvaluate("2+3");
                 if (p.getResult() != "5") {
                   why = "setup failed";
                   return false;
                 }
                 p.parseAndEvaluate("unknownVar");
                 if (p.getError().empty()) {
                   why = "expected error";
                   return false;
                 }
                 if (!p.getResult().empty()) {
                   why = "result should be empty on error";
                   return false;
                 }
                 return true;
               }});
  t.push_back({"unit/getRawResult scalar int/float", [](std::string& why) {
                 MathParser p;
                 p.parseAndEvaluate("2+3");
                 if (!p.getError().empty()) {
                   why = p.getError();
                   return false;
                 }
                 MathParser::RawResult r = p.getRawResult();
                 if (!r.isScalar() || !r.scalar.real.isInt64() || r.scalar.real.intValue != 5LL) {
                   why = "expected scalar int64=5";
                   return false;
                 }

                 p.parseAndEvaluate("1/2");
                 if (!p.getError().empty()) {
                   why = p.getError();
                   return false;
                 }
                 r = p.getRawResult();
                 if (!r.isScalar() || !r.scalar.real.isFloatingPoint() ||
                     std::fabs(r.scalar.real.floatingPoint - 0.5) > 1e-12) {
                   why = "expected scalar float=0.5";
                   return false;
                 }
                 return true;
               }});
#if SMARTMATH_TIME_VALUES
  t.push_back({"unit/getRawResult scalar time (ms in intValue, parity with Basic RSK_TIME)",
               [](std::string& why) {
                 MathParser p;
                 p.setSupportTimeValues(true);
                 p.parseAndEvaluate("1:30");
                 if (!p.getError().empty()) {
                   why = p.getError();
                   return false;
                 }
                 MathParser::RawResult r = p.getRawResult();
                 if (!r.isScalar() || !r.scalar.isTime() || r.scalar.real.intValue != 90000LL) {
                   why = "expected scalar time 1:30 -> 90000 ms";
                   return false;
                 }

                 p.parseAndEvaluate("1h");
                 if (!p.getError().empty()) {
                   why = p.getError();
                   return false;
                 }
                 r = p.getRawResult();
                 if (!r.isScalar() || !r.scalar.real.isTime() || r.scalar.real.intValue != 3600000LL) {
                   why = "expected scalar time 1h -> 3600000 ms";
                   return false;
                 }

                 p.parseAndEvaluate("milliseconds((0:01:00,0:02:00))");
                 if (!p.getError().empty()) {
                   why = p.getError();
                   return false;
                 }
                 r = p.getRawResult();
                 if (!r.isArray() || r.array.size() != 2 || !r.array[0].real.isInt64() || r.array[0].real.intValue != 60000LL ||
                     !r.array[1].real.isInt64() || r.array[1].real.intValue != 120000LL) {
                   why = "expected time array converted to int64 ms via milliseconds()";
                   return false;
                 }
                 return true;
               }});
#endif
  t.push_back({"unit/getRawResult array and error state", [](std::string& why) {
                 MathParser p;
                 p.parseAndEvaluate("(1,2.5,3)");
                 if (!p.getError().empty()) {
                   why = p.getError();
                   return false;
                 }
                 MathParser::RawResult r = p.getRawResult();
                 if (!r.isArray() || r.array.size() != 3) {
                   why = "expected array[3]";
                   return false;
                 }
                 if (!r.array[0].real.isInt64() || r.array[0].real.intValue != 1LL) {
                   why = "array[0] mismatch";
                   return false;
                 }
                 if (!r.array[1].real.isFloatingPoint() ||
                     std::fabs(r.array[1].real.floatingPoint - 2.5) > 1e-12) {
                   why = "array[1] mismatch";
                   return false;
                 }
                 if (!r.array[2].real.isInt64() || r.array[2].real.intValue != 3LL) {
                   why = "array[2] mismatch";
                   return false;
                 }

                 p.parseAndEvaluate("unknownVar");
                 r = p.getRawResult();
                 if (r.hasValue()) {
                   why = "expected no raw value on error";
                   return false;
                 }
                 return true;
               }});
  t.push_back({"unit/getRawResult rational from ratio()",
               [](std::string& why) {
                 MathParser p;
                 p.parseAndEvaluate("ratio(0.5)");
                 if (!p.getError().empty()) {
                   why = p.getError();
                   return false;
                 }
                 MathParser::RawResult r = p.getRawResult();
                 if (!r.isScalar() || !r.scalar.isRational() || r.scalar.real.rational.numerator != 1LL ||
                     r.scalar.real.rational.denominator != 2ULL) {
                   why = "expected scalar rational 1/2";
                   return false;
                 }

                 p.parseAndEvaluate("ratio(sqrt(2))");
                 if (!p.getError().empty()) {
                   why = p.getError();
                   return false;
                 }
                 r = p.getRawResult();
                 if (!r.isScalar() || !r.scalar.isRational() || r.scalar.real.rational.numerator != 13250218LL ||
                     r.scalar.real.rational.denominator != 9369319ULL) {
                   why = "expected scalar rational sqrt(2) convergent";
                   return false;
                 }

                 p.parseAndEvaluate("ratio((0.5, 1/3))");
                 if (!p.getError().empty()) {
                   why = p.getError();
                   return false;
                 }
                 r = p.getRawResult();
                 if (!r.isArray() || r.array.size() != 2U) {
                   why = "expected rational array[2]";
                   return false;
                 }
                 if (!r.array[0].isRational() || r.array[0].real.rational.numerator != 1LL ||
                     r.array[0].real.rational.denominator != 2ULL) {
                   why = "array[0] expected 1/2";
                   return false;
                 }
                 if (!r.array[1].isRational() || r.array[1].real.rational.numerator != 1LL ||
                     r.array[1].real.rational.denominator != 3ULL) {
                   why = "array[1] expected 1/3";
                   return false;
                 }
                 return true;
               }});
  t.push_back({"unit/parseAndEvaluateRaw returns raw and empty on error", [](std::string& why) {
                 MathParser p;
                 MathParser::RawResult r = p.parseAndEvaluateRaw("10+20");
                 if (!p.getError().empty()) {
                   why = p.getError();
                   return false;
                 }
                 if (!r.isScalar() || !r.scalar.real.isInt64() || r.scalar.real.intValue != 30LL) {
                   why = "expected scalar int64=30";
                   return false;
                 }

                 r = p.parseAndEvaluateRaw("unknownVar");
                 if (!r.hasValue()) {
                   return true;
                 }
                 why = "expected empty raw result on error";
                 return false;
               }});

  // --- Scalar-vs-array ambiguity locks (parentheses, logical ops, bitwise ops) ---
  t.push_back({"unit/grouping parentheses are scalar: ((5))", [](std::string& why) {
                 MathParser p;
                 return expectEval(p, "((5))", "5", why);
               }});

  t.push_back({"unit/logical NOT precedence vs equality: !2==1", [](std::string& why) {
                 MathParser p;
                 return expectEval(p, "!2==1", "0", why);
               }});
  t.push_back({"unit/logical NOT precedence vs equality: not 2==1", [](std::string& why) {
                 MathParser p;
                 return expectEval(p, "not 2==1", "1", why);
               }});

  t.push_back({"unit/logical NOT with non-empty array is truthy", [](std::string& why) {
                 MathParser p;
                 return expectEval(p, "!(0,0)", "0", why);
               }});
  t.push_back({"unit/logical NOT with scalar is normal", [](std::string& why) {
                 MathParser p;
                 return expectEval(p, "!0", "1", why);
               }});

  t.push_back({"unit/logical NOT scalar: !5", [](std::string& why) {
                 MathParser p;
                 return expectEval(p, "!5", "0", why);
               }});

  t.push_back({"unit/logical NOT keyword scalar: not 0", [](std::string& why) {
                 MathParser p;
                 return expectEval(p, "not 0", "1", why);
               }});

  t.push_back({"unit/logical NOT keyword scalar: not 2", [](std::string& why) {
                 MathParser p;
                 return expectEval(p, "not 2", "0", why);
               }});

  t.push_back({"unit/logical NOT with grouped scalar: !(0)", [](std::string& why) {
                 MathParser p;
                 return expectEval(p, "!(0)", "1", why);
               }});

  t.push_back({"unit/logical NOT keyword with non-empty array: NOT (0,0)", [](std::string& why) {
                 MathParser p;
                 return expectEval(p, "NOT (0,0)", "0", why);
               }});

  t.push_back({"unit/logical AND with scalar falsy left side", [](std::string& why) {
                 MathParser p;
                 return expectEval(p, "0&&(0,0)", "0", why);
               }});

  t.push_back({"unit/logical AND with grouped scalar: ((0))&&1", [](std::string& why) {
                 MathParser p;
                 return expectEval(p, "((0))&&1", "0", why);
               }});

  t.push_back({"unit/logical AND with non-empty array operand is truthy", [](std::string& why) {
                 MathParser p;
                 return expectEval(p, "(0,0)&&1", "1", why);
               }});

  t.push_back({"unit/logical OR with non-empty array operand is truthy", [](std::string& why) {
                 MathParser p;
                 return expectEval(p, "0||(0,0)", "1", why);
               }});

  t.push_back({"unit/logical AND scalar: 1&&1", [](std::string& why) {
                 MathParser p;
                 return expectEval(p, "1&&1", "1", why);
               }});
  t.push_back({"unit/logical AND scalar: 1&&0", [](std::string& why) {
                 MathParser p;
                 return expectEval(p, "1&&0", "0", why);
               }});
  t.push_back({"unit/logical OR scalar: 1||0", [](std::string& why) {
                 MathParser p;
                 return expectEval(p, "1||0", "1", why);
               }});
  t.push_back({"unit/logical OR scalar: 0||0", [](std::string& why) {
                 MathParser p;
                 return expectEval(p, "0||0", "0", why);
               }});

  t.push_back({"unit/logical AND scalar aliases: 1 and 1", [](std::string& why) {
                 MathParser p;
                 return expectEval(p, "1 and 1", "1", why);
               }});
  t.push_back({"unit/logical AND scalar aliases: 1 and 0", [](std::string& why) {
                 MathParser p;
                 return expectEval(p, "1 and 0", "0", why);
               }});
  t.push_back({"unit/logical OR scalar aliases: 1 or 0", [](std::string& why) {
                 MathParser p;
                 return expectEval(p, "1 or 0", "1", why);
               }});
  t.push_back({"unit/logical OR scalar aliases: 0 or 0", [](std::string& why) {
                 MathParser p;
                 return expectEval(p, "0 or 0", "0", why);
               }});

  t.push_back({"unit/logical precedence: 1 || 0 && 0", [](std::string& why) {
                 MathParser p;
                 return expectEval(p, "1 || 0 && 0", "1", why);
               }});

  t.push_back({"unit/logical NOT keyword with non-empty array is truthy", [](std::string& why) {
                 MathParser p;
                 return expectEval(p, "not (0,0)", "0", why);
               }});

  t.push_back({"unit/logical precedence: not 1<0", [](std::string& why) {
                 MathParser p;
                 return expectEval(p, "not 1<0", "1", why);
               }});
  t.push_back({"unit/logical precedence: !1<0", [](std::string& why) {
                 MathParser p;
                 return expectEval(p, "!1<0", "0", why);
               }});
  t.push_back({"unit/logical precedence: !1=0", [](std::string& why) {
                 MathParser p;
                 return expectEval(p, "!1=0", "1", why);
               }});

  t.push_back({"unit/bitwise &: grouping scalar stays scalar", [](std::string& why) {
                 MathParser p;
                 return expectEval(p, "((0))&((0))", "0", why);
               }});
  t.push_back({"unit/bitwise &: array & array is element-wise array", [](std::string& why) {
                 MathParser p;
                 return expectEval(p, "(0,0)&(0,0)", "(0,0)", why);
               }});

  t.push_back({"unit/bitwise |: array & scalar is element-wise array", [](std::string& why) {
                 MathParser p;
                 return expectEval(p, "(0,0)|1", "(1,1)", why);
               }});

  t.push_back({"unit/bitwise |: grouping scalar is scalar", [](std::string& why) {
                 MathParser p;
                 return expectEval(p, "((0))|((0))", "0", why);
               }});

  t.push_back({"unit/bitwise |: array | array is element-wise array", [](std::string& why) {
                 MathParser p;
                 return expectEval(p, "(0,0)|(0,0)", "(0,0)", why);
               }});

  return t;
}

std::vector<TestCase> buildEdgeIntFloatCases() {
  std::vector<TestCase> t;
  const long long kLLMax = std::numeric_limits<long long>::max();
  const long long kLLMin = std::numeric_limits<long long>::min();
  const long long kPow53 = 9007199254740992LL;  // 2^53

  t.push_back({"edge/ll const 0 add and hex", [](std::string& why) {
                 MathParser p;
                 addConstTracked(p, "k", 0LL);
                 if (!expectEval(p, "k+0", "0", why)) {
                   return false;
                 }
                 return expectEval(p, "hex(k)", "0x0", why);
               }});
  t.push_back({"edge/ll const 1 square and hex", [](std::string& why) {
                 MathParser p;
                 addConstTracked(p, "k", 1LL);
                 if (!expectEval(p, "k*k", "1", why)) {
                   return false;
                 }
                 return expectEval(p, "hex(k)", "0x1", why);
               }});
  t.push_back({"edge/ll const 100 add and hex", [](std::string& why) {
                 MathParser p;
                 addConstTracked(p, "k", 100LL);
                 if (!expectEval(p, "k+1", "101", why)) {
                   return false;
                 }
                 return expectEval(p, "hex(k)", "0x64", why);
               }});
  t.push_back({"edge/ll const 1000000 add and hex", [](std::string& why) {
                 MathParser p;
                 addConstTracked(p, "k", 1000000LL);
                 if (!expectEval(p, "k+1", "1000001", why)) {
                   return false;
                 }
                 return expectEval(p, "hex(k)", "0xF4240", why);
               }});

  t.push_back({"edge/double const 0 keeps integer-valued hex()", [](std::string& why) {
                 MathParser p;
                 addConstTracked(p, "k", 0.0);
                 return expectEval(p, "hex(k)", "0x0", why);
               }});
  t.push_back({"edge/double const 1 keeps integer-valued hex()", [](std::string& why) {
                 MathParser p;
                 addConstTracked(p, "k", 1.0);
                 return expectEval(p, "hex(k)", "0x1", why);
               }});
  t.push_back({"edge/double const 100 keeps integer-valued hex()", [](std::string& why) {
                 MathParser p;
                 addConstTracked(p, "k", 100.0);
                 return expectEval(p, "hex(k)", "0x64", why);
               }});
  t.push_back({"edge/double const 1000000 keeps integer-valued hex()", [](std::string& why) {
                 MathParser p;
                 addConstTracked(p, "k", 1000000.0);
                 return expectEval(p, "hex(k)", "0xF4240", why);
               }});
  t.push_back({"edge/hex array with non-integer element rejects", [](std::string& why) {
                MathParser p;
                return expectEvalErrorContains(p, "hex((1,2.5,3))", "hex() expects integer values", why);
              }});

  t.push_back({"edge/double const 0 same add as ll", [](std::string& why) {
                 MathParser pD;
                 MathParser pI;
                 addConstTracked(pD, "k", 0.0);
                 addConstTracked(pI, "k", 0LL);
                 pD.parseAndEvaluate("k+1");
                 pI.parseAndEvaluate("k+1");
                 if (!pD.getError().empty() || pD.getResult() != "1") {
                   why = "double const: " + pD.getError() + " / " + pD.getResult();
                   return false;
                 }
                 if (!pI.getError().empty() || pI.getResult() != "1") {
                   why = "ll const: " + pI.getError() + " / " + pI.getResult();
                   return false;
                 }
                 return true;
               }});
  t.push_back({"edge/double const 1 same mul as ll", [](std::string& why) {
                 MathParser pD;
                 MathParser pI;
                 addConstTracked(pD, "k", 1.0);
                 addConstTracked(pI, "k", 1LL);
                 pD.parseAndEvaluate("k*42");
                 pI.parseAndEvaluate("k*42");
                 return pD.getResult() == pI.getResult() && pD.getResult() == "42" && pD.getError().empty() &&
                        pI.getError().empty();
               }});
  t.push_back({"edge/double const 100 same add1 as ll", [](std::string& why) {
                 MathParser pD;
                 MathParser pI;
                 addConstTracked(pD, "k", 100.0);
                 addConstTracked(pI, "k", 100LL);
                 pD.parseAndEvaluate("k+1");
                 pI.parseAndEvaluate("k+1");
                 if (pD.getResult() != "101" || pI.getResult() != "101") {
                   why = "mismatch double vs ll";
                   return false;
                 }
                 return pD.getError().empty() && pI.getError().empty();
               }});
  t.push_back({"edge/double const 1000000 same add1 as ll", [](std::string& why) {
                 MathParser pD;
                 MathParser pI;
                 addConstTracked(pD, "k", 1000000.0);
                 addConstTracked(pI, "k", 1000000LL);
                 pD.parseAndEvaluate("k+1");
                 pI.parseAndEvaluate("k+1");
                 if (pD.getResult() != "1000001" || pI.getResult() != "1000001") {
                   why = "mismatch double vs ll";
                   return false;
                 }
                 return pD.getError().empty() && pI.getError().empty();
               }});

  t.push_back({"edge/ll max plus zero", [kLLMax](std::string& why) {
                 MathParser p;
                 addConstTracked(p, "k", kLLMax);
                 return expectEval(p, "k+0", "9223372036854775807", why);
               }});
  t.push_back({"edge/ll min plus zero", [kLLMin](std::string& why) {
                 MathParser p;
                 addConstTracked(p, "k", kLLMin);
                 return expectEval(p, "k+0", "-9223372036854775808", why);
               }});
  t.push_back({"edge/ll max minus one plus one", [kLLMax](std::string& why) {
                 MathParser p;
                 addConstTracked(p, "k", kLLMax);
                 return expectEval(p, "(k-1)+1", "9223372036854775807", why);
               }});

  t.push_back({"edge/ll 2^53 plus 1 hex", [kPow53](std::string& why) {
                 MathParser p;
                 addConstTracked(p, "k", kPow53);
                 return expectEval(p, "hex(k+1)", "0x20000000000001", why);
               }});
  t.push_back({"edge/double 2^53 plus 1 rounds in IEEE double", [kPow53](std::string& why) {
                 MathParser p;
                 addConstTracked(p, "k", static_cast<double>(kPow53));
                 // 2^53+1 is not representable; +1 rounds back to 2^53.
                 return expectEval(p, "k+1", "9007199254740992", why);
               }});

  t.push_back({"edge/gcd ll const integer path", [](std::string& why) {
                 MathParser p;
                 addConstTracked(p, "a", 100LL);
                 return expectEval(p, "gcd(a,25)", "25", why);
               }});
  t.push_back({"edge/gcd uint64-above-signed-range accepted", [](std::string& why) {
                MathParser p;
                 return expectEval(p, "gcd(18446744073709551615,3)", "3", why);
              }});
  t.push_back({"edge/gcd double const integer-valued", [](std::string& why) {
                 MathParser p;
                 addConstTracked(p, "a", 100.0);
                 return expectEval(p, "gcd(a,25)", "25", why);
               }});
  t.push_back({"edge/lcm uint64-above-signed-range accepted", [](std::string& why) {
                MathParser p;
                 return expectEval(p, "lcm(18446744073709551615,3)", "18446744073709551615", why);
              }});
  t.push_back({"edge/sum min max uint64 exact before double", [](std::string& why) {
                MathParser p;
                if (!expectEval(p, "sum(18446744073709551614,1)", "18446744073709551615", why)) return false;
                if (!expectEval(p, "max(3,18446744073709551615)", "18446744073709551615", why)) return false;
                if (!expectEval(p, "min(18446744073709551615,5)", "5", why)) return false;
                return true;
              }});
  t.push_back({"edge/sum max int64 negative exact before double", [](std::string& why) {
                MathParser p;
                if (!expectEval(p, "sum(-9223372036854775807,-1)", "-9223372036854775808", why)) return false;
                if (!expectEval(p, "max(-9,-1)", "-1", why)) return false;
                return true;
              }});
  t.push_back({"edge/product uint64 and int64 exact before double", [](std::string& why) {
                MathParser p;
                if (!expectEval(p, "uhex(product(18446744073709551615,1))", "0xFFFFFFFFFFFFFFFF", why)) return false;
                if (!expectEval(p, "prod(-9223372036854775807,1)&1", "1", why)) return false;
                if (!expectEval(p, "uhex(product(-9223372036854775807-1,-1))", "0x8000000000000000", why)) return false;
                return true;
              }});
  t.push_back({"edge/uint64 max plus one rounds via float like int(1e20)", [](std::string& why) {
                MathParser p;
                const char* p64 = "1.844674407370955e+019";
                if (!expectEval(p, "int(0xFFFFFFFFFFFFFFFF+1)", p64, why)) return false;
                if (!expectEval(p, "trunc(0xFFFFFFFFFFFFFFFF+1)", p64, why)) return false;
                if (!expectEval(p, "floor(0xFFFFFFFFFFFFFFFF+1)", p64, why)) return false;
                if (!expectEval(p, "ceil(0xFFFFFFFFFFFFFFFF+1)", p64, why)) return false;
                if (!expectEval(p, "0xFFFFFFFFFFFFFFFF+1", p64, why)) return false;
                return true;
              }});
  t.push_back({"edge/ncr basic", [](std::string& why) {
                MathParser p;
                return expectEval(p, "ncr(5,2)", "10", why);
              }});
  t.push_back({"edge/npr basic", [](std::string& why) {
                MathParser p;
                return expectEval(p, "npr(5,2)", "20", why);
              }});
  t.push_back({"edge/ncr invalid domain", [](std::string& why) {
                MathParser p;
                return expectEvalErrorContains(p, "ncr(5,7)", "numeric error in ncr()", why);
              }});
  t.push_back({"edge/npr rejects inf integer-only", [](std::string& why) {
                MathParser p;
                return expectEvalErrorContains(p, "npr(inf,2)", "npr() expects integer values", why);
              }});
  t.push_back({"edge/ncr missing call hint", [](std::string& why) {
                MathParser p;
                return expectEvalErrorContains(p, "ncr", "function: ncr(n, r)", why);
              }});
  t.push_back({"edge/mod ll const", [](std::string& why) {
                 MathParser p;
                 addConstTracked(p, "a", 100LL);
                 return expectEval(p, "mod(a,3)", "1", why);
               }});
  t.push_back({"edge/mod operator uint64-above-signed-range accepted", [](std::string& why) {
                MathParser p;
                return expectEval(p, "18446744073709551615%3", "0", why);
              }});
  t.push_back({"edge/mod double const integer-valued", [](std::string& why) {
                 MathParser p;
                 addConstTracked(p, "a", 100.0);
                 return expectEval(p, "mod(a,3)", "1", why);
               }});
  t.push_back({"edge/mod double const fractional rejects", [](std::string& why) {
                 MathParser p;
                 addConstTracked(p, "a", 100.5);
                 return expectEvalErrorContains(p, "mod(a,3)", "mod() expects integer values", why);
               }});
  t.push_back({"edge/mod builtin uint64-above-signed-range accepted", [](std::string& why) {
                MathParser p;
                return expectEval(p, "mod(18446744073709551615,3)", "0", why);
              }});
  t.push_back({"edge/uhex decimal uint64 max formats unsigned", [](std::string& why) {
                MathParser p;
                return expectEval(p, "uhex(18446744073709551615)", "0xFFFFFFFFFFFFFFFF", why);
              }});

  t.push_back({"edge/bitand ll const", [](std::string& why) {
                 MathParser p;
                 addConstTracked(p, "a", 100LL);
                 return expectEval(p, "a&7", "4", why);
               }});
  t.push_back({"edge/bitand double const integer-valued ok", [](std::string& why) {
                 MathParser p;
                 addConstTracked(p, "a", 100.0);
                 return expectEval(p, "a&7", "4", why);
               }});
  t.push_back({"edge/bitand double const fractional rejects", [](std::string& why) {
                 MathParser p;
                 addConstTracked(p, "a", 100.5);
                 return expectEvalErrorContains(p, "a&7", "bitwise operands must be integer values", why);
               }});

  t.push_back({"edge/int() strips double const fraction", [](std::string& why) {
                 MathParser p;
                 addConstTracked(p, "x", 100.7);
                 return expectEval(p, "int(x)", "100", why);
               }});
  t.push_back({"edge/int() ll const still exact", [](std::string& why) {
                 MathParser p;
                 addConstTracked(p, "x", 100LL);
                 return expectEval(p, "int(x)", "100", why);
               }});

  t.push_back({"edge/double const 1e12 add matches literal", [](std::string& why) {
                 MathParser p;
                 addConstTracked(p, "k", 1e12);
                 p.parseAndEvaluate("k+1");
                 if (!p.getError().empty()) {
                   why = p.getError();
                   return false;
                 }
                 if (p.getResult() != "1000000000001") {
                   why = "got " + p.getResult();
                   return false;
                 }
                 return true;
               }});

  t.push_back({"edge/ll const near double int limit", [](std::string& why) {
                 MathParser p;
                 addConstTracked(p, "k", 9007199254740991LL);  // 2^53 - 1
                 return expectEval(p, "k+1", "9007199254740992", why);
               }});

  // --- LLONG_MAX / LLONG_MIN neighbors (exact int path) ---
  t.push_back({"edge/ll const max minus one plus one", [kLLMax](std::string& why) {
                 MathParser p;
                 addConstTracked(p, "k", kLLMax - 1);
                 return expectEval(p, "k+1", "9223372036854775807", why);
               }});
  t.push_back({"edge/ll const max minus one identity", [kLLMax](std::string& why) {
                 MathParser p;
                 addConstTracked(p, "k", kLLMax - 1);
                 return expectEval(p, "k+0", "9223372036854775806", why);
               }});
  t.push_back({"edge/ll const min plus one minus one", [kLLMin](std::string& why) {
                 MathParser p;
                 addConstTracked(p, "k", kLLMin + 1);
                 return expectEval(p, "k-1", "-9223372036854775808", why);
               }});
  t.push_back({"edge/ll const min plus one identity", [kLLMin](std::string& why) {
                 MathParser p;
                 addConstTracked(p, "k", kLLMin + 1);
                 return expectEval(p, "k+0", "-9223372036854775807", why);
               }});

  t.push_back({"edge/literal max minus one plus one", [](std::string& why) {
                 MathParser p;
                 return expectEval(p, "9223372036854775806+1", "9223372036854775807", why);
               }});
  t.push_back({"edge/literal min plus one minus one", [](std::string& why) {
                 MathParser p;
                 return expectEval(p, "-9223372036854775807-1", "-9223372036854775808", why);
               }});
  t.push_back({"edge/literal min plus one plus zero", [](std::string& why) {
                 MathParser p;
                 return expectEval(p, "-9223372036854775807+0", "-9223372036854775807", why);
               }});

  t.push_back({"edge/hex ll max minus one", [kLLMax](std::string& why) {
                 MathParser p;
                 addConstTracked(p, "k", kLLMax - 1);
                 return expectEval(p, "hex(k)", "0x7FFFFFFFFFFFFFFE", why);
               }});
  t.push_back({"edge/hex ll min plus one", [kLLMin](std::string& why) {
                 MathParser p;
                 addConstTracked(p, "k", kLLMin + 1);
                 return expectEval(p, "hex(k)", "-0x7FFFFFFFFFFFFFFF", why);
               }});
  t.push_back({"edge/uhex ll min plus one", [kLLMin](std::string& why) {
                 MathParser p;
                 addConstTracked(p, "k", kLLMin + 1);
                 return expectEval(p, "uhex(k)", "0x8000000000000001", why);
               }});

  // --- Double vs ll at 2^53: (2^53-1)+2 is 2^53+1, not representable in double; rounds to 2^53 ---
  t.push_back({"edge/double 2^53 minus 1 plus 2 rounds to 2^53", [kPow53](std::string& why) {
                 MathParser p;
                 addConstTracked(p, "k", static_cast<double>(kPow53 - 1));
                 return expectEval(p, "k+2", "9007199254740992", why);
               }});
  t.push_back({"edge/ll 2^53 minus 1 plus 2 exact", [kPow53](std::string& why) {
                 MathParser p;
                 addConstTracked(p, "k", kPow53 - 1);
                 return expectEval(p, "k+2", "9007199254740993", why);
               }});
  t.push_back({"edge/double 2^53 plus 2 no rounding loss", [kPow53](std::string& why) {
                 MathParser p;
                 addConstTracked(p, "k", static_cast<double>(kPow53));
                 return expectEval(p, "k+2", "9007199254740994", why);
               }});
  t.push_back({"edge/double vs ll 2^53 minus 1 plus 2 differ (IEEE)", [kPow53](std::string& why) {
                 MathParser pD;
                 MathParser pI;
                 addConstTracked(pD, "k", static_cast<double>(kPow53 - 1));
                 addConstTracked(pI, "k", kPow53 - 1);
                 pD.parseAndEvaluate("k+2");
                 pI.parseAndEvaluate("k+2");
                 if (!pD.getError().empty() || !pI.getError().empty()) {
                   why = pD.getError() + " / " + pI.getError();
                   return false;
                 }
                 if (pD.getResult() == pI.getResult()) {
                   why = "expected double path to round differently from exact ll";
                   return false;
                 }
                 if (pD.getResult() != "9007199254740992" || pI.getResult() != "9007199254740993") {
                   why = "double " + pD.getResult() + " vs ll " + pI.getResult();
                   return false;
                 }
                 return true;
               }});

  // LLONG_MAX as double loses precision vs ll const (platform IEEE round)
  t.push_back({"edge/double ll max const differs from ll max", [kLLMax](std::string& why) {
                 MathParser pLl;
                 MathParser pD;
                 addConstTracked(pLl, "k", kLLMax);
                 addConstTracked(pD, "k", static_cast<double>(kLLMax));
                 pLl.parseAndEvaluate("k+0");
                 pD.parseAndEvaluate("k+0");
                 if (pLl.getError().empty() && pD.getError().empty() && pLl.getResult() == pD.getResult()) {
                   why = "expected double const to differ from ll max (precision loss)";
                   return false;
                 }
                 return pLl.getResult() == "9223372036854775807" && pD.getError().empty();
               }});

  t.push_back({"edge/int double 2^53 minus 1", [kPow53](std::string& why) {
                 MathParser p;
                 addConstTracked(p, "k", static_cast<double>(kPow53 - 1));
                 return expectEval(p, "int(k)", "9007199254740991", why);
               }});
  t.push_back({"edge/int ll 2^53 minus 1", [kPow53](std::string& why) {
                 MathParser p;
                 addConstTracked(p, "k", kPow53 - 1);
                 return expectEval(p, "int(k)", "9007199254740991", why);
               }});

  // Exact-int recovery after float-path operations (scalar).
  t.push_back({"edge/sqr exact int preserves metadata for hex", [](std::string& why) {
                 MathParser p;
                 return expectEval(p, "hex(sqr(9))", "0x51", why);
               }});
  t.push_back({"edge/sqr exact int near int64 square limit", [](std::string& why) {
                 MathParser p;
                 p.parseAndEvaluate("sqr(3037000499)");
                 if (!p.getError().empty()) {
                   why = p.getError();
                   return false;
                 }
                 const std::string actual = p.getResult();
                 const MathParser::RawResult raw = p.getRawResult();
                 if (!raw.isScalar() || !raw.scalar.real.isInt64() ||
                     raw.scalar.real.intValue != 9223372030926249001LL) {
                   why = "expected exact int64 square in raw, got result=\"" + actual + "\" kind=" +
                         std::to_string(static_cast<int>(raw.scalar.real.kind));
                   if (raw.isScalar() && raw.scalar.real.isInt64()) {
                     why += " intValue=" + std::to_string(raw.scalar.real.intValue);
                   } else if (raw.isScalar() && raw.scalar.real.isFloatingPoint()) {
                     why += " fp=" + std::to_string(raw.scalar.real.floatingPoint);
                   }
                   return false;
                 }
                 if (actual != "9223372030926249001") {
                   why = "expected \"9223372030926249001\", got \"" + actual + "\"";
                   return false;
                 }
                 return true;
               }});
  t.push_back({"edge/hypot exact int squares then library sqrt", [](std::string& why) {
                 MathParser p;
                 return expectEval(p, "hypot(3037000499,0)", "3037000499", why);
               }});
  t.push_back({"edge/sqr float input uses floating path", [](std::string& why) {
                 MathParser p;
                 return expectEval(p, "sqr(1.5)", "2.25", why);
               }});
  t.push_back({"edge/scalar sqrt perfect square keeps int metadata for bitwise", [](std::string& why) {
                 MathParser p;
                 return expectEval(p, "sqrt(81)&7", "1", why);
               }});
  t.push_back({"edge/scalar sqrt perfect square keeps int metadata for hex", [](std::string& why) {
                 MathParser p;
                 return expectEval(p, "hex(sqrt(81))", "0x9", why);
               }});
  t.push_back({"edge/sqrt non-perfect square must not fake exact int64", [](std::string& why) {
                 MathParser p;
                 p.parseAndEvaluate("sqrt(4611686014132420611)");
                 if (!p.getError().empty()) {
                   why = p.getError();
                   return false;
                 }
                 const MathParser::RawResult raw = p.getRawResult();
                 if (!raw.isScalar() || raw.scalar.real.isInt64() || raw.scalar.real.isUInt64()) {
                   why = "expected floating raw for non-perfect-square sqrt";
                   return false;
                 }
                 if (p.getResult() != "2147483647") {
                   why = "expected \"2147483647\", got \"" + p.getResult() + "\"";
                   return false;
                 }
                 return true;
               }});
  t.push_back({"edge/sqrt perfect square raw is int64", [](std::string& why) {
                 MathParser p;
                 p.parseAndEvaluate("sqrt(81)");
                 if (!p.getError().empty()) {
                   why = p.getError();
                   return false;
                 }
                 const MathParser::RawResult raw = p.getRawResult();
                 if (!raw.isScalar() || !raw.scalar.real.isInt64() || raw.scalar.real.intValue != 9LL) {
                   why = "expected int64 9 in raw for sqrt(81)";
                   return false;
                 }
                 return true;
               }});
  t.push_back({"edge/sqrt non-perfect square product differs from radicand", [](std::string& why) {
                 MathParser p;
                 return expectEval(p, "sqrt(4611686014132420611)*sqrt(4611686014132420611)",
                                   "4611686014132420609", why);
               }});
  t.push_back({"edge/sqrt non-perfect square sqr detects no fake exact int", [](std::string& why) {
                 MathParser p;
                 return expectEval(p, "sqr(sqrt(4611686014132420611))", "4611686014132420609", why);
               }});
  t.push_back({"edge/pow verified fractional root exact int", [](std::string& why) {
                 MathParser p;
                 if (!expectEval(p, "125**(1/3)", "5", why)) return false;
                 return expectEval(p, "823543**(1/7)", "7", why);
               }});
  t.push_back({"edge/pow negative base odd exponent and root", [](std::string& why) {
                 MathParser p;
                 p.setSupportComplexNumbers(true);
                 if (!expectEval(p, "(-5)**3", "-125", why)) return false;
                 if (!expectEval(p, "(-125)**(1/3)", "-5", why)) return false;
                 if (!expectEval(p, "pow(-5,3)", "-125", why)) return false;
                 if (!expectEval(p, "pow(-125,1/3)", "-5", why)) return false;
                 if (!expectEval(p, "pow(pow(-3,39), 1/39)", "-3", why)) return false;
                 return expectEval(p, "pow(pow(-3,45), 1/45)", "-3", why);
               }});
  t.push_back({"edge/pow base 1 and integer exponent verify", [](std::string& why) {
                 MathParser p;
                 if (!expectEval(p, "pow(1,99)", "1", why)) return false;
                 return expectEval(p, "2**10", "1024", why);
               }});
  t.push_back({"edge/pow non-perfect root stays float", [](std::string& why) {
                 MathParser p;
                 return expectEval(p, "126**(1/3)", "5.013297934964584", why);
               }});
  t.push_back({"edge/scalar abs keeps int metadata for modulo", [](std::string& why) {
                 MathParser p;
                 return expectEval(p, "mod(abs(-14),5)", "4", why);
               }});
  t.push_back({"edge/scalar pow exact int keeps metadata for bitwise", [](std::string& why) {
                 MathParser p;
                 return expectEval(p, "pow(3,2)&7", "1", why);
               }});
  t.push_back({"edge/power uint64 exact before double", [](std::string& why) {
                 MathParser p;
                 if (!expectEval(p, "uhex(2**63)", "0x8000000000000000", why)) return false;
                 if (!expectEval(p, "uhex(pow(2,63))", "0x8000000000000000", why)) return false;
                 if (!expectEval(p, "uhex(pow((2,3),(63,2)))", "(0x8000000000000000,0x9)", why)) return false;
                 if (!expectEval(p, "uhex((-2)**63)", "0x8000000000000000", why)) return false;
                 return true;
               }});
  t.push_back({"edge/scalar div-by-one keeps int metadata for hex", [kPow53](std::string& why) {
                 MathParser p;
                 addConstTracked(p, "k", kPow53 + 2);  // 2^53 + 2 is exactly representable in double
                 return expectEval(p, "hex(int(k/1))", "0x20000000000002", why);
               }});
  t.push_back({"edge/scalar int() after float path keeps metadata", [kPow53](std::string& why) {
                 MathParser p;
                 addConstTracked(p, "k", kPow53 + 2);
                 return expectEval(p, "hex(int((k/2)+0.0))", "0x10000000000001", why);
               }});

  // Exact-int recovery after float-path operations (arrays).
  t.push_back({"edge/array sqrt perfect squares index keeps bitwise int path", [](std::string& why) {
                 MathParser p;
                 return expectEval(p, "a=sqrt((81,16,25)); a[0]&3", "1", why);
               }});
  t.push_back({"edge/array sqrt perfect squares index keeps hex int path", [](std::string& why) {
                 MathParser p;
                 return expectEval(p, "a=sqrt((81,16,25)); hex(a[2])", "0x5", why);
               }});
  t.push_back({"edge/array int() on mixed signs keeps index bitwise int path", [](std::string& why) {
                 MathParser p;
                 return expectEval(p, "a=int((2.9,-2.9,7.1)); a[1]&1", "0", why);
               }});
  t.push_back({"edge/array int() on mixed signs keeps index hex int path", [](std::string& why) {
                 MathParser p;
                 return expectEval(p, "a=int((2.9,-2.9,7.1)); hex(a[0])", "0x2", why);
               }});
  t.push_back({"edge/array div-by-one keeps large-int metadata for indexed hex", [kPow53](std::string& why) {
                 MathParser p;
                 addConstTracked(p, "k", kPow53);
                 return expectEval(p, "a=int(((k+2),(k+6))/1); hex(a[0])", "0x20000000000002", why);
               }});
  t.push_back({"edge/array div-by-one keeps large-int metadata for indexed mod", [kPow53](std::string& why) {
                 MathParser p;
                 addConstTracked(p, "k", kPow53);
                 return expectEval(p, "a=int(((k+2),(k+6))/1); mod(a[1],4)", "2", why);
               }});
  t.push_back({"edge/array half-then-int keeps metadata for indexed bitwise", [kPow53](std::string& why) {
                 MathParser p;
                 addConstTracked(p, "k", kPow53);
                 return expectEval(p, "a=int(((k+2),(k+6))/2); a[0]&1", "1", why);
               }});
  t.push_back({"edge/array half-then-int keeps metadata for indexed hex", [kPow53](std::string& why) {
                 MathParser p;
                 addConstTracked(p, "k", kPow53);
                 return expectEval(p, "a=int(((k+2),(k+6))/2); hex(a[1])", "0x10000000000003", why);
               }});
  t.push_back({"edge/array int over arithmetic expression keeps metadata", [](std::string& why) {
                 MathParser p;
                 return expectEval(p, "a=int(((5.9+0.1),(9.2+0.8))); mod(a[1],4)", "2", why);
               }});
  t.push_back({"edge/array int over arithmetic expression keeps indexed hex", [](std::string& why) {
                 MathParser p;
                 return expectEval(p, "a=int(((5.9+0.1),(9.2+0.8))); hex(a[0])", "0x6", why);
               }});
  t.push_back({"edge/array left-shift keeps int64 exact near max then index hex", [](std::string& why) {
                 MathParser p;
                 return expectEval(p, "a=(4611686018427387903,5)<<1; hex(a[0])", "0x7FFFFFFFFFFFFFFE", why);
               }});
  t.push_back({"edge/array left-shift keeps int64 exact near max then index mod", [](std::string& why) {
                 MathParser p;
                 return expectEval(p, "a=(4611686018427387903,5)<<1; mod(a[0],7)", "6", why);
               }});
  t.push_back({"edge/array right-shift keeps int64 exact after big left-shift", [](std::string& why) {
                 MathParser p;
                 return expectEval(p, "a=(4611686018427387903,5)<<1; b=a>>1; hex(b[0])", "0x3FFFFFFFFFFFFFFF", why);
               }});
  t.push_back({"edge/array floor keeps exact-int context beyond 2^53", [](std::string& why) {
                MathParser p;
                return expectEval(p, "a=floor((9007199254740993,2.9)); hex(a[0])", "0x20000000000001", why);
              }});
  t.push_back({"edge/aggregate sum keeps scalar exact-int context beyond 2^53", [](std::string& why) {
                MathParser p;
                return expectEval(p, "a=(9007199254740993,2); hex(sum(a[0]))", "0x20000000000001", why);
              }});
  t.push_back({"edge/array abs keeps uint64 context beyond signed range", [](std::string& why) {
                MathParser p;
                return expectEval(p, "a=abs((18446744073709551615,-7)); uhex(a[0])", "0xFFFFFFFFFFFFFFFF", why);
              }});
  t.push_back({"edge/sort keeps uint64 context beyond signed range", [](std::string& why) {
                MathParser p;
                return expectEval(p, "a=sort((18446744073709551615,2)); uhex(a[1])", "0xFFFFFFFFFFFFFFFF", why);
              }});
  t.push_back({"edge/reverse keeps uint64 context beyond signed range", [](std::string& why) {
                MathParser p;
                return expectEval(p, "a=reverse((18446744073709551615,2)); uhex(a[1])", "0xFFFFFFFFFFFFFFFF", why);
              }});
  t.push_back({"edge/unique keeps uint64 context beyond signed range", [](std::string& why) {
                MathParser p;
                return expectEval(p, "a=unique((18446744073709551615,18446744073709551615,2)); uhex(a[0])", "0xFFFFFFFFFFFFFFFF", why);
              }});
  t.push_back({"edge/array bitwise and keeps int64 exact near max", [](std::string& why) {
                 MathParser p;
                 return expectEval(p, "a=(9223372036854775806,15); b=a&7; hex(b[0])", "0x6", why);
               }});
  t.push_back({"edge/array bitwise or keeps int64 exact near max", [](std::string& why) {
                 MathParser p;
                 return expectEval(p, "a=(9223372036854775800,1); b=a|7; hex(b[0])", "0x7FFFFFFFFFFFFFFF", why);
               }});
  t.push_back({"edge/scalar-array modulo supports elementwise output", [](std::string& why) {
                MathParser p;
                return expectEval(p, "mod(7,(2,3))", "(1,1)", why);
              }});
  t.push_back({"edge/array-scalar modulo supports elementwise output", [](std::string& why) {
                MathParser p;
                return expectEval(p, "mod((7,8),3)", "(1,2)", why);
              }});
  t.push_back({"edge/array shift rejects out-of-range count", [](std::string& why) {
                MathParser p;
                return expectEvalErrorContains(p, "a=(1,2)<<64", "incompatible operands", why);
              }});
  t.push_back({"edge/array add uint64 max promotes to floating", [](std::string& why) {
                MathParser p;
                return expectEvalErrorContains(p, "a=(18446744073709551615,1); b=(1,2); c=a+b; c[0]&1", "bitwise operands must be integer values", why);
              }});
  t.push_back({"edge/array mul uint64 max promotes to floating", [](std::string& why) {
                MathParser p;
                return expectEvalErrorContains(p, "a=(18446744073709551615,3); b=(3,4); c=a*b; mod(c[0],5)", "mod() expects integer values", why);
              }});
  t.push_back({"edge/scalar 2^53 float path restores signed exact int", [](std::string& why) {
                MathParser p;
                return expectEval(p, "k=9007199254740992; v=(k+2)-2; hex(int(v))", "0x20000000000000", why);
              }});
  t.push_back({"edge/scalar uint64 max plus one promotes to floating", [](std::string& why) {
                MathParser p;
                return expectEvalErrorContains(p, "k=18446744073709551615; v=k+1; uhex(v)", "uhex() expects integer values", why);
              }});
  t.push_back({"edge/array uint64 max plus one promotes per element", [](std::string& why) {
                MathParser p;
                return expectEvalErrorContains(p, "a=(18446744073709551615,1); b=(1,2); c=a+b; uhex(c[0])", "uhex() expects integer values", why);
              }});
  t.push_back({"edge/array uint64 max times three promotes per element", [](std::string& why) {
                MathParser p;
                return expectEvalErrorContains(p, "a=(18446744073709551615,3); b=(3,4); c=a*b; uhex(c[0])", "uhex() expects integer values", why);
              }});
  t.push_back({"edge/array near-max stored then int-op preserves exactness", [](std::string& why) {
                 MathParser p;
                 return expectEval(p, "a=(9223372036854775806,9223372036854775805); mod(a[1],5)", "0", why);
               }});
  t.push_back({"edge/array near-max stored then shift preserves exactness", [](std::string& why) {
                 MathParser p;
                 return expectEval(p, "a=(9223372036854775806,9223372036854775805); b=a>>2; hex(b[0])", "0x1FFFFFFFFFFFFFFF", why);
               }});

  return t;
}

std::vector<TestCase> buildNanInfCases() {
  std::vector<TestCase> t;
  const double kQNaN = std::numeric_limits<double>::quiet_NaN();
  const double kPosInf = std::numeric_limits<double>::infinity();
  const double kNegInf = -std::numeric_limits<double>::infinity();

  t.push_back({"naninf/quiet NaN dec identity", [kQNaN](std::string& why) {
                 MathParser p;
                 addConstTracked(p, "x", kQNaN);
                 return expectEval(p, "x+0", "nan", why);
               }});
  t.push_back({"naninf/+inf dec identity", [kPosInf](std::string& why) {
                 MathParser p;
                 addConstTracked(p, "x", kPosInf);
                 return expectEval(p, "x+0", "inf", why);
               }});
  t.push_back({"naninf/-inf dec identity", [kNegInf](std::string& why) {
                 MathParser p;
                 addConstTracked(p, "x", kNegInf);
                 return expectEval(p, "x+0", "-inf", why);
               }});

  t.push_back({"naninf/inf plus finite", [kPosInf](std::string& why) {
                 MathParser p;
                 addConstTracked(p, "x", kPosInf);
                 return expectEval(p, "x+1", "inf", why);
               }});
  t.push_back({"naninf/-inf minus finite", [kNegInf](std::string& why) {
                 MathParser p;
                 addConstTracked(p, "x", kNegInf);
                 return expectEval(p, "x-1", "-inf", why);
               }});
  t.push_back({"naninf/nan plus finite stays nan", [kQNaN](std::string& why) {
                 MathParser p;
                 addConstTracked(p, "x", kQNaN);
                 return expectEval(p, "x+1", "nan", why);
               }});

  t.push_back({"naninf/abs -inf", [kNegInf](std::string& why) {
                 MathParser p;
                 addConstTracked(p, "x", kNegInf);
                 return expectEval(p, "abs(x)", "inf", why);
               }});
  t.push_back({"naninf/abs nan", [kQNaN](std::string& why) {
                 MathParser p;
                 addConstTracked(p, "x", kQNaN);
                 return expectEval(p, "abs(x)", "nan", why);
               }});
  t.push_back({"naninf/sign +inf", [kPosInf](std::string& why) {
                 MathParser p;
                 addConstTracked(p, "x", kPosInf);
                 return expectEval(p, "sign(x)", "1", why);
               }});
  t.push_back({"naninf/sign -inf", [kNegInf](std::string& why) {
                 MathParser p;
                 addConstTracked(p, "x", kNegInf);
                 return expectEval(p, "sign(x)", "-1", why);
               }});
  t.push_back({"naninf/sign nan", [kQNaN](std::string& why) {
                 MathParser p;
                 addConstTracked(p, "x", kQNaN);
                 return expectEval(p, "sign(x)", "0", why);
               }});

  t.push_back({"naninf/ln nan", [kQNaN](std::string& why) {
                 MathParser p;
                 addConstTracked(p, "x", kQNaN);
                 return expectEval(p, "ln(x)", "nan", why);
               }});
  t.push_back({"naninf/ln +inf", [kPosInf](std::string& why) {
                 MathParser p;
                 addConstTracked(p, "x", kPosInf);
                 return expectEval(p, "ln(x)", "inf", why);
               }});
  t.push_back({"naninf/ln zero to -inf", [](std::string& why) {
                 MathParser p;
                 addConstTracked(p, "z", 0.0);
                 return expectEval(p, "ln(z)", "-inf", why);
               }});
  t.push_back({"naninf/sqrt nan", [kQNaN](std::string& why) {
                 MathParser p;
                 addConstTracked(p, "x", kQNaN);
                 return expectEval(p, "sqrt(x)", "nan", why);
               }});
  t.push_back({"naninf/sin +inf", [kPosInf](std::string& why) {
                 MathParser p;
                 addConstTracked(p, "x", kPosInf);
                 return expectEvalErrorContains(p, "sin(x)", "numeric error in sin()", why);
               }});

  t.push_back({"naninf/frac +inf", [kPosInf](std::string& why) {
                 MathParser p;
                 addConstTracked(p, "x", kPosInf);
                 return expectEval(p, "frac(x)", "nan", why);
               }});
  t.push_back({"naninf/frac nan", [kQNaN](std::string& why) {
                 MathParser p;
                 addConstTracked(p, "x", kQNaN);
                 return expectEval(p, "frac(x)", "nan", why);
               }});

  t.push_back({"naninf/hex NaN preserved", [kQNaN](std::string& why) {
                 MathParser p;
                 addConstTracked(p, "x", kQNaN);
                 return expectEval(p, "hex(x)", "nan", why);
               }});
  t.push_back({"naninf/hex +inf preserved", [kPosInf](std::string& why) {
                 MathParser p;
                 addConstTracked(p, "x", kPosInf);
                 return expectEval(p, "hex(x)", "inf", why);
               }});
  t.push_back({"naninf/oct -inf preserved", [kNegInf](std::string& why) {
                 MathParser p;
                 addConstTracked(p, "x", kNegInf);
                 return expectEval(p, "oct(x)", "-inf", why);
               }});
  t.push_back({"naninf/uhex +inf preserved", [kPosInf](std::string& why) {
                 MathParser p;
                 addConstTracked(p, "x", kPosInf);
                 return expectEval(p, "uhex(x)", "inf", why);
               }});
  t.push_back({"naninf/mod NaN rejects", [kQNaN](std::string& why) {
                 MathParser p;
                 addConstTracked(p, "x", kQNaN);
                 return expectEvalErrorContains(p, "mod(x,3)", "mod() expects integer values", why);
               }});
  t.push_back({"naninf/mod +inf rejects", [kPosInf](std::string& why) {
                 MathParser p;
                 addConstTracked(p, "x", kPosInf);
                 return expectEvalErrorContains(p, "mod(x,3)", "mod() expects integer values", why);
               }});
  t.push_back({"naninf/gcd NaN rejects", [kQNaN](std::string& why) {
                 MathParser p;
                 addConstTracked(p, "x", kQNaN);
                 return expectEvalErrorContains(p, "gcd(x,1)", "gcd() expects integer values", why);
               }});
  t.push_back({"naninf/bitand NaN rejects", [kQNaN](std::string& why) {
                 MathParser p;
                 addConstTracked(p, "x", kQNaN);
                 return expectEvalErrorContains(p, "x&1", "bitwise operands must be integer values", why);
               }});
  t.push_back({"naninf/bitand +inf rejects", [kPosInf](std::string& why) {
                 MathParser p;
                 addConstTracked(p, "x", kPosInf);
                 return expectEvalErrorContains(p, "x&1", "bitwise operands must be integer values", why);
               }});
  t.push_back({"naninf/sum with +inf rejects", [kPosInf](std::string& why) {
                 MathParser p;
                 addConstTracked(p, "x", kPosInf);
                 return expectEval(p, "sum(1,x)", "inf", why);
               }});

  t.push_back({"naninf/not NaN is true (NaN is falsy)", [kQNaN](std::string& why) {
                 MathParser p;
                 addConstTracked(p, "x", kQNaN);
                 return expectEval(p, "not x", "1", why);
               }});
  t.push_back({"naninf/not +inf is false", [kPosInf](std::string& why) {
                 MathParser p;
                 addConstTracked(p, "x", kPosInf);
                 return expectEval(p, "not x", "0", why);
               }});
  t.push_back({"naninf/NaN and 1", [kQNaN](std::string& why) {
                 MathParser p;
                 addConstTracked(p, "x", kQNaN);
                 return expectEval(p, "x&&1", "0", why);
               }});
  t.push_back({"naninf/NaN or 1", [kQNaN](std::string& why) {
                 MathParser p;
                 addConstTracked(p, "x", kQNaN);
                 return expectEval(p, "x||1", "1", why);
               }});

  t.push_back({"naninf/NaN == NaN is false (IEEE unordered)", [kQNaN](std::string& why) {
                 MathParser p;
                 addConstTracked(p, "x", kQNaN);
                 return expectEval(p, "x==x", "0", why);
               }});
  t.push_back({"naninf/nan==nan literal false", [](std::string& why) {
                 MathParser p;
                 return expectEval(p, "nan==nan", "0", why);
               }});
  t.push_back({"naninf/nan!=nan literal true", [](std::string& why) {
                 MathParser p;
                 return expectEval(p, "nan!=nan", "1", why);
               }});
  t.push_back({"naninf/nan<>nan literal true", [](std::string& why) {
                 MathParser p;
                 return expectEval(p, "nan<>nan", "1", why);
               }});
  t.push_back({"naninf/nan<nan literal false", [](std::string& why) {
                 MathParser p;
                 return expectEval(p, "nan<nan", "0", why);
               }});
  t.push_back({"naninf/nan>nan literal false", [](std::string& why) {
                 MathParser p;
                 return expectEval(p, "nan>nan", "0", why);
               }});
  t.push_back({"naninf/nan<=nan literal false", [](std::string& why) {
                 MathParser p;
                 return expectEval(p, "nan<=nan", "0", why);
               }});
  t.push_back({"naninf/nan>=nan literal false", [](std::string& why) {
                 MathParser p;
                 return expectEval(p, "nan>=nan", "0", why);
               }});
  t.push_back({"naninf/+inf == +inf", [kPosInf](std::string& why) {
                 MathParser p;
                 addConstTracked(p, "x", kPosInf);
                 return expectEval(p, "x==x", "1", why);
               }});
  t.push_back({"naninf/array arg sum with +inf", [kPosInf](std::string& why) {
                 MathParser p;
                 addConstTracked(p, "x", kPosInf);
                 return expectEval(p, "sum((x,2,3))", "inf", why);
               }});
  t.push_back({"naninf/sort orders NaN before infinities and numbers", [](std::string& why) {
                 MathParser p;
                 if (!expectEval(p, "sort(0xFFFFFFFFFFFFFFFF,nan,0x7FFFFFFFFFFFFFFF);hex", "(nan,0x7FFFFFFFFFFFFFFF,0xFFFFFFFFFFFFFFFF)", why)) return false;
                 if (!expectEval(p, "sort(-5,nan,3,-inf,inf,0)", "(nan,-inf,-5,0,3,inf)", why)) return false;
                 if (!expectEval(p, "sort(nan,inf,2,-inf,nan,-2)", "(nan,nan,-inf,-2,2,inf)", why)) return false;
                 if (!expectEval(p, "sorted((nan,-3,4,inf,-inf))", "(nan,-inf,-3,4,inf)", why)) return false;
                 return true;
               }});
  t.push_back({"naninf/array arg mod rejects +inf", [kPosInf](std::string& why) {
                 MathParser p;
                 addConstTracked(p, "x", kPosInf);
                 return expectEvalErrorContains(p, "mod((x,2),3)", "mod() expects integer values", why);
               }});
  t.push_back({"naninf/2arg function hypot(+inf,3)", [kPosInf](std::string& why) {
                 MathParser p;
                 addConstTracked(p, "x", kPosInf);
                 return expectEval(p, "hypot(x,3)", "inf", why);
               }});
  t.push_back({"naninf/3arg function clamp(+inf,0,7)", [kPosInf](std::string& why) {
                 MathParser p;
                 addConstTracked(p, "x", kPosInf);
                 return expectEval(p, "clamp(x,0,7)", "7", why);
               }});
  t.push_back({"naninf/udf 1 arg with +inf", [kPosInf](std::string& why) {
                 MathParser p;
                 addConstTracked(p, "x", kPosInf);
                 return expectEval(p, "f(a)=a+1; f(x)", "inf", why);
               }});
  t.push_back({"naninf/udf 2 args with +inf", [kPosInf](std::string& why) {
                 MathParser p;
                 addConstTracked(p, "x", kPosInf);
                 return expectEval(p, "f(a,b)=a*b; f(x,2)", "inf", why);
               }});
  t.push_back({"naninf/operator arithmetic multiply", [kPosInf](std::string& why) {
                 MathParser p;
                 addConstTracked(p, "x", kPosInf);
                 return expectEval(p, "x*2", "inf", why);
               }});
  t.push_back({"naninf/operator bitwise or rejects +inf", [kPosInf](std::string& why) {
                 MathParser p;
                 addConstTracked(p, "x", kPosInf);
                 return expectEvalErrorContains(p, "x|1", "bitwise operands must be integer values", why);
               }});
  t.push_back({"naninf/operator logical and", [kPosInf](std::string& why) {
                 MathParser p;
                 addConstTracked(p, "x", kPosInf);
                 return expectEval(p, "x&&0", "0", why);
               }});
  t.push_back({"naninf/operator comparison greater", [kPosInf](std::string& why) {
                 MathParser p;
                 addConstTracked(p, "x", kPosInf);
                 return expectEval(p, "x>1", "1", why);
               }});
  t.push_back({"naninf/variadic sum with +inf", [kPosInf](std::string& why) {
                 MathParser p;
                 addConstTracked(p, "x", kPosInf);
                 return expectEval(p, "sum(1,2,x)", "inf", why);
               }});
  t.push_back({"naninf/variadic max with +inf", [kPosInf](std::string& why) {
                 MathParser p;
                 addConstTracked(p, "x", kPosInf);
                 return expectEval(p, "max(1,2,x)", "inf", why);
               }});
  t.push_back({"naninf/variadic avg with +inf", [kPosInf](std::string& why) {
                 MathParser p;
                 addConstTracked(p, "x", kPosInf);
                 return expectEval(p, "avg(1,x,3)", "inf", why);
               }});
  t.push_back({"naninf/operator compare +inf == +inf", [kPosInf](std::string& why) {
                 MathParser p;
                 addConstTracked(p, "x", kPosInf);
                 addConstTracked(p, "y", kPosInf);
                 return expectEval(p, "x==y", "1", why);
               }});

  t.push_back({"naninf/array arg sum with -inf", [kNegInf](std::string& why) {
                 MathParser p;
                 addConstTracked(p, "x", kNegInf);
                 return expectEval(p, "sum((x,2,3))", "-inf", why);
               }});
  t.push_back({"naninf/array arg mod rejects -inf", [kNegInf](std::string& why) {
                 MathParser p;
                 addConstTracked(p, "x", kNegInf);
                 return expectEvalErrorContains(p, "mod((x,2),3)", "mod() expects integer values", why);
               }});
  t.push_back({"naninf/2arg function hypot(-inf,3)", [kNegInf](std::string& why) {
                 MathParser p;
                 addConstTracked(p, "x", kNegInf);
                 return expectEval(p, "hypot(x,3)", "inf", why);
               }});
  t.push_back({"naninf/3arg function clamp(-inf,0,7)", [kNegInf](std::string& why) {
                 MathParser p;
                 addConstTracked(p, "x", kNegInf);
                 return expectEval(p, "clamp(x,0,7)", "0", why);
               }});
  t.push_back({"naninf/udf 1 arg with -inf", [kNegInf](std::string& why) {
                 MathParser p;
                 addConstTracked(p, "x", kNegInf);
                 return expectEval(p, "f(a)=a+1; f(x)", "-inf", why);
               }});
  t.push_back({"naninf/udf 2 args with -inf", [kNegInf](std::string& why) {
                 MathParser p;
                 addConstTracked(p, "x", kNegInf);
                 return expectEval(p, "f(a,b)=a*b; f(x,2)", "-inf", why);
               }});
  t.push_back({"naninf/operator arithmetic multiply -inf", [kNegInf](std::string& why) {
                 MathParser p;
                 addConstTracked(p, "x", kNegInf);
                 return expectEval(p, "x*2", "-inf", why);
               }});
  t.push_back({"naninf/operator bitwise or rejects -inf", [kNegInf](std::string& why) {
                 MathParser p;
                 addConstTracked(p, "x", kNegInf);
                 return expectEvalErrorContains(p, "x|1", "bitwise operands must be integer values", why);
               }});
  t.push_back({"naninf/operator logical and -inf", [kNegInf](std::string& why) {
                 MathParser p;
                 addConstTracked(p, "x", kNegInf);
                 return expectEval(p, "x&&0", "0", why);
               }});
  t.push_back({"naninf/operator comparison -inf not greater than 1", [kNegInf](std::string& why) {
                 MathParser p;
                 addConstTracked(p, "x", kNegInf);
                 return expectEval(p, "x>1", "0", why);
               }});
  t.push_back({"naninf/variadic sum with -inf", [kNegInf](std::string& why) {
                 MathParser p;
                 addConstTracked(p, "x", kNegInf);
                 return expectEval(p, "sum(1,2,x)", "-inf", why);
               }});
  t.push_back({"naninf/variadic max with -inf", [kNegInf](std::string& why) {
                 MathParser p;
                 addConstTracked(p, "x", kNegInf);
                 return expectEval(p, "max(1,2,x)", "2", why);
               }});
  t.push_back({"naninf/variadic avg with -inf", [kNegInf](std::string& why) {
                 MathParser p;
                 addConstTracked(p, "x", kNegInf);
                 return expectEval(p, "avg(1,x,3)", "-inf", why);
               }});
  t.push_back({"naninf/operator compare -inf == -inf", [kNegInf](std::string& why) {
                 MathParser p;
                 addConstTracked(p, "x", kNegInf);
                 addConstTracked(p, "y", kNegInf);
                 return expectEval(p, "x==y", "1", why);
               }});

  t.push_back({"naninf/** inf**2", [kPosInf](std::string& why) {
                 MathParser p;
                 addConstTracked(p, "x", kPosInf);
                 return expectEval(p, "x**2", "inf", why);
               }});
  t.push_back({"naninf/** 2**inf", [kPosInf](std::string& why) {
                 MathParser p;
                 addConstTracked(p, "x", kPosInf);
                 return expectEval(p, "2**x", "inf", why);
               }});
  t.push_back({"naninf/** inf**inf", [kPosInf](std::string& why) {
                 MathParser p;
                 addConstTracked(p, "x", kPosInf);
                 return expectEval(p, "x**x", "inf", why);
               }});
  t.push_back({"naninf/pow(inf,2)", [kPosInf](std::string& why) {
                 MathParser p;
                 addConstTracked(p, "x", kPosInf);
                 return expectEval(p, "pow(x,2)", "inf", why);
               }});
  t.push_back({"naninf/pow(2,inf)", [kPosInf](std::string& why) {
                 MathParser p;
                 addConstTracked(p, "x", kPosInf);
                 return expectEval(p, "pow(2,x)", "inf", why);
               }});
  t.push_back({"naninf/pow(inf,inf)", [kPosInf](std::string& why) {
                 MathParser p;
                 addConstTracked(p, "x", kPosInf);
                 return expectEval(p, "pow(x,x)", "inf", why);
               }});

  t.push_back({"naninf/atan2(+inf,1)", [kPosInf](std::string& why) {
                 MathParser p;
                 addConstTracked(p, "x", kPosInf);
                 return expectEval(p, "atan2(x,1)", "1.570796326794897", why);
               }});
  t.push_back({"naninf/atan2(-inf,1)", [kNegInf](std::string& why) {
                 MathParser p;
                 addConstTracked(p, "x", kNegInf);
                 return expectEval(p, "atan2(x,1)", "-1.570796326794897", why);
               }});
  t.push_back({"naninf/atan2(1,-inf)", [kNegInf](std::string& why) {
                 MathParser p;
                 addConstTracked(p, "x", kNegInf);
                 return expectEval(p, "atan2(1,x)", "3.141592653589793", why);
               }});
  t.push_back({"naninf/sin -inf", [kNegInf](std::string& why) {
                 MathParser p;
                 addConstTracked(p, "x", kNegInf);
                 return expectEvalErrorContains(p, "sin(x)", "numeric error in sin()", why);
               }});

  t.push_back({"naninf/int accepts -inf", [kNegInf](std::string& why) {
                 MathParser p;
                 addConstTracked(p, "x", kNegInf);
                 return expectEval(p, "int(x)", "-inf", why);
               }});
  t.push_back({"naninf/int accepts inf", [kPosInf](std::string& why) {
                 MathParser p;
                 addConstTracked(p, "x", kPosInf);
                 return expectEval(p, "int(x)", "inf", why);
               }});
  t.push_back({"naninf/int accepts -nan", [kQNaN](std::string& why) {
                 MathParser p;
                 addConstTracked(p, "x", -kQNaN);
                 return expectEval(p, "int(x)", "nan", why);
               }});
  t.push_back({"naninf/int accepts nan", [kQNaN](std::string& why) {
                 MathParser p;
                 addConstTracked(p, "x", kQNaN);
                 return expectEval(p, "int(x)", "nan", why);
               }});
  t.push_back({"naninf/ceil accepts -inf", [kNegInf](std::string& why) {
                 MathParser p;
                 addConstTracked(p, "x", kNegInf);
                 return expectEval(p, "ceil(x)", "-inf", why);
               }});
  t.push_back({"naninf/floor accepts inf", [kPosInf](std::string& why) {
                 MathParser p;
                 addConstTracked(p, "x", kPosInf);
                 return expectEval(p, "floor(x)", "inf", why);
               }});
  t.push_back({"naninf/round accepts -nan", [kQNaN](std::string& why) {
                 MathParser p;
                 addConstTracked(p, "x", -kQNaN);
                 return expectEval(p, "round(x)", "nan", why);
               }});
  t.push_back({"naninf/trunc accepts nan", [kQNaN](std::string& why) {
                 MathParser p;
                 addConstTracked(p, "x", kQNaN);
                 return expectEval(p, "trunc(x)", "nan", why);
               }});

  return t;
}

std::vector<TestCase> buildRegressionCases() {
  std::vector<TestCase> t;

  // compile/evaluate lifecycle and AST reuse regressions.
  t.push_back({"regression/compile then evaluate", [](std::string& why) {
                 MathParser p;
                 if (!p.compile("a=10; a*2")) {
                   why = "compile failed: " + p.getError();
                   return false;
                 }
                 p.evaluate();
                 if (!p.getError().empty()) {
                   why = "evaluate failed: " + p.getError();
                   return false;
                 }
                 if (p.getResult() != "20") {
                   why = "expected 20, got " + p.getResult();
                   return false;
                 }
                 return true;
               }});
  t.push_back({"regression/evaluate without compile", [](std::string& why) {
                 MathParser p;
                 p.evaluate();
                 if (p.getError().find("compile") == std::string::npos) {
                   why = "expected compile-first error, got: " + p.getError();
                   return false;
                 }
                 return true;
               }});
  t.push_back({"regression/re-evaluate after const update", [](std::string& why) {
                 MathParser p;
                 addConstTracked(p, "k", 3LL);
                 if (!p.compile("k*10")) {
                   why = "compile failed";
                   return false;
                 }
                 p.evaluate();
                 if (p.getResult() != "30") {
                   why = "expected 30, got " + p.getResult();
                   return false;
                 }
                 addConstTracked(p, "k", 9LL);
                 p.evaluate();
                 if (p.getResult() != "90") {
                   why = "expected 90 after const update, got " + p.getResult();
                   return false;
                 }
                 return true;
               }});
  t.push_back({"regression/re-evaluate keeps compiled function definition", [](std::string& why) {
                 MathParser p;
                 if (!p.compile("f(x)=x*2; f(10)")) {
                   why = p.getError();
                   return false;
                 }
                 p.evaluate();
                 if (p.getResult() != "20") {
                   why = "expected 20, got " + p.getResult();
                   return false;
                 }
                 const std::string err = p.addUserFunction("f(x)=x*3");
                 if (!err.empty()) {
                   why = err;
                   return false;
                 }
                 p.evaluate();
                 if (p.getResult() != "20") {
                   why = "expected compiled definition to win (20), got " + p.getResult();
                   return false;
                 }
                 return true;
               }});
  t.push_back({"regression/function override affects call-only compiled expression", [](std::string& why) {
                 MathParser p;
                 const std::string err1 = p.addUserFunction("f(x)=x*2");
                 if (!err1.empty()) {
                   why = err1;
                   return false;
                 }
                 if (!p.compile("f(10)")) {
                   why = p.getError();
                   return false;
                 }
                 p.evaluate();
                 if (p.getResult() != "20") {
                   why = "expected 20, got " + p.getResult();
                   return false;
                 }
                 const std::string err2 = p.addUserFunction("f(x)=x*3");
                 if (!err2.empty()) {
                   why = err2;
                   return false;
                 }
                 p.evaluate();
                 if (p.getResult() != "30") {
                   why = "expected 30 after override, got " + p.getResult();
                   return false;
                 }
                 return true;
               }});
  t.push_back({"regression/failed compile drops old program", [](std::string& why) {
                 MathParser p;
                 if (!p.compile("2+3")) {
                   why = p.getError();
                   return false;
                 }
                 p.evaluate();
                 if (p.getResult() != "5") {
                   why = "setup result mismatch";
                   return false;
                 }
                 if (p.compile("pow(2,")) {
                   why = "compile unexpectedly succeeded";
                   return false;
                 }
                 p.evaluate();
                 if (p.getError().find("compile") == std::string::npos) {
                   why = "expected evaluate compile-first after failed compile, got: " + p.getError();
                   return false;
                 }
                 return true;
               }});
  t.push_back({"regression/percentage precedence over multiply", [](std::string& why) {
                 MathParser p;
                return expectEval(p, "200 + 2*15%", "200.3", why);
               }});
  t.push_back({"regression/percentage on grouped product", [](std::string& why) {
                MathParser p;
                return expectEval(p, "200 + (2*15)%", "260", why);
              }});
  t.push_back({"regression/unpack requires at least one argument", [](std::string& why) {
                MathParser p;
                return expectEvalErrorContains(p, "unpack()", "unpack() expects at least 1 argument", why);
              }});
  t.push_back({"regression/unpack single argument is passthrough", [](std::string& why) {
                MathParser p;
                return expectEval(p, "unpack(5)", "5", why);
              }});
  t.push_back({"regression/unpack flattens multi-argument values", [](std::string& why) {
                MathParser p;
                return expectEval(p, "unpack((1,2),3)", "(1,2,3)", why);
              }});
  t.push_back({"regression/unpack preserves left-to-right flatten order", [](std::string& why) {
                MathParser p;
                return expectEval(p, "unpack(1,(2,3))", "(1,2,3)", why);
              }});
  t.push_back({"regression/unpack single array keeps array value", [](std::string& why) {
                MathParser p;
                return expectEval(p, "unpack((1,2))", "(1,2)", why);
              }});
  t.push_back({"regression/deg requires at least one argument", [](std::string& why) {
                MathParser p;
                return expectEvalErrorContains(p, "deg()", "deg() expects at least 1 argument", why);
              }});
  t.push_back({"regression/deg array preserves elementwise conversion", [](std::string& why) {
                MathParser p;
                return expectEval(p, "deg((pi,pi/2))", "(180,90)", why);
              }});
  t.push_back({"regression/rad requires at least one argument", [](std::string& why) {
                MathParser p;
                return expectEvalErrorContains(p, "rad()", "rad() expects at least 1 argument", why);
              }});
  t.push_back({"regression/rad array preserves elementwise conversion", [](std::string& why) {
                MathParser p;
                return expectEval(p, "rad((180,90))", "(3.141592653589793,1.570796326794896)", why);
              }});
  t.push_back({"regression/rad zero array stays zero", [](std::string& why) {
                MathParser p;
                return expectEval(p, "rad((0,0))", "(0,0)", why);
              }});
  t.push_back({"regression/clamp requires exactly three arguments", [](std::string& why) {
                MathParser p;
                return expectEvalErrorContains(p, "clamp(1,2)", "clamp() expects 3 argument(s)", why);
              }});
  t.push_back({"regression/clamp rejects extra arguments", [](std::string& why) {
                MathParser p;
                return expectEvalErrorContains(p, "clamp(1,2,3,4)", "clamp() expects 3 argument(s)", why);
              }});
  t.push_back({"regression/clamp scalar stays scalar result", [](std::string& why) {
                MathParser p;
                return expectEval(p, "clamp(9,0,7)", "7", why);
              }});
  t.push_back({"regression/clamp array value with scalar bounds stays elementwise", [](std::string& why) {
                MathParser p;
                return expectEval(p, "clamp((1,9),0,7)", "(1,7)", why);
              }});
  t.push_back({"regression/clamp rejects non-scalar min", [](std::string& why) {
                MathParser p;
                return expectEvalErrorContains(p, "clamp((1,2),(3,4),4)", "clamp() expects scalar min/max", why);
              }});
  t.push_back({"regression/clamp rejects non-scalar max", [](std::string& why) {
                MathParser p;
                return expectEvalErrorContains(p, "clamp(5,1,(4,7))", "clamp() expects scalar min/max", why);
              }});
  t.push_back({"regression/atan2 supports array-scalar elementwise", [](std::string& why) {
                MathParser p;
               return expectEval(p, "atan2((1,2),3)", "(0.3217505543966422,0.5880026035475675)", why);
              }});
  t.push_back({"regression/hypot supports array-scalar elementwise", [](std::string& why) {
                MathParser p;
               return expectEval(p, "hypot((3,4),5)", "(5.830951894845301,6.403124237432849)", why);
              }});
  t.push_back({"regression/random non-scalar keeps scalar-values contract", [](std::string& why) {
                MathParser p;
                return expectEvalErrorContains(p, "random((1,2),3)", "random() expects scalar values", why);
              }});
  t.push_back({"regression/random requires exactly two arguments", [](std::string& why) {
                MathParser p;
                return expectEvalErrorContains(p, "random(1)", "random() expects 2 argument(s)", why);
              }});
  t.push_back({"regression/random equal bounds returns that bound", [](std::string& why) {
                MathParser p;
                return expectEval(p, "random(5,5)", "5", why);
              }});
  t.push_back({"regression/rand requires zero arguments", [](std::string& why) {
                MathParser p;
                return expectEvalErrorContains(p, "rand(1)", "rand() expects 0 argument(s)", why);
              }});
  t.push_back({"regression/factorial rejects non-scalar input", [](std::string& why) {
                MathParser p;
                return expectEvalErrorContains(p, "fact((3,4))", "fact() expects a non-negative integer", why);
              }});
  t.push_back({"regression/factorial rejects negative input", [](std::string& why) {
                MathParser p;
                return expectEvalErrorContains(p, "fact(-1)", "fact() expects a non-negative integer", why);
              }});
  t.push_back({"regression/factorial above 20 switches to floating output", [](std::string& why) {
                MathParser p;
                return expectEval(p, "fact(21)", "5.109094217170944e+019", why);
              }});
  t.push_back({"regression/factorial zero equals one", [](std::string& why) {
                MathParser p;
                return expectEval(p, "fact(0)", "1", why);
              }});
  t.push_back({"regression/factorial stops once double overflows", [](std::string& why) {
                MathParser p;
                if (!expectEval(p, "fact(171)", "inf", why)) {
                  return false;
                }
                return expectEval(p, "fact(170000000)", "inf", why);
              }});
  t.push_back({"regression/comparison operators keep boolean semantics", [](std::string& why) {
                MathParser p;
                return expectEval(p, "(2<3)+(2>=3)+(5==5)+(5<>4)", "3", why);
              }});
  t.push_back({"regression/comparison <> alias matches !=", [](std::string& why) {
                MathParser p;
                return expectEval(p, "(7<>8)+(7!=8)", "2", why);
              }});
  t.push_back({"regression/comparison strict and non-strict mix", [](std::string& why) {
                MathParser p;
                return expectEval(p, "(9>8)+(9<=8)+(8>=8)", "2", why);
              }});
  t.push_back({"regression/comparison equality and inequality mix", [](std::string& why) {
                MathParser p;
                return expectEval(p, "(3==3)+(3!=4)+(3<>3)", "2", why);
              }});
  t.push_back({"regression/builtin call then == is comparison not UDF starter", [](std::string& why) {
                MathParser p;
                return expectEval(p, "sin(0)==0", "1", why);
              }});
  t.push_back({"regression/variable then sin(x)==x compares equality", [](std::string& why) {
                MathParser p;
                return expectEval(p, "x=0; sin(x)==x", "1", why);
              }});
  t.push_back({"regression/UDF body uses single equals call then ==", [](std::string& why) {
                MathParser p;
                return expectEval(p, "f(t)=t*3; f(2)==6", "1", why);
              }});
  t.push_back({"regression/sum requires at least one argument", [](std::string& why) {
                MathParser p;
                return expectEvalErrorContains(p, "sum()", "sum() expects at least 1 argument", why);
              }});
  t.push_back({"regression/hex requires at least one argument", [](std::string& why) {
                MathParser p;
                return expectEvalErrorContains(p, "hex()", "hex() expects at least 1 argument", why);
              }});
  t.push_back({"regression/ubin formats scalar as unsigned binary", [](std::string& why) {
                MathParser p;
                return expectEval(p, "ubin(5)", "0b101", why);
              }});
  t.push_back({"regression/mod requires exactly two arguments", [](std::string& why) {
                MathParser p;
                return expectEvalErrorContains(p, "mod(5)", "mod() expects 2 argument(s)", why);
              }});
  t.push_back({"regression/pow requires exactly two arguments", [](std::string& why) {
                MathParser p;
                return expectEvalErrorContains(p, "pow(5)", "pow() expects 2 argument(s)", why);
              }});
  t.push_back({"regression/operator caret is bitwise xor in compiled path", [](std::string& why) {
                MathParser p;
                return expectEval(p, "2^3", "1", why);
              }});
  t.push_back({"regression/operator double-star is power in compiled path", [](std::string& why) {
                MathParser p;
                return expectEval(p, "2**3", "8", why);
              }});
  t.push_back({"regression/operator caret is not power (3^2)", [](std::string& why) {
                MathParser p;
                return expectEval(p, "3^2", "1", why);
              }});
  t.push_back({"regression/operator double-star power (3**2)", [](std::string& why) {
                MathParser p;
                return expectEval(p, "3**2", "9", why);
              }});
  t.push_back({"regression/log requires exactly two arguments", [](std::string& why) {
                MathParser p;
                return expectEvalErrorContains(p, "log(5)", "log() expects 2 argument(s)", why);
              }});
  t.push_back({"regression/builtin arity table: ratio wrong count", [](std::string& why) {
                MathParser p;
                return expectEvalErrorContains(p, "ratio(1,2)", "ratio() expects 1 argument(s)", why);
              }});
  t.push_back({"regression/builtin arity table: sortby wrong counts are parse errors", [](std::string& why) {
                MathParser p;
                if (!expectEvalErrorContains(
                        p, "sortby()", "sortby expects a function that takes 1 parameter", why)) {
                  return false;
                }
                if (!expectEvalErrorContains(
                        p, "sortby(1)", "sortby expects a function that takes 1 parameter", why)) {
                  return false;
                }
                return expectEvalErrorContains(p, "sortby(1,2,3)", "sortby expects", why);
              }});
  t.push_back({"regression/builtin arity table: lcm and npr require two args", [](std::string& why) {
                MathParser p;
                if (!expectEvalErrorContains(p, "lcm()", "lcm() expects 2 argument(s)", why)) {
                  return false;
                }
                return expectEvalErrorContains(p, "npr(1)", "npr() expects 2 argument(s)", why);
              }});
  t.push_back({"regression/builtin arity table: variadic deg hex sum unpack", [](std::string& why) {
                MathParser p;
                if (!expectEval(p, "deg(1,2)", "(57.29577951308232, 114.5915590261646)", why)) {
                  return false;
                }
                if (!expectEval(p, "hex(1,2,3)", "(0x1,0x2,0x3)", why)) {
                  return false;
                }
                if (!expectEval(p, "sum(1,2,3)", "6", why)) {
                  return false;
                }
                return expectEval(p, "unpack(1,2)", "(1,2)", why);
              }});
  t.push_back({"regression/abs preserves exact uint64 above int64 max", [](std::string& why) {
                MathParser p;
                if (!expectEval(p, "abs(0x7FFFFFFFFFFFFFFF+20)", "9223372036854775827", why)) {
                  return false;
                }
                return expectEval(p, "hex(abs(0x7FFFFFFFFFFFFFFF+20))", "0x8000000000000013", why);
              }});
  t.push_back({"regression/abs preserves exact signed int64", [](std::string& why) {
                MathParser p;
                if (!expectEval(p, "abs(-14)", "14", why)) {
                  return false;
                }
                return expectEval(p, "abs(-9223372036854775808)", "9223372036854775808", why);
              }});
#if SMARTMATH_TIME_VALUES
  t.push_back({"regression/builtin arity table: time converters exact one arg", [](std::string& why) {
                MathParser p;
                p.setSupportTimeValues(true);
                if (!expectEvalErrorContains(p, "milliseconds()", "milliseconds() expects 1 argument(s)", why)) {
                  return false;
                }
                return expectEvalErrorContains(
                    p, "seconds(1:00,2:00)", "seconds() expects 1 argument(s)", why);
              }});
#endif
  t.push_back({"regression/log array with scalar base stays elementwise", [](std::string& why) {
                MathParser p;
                return expectEval(p, "log((8,64),2)", "(3,6)", why);
              }});
  t.push_back({"regression/multiplicative mismatched arrays keep incompatible-operands error", [](std::string& why) {
                MathParser p;
                return expectEvalErrorContains(p, "(1,2)*(3,4,5)", "incompatible operands", why);
              }});
  t.push_back({"regression/additive mismatched arrays keep incompatible-operands error", [](std::string& why) {
                MathParser p;
                return expectEvalErrorContains(p, "(1,2)+(3,4,5)", "incompatible operands", why);
              }});
  t.push_back({"regression/formatter stability large integers", [](std::string& why) {
                 MathParser p;
                 p.parseAndEvaluate("9223372036854775807");
                 if (!p.getError().empty()) {
                   why = p.getError();
                   return false;
                 }
                 const std::string d = p.getResultAsDec();
                 if (d.empty()) {
                   why = "empty decimal result";
                   return false;
                 }
                 return true;
               }});
  t.push_back({"regression/unary minus on LLONG_MIN (overflow-safe)", [](std::string& why) {
                 MathParser p;
                 if (!p.compile("-((-9223372036854775807)-1)")) {
                   why = "compile failed: " + p.getError();
                   return false;
                 }
                 p.evaluate();
                 if (!p.getError().empty()) {
                   why = p.getError();
                   return false;
                 }
                 if (p.getResultAsHex() != "0x8000000000000000") {
                   why = "expected 0x8000000000000000, got " + p.getResultAsHex();
                   return false;
                 }
                 return true;
               }});
  t.push_back({"regression/ans exact after user function call", [](std::string& why) {
                 MathParser p;
                 if (!p.compile("f(x)=9007199254740993+x; f(0); hex(ans)")) {
                   why = "compile failed: " + p.getError();
                   return false;
                 }
                 p.evaluate();
                 if (!p.getError().empty()) {
                   why = p.getError();
                   return false;
                 }
                 if (p.getResult() != "0x20000000000001") {
                   why = "expected 0x20000000000001, got " + p.getResult();
                   return false;
                 }
                 return true;
               }});
  t.push_back({"regression/ans exact across program statements", [](std::string& why) {
                 MathParser p;
                 if (!p.compile("9007199254740992+1; hex(ans)")) {
                   why = "compile failed: " + p.getError();
                   return false;
                 }
                 p.evaluate();
                 if (!p.getError().empty()) {
                   why = p.getError();
                   return false;
                 }
                 if (p.getResult() != "0x20000000000001") {
                   why = "expected 0x20000000000001, got " + p.getResult();
                   return false;
                 }
                 return true;
               }});
  t.push_back({"regression/statement expression updates ans", [](std::string& why) {
                MathParser p;
                return expectEval(p, "a=1; 2; ans", "2", why);
              }});
  t.push_back({"regression/ans resets to zero after failed evaluation", [](std::string& why) {
                MathParser p;
                if (!expectEval(p, "10+5", "15", why)) {
                  return false;
                }
                p.parseAndEvaluate("unknownFunc(1)");
                if (p.getError().empty()) {
                  why = "expected failure for unknownFunc(1)";
                  return false;
                }
                return expectEval(p, "ans", "0", why);
              }});
  t.push_back({"regression/exact scalar metadata survives variable copy", [](std::string& why) {
                MathParser p;
                return expectEval(p, "a=9007199254740992+1; b=a; hex(b)", "0x20000000000001", why);
              }});
  t.push_back({"regression/uhex array keeps unsigned formatting", [](std::string& why) {
                MathParser p;
                return expectEval(p, "uhex((1,-1))", "(0x1,0xFFFFFFFFFFFFFFFF)", why);
              }});
  t.push_back({"regression/sum unpack mixed inputs keeps flatten order", [](std::string& why) {
                MathParser p;
                return expectEval(p, "sum(unpack((1,2),(3,4),5))", "15", why);
              }});
  t.push_back({"regression/array-array add keeps second element", [](std::string& why) {
                MathParser p;
                return expectEval(p, "a=(1,2)+(3,4); hex(a[1])", "0x6", why);
              }});
  t.push_back({"regression/not keyword remains case-insensitive", [](std::string& why) {
                MathParser p;
                return expectEval(p, "NoT 0", "1", why);
              }});
  t.push_back({"regression/comparison chain remains left-associative", [](std::string& why) {
                MathParser p;
                return expectEval(p, "(5>=4)<2", "1", why);
              }});
  t.push_back({"regression/round array keeps integer exactness on index", [](std::string& why) {
                MathParser p;
                return expectEval(p, "a=round((1.2,2.8)); hex(a[1])", "0x3", why);
              }});
  t.push_back({"regression/array-array multiply keeps second element exact", [](std::string& why) {
                MathParser p;
                return expectEval(p, "a=(5,6)*(7,8); hex(a[1])", "0x30", why);
              }});
  t.push_back({"regression/array-array subtraction keeps second element exact", [](std::string& why) {
                MathParser p;
                return expectEval(p, "a=(10,20)-(3,4); hex(a[1])", "0x10", why);
              }});
  t.push_back({"regression/array-array modulo keeps second element exact", [](std::string& why) {
                MathParser p;
                return expectEval(p, "a=(20,21)%(3,4); hex(a[1])", "0x1", why);
              }});
  t.push_back({"regression/log scalar value with array bases stays elementwise", [](std::string& why) {
                MathParser p;
                return expectEval(p, "log(8,(2,4))", "(3,1.5)", why);
              }});
  t.push_back({"regression/lcm scalar pair remains exact", [](std::string& why) {
                MathParser p;
                return expectEval(p, "lcm(21,6)", "42", why);
              }});
  t.push_back({"regression/exact integer division quotient", [](std::string& why) {
                MathParser p;
                if (!expectEval(p, "10/2", "5", why)) return false;
                if (!expectEval(p, "-12/4", "-3", why)) return false;
                if (!expectEval(p, "1955685900012/338822921", "5772", why)) return false;
                if (!expectEval(p, "111*17618791892/5772", "338822921", why)) return false;
                if (!expectEval(p, "gcd(111*17618791892,222*76568758786)", "5772", why)) return false;
                if (!expectEval(p, "(2**63 - 3)/5", "1844674407370955161", why)) return false;
                if (!expectEval(p, "(-2**63 + 3)/5", "-1844674407370955161", why)) return false;
                if (!expectEval(p, "(0xFFFFFFFFFFFFFFFF - 5)/5", "3689348814741910322", why)) return false;
                if (!expectEval(p, "(2**52 - 1)/5", "900719925474099", why)) return false;
                if (!expectEval(p, "(-2**52 + 1)/5", "-900719925474099", why)) return false;
                if (!expectEval(p, "1317624576693539401/(1/11)", "14493870343628933411", why)) return false;
                if (!expectEval(p, "1317624576693539401/(1/7)", "9223372036854775807", why)) return false;
                if (!expectEval(p, "-1317624576693539401/(1/7)", "-9223372036854775807", why)) return false;
                if (!expectEval(p, "1317624576693539401/(-1/7)", "-9223372036854775807", why)) return false;
                if (!expectEval(p, "0x7FFFFFFFFFFFFFFF/(1/8)", "7.37869762948382e+019", why)) return false;
                if (!expectEval(p, "2**64/(1/2)", "3.68934881474191e+019", why)) return false;
                if (!expectEval(p, "1317624576693539401*(1/11)", "1.197840524266854e+017", why)) return false;
                if (!expectEval(p, "0x7FFFFFFFFFFFFFFF*(-1/8)", "-1.152921504606847e+018", why)) return false;
                if (!expectEval(p, "100/(5/2)", "40", why)) return false;
                if (!expectEval(p, "100/(10/3+1e-15)", "29.99999999999999", why)) return false;
                if (!expectEval(p, "-50/(5/2)", "-20", why)) return false;
                if (!expectEval(p, "-50/((10/3)+1e-9)", "-14.9999999955", why)) return false;
                if (!expectEval(p, "700000000/(5/2)", "280000000", why)) return false;
                if (!expectEval(p, "700000000/(10/3+1e-15)", "209999999.9999999", why)) return false;
                if (!expectEval(p, "-300000000/(5/2)", "-120000000", why)) return false;
                if (!expectEval(p, "-300000000/(10/3+1e-15)", "-89999999.99999997", why)) return false;
                if (!expectEval(p, "(2**51-8)/(5/2)", "900719925474096", why)) return false;
                if (!expectEval(p, "(2**51-8)/((5/2)+1e-12)", "900719925473735.6", why)) return false;
                if (!expectEval(p, "-(2**50-1)/(3/2)", "-750599937895082", why)) return false;
                if (!expectEval(p, "-(2**50-1)/((3/2)+1e-12)", "-750599937894581.5", why)) return false;
                if (!expectEval(p, "100*(5/2)", "250", why)) return false;
                if (!expectEval(p, "100*((5/2)+1e-12)", "250.0000000001", why)) return false;
                if (!expectEval(p, "-50*(5/2)", "-125", why)) return false;
                if (!expectEval(p, "-50*((5/2)+1e-9)", "-125.00000005", why)) return false;
                if (!expectEval(p, "700000000*(5/2)", "1750000000", why)) return false;
                if (!expectEval(p, "700000000*((5/2)+1e-12)", "1750000000.0007", why)) return false;
                if (!expectEval(p, "-300000000*(5/2)", "-750000000", why)) return false;
                if (!expectEval(p, "-300000000*((5/2)+1e-12)", "-750000000.0003", why)) return false;
                if (!expectEval(p, "(2**51-8)*(5/2)", "5629499534213100", why)) return false;
                if (!expectEval(p, "(2**51-8)*((5/2)+1e-12)", "5629499534215352", why)) return false;
                if (!expectEval(p, "-(2**50-1)*2", "-2251799813685246", why)) return false;
                if (!expectEval(p, "-(2**50-1)*(2.0001)", "-225191240367593", why)) return false;
                p.parseAndEvaluate("7/3");
                if (p.getResult().find('.') == std::string::npos && p.getResult().find('/') == std::string::npos) {
                  why = "7/3 must stay non-integer float, got \"" + p.getResult() + "\"";
                  return false;
                }
                return true;
              }});
  t.push_back({"regression/array-array bitwise or keeps second element exact", [](std::string& why) {
                MathParser p;
                return expectEval(p, "a=(1,2)|(4,8); hex(a[1])", "0xA", why);
              }});
  t.push_back({"regression/array-array bitwise xor keeps second element exact", [](std::string& why) {
                MathParser p;
                return expectEval(p, "a=(1,2)^(4,8); hex(a[1])", "0xA", why);
              }});
  t.push_back({"regression/array literal expression keeps second element exact", [](std::string& why) {
                MathParser p;
                return expectEval(p, "a=(1,(2+3),7); hex(a[1])", "0x5", why);
              }});
  t.push_back({"regression/abs array keeps second element exact", [](std::string& why) {
                MathParser p;
                return expectEval(p, "a=abs((-1,-2)); hex(a[1])", "0x2", why);
              }});
  t.push_back({"regression/unary minus array keeps second element exact", [](std::string& why) {
                MathParser p;
                return expectEval(p, "a=-(-1,-2); hex(a[1])", "0x2", why);
              }});
  t.push_back({"regression/array shift-right keeps second element exact", [](std::string& why) {
                MathParser p;
                return expectEval(p, "a=(8,9)>>1; hex(a[1])", "0x4", why);
              }});
  t.push_back({"regression/array shift-left keeps second element exact", [](std::string& why) {
                MathParser p;
                return expectEval(p, "a=(8,9)<<1; hex(a[1])", "0x12", why);
              }});
  t.push_back({"regression/array bitand pair keeps second element exact", [](std::string& why) {
                MathParser p;
                return expectEval(p, "a=(9,10)&(3,12); hex(a[1])", "0x8", why);
              }});
  t.push_back({"regression/array bitor pair keeps second element exact", [](std::string& why) {
                MathParser p;
                return expectEval(p, "a=(9,10)|(3,12); hex(a[1])", "0xE", why);
              }});
  t.push_back({"regression/array bitxor pair keeps first element exact", [](std::string& why) {
                MathParser p;
                return expectEval(p, "a=(9,10)^(3,12); hex(a[0])", "0xA", why);
              }});
  t.push_back({"regression/abs array keeps first element exact", [](std::string& why) {
                MathParser p;
                return expectEval(p, "a=abs((-3,-4)); hex(a[0])", "0x3", why);
              }});
  t.push_back({"regression/array literal expression keeps third element exact", [](std::string& why) {
                MathParser p;
                return expectEval(p, "a=(1,(2+3),4); hex(a[2])", "0x4", why);
              }});
  t.push_back({"regression/int array keeps second element exact", [](std::string& why) {
                MathParser p;
                return expectEval(p, "a=int((7.9,8.1)); hex(a[1])", "0x8", why);
              }});
  t.push_back({"regression/array modulo keeps first element exact", [](std::string& why) {
                MathParser p;
                return expectEval(p, "a=(20,21)%(6,4); hex(a[0])", "0x2", why);
              }});
  t.push_back({"regression/array add keeps second element exact 2", [](std::string& why) {
                MathParser p;
                return expectEval(p, "a=(2,3)+(4,5); hex(a[1])", "0x8", why);
              }});
  t.push_back({"regression/array subtract keeps first element exact 2", [](std::string& why) {
                MathParser p;
                return expectEval(p, "a=(9,7)-(4,5); hex(a[0])", "0x5", why);
              }});
  t.push_back({"regression/array multiply keeps first element exact 2", [](std::string& why) {
                MathParser p;
                return expectEval(p, "a=(3,4)*(5,6); hex(a[0])", "0xF", why);
              }});
  t.push_back({"regression/array divide-by-one keeps second element exact 2", [](std::string& why) {
                MathParser p;
                return expectEval(p, "a=(8,9)/1; hex(a[1])", "0x9", why);
              }});
  t.push_back({"regression/array int keeps first element exact 2", [](std::string& why) {
                MathParser p;
                return expectEval(p, "a=int((8.9,9.1)); hex(a[0])", "0x8", why);
              }});
  t.push_back({"regression/array round keeps first element exact 2", [](std::string& why) {
                MathParser p;
                return expectEval(p, "a=round((8.2,9.8)); hex(a[0])", "0x8", why);
              }});
  t.push_back({"regression/array unary minus keeps first element exact 2", [](std::string& why) {
                MathParser p;
                return expectEval(p, "a=-(-5,-6); hex(a[0])", "0x5", why);
              }});
  t.push_back({"regression/array abs keeps third element exact", [](std::string& why) {
                MathParser p;
                return expectEval(p, "a=abs((-1,-2,-3)); hex(a[2])", "0x3", why);
              }});
  t.push_back({"regression/array bitor scalar keeps first element exact 2", [](std::string& why) {
                MathParser p;
                return expectEval(p, "a=(8,9)|1; hex(a[0])", "0x9", why);
              }});
  t.push_back({"regression/array bitand scalar keeps first element exact 2", [](std::string& why) {
                MathParser p;
                return expectEval(p, "a=(8,9)&7; hex(a[0])", "0x0", why);
              }});
  t.push_back({"regression/array shift-left keeps first element exact 2", [](std::string& why) {
                MathParser p;
                return expectEval(p, "a=(8,9)<<1; hex(a[0])", "0x10", why);
              }});
  t.push_back({"regression/array shift-right keeps first element exact 2", [](std::string& why) {
                MathParser p;
                return expectEval(p, "a=(8,9)>>1; hex(a[0])", "0x4", why);
              }});
  t.push_back({"regression/array modulo keeps second element exact 2", [](std::string& why) {
                MathParser p;
                return expectEval(p, "a=(20,21)%(6,4); hex(a[1])", "0x1", why);
              }});
  t.push_back({"regression/array int keeps third element exact", [](std::string& why) {
                MathParser p;
                return expectEval(p, "a=int((5.9,6.1,7.2)); hex(a[2])", "0x7", why);
              }});
  t.push_back({"regression/array round keeps third element exact", [](std::string& why) {
                MathParser p;
                return expectEval(p, "a=round((5.2,6.8,7.1)); hex(a[2])", "0x7", why);
              }});
  t.push_back({"regression/array abs keeps second element exact 2", [](std::string& why) {
                MathParser p;
                return expectEval(p, "a=abs((-5,-6,-7)); hex(a[1])", "0x6", why);
              }});
  t.push_back({"regression/array bitor scalar keeps second element exact 2", [](std::string& why) {
                MathParser p;
                return expectEval(p, "a=(8,9)|2; hex(a[1])", "0xB", why);
              }});
  t.push_back({"regression/array bitand scalar keeps second element exact 2", [](std::string& why) {
                MathParser p;
                return expectEval(p, "a=(8,9)&3; hex(a[1])", "0x1", why);
              }});
  t.push_back({"regression/array xor scalar keeps second element exact 2", [](std::string& why) {
                MathParser p;
                return expectEval(p, "a=(8,9)^3; hex(a[1])", "0xA", why);
              }});
  t.push_back({"regression/array add keeps first element exact 3", [](std::string& why) {
                MathParser p;
                return expectEval(p, "a=(10,11)+(1,2); hex(a[0])", "0xB", why);
              }});
  t.push_back({"regression/array subtract keeps second element exact 3", [](std::string& why) {
                MathParser p;
                return expectEval(p, "a=(10,11)-(1,2); hex(a[1])", "0x9", why);
              }});
  t.push_back({"regression/array multiply keeps second element exact 3", [](std::string& why) {
                MathParser p;
                return expectEval(p, "a=(10,11)*(2,3); hex(a[1])", "0x21", why);
              }});
  t.push_back({"regression/array divide-by-one keeps first element exact 3", [](std::string& why) {
                MathParser p;
                return expectEval(p, "a=(10,11)/1; hex(a[0])", "0xA", why);
              }});
  t.push_back({"regression/array int keeps second element exact 3", [](std::string& why) {
                MathParser p;
                return expectEval(p, "a=int((10.9,11.1)); hex(a[1])", "0xB", why);
              }});
  t.push_back({"regression/array round keeps second element exact 3", [](std::string& why) {
                MathParser p;
                return expectEval(p, "a=round((10.2,11.8)); hex(a[1])", "0xC", why);
              }});
  t.push_back({"regression/array abs keeps first element exact 3", [](std::string& why) {
                MathParser p;
                return expectEval(p, "a=abs((-10,-11)); hex(a[0])", "0xA", why);
              }});
  t.push_back({"regression/array unary minus keeps second element exact 3", [](std::string& why) {
                MathParser p;
                return expectEval(p, "a=-(-10,-11); hex(a[1])", "0xB", why);
              }});
  t.push_back({"regression/array bitor scalar keeps third element exact", [](std::string& why) {
                MathParser p;
                return expectEval(p, "a=(8,9,10)|1; hex(a[2])", "0xB", why);
              }});
  t.push_back({"regression/array bitand scalar keeps third element exact", [](std::string& why) {
                MathParser p;
                return expectEval(p, "a=(8,9,10)&3; hex(a[2])", "0x2", why);
              }});
  t.push_back({"regression/array xor scalar keeps third element exact", [](std::string& why) {
                MathParser p;
                return expectEval(p, "a=(8,9,10)^3; hex(a[2])", "0x9", why);
              }});
  t.push_back({"regression/array shift-right keeps second element exact 3", [](std::string& why) {
                MathParser p;
                return expectEval(p, "a=(10,11)>>1; hex(a[1])", "0x5", why);
              }});
  t.push_back({"regression/unique keeps first-seen order 1", [](std::string& why) {
                MathParser p;
                return expectEval(p, "unique((5,3,5,2,3))", "(5,3,2)", why);
              }});
  t.push_back({"regression/unique keeps first-seen order 2", [](std::string& why) {
                MathParser p;
                return expectEval(p, "unique((2,1,2,1,3,2))", "(2,1,3)", why);
              }});
  t.push_back({"regression/unique scalar returns single-item array", [](std::string& why) {
                MathParser p;
                return expectEval(p, "unique(7)", "(7)", why);
              }});
  t.push_back({"regression/unique mixed args flatten in order", [](std::string& why) {
                MathParser p;
                return expectEval(p, "unique((1,2),2,1,3)", "(1,2,3)", why);
              }});
  t.push_back({"regression/unique zeros collapse", [](std::string& why) {
                MathParser p;
                return expectEval(p, "unique((0,-0,0,0))", "(0)", why);
              }});
  t.push_back({"regression/unique preserves first negative sign representative", [](std::string& why) {
                MathParser p;
                return expectEval(p, "unique((-0,0,1))", "(0,1)", why);
              }});
  t.push_back({"regression/unique decimal duplicates collapse", [](std::string& why) {
                MathParser p;
                return expectEval(p, "unique((1.5,1.5,2.5,1.5))", "(1.5,2.5)", why);
              }});
  t.push_back({"regression/unique integer-valued doubles collapse", [](std::string& why) {
                MathParser p;
                return expectEval(p, "unique((2.0,2,2.000,3))", "(2,3)", why);
              }});
  t.push_back({"regression/unique with unpack keeps order", [](std::string& why) {
                MathParser p;
                return expectEval(p, "unique(unpack((4,1),(4,2),1))", "(4,1,2)", why);
              }});
  t.push_back({"regression/unique long run keeps first occurrence", [](std::string& why) {
                MathParser p;
                return expectEval(p, "unique((9,8,7,9,8,7,6,5,6))", "(9,8,7,6,5)", why);
              }});
  t.push_back({"regression/unique repeated clusters keep order", [](std::string& why) {
                MathParser p;
                return expectEval(p, "unique((3,3,3,3,2,2,1))", "(3,2,1)", why);
              }});
  t.push_back({"regression/median odd unsorted stays stable", [](std::string& why) {
                MathParser p;
                return expectEval(p, "median((9,1,5,3,7))", "5", why);
              }});
  t.push_back({"regression/median even unsorted stays stable", [](std::string& why) {
                MathParser p;
                return expectEval(p, "median((10,1,7,3))", "5", why);
              }});
  t.push_back({"regression/variance mixed signs stays stable", [](std::string& why) {
                MathParser p;
                return expectEval(p, "variance((-3,0,3,6))", "11.25", why);
              }});
  t.push_back({"regression/stddev mixed signs stays stable", [](std::string& why) {
                MathParser p;
                return expectEval(p, "stddev((-3,0,3,6))", "3.354101966249684", why);
              }});
  t.push_back({"regression/factorial 20 keeps exact integer formatting", [](std::string& why) {
                MathParser p;
                return expectEval(p, "hex(fact(20))", "0x21C3677C82B40000", why);
              }});
  t.push_back({"regression/factorial 30 remains available in floating mode", [](std::string& why) {
                MathParser p;
                return expectEval(p, "factorial(30)", "2.65252859812191e+032", why);
              }});
  t.push_back({"regression/sort keeps duplicates and order contract", [](std::string& why) {
                MathParser p;
                return expectEval(p, "sort((5,1,5,2,2,9,1))", "(1,1,2,2,5,5,9)", why);
              }});
  t.push_back({"regression/sorted alias keeps duplicates and order contract", [](std::string& why) {
                MathParser p;
                return expectEval(p, "sorted((4,4,3,2,2,1))", "(1,2,2,3,4,4)", why);
              }});
  t.push_back({"regression/aggregate avg mixed scalars-arrays stays stable", [](std::string& why) {
                MathParser p;
                return expectEval(p, "avg((2,4),6,8)", "5", why);
              }});
  t.push_back({"regression/aggregate min mixed scalars-arrays stays stable", [](std::string& why) {
                MathParser p;
                return expectEval(p, "min((5,7),3,9)", "3", why);
              }});
  t.push_back({"regression/long call argument list parses and evaluates", [](std::string& why) {
                MathParser p;
                return expectEval(p, "sum(1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16)", "136", why);
              }});
  t.push_back({"regression/aggregate max mixed scalars-arrays stays stable", [](std::string& why) {
                MathParser p;
                return expectEval(p, "max((5,7),3,9)", "9", why);
              }});
  t.push_back({"regression/unpack-heavy call list keeps evaluation order", [](std::string& why) {
                MathParser p;
                return expectEval(p, "sum(unpack((1,2,3),(4,5),6,7))", "28", why);
              }});
  t.push_back({"regression/sorted single-array duplicates stays stable", [](std::string& why) {
                MathParser p;
                return expectEval(p, "sorted((3,3,1,2,1))", "(1,1,2,3,3)", why);
              }});
  t.push_back({"regression/reversed single-array alias stays stable", [](std::string& why) {
                MathParser p;
                return expectEval(p, "reversed((1,2,3,4,5))", "(5,4,3,2,1)", why);
              }});
  t.push_back({"regression/call without unpack keeps args intact", [](std::string& why) {
                MathParser p;
                return expectEval(p, "f(x,y)=x+y; f(2,3)", "5", why);
              }});
  t.push_back({"regression/UDF parameter shadows global during body validation", [](std::string& why) {
                MathParser p;
                if (!expectEval(p, "x=0.5; f(x)=x<<2; f(3)", "12", why)) return false;
                if (!expectEvalErrorContains(p, "f(x)=x*2; f=10; f(3)", "unknown function: f", why))
                  return false;
                if (!expectEval(p, "f(x)=x*2; f=10; f", "10", why)) return false;
                if (!expectEvalErrorContains(p, "f=10; f(x)=x*2; f", "user-defined function: f(x)", why))
                  return false;
                if (!expectEvalErrorContains(p, "f(x)=x+1; f", "user-defined function: f(x)", why))
                  return false;
#if SMARTMATH_TIME_VALUES
                if (!expectEvalErrorContains(
                        p, "a(x)=1/x; sortby((0:30, 1:00, 0:45), a)", "incompatible operands", why))
                  return false;
#endif
                if (!expectEvalErrorContains(
                        p, "a(x)=1/a; sortby((0:30, 1:00, 0:45), a)", "unknown variable: a", why))
                  return false;
                return true;
              }});
  t.push_back({"regression/underscore reads as 1 before assignment", [](std::string& why) {
                MathParser p;
                if (!expectEval(p, "_", "1", why)) return false;
                return expectEval(p, "_+5", "6", why);
              }});
  t.push_back({"regression/UDF formal validation probe from _ default 1", [](std::string& why) {
                MathParser p;
                if (!p.addUserFunction("f(x)=1%x").empty()) {
                  why = "expected addUserFunction(f(x)=1%x) to succeed with default probe 1";
                  return false;
                }
                if (!expectEval(p, "f(7)", "1", why)) return false;
                MathParser p2;
                if (!expectEval(p2, "_=10", "10", why)) return false;
                if (!p2.addUserFunction("g(x)=1%(x-1)").empty()) {
                  why = "expected addUserFunction after _=10 to validate 1%(x-1)";
                  return false;
                }
                if (!expectEval(p2, "g(5)", "1", why)) return false;
                return true;
              }});
  t.push_back({"regression/late binding newfn definition then missing call then define newfn", [](std::string& why) {
                MathParser p;
                if (!expectEval(p, "f(x)=x*newfn(x)", "0", why)) return false;
                if (!expectEvalErrorContains(p, "f(2)", "unknown function: newfn", why)) return false;
                return expectEval(p, "newfn(x)=x**(1/3); f(8)", "16", why);
              }});
#if SMARTMATH_LAMBDA_FUNCTIONS
  t.push_back({"regression/udf reject empty tuple body", [](std::string& why) {
                MathParser p;
                if (!expectEvalErrorContains(p, "f=():()", "function body is empty", why)) return false;
                if (!expectEvalErrorContains(p, "f()=( )", "function body is empty", why)) return false;
                if (!expectEval(p, "f=():100; f()", "100", why)) return false;
                p.parseAndEvaluate("g=():(); g()");
                if (p.getError().empty()) {
                  why = "expected error after defining g=():()";
                  return false;
                }
                return true;
              }});
#endif
  t.push_back({"regression/udf formal parameter array indexing", [](std::string& why) {
                MathParser p;
                if (!expectEval(p, "_=(1,2,3,4); ff(a)=(a[1],a[-1]); ff(_)", "(2,4)", why)) return false;
                return expectEval(p, "t=(1,2,3,4,5,6); fff(a)=(a[0],a[-3]); fff(t)", "(1,4)", why);
              }});
  t.push_back({"regression/ident with digit: atan2 log10 bare hint and UDF f1", [](std::string& why) {
                MathParser p;
                if (!expectEvalErrorContains(p, "atan2", "function:", why)) return false;
                if (!expectEvalErrorContains(p, "atan2+1", "unknown variable: atan2", why)) return false;
                if (!expectEvalErrorContains(p, "log10", "function: log10(value)", why)) return false;
                if (!expectEvalErrorContains(p, "log10+1", "unknown variable: log10", why)) return false;
                if (!expectEval(p, "f1(x)=x+1", "0", why)) return false;
                if (!expectEvalErrorContains(p, "f1", "user-defined function:", why)) return false;
                if (!expectEvalErrorContains(p, "f1+2", "unknown variable: f1", why)) return false;
                if (!expectEvalErrorContains(p, "x_1+2", "unknown variable: x_1", why)) return false;
                return true;
              }});
  t.push_back({"regression/gcd lcm ncr npr array broadcast", [](std::string& why) {
                MathParser p;
                if (!expectEval(p, "gcd((84,30),6)", "(6,6)", why)) return false;
                if (!expectEval(p, "gcd(6,(84,30))", "(6,6)", why)) return false;
                if (!expectEval(p, "lcm((6,8),3)", "(6,24)", why)) return false;
                if (!expectEval(p, "lcm((6,8),(3,5))", "(6,40)", why)) return false;
                if (!expectEval(p, "ncr((5,6),(2,3))", "(10,20)", why)) return false;
                if (!expectEval(p, "npr((5,6),(2,3))", "(20,120)", why)) return false;
                if (!expectEvalErrorContains(p, "gcd((1,2),(3,4,5))", "incompatible operands", why))
                  return false;
                if (!expectEval(p, "6*(1/2)", "3", why)) return false;
                if (!expectEval(p, "6*(1/3)", "2", why)) return false;
                if (!expectEval(p, "5*(1/3)", "1.666666666666667", why)) return false;
                if (!expectEval(p, "0xFFFFFFFFFFFFFFFF*(1/5)", "3689348814741910323", why)) return false;
                if (!expectEval(p, "0x7FFFFFFFFFFFFFFF*(1/7)", "1317624576693539401", why)) return false;
                if (!expectEval(p, "2**64*(1/2)", "9.223372036854776e+018", why)) return false;
                if (!expectEval(p, "0x7FFFFFFFFFFFFFFF*(1/8)", "1.152921504606847e+018", why)) return false;
                if (!expectEval(p, "2**52*(1/3)", "1.501199875790165e+015", why)) return false;
                if (!expectEval(p, "-0x7FFFFFFFFFFFFFFF*(1/7)", "-1317624576693539401", why)) return false;
                if (!expectEval(p, "0x7FFFFFFFFFFFFFFF*(-1/7)", "-1317624576693539401", why)) return false;
                if (!expectEval(p, "-2**52*(1/2)", "-2251799813685248", why)) return false;
                if (!expectEval(p, "2**52*(-1/2)", "-2251799813685248", why)) return false;
                if (!expectEval(p, "1317624576693539401/(1/11)", "14493870343628933411", why)) return false;
                if (!expectEval(p, "1317624576693539401/(1/7)", "9223372036854775807", why)) return false;
                if (!expectEval(p, "-1317624576693539401/(1/7)", "-9223372036854775807", why)) return false;
                if (!expectEval(p, "1317624576693539401/(-1/7)", "-9223372036854775807", why)) return false;
                if (!expectEval(p, "1317624576693539401/(1/8)", "10540996613548315208", why)) return false;
                if (!expectEval(p, "5/(1/3)", "15", why)) return false;
                if (!expectEval(p, "-5/(1/3)", "-15", why)) return false;
                if (!expectEval(p, "5/(-1/3)", "-15", why)) return false;
                if (!expectEval(p, "0x7FFFFFFFFFFFFFFF/(1/8)", "7.37869762948382e+019", why)) return false;
                if (!expectEval(p, "2**64/(1/2)", "3.68934881474191e+019", why)) return false;
                if (!expectEval(p, "1317624576693539401*(1/11)", "1.197840524266854e+017", why)) return false;
                if (!expectEval(p, "0x7FFFFFFFFFFFFFFF*(-1/8)", "-1.152921504606847e+018", why)) return false;
                return expectEvalErrorContains(p, "ncr((5,6),(2,7))", "numeric error in ncr()", why);
              }});
#if SMARTMATH_FACTORINT
  t.push_back({"regression/factorint prime decomposition", [](std::string& why) {
                MathParser p;
                if (!expectEval(p, "factorint(13)", "(13)", why)) return false;
                if (!expectEval(p, "factorint(0)", "(0)", why)) return false;
                if (!expectEval(p, "factorint(-1)", "(-1)", why)) return false;
                if (!expectEval(p, "factorint(-12)", "(-2**2, 3)", why)) return false;
                if (!expectEval(p, "factorint(2**63-1)", "(7**2, 73, 127, 337, 92737, 649657)", why)) return false;
                if (!expectEval(p, "factorint(18446744073709551615)", "(3, 5, 17, 257, 641, 65537, 6700417)", why))
                  return false;
                if (!expectEval(p, "factorint(90)", "(2, 3**2, 5)", why)) return false;
                if (!expectEval(p, "factorint(9007)", "(9007)", why)) return false;
                if (!expectEval(p, "factorint(900719)", "(900719)", why)) return false;
                if (!expectEval(p, "factorint(90071992)", "(2**3, 11258999)", why)) return false;
                if (!expectEval(p, "factorint(9007199254)", "(2, 89, 50602243)", why)) return false;
                if (!expectEval(p, "factorint(900719925474)", "(2, 3, 12907, 11630897)", why)) return false;
                if (!expectEval(p, "factorint(76568758722)", "(2, 3**2, 47, 101, 896107)", why)) return false;
                if (!expectEval(p, "factorint(76568758722112367)", "(113, 677599634708959)", why)) return false;
                if (!expectEval(p, "factorint(-3333*9)", "(-3**3, 11, 101)", why)) return false;
                if (!expectEval(p, "factorint(-9999)", "(-3**2, 11, 101)", why)) return false;
                if (!expectEval(p, "prod(factorint(-33))", "-33", why)) return false;
                if (!expectEval(p, "prod(factorint(2**52))", "4503599627370496", why)) return false;
                if (!expectEvalErrorContains(p, "factorint(33.5)", "factorint() expects integer values", why))
                  return false;
                return true;
              }});
#endif
  t.push_back({"regression/binary builtin array length mismatch", [](std::string& why) {
                MathParser p;
                const char* exprs[] = {
                    "atan2((1,2),(3,4,5))",
                    "gcd((1,2),(3,4,5))",
                    "hypot((1,2),(3,4,5))",
                    "lcm((1,2),(3,4,5))",
                    "log((1,2),(3,4,5))",
                    "mod((1,2),(3,4,5))",
                    "ncr((1,2),(3,4,5))",
                    "npr((1,2),(3,4,5))",
                    "pow((1,2),(3,4,5))",
                };
                for (const char* expr : exprs) {
                  if (!expectEvalErrorContains(p, expr, "incompatible operands", why)) return false;
                }
                return true;
              }});
  t.push_back({"regression/incomplete function call open paren hints", [](std::string& why) {
                MathParser p;
                if (!expectEval(p, "f(x)=x", "0", why)) return false;
                if (!expectEvalErrorContains(p, "f(", "user-defined function: f(x)", why)) return false;
                if (!expectEvalErrorContains(p, "f (", "user-defined function: f(x)", why)) return false;
                if (!expectEvalErrorContains(p, "abs(", "function: abs(value)", why)) return false;
                if (!expectEvalErrorContains(p, "abs (", "function: abs(value)", why)) return false;
                if (!expectEvalErrorContains(p, "log(", "function: log(value, base)", why)) return false;
                if (!expectEvalErrorContains(p, "log (", "function: log(value, base)", why)) return false;
                if (!expectEvalErrorContains(p, "log10(", "function: log10(value)", why)) return false;
                if (!expectEvalErrorContains(p, "log10 (", "function: log10(value)", why)) return false;
                if (!expectEvalErrorContains(
                        p, "sortby((1,2,3), f(", "sortby expects exactly one function", why))
                  return false;
                if (!expectEvalErrorContains(
                        p, "sortby((1,2,3), abs(", "sortby expects exactly one function", why))
                  return false;
                return expectEvalErrorContains(
                    p, "sortby((1,2,3), log(", "sortby expects exactly one function", why);
              }});
  t.push_back({"regression/bare function name followed by immediate closer", [](std::string& why) {
                MathParser p;
                if (!expectEvalErrorContains(p, "sin)", "mismatched closing parenthesis", why)) return false;
                if (p.getError().find("function:") != std::string::npos) {
                  why = "sin): must not emit function hint";
                  return false;
                }
                if (!expectEvalErrorContains(p, "sin]", "mismatched closing bracket", why)) return false;
                if (p.getError().find("function:") != std::string::npos) {
                  why = "sin]: must not emit function hint";
                  return false;
                }
                if (!expectEvalErrorContains(p, "sin}", "mismatched closing brace", why)) return false;
                if (p.getError().find("function:") != std::string::npos) {
                  why = "sin}: must not emit function hint";
                  return false;
                }
                if (!expectEval(p, "f(x)=x", "0", why)) return false;
                if (!expectEvalErrorContains(p, "f(x)=x; f)", "mismatched closing parenthesis", why)) return false;
                if (p.getError().find("user-defined function:") != std::string::npos) {
                  why = "f(x)=x; f): must not emit UDF signature hint";
                  return false;
                }
                if (!expectEvalErrorContains(p, "sin)(1)", "mismatched closing parenthesis", why)) return false;
                if (p.getError().find("function:") != std::string::npos) {
                  why = "sin)(1): must not emit function hint";
                  return false;
                }
                return true;
              }});
  t.push_back({"regression/bare function name before semicolon is syntax error", [](std::string& why) {
                MathParser p;
                if (!expectEvalErrorContains(p, "atan2;", "unexpected token", why)) return false;
                if (!expectEvalErrorContains(p, "atan2;(2,3)", "unexpected token", why)) return false;
                if (!expectEval(p, "nnnn=5", "5", why)) return false;
                if (!expectEvalErrorContains(p, "atan2;", "unexpected token", why)) return false;
                if (p.getError().find("nnnn") != std::string::npos) {
                  why = "atan2; error must not reference a prior expression";
                  return false;
                }
                if (!expectEval(p, "f=():100; f()", "100", why)) return false;
                p.parseAndEvaluate("f=():100; f;");
                if (p.getError().find("unexpected token") == std::string::npos) {
                  why = "f; after define: expected unexpected token, got \"" + p.getError() + "\"";
                  return false;
                }
                return true;
              }});
#if SMARTMATH_LAMBDA_FUNCTIONS
  t.push_back({"regression/bare function name tail hint vs middle unknown variable", [](std::string& why) {
                MathParser p;
                if (!expectEvalErrorContains(p, "polar", "function:", why)) return false;
                if (!expectEvalErrorContains(p, "polar+123", "unknown variable: polar", why)) return false;
                if (!expectEvalErrorContains(p, "1+2; polar", "function:", why)) return false;
                if (!expectEvalErrorContains(p, "1+2; polar+2", "unknown variable: polar", why)) return false;
                if (!expectEvalErrorContains(p, "qqq(x)=polar", "function:", why)) return false;
                if (!expectEvalErrorContains(p, "qqq=x:polar", "function:", why)) return false;
                if (!expectEvalErrorContains(p, "qqq(x)=polar+2", "unknown variable: polar", why)) return false;
                if (!expectEvalErrorContains(p, "qqq(x)=sin+cos", "unknown variable: sin", why)) return false;
#if SMARTMATH_TIME_VALUES
                if (!expectEvalErrorContains(
                        p, "sortby((1:30,1:00),x:polar)", "unknown variable: polar", why))
                  return false;
                if (!expectEvalErrorContains(
                        p, "sortby((1:30,1:00),x:polar+2)", "unknown variable: polar", why))
                  return false;
#endif
                MathParser pf;
                pf.parseAndEvaluate("h(x)=x+1; h+2");
                if (pf.getError().find("unknown variable: h") == std::string::npos) {
                  why = "h(x)=x+1; h+2: expected unknown variable: h, got \"" + pf.getError() +
                        "\" res=\"" + pf.getResult() + "\"";
                  return false;
                }
                pf.parseAndEvaluate("h(x)=h+1");
                if (pf.getError().find("unknown variable: h") == std::string::npos) {
                  why = "h(x)=h+1: expected unknown variable: h, got \"" + pf.getError() + "\"";
                  return false;
                }
                MathParser pg;
                if (!expectEvalErrorContains(
                        pg, "g(x)=x; sortby((3,1,2),x:g)", "unknown variable: g", why))
                  return false;
                if (!expectEvalErrorContains(
                        pg, "g(x)=x; sortby((3,1,2),x:g+1)", "unknown variable: g", why))
                  return false;
                return expectEvalErrorContains(
                    pg, "sortby((1,2,3),x:x+sin)", "unknown variable: sin", why);
              }});
#endif
  t.push_back({"regression/bare function hint at expression end", [](std::string& why) {
                MathParser p;
                if (!expectEvalErrorContains(p, "1+sin", "function: sin(angle)", why)) return false;
                if (!expectEvalErrorContains(p, "1+sum", "function: sum(...)", why)) return false;
                if (!expectEvalErrorContains(p, "(1+2)/sum", "function: sum(...)", why)) return false;
                if (!expectEvalErrorContains(p, "1+sin(", "function: sin(angle)", why)) return false;
                if (!expectEvalErrorContains(p, "1+sin(rad(", "function: rad(...)", why)) return false;
                if (!expectEvalErrorContains(
                        p, "1+sin(rad(45)", "missing closing parenthesis", why))
                  return false;
                if (!expectEvalErrorContains(p, "sortby((1,2,3, abs", "function: abs(value)", why))
                  return false;
                if (!expectEval(p, "f(x)=x", "0", why)) return false;
                return expectEvalErrorContains(p, "f(x)=x; 1+f", "user-defined function: f(x)", why);
              }});
  t.push_back({"regression/semicolon suppresses function hints", [](std::string& why) {
                MathParser p;
                const auto mustNotHint = [&](const char* expr, const char* hintPart) -> bool {
                  p.parseAndEvaluate(expr);
                  const std::string err = p.getError();
                  if (err.empty()) {
                    why = std::string("expected error for \"") + expr + "\"";
                    return false;
                  }
                  if (err.find(hintPart) != std::string::npos) {
                    why = std::string("\"") + expr + "\" must not contain \"" + hintPart + "\", got \"" + err +
                          "\"";
                    return false;
                  }
                  return true;
                };
                if (!mustNotHint("sin(1)/cos;", "function: cos")) return false;
                if (!mustNotHint("sin(1)/cos; 2", "function: cos")) return false;
                if (!mustNotHint("1+sin;", "function: sin")) return false;
                if (!mustNotHint("1+sin(;", "function: sin")) return false;
                if (!mustNotHint("sin(;", "function: sin")) return false;
                if (!mustNotHint("cos;", "function: cos")) return false;
                if (!mustNotHint("sum(;", "function: sum")) return false;
                if (!expectEval(p, "f(x)=x", "0", why)) return false;
                if (!mustNotHint("f(x)=x; f(;", "user-defined function:")) return false;
                if (!expectEvalErrorContains(p, "f(x)=x; f(", "user-defined function: f(x)", why)) return false;
                return expectEvalErrorContains(p, "f(x)=x; 1+f", "user-defined function: f(x)", why);
              }});
  t.push_back({"regression/multi-statement trailing semicolon", [](std::string& why) {
                MathParser p;
                if (!expectEval(p, "a=1; b=2; c=3; a+b+c;", "6", why)) return false;
                if (!expectEval(p, "r=10; r;", "10", why)) return false;
                if (!expectEval(p, "r()=5; r();", "5", why)) return false;
                if (!expectEval(p, "1+2; 3+4", "7", why)) return false;
                if (!expectEval(p, "1+2; 3+4;", "7", why)) return false;
                if (!expectEval(p, "f(x)=x+1; f(3)", "4", why)) return false;
                if (!expectEval(p, "f(x)=x+1; f(3);", "4", why)) return false;
                if (!expectEval(p, "0xAA; hex", "0xAA", why)) return false;
                if (!expectEval(p, "0xAA; hex;", "0xAA", why)) return false;
                if (!expectEval(p, "(10,20,30); hex()", "(0xA,0x14,0x1E)", why)) return false;
                if (!expectEval(p, "(10,20,30); hex;", "(0xA,0x14,0x1E)", why)) return false;
                if (!expectEvalErrorContains(p, "(10,20,30); cos;", "unexpected token", why))
                  return false;
                return expectEvalErrorContains(p, "1;;2", "empty statement", why);
              }});
#if SMARTMATH_LAMBDA_FUNCTIONS
  t.push_back({"regression/sortby inline lambda key at EOF missing closing paren", [](std::string& why) {
                MathParser p;
                if (!expectEvalErrorContains(
                        p, "sortby((1,2,3),x:x", "missing closing parenthesis", why))
                  return false;
                if (!expectEvalErrorContains(
                        p, "sortby((1,2,3),x:x+sin", "missing closing parenthesis", why))
                  return false;
                return expectEvalErrorContains(
                    p, "sortby((1,2,3),x:x+sin(", "missing closing parenthesis", why);
              }});
#endif
  t.push_back({"regression/late binding unresolved referenced UDF reports unknown function", [](std::string& why) {
                MathParser p;
                return expectEvalErrorContains(p, "f(x)=x*p(x); f(2)", "unknown function", why);
              }});
  t.push_back({"regression/late binding resolves after referenced UDF definition", [](std::string& why) {
                MathParser p;
                return expectEval(p, "f(x)=x*p(x); p(x)=x+5; f(10)", "150", why);
              }});
  t.push_back({"regression/late binding with nonlinear referenced UDF", [](std::string& why) {
                MathParser p;
                return expectEval(p, "f(x)=x*p(x); p(x)=x**(1/3); f(8)", "16", why);
              }});
  t.push_back({"regression/late binding uses latest referenced UDF definition", [](std::string& why) {
                MathParser p;
                return expectEval(p, "f(x)=x*p(x); p(x)=x+5; p(x)=x**(1/3); f(8)", "16", why);
              }});
  t.push_back({"regression/UDF self-reference rejected at runtime", [](std::string& why) {
                MathParser p;
                return expectEvalErrorContains(p, "x(a)=x(a); x(1)", "recursive function call: x", why);
              }});
  t.push_back({"regression/mutual recursion y<->g rejected at runtime", [](std::string& why) {
                MathParser p;
                return expectEvalErrorContains(
                    p, "g(a)=y(a)+1; y(a)=g(a)+2; y(5)", "recursive function call", why);
              }});
  t.push_back({"regression/mutual recursion cycle a..d rejected at runtime", [](std::string& why) {
                MathParser p;
                return expectEvalErrorContains(
                    p,
                    "a(x)=b(x); b(x)=c(x); c(x)=d(x); d(x)=b(x); a(1)",
                    "recursive function call",
                    why);
              }});
  t.push_back({"regression/unique single-array variable keeps order", [](std::string& why) {
                MathParser p;
                return expectEval(p, "a=(5,2,5,9,2); unique(a)", "(5,2,9)", why);
              }});
  t.push_back({"regression/aggregate single scalar sum stays stable", [](std::string& why) {
                MathParser p;
                return expectEval(p, "sum(42)", "42", why);
              }});
  t.push_back({"regression/sorted single-array variable stays stable", [](std::string& why) {
                MathParser p;
                return expectEval(p, "a=(3,1,2); sorted(a)", "(1,2,3)", why);
              }});
  t.push_back({"regression/unique clustered duplicates stays stable", [](std::string& why) {
                MathParser p;
                return expectEval(p, "unique((1,1,2,2,3,3))", "(1,2,3)", why);
              }});
  t.push_back({"regression/aggregate unpack stream traversal stays stable", [](std::string& why) {
                MathParser p;
                return expectEval(p, "avg(unpack((1,2),(3,4)),10)", "4", why);
              }});
  t.push_back({"regression/reverse mixed flatten traversal stays stable", [](std::string& why) {
                MathParser p;
                return expectEval(p, "reverse((1,2),3,(4,5))", "(5,4,3,2,1)", why);
              }});
  t.push_back({"regression/user function unpack args normalization stays stable", [](std::string& why) {
                MathParser p;
                return expectEval(p, "f(a,b)=a*b; f(unpack((3,4)))", "12", why);
              }});
  t.push_back({"regression/unpack expansion keeps exact integer metadata", [](std::string& why) {
                MathParser p;
                return expectEval(p, "a=(9007199254740993); hex(unpack(a))", "0x20000000000001", why);
              }});
  t.push_back({"regression/unary minus array keeps large exact integer metadata", [](std::string& why) {
                MathParser p;
                return expectEval(p, "a=-(9007199254740991,5); uhex(a[0])", "0xFFE0000000000001", why);
              }});
  t.push_back({"regression/hex array rejects non-integer element contract", [](std::string& why) {
                MathParser p;
                return expectEvalErrorContains(p, "hex((1,2.5))", "hex() expects integer values", why);
              }});
  t.push_back({"regression/unique single-array fast path clustered order stable", [](std::string& why) {
                MathParser p;
                return expectEval(p, "unique((9,9,8,8,7))", "(9,8,7)", why);
              }});
  t.push_back({"regression/reverse single-array fast path order stable", [](std::string& why) {
                MathParser p;
                return expectEval(p, "reverse((9,8,7,6))", "(6,7,8,9)", why);
              }});
  t.push_back({"regression/deg variadic mixed args order stable", [](std::string& why) {
                MathParser p;
                return expectEval(p, "deg((pi/2),pi)", "(90,180)", why);
              }});
  t.push_back({"regression/uhex variadic array-scalar order stable", [](std::string& why) {
                MathParser p;
                return expectEval(p, "uhex((1,-1),2)", "(0x1,0xFFFFFFFFFFFFFFFF,0x2)", why);
              }});
  t.push_back({"regression/uhex unpack variadic array-scalar order stable", [](std::string& why) {
                MathParser p;
                return expectEval(p, "uhex(unpack((1,-1),2))", "(0x1,0xFFFFFFFFFFFFFFFF,0x2)", why);
              }});
  t.push_back({"regression/variance mixed variadic array-scalar stable", [](std::string& why) {
                MathParser p;
                return expectEval(p, "variance((1,2),3,4)", "1.25", why);
              }});
  t.push_back({"regression/stddev mixed variadic array-scalar stable", [](std::string& why) {
                MathParser p;
                return expectEval(p, "stddev((1,2),3,4)", "1.118033988749894", why);
              }});
  t.push_back({"regression/median single scalar fast path stable", [](std::string& why) {
                MathParser p;
                return expectEval(p, "median(42)", "42", why);
              }});
  t.push_back({"regression/median single array item fast path stable", [](std::string& why) {
                MathParser p;
                return expectEval(p, "median((42))", "42", why);
              }});

  return t;
}

std::vector<TestCase> buildSortbyRatioCases() {
  std::vector<TestCase> t;
  t.push_back({"sortby/ratio: ratio integers and zero",
               [](std::string& why) {
                 MathParser p;
                 if (!expectEval(p, "ratio(5)", "5", why)) return false;
                 if (!expectEval(p, "ratio(0)", "0", why)) return false;
                 if (!expectEval(p, "ratio(-4)", "-4", why)) return false;
                 return true;
               }});
  t.push_back({"sortby/ratio: ratio approximate and exact fractions",
               [](std::string& why) {
                 MathParser p;
                 if (!expectEval(p, "ratio(0.5)", "1/2", why)) return false;
                 if (!expectEval(p, "ratio(0.3333333333333333)", "1/3", why)) return false;
                 if (!expectEval(p, "ratio(1.5)", "3/2", why)) return false;
                 if (!expectEval(p, "ratio(1e-7)", "1/10000000", why)) return false;
                 if (!expectEval(p, "ratio(1e-8)", "1/100000000", why)) return false;
                 if (!expectEval(p, "ratio(sqrt(2))", "13250218/9369319", why)) return false;
                 if (!expectEval(p, "ratio((sqrt(2), e, pi))",
                                 "(13250218/9369319, 14665106/5394991, 5419351/1725033)",
                                 why)) {
                   return false;
                 }
                 return true;
               }});
  t.push_back({"sortby/ratio: ratio nan inf and array",
               [](std::string& why) {
                 MathParser p;
                 if (!expectEval(p, "ratio(nan)", "nan", why)) return false;
                 if (!expectEval(p, "ratio(inf)", "inf", why)) return false;
                 if (!expectEval(p, "ratio((1,2,3))", "(1,2,3)", why)) return false;
                 if (!expectEval(p, "ratio((0.5, 1/3))", "(1/2, 1/3)", why)) return false;
                 return true;
               }});
  t.push_back({"sortby/ratio: ratio edge cases (cross-multiply error path)",
               [](std::string& why) {
                 MathParser p;
                 if (!expectEval(p, "ratio(0.14285714285714285)", "1/7", why)) return false;
                 if (!expectEval(p, "ratio(-0.3333333333333333)", "-1/3", why)) return false;
                 if (!expectEval(p, "ratio(1234567)", "1234567", why)) return false;
                 if (!expectEval(p, "ratio(0.0000001)", "1/10000000", why)) return false;
                 if (!expectEval(p, "ratio(0.4142135623730951)", "3880899/9369319", why)) return false;
                 if (!expectEval(p, "ratio((0.14285714285714285, sqrt(2)))", "(1/7, 13250218/9369319)", why))
                   return false;
                 if (!expectEval(p, "ratio(0.9999999)", "9999999/10000000", why)) return false;
                 if (!expectEval(p, "ratio(0.123456789012345)", "10/81", why)) return false;
                 return expectEval(p, "ratio(1/7)", "1/7", why);
               }});
#if SMARTMATH_TIME_VALUES
  t.push_back({"sortby/ratio: ratio rejects time",
               [](std::string& why) {
                 MathParser p;
                 return expectEvalErrorContains(p, "ratio(1:00)", "incompatible operands", why);
               }});
#endif
  t.push_back({"sortby/ratio: sortby stable by abs",
               [](std::string& why) {
                 MathParser p;
                 if (!expectEval(p, "sortby((3,-1,2), abs)", "(-1,2,3)", why)) return false;
                 if (!expectEval(p, "sortby((5,1,5,2,2,9,1), abs)", "(1,1,2,2,5,5,9)", why)) return false;
                 return true;
               }});
  t.push_back({"sortby/ratio: empty array rejects like sort",
               [](std::string& why) {
                 MathParser p;
                 if (!expectEvalErrorContains(p, "sortby((), abs)", "expects at least 1 argument", why)) return false;
                 return expectEvalErrorContains(p, "ratio(())", "expects at least 1 argument", why);
               }});
  t.push_back({"sortby/ratio: sortby parse and arity errors",
               [](std::string& why) {
                 MathParser p;
                 if (!expectEvalErrorContains(p, "sortby((1,2), abs())", "sortby expects exactly one function", why))
                   return false;
                 if (!expectEvalErrorContains(p, "sortby((1,2))", "sortby expects a function that takes 1 parameter", why))
                   return false;
                 if (!expectEvalErrorContains(p, "sortby((1,2), 3)", "sortby expects a function that takes 1 parameter", why))
                   return false;
                 if (!expectEvalErrorContains(p, "sortby((1,2), (1,2))", "sortby expects exactly one function", why))
                   return false;
                 if (!expectEvalErrorContains(p, "sortby((1,2), pow)", "sortby expects a function that takes 1 parameter", why))
                   return false;
                 if (!expectEvalErrorContains(p, "sortby((1,2), rand)", "sortby expects a function that takes 1 parameter", why))
                   return false;
                 return true;
               }});
  t.push_back({"sortby/ratio: sortby unknown and UDF unary",
               [](std::string& why) {
                 MathParser p;
                 if (!expectEvalErrorContains(p, "sortby((1,2), nosuchfn)", "unknown function: nosuchfn", why))
                   return false;
#if SMARTMATH_TIME_VALUES
                if (!expectEvalErrorContains(
                        p, "x=10; sortby((0:30, 1:00, 0:45), x)",
                        "sortby expects a function that takes 1 parameter", why))
                  return false;
#endif
                 p.addUserFunction("f(x)=x*x");
                 if (!expectEval(p, "sortby((3,1,2), f)", "(1,2,3)", why)) return false;
                 return true;
               }});
  t.push_back({"sortby/ratio: sortby polar and UDF array keys (lexicographic compare)",
               [](std::string& why) {
                 MathParser p;
                 if (!expectEval(p, "sortby((2,1), polar)", "(1,2)", why)) return false;
                 if (!expectEval(p, "f(x)=x*(10,20); sortby((3,1,2), f)", "(1,2,3)", why)) return false;
                 return true;
               }});
#if SMARTMATH_LAMBDA_FUNCTIONS
  t.push_back({"sortby/ratio: lambda assignment and sortby anonymous lambdas",
               [](std::string& why) {
                 MathParser p;
                 if (!expectEval(p, "f=x:x+2; f(3)", "5", why)) return false;
                 if (!expectEval(p, "f=(x):x+2; f(4)", "6", why)) return false;
                 if (!expectEval(p, "f=():42; f()", "42", why)) return false;
                 if (!expectEval(p, "f=():100; f()", "100", why)) return false;
                 if (!expectEval(p, "f=x,y:x+y; f(1,2)", "3", why)) return false;
                 if (!expectEval(p, "f=x,y,z:x+y+z; f(1,2,3)", "6", why)) return false;
                 if (!expectEval(p, "f=(x,y,z):(x+y+z); f(1,2,3)", "6", why)) return false;
                 if (!expectEval(p, "a=10; a", "10", why)) return false;
                 if (!expectEval(p, "sortby((3,1,2), x:-x)", "(3,2,1)", why)) return false;
                 if (!expectEvalErrorContains(
                         p, "sortby((5,1), ():1)", "sortby expects a function that takes 1 parameter", why))
                   return false;
                 if (!expectEvalErrorContains(
                         p, "sortby((1,2), x,y:x+y)", "sortby expects a function that takes 1 parameter", why))
                   return false;
                 if (!expectEval(p, "g=((x):(1/x)); g(2)", "0.5", why)) return false;
                 if (!expectEval(p, "sortby((2,1), (x):(1/x))", "(2,1)", why)) return false;
#if SMARTMATH_TIME_VALUES
                if (!expectEvalErrorContains(
                        p, "sortby((1:30,1:00),x:)", "function body is empty", why))
                  return false;
#endif
                 return true;
              }});
#endif
#if SMARTMATH_TIME_VALUES
  t.push_back({"sortby/ratio: sortby time and milliseconds keys",
               [](std::string& why) {
                 MathParser p;
                 if (!expectEval(p, "sortby((1:30,0:30,1:00), milliseconds)", "(00:30,01:00,01:30)", why)) return false;
                 return true;
               }});
  t.push_back({"sortby/ratio: UDF divide duration via sortby",
               [](std::string& why) {
                 MathParser p;
                 return expectEvalErrorContains(
                     p, "f(x)=1/x; sortby((0:30, 1:00, 0:45), f)", "incompatible operands", why);
               }});
#endif
  return t;
}

std::vector<TestCase> buildRatioInExpressionCases() {
  std::vector<TestCase> t;
  t.push_back({"ratio/expr: arithmetic and unary on ratio values",
               [](std::string& why) {
                 MathParser p;
                 if (!expectEval(p, "ratio(1.3456)+1", "2.3456", why)) return false;
                 if (!expectEval(p, "ratio(1.3456)*10", "13.456", why)) return false;
                 if (!expectEval(p, "ratio(1.3456)-1", "0.3456", why)) return false;
                 if (!expectEval(p, "int(ratio(1.3456))", "1", why)) return false;
                 if (!expectEval(p, "trunc(ratio(1.3456))", "1", why)) return false;
                 if (!expectEval(p, "floor(ratio(1.3456))", "1", why)) return false;
                 if (!expectEval(p, "round(ratio(1.3456))", "1", why)) return false;
                 if (!expectEval(p, "ratio(1.3456)-ratio(1.3456)", "0", why)) return false;
                 return true;
               }});
  t.push_back({"ratio/expr: aggregates, builtins, and UDF",
               [](std::string& why) {
                 MathParser p;
                 if (!expectEval(p, "prod(ratio(1.3456),1)", "1.3456", why)) return false;
                 if (!expectEval(p, "product(ratio(1.3456),1)", "1.3456", why)) return false;
                 if (!expectEval(p, "sum(ratio(0.5),ratio(0.5))", "1", why)) return false;
                 if (!expectEval(p, "min(ratio(1.3456), 2)", "841/625", why)) return false;
                 if (!expectEval(p, "max(ratio(1.3456), 2)", "2", why)) return false;
                 if (!expectEval(p, "abs(ratio(-0.5))", "0.5", why)) return false;
                 if (!expectEval(p, "sign(ratio(-0.5))", "-1", why)) return false;
                 if (!expectEval(p, "clamp(ratio(1.3456),0,2)", "841/625", why)) return false;
                 p.addUserFunction("f(x)=x-1");
                 if (!expectEval(p, "f(ratio(1.3456))", "0.3456", why)) return false;
                 p.addUserFunction("g(x)=x*2");
                 return expectEval(p, "g(ratio(0.25))", "0.5", why);
               }});
  t.push_back({"ratio/expr: structure-preserving and sort paths",
               [](std::string& why) {
                 MathParser p;
                 if (!expectEval(p, "reverse(ratio(1/2),ratio(1))", "(1, 1/2)", why)) return false;
                 if (!expectEval(p, "unpack(ratio(1),ratio(1/2))", "(1, 1/2)", why)) return false;
                 if (!expectEval(p, "unique(ratio(1),ratio(0.5),ratio(1))", "(1, 1/2)", why)) return false;
                 if (!expectEval(p, "sort((ratio(3),ratio(1),ratio(1/2)))", "(1/2, 1, 3)", why)) return false;
#if SMARTMATH_LAMBDA_FUNCTIONS
                 if (!expectEval(p, "sortby((ratio(3),ratio(1/2)),x:x)", "(1/2, 3)", why)) return false;
#endif
                 if (!expectEval(p, "sortby((ratio(3),ratio(1/2)),abs)", "(1/2, 3)", why)) return false;
                 if (!expectEval(p, "(ratio(0.5),ratio(0.25))", "(1/2, 1/4)", why)) return false;
#if SMARTMATH_COMPLEX_NUMBERS
                 p.setSupportComplexNumbers(true);
                 if (!expectEval(p, "ratio(0.5+0.25i)+1", "1.5+0.25i", why)) return false;
#endif
                 return true;
               }});
  return t;
}

std::vector<TestCase> buildMinMaxPreserveWinnerCases() {
  std::vector<TestCase> t;
  t.push_back({"minmax/winner: mixed float keeps exact integer display",
               [](std::string& why) {
                 MathParser p;
                 if (!expectEval(p, "max(1,2**54,2)", "18014398509481984", why)) return false;
                 if (!expectEval(p, "max(1.1,2**54,2)", "18014398509481984", why)) return false;
                 if (!expectEval(p, "max(2**54,1.1,2)", "18014398509481984", why)) return false;
                 if (!expectEval(p, "min(3,1,2.1)", "1", why)) return false;
                 if (!expectEval(p, "min(1.0,1)", "1", why)) return false;
                 return expectEval(p, "min(1,1.0)", "1", why);
               }});
  t.push_back({"minmax/winner: arrays and ratio operands",
               [](std::string& why) {
                 MathParser p;
                 if (!expectEval(p, "min((3,1,2))", "1", why)) return false;
                 if (!expectEval(p, "max((1,2**54,2))", "18014398509481984", why)) return false;
                 if (!expectEval(p, "min(ratio(0.5),1)", "1/2", why)) return false;
                 return expectEval(p, "max(ratio(1.3456),2)", "2", why);
               }});
  t.push_back({"minmax/winner: NaN skipped when non-NaN present",
               [](std::string& why) {
                 MathParser p;
                 if (!expectEval(p, "max(1,nan)", "1", why)) return false;
                 if (!expectEval(p, "min(nan,1)", "1", why)) return false;
                 if (!expectEval(p, "min(nan,1,2)", "1", why)) return false;
                 if (!expectEval(p, "max(nan,1,2)", "2", why)) return false;
                 if (!expectEval(p, "min(nan,nan)", "nan", why)) return false;
                 return expectEval(p, "max(nan,nan,nan)", "nan", why);
               }});
  t.push_back({"minmax/winner: all-integer ties",
               [](std::string& why) {
                 MathParser p;
                 return expectEval(p, "min(1,1,1)", "1", why);
               }});
  return t;
}

std::vector<TestCase> buildClampPreserveIntegerMetadataCases() {
  std::vector<TestCase> t;
  t.push_back({"clamp/integer: array and scalar bounds preserve exact integers",
               [](std::string& why) {
                 MathParser p;
                 if (!expectEval(p, "clamp((0,1,2,3),1,2)", "(1, 1, 2, 2)", why)) return false;
                 if (!expectEval(p, "clamp(0,1,2)", "1", why)) return false;
                 if (!expectEval(p, "clamp(3,1,2)", "2", why)) return false;
                 if (!expectEval(p, "clamp(1,1,2)", "1", why)) return false;
                 return expectEval(p, "clamp(1.5,1,2)", "1.5", why);
               }});
  return t;
}

std::vector<TestCase> buildClampInvertedBoundsCases() {
  std::vector<TestCase> t;
  t.push_back({"clamp/inverted bounds: value at min when min > max",
               [](std::string& why) {
                 MathParser p;
                 if (!expectEval(p, "clamp(4,4,2)", "4", why)) return false;
                 if (!expectEval(p, "clamp(5,4,2)", "2", why)) return false;
                 return expectEval(p, "clamp(4.01,4,2)", "2", why);
               }});
  return t;
}

std::vector<TestCase> buildClampNanBoundCases() {
  std::vector<TestCase> t;
  t.push_back({"clamp/NaN bounds: absent lower bound, unknown upper bound",
               [](std::string& why) {
                 MathParser p;
                 if (!expectEval(p, "clamp(nan,1,2)", "1", why)) return false;
                 if (!expectEval(p, "clamp(nan,nan,2)", "nan", why)) return false;
                 if (!expectEval(p, "clamp(nan,1,nan)", "1", why)) return false;
                 if (!expectEval(p, "clamp(0,1,nan)", "1", why)) return false;
                 if (!expectEval(p, "clamp(2,1,nan)", "nan", why)) return false;
                 if (!expectEval(p, "clamp(1,nan,2)", "1", why)) return false;
                 if (!expectEval(p, "clamp(3,nan,2)", "2", why)) return false;
                 if (!expectEval(p, "clamp(2,nan,2)", "2", why)) return false;
                 if (!expectEval(p, "clamp(-inf,nan,2)", "-inf", why)) return false;
                 if (!expectEval(p, "clamp(-inf,2,nan)", "2", why)) return false;
                 if (!expectEval(p, "clamp(inf,2,nan)", "nan", why)) return false;
                 if (!expectEval(p, "clamp(nan,inf,inf)", "inf", why)) return false;
                 if (!expectEval(p, "clamp(nan,-inf,1)", "-inf", why)) return false;
                 if (!expectEval(p, "clamp(nan,-inf,inf)", "-inf", why)) return false;
                 if (!expectEval(p, "clamp(inf,nan,2)", "2", why)) return false;
                 if (!expectEval(p, "clamp(-inf,-inf,nan)", "-inf", why)) return false;
                 if (!expectEval(p, "clamp(inf,inf,nan)", "inf", why)) return false;
                 return expectEval(p, "clamp(1,1,nan)", "1", why);
               }});
  return t;
}

// Parity port from Basic smoke tests (SmokeTest_MathParser.bas) for expression-evaluation coverage.
// This keeps C++ independent: only literal expr/expected/error-substring strings are copied in.
struct ParityBasicCase {
  enum class Kind { Expected, ErrorContains } kind;
  const char* expr;
  const char* payload;
};

static const ParityBasicCase kParityBasicFromSmokeCases[] = {
    {ParityBasicCase::Kind::Expected, "16**-0.5", "0.25"} ,
    {ParityBasicCase::Kind::Expected, "+5", "5"} ,
    {ParityBasicCase::Kind::Expected, "-5", "-5"} ,
    {ParityBasicCase::Kind::Expected, "~5", "-6"} ,
    {ParityBasicCase::Kind::Expected, "5%3", "2"} ,
    {ParityBasicCase::Kind::Expected, "200 + 15%", "230"} ,
    {ParityBasicCase::Kind::Expected, "200 - 15%", "170"} ,
    {ParityBasicCase::Kind::Expected, "8>>1", "4"} ,
    {ParityBasicCase::Kind::Expected, "3<<2", "12"} ,
    {ParityBasicCase::Kind::Expected, "6&3", "2"} ,
    {ParityBasicCase::Kind::Expected, "6^3", "5"} ,
    {ParityBasicCase::Kind::Expected, "6|3", "7"} ,
    {ParityBasicCase::Kind::Expected, "2(3+4)", "14"} ,
    {ParityBasicCase::Kind::Expected, "2(3+4)**2", "98"} ,
    {ParityBasicCase::Kind::Expected, "2+3<<1", "10"} ,
    {ParityBasicCase::Kind::Expected, "1|2^3&6<<1", "3"} ,
    {ParityBasicCase::Kind::Expected, "2(1+2)%4", "2"} ,
    {ParityBasicCase::Kind::ErrorContains, "5.5&1", "bitwise operands must be integer values"} ,
    {ParityBasicCase::Kind::ErrorContains, "5|1.1", "bitwise operands must be integer values"} ,
    {ParityBasicCase::Kind::ErrorContains, "3.2^1", "bitwise operands must be integer values"} ,
    {ParityBasicCase::Kind::ErrorContains, "8.1>>1", "bitwise operands must be integer values"} ,
    {ParityBasicCase::Kind::ErrorContains, "8<<1.2", "bitwise operands must be integer values"} ,
    {ParityBasicCase::Kind::ErrorContains, "~2.5", "bitwise operands must be integer values"} ,
    {ParityBasicCase::Kind::ErrorContains, "5.5%2", "modulo operands must be integer values"} ,
    {ParityBasicCase::Kind::ErrorContains, "5%2.2", "modulo operands must be integer values"} ,
    {ParityBasicCase::Kind::ErrorContains, "2(1+2.5)%4.2", "modulo operands must be integer values"} ,
    {ParityBasicCase::Kind::Expected, "pow(2,3)", "8"} ,
    {ParityBasicCase::Kind::Expected, "prod(2,3,4)", "24"} ,
    {ParityBasicCase::Kind::ErrorContains, "pow", "function: pow("} ,
    {ParityBasicCase::Kind::ErrorContains, "sin", "function: sin(angle)"} ,
    {ParityBasicCase::Kind::ErrorContains, "sum", "function: sum(...)" } ,
    {ParityBasicCase::Kind::Expected, "sqr(5)", "25"} ,
    {ParityBasicCase::Kind::ErrorContains, "sqr", "function: sqr(value)"} ,
    {ParityBasicCase::Kind::Expected, "1 + 2 # calculates 1 + 2", "3"} ,
    {ParityBasicCase::Kind::Expected, "// this entire line is a comment", ""} ,
    {ParityBasicCase::Kind::Expected, "sin(pi/2) // calculates sin(pi/2)", "1"} ,
    {ParityBasicCase::Kind::Expected, "sin(-pi/2) // calculates sin(-pi/2)", "-1"} ,
    {ParityBasicCase::Kind::Expected, "sin(77777*pi/2) // calculates sin(77777*pi/2)", "1"} ,
    {ParityBasicCase::Kind::Expected, "sin(-77777*pi/2) // calculates sin(-77777*pi/2)", "-1"} ,
    {ParityBasicCase::Kind::Expected, "sin(0) // calculates sin(0)", "0"} ,
    {ParityBasicCase::Kind::Expected, "sin(pi) // calculates sin(pi)", "0"} ,
    {ParityBasicCase::Kind::Expected, "sin(-pi) // calculates sin(-pi)", "0"} ,
    {ParityBasicCase::Kind::Expected, "sin(2*pi) // calculates sin(2*pi)", "0"} ,
    {ParityBasicCase::Kind::Expected, "sin(-2*pi) // calculates sin(-2*pi)", "0"} ,
    {ParityBasicCase::Kind::Expected, "sin(77777*pi) // calculates sin(77777*pi)", "0"} ,
    {ParityBasicCase::Kind::Expected, "sin(-77777*pi) // calculates sin(-77777*pi)", "0"} ,
    {ParityBasicCase::Kind::Expected, "sin(77778*pi) // calculates sin(77778*pi)", "0"} ,
    {ParityBasicCase::Kind::Expected, "sin(-77778*pi) // calculates sin(-77778*pi)", "0"} ,
    {ParityBasicCase::Kind::Expected, "sin((-77778*pi, 77778*pi, -77777*pi, 77777*pi, -2*pi, 2*pi, -pi, pi, 0, -77777*pi/2, 77777*pi/2, -pi/2, pi/2))", "(0, 0, 0, 0, 0, 0, 0, 0, 0, -1, 1, -1, 1)"} ,
    {ParityBasicCase::Kind::Expected, "cos(pi/2) // calculates cos(pi/2)", "0"} ,
    {ParityBasicCase::Kind::Expected, "cos(-pi/2) // calculates cos(-pi/2)", "0"} ,
    {ParityBasicCase::Kind::Expected, "cos(77777*pi/2) // calculates cos(77777*pi/2)", "0"} ,
    {ParityBasicCase::Kind::Expected, "cos(-77777*pi/2) // calculates cos(-77777*pi/2)", "0"} ,
    {ParityBasicCase::Kind::Expected, "cos(0) // calculates cos(0)", "1"} ,
    {ParityBasicCase::Kind::Expected, "cos(pi) // calculates cos(pi)", "-1"} ,
    {ParityBasicCase::Kind::Expected, "cos(-pi) // calculates cos(-pi)", "-1"} ,
    {ParityBasicCase::Kind::Expected, "cos(2*pi) // calculates cos(2*pi)", "1"} ,
    {ParityBasicCase::Kind::Expected, "cos(-2*pi) // calculates cos(-2*pi)", "1"} ,
    {ParityBasicCase::Kind::Expected, "cos(77777*pi) // calculates cos(77777*pi)", "-1"} ,
    {ParityBasicCase::Kind::Expected, "cos(-77777*pi) // calculates cos(-77777*pi)", "-1"} ,
    {ParityBasicCase::Kind::Expected, "cos(77778*pi) // calculates cos(77778*pi)", "1"} ,
    {ParityBasicCase::Kind::Expected, "cos(-77778*pi) // calculates cos(-77778*pi)", "1"} ,
    {ParityBasicCase::Kind::Expected, "cos((-77778*pi, 77778*pi, -77777*pi, 77777*pi, -2*pi, 2*pi, -pi, pi, 0, -77777*pi/2, 77777*pi/2, -pi/2, pi/2))", "(1, 1, -1, -1, 1, 1, -1, -1, 1, 0, 0, 0, 0)"} ,
    {ParityBasicCase::Kind::Expected, "tan(pi/2) // calculates tan(pi/2)", "Inf"} ,
    {ParityBasicCase::Kind::Expected, "tan(-pi/2) // calculates tan(-pi/2)", "-Inf"} ,
    {ParityBasicCase::Kind::Expected, "tan(77777*pi/2) // calculates tan(77777*pi/2)", "Inf"} ,
    {ParityBasicCase::Kind::Expected, "tan(-77777*pi/2) // calculates tan(-77777*pi/2)", "-Inf"} ,
    {ParityBasicCase::Kind::Expected, "tan(0) // calculates tan(0)", "0"} ,
    {ParityBasicCase::Kind::Expected, "tan(pi) // calculates tan(pi)", "0"} ,
    {ParityBasicCase::Kind::Expected, "tan(-pi) // calculates tan(-pi)", "0"} ,
    {ParityBasicCase::Kind::Expected, "tan(2*pi) // calculates tan(2*pi)", "0"} ,
    {ParityBasicCase::Kind::Expected, "tan(-2*pi) // calculates tan(-2*pi)", "0"} ,
    {ParityBasicCase::Kind::Expected, "tan(77777*pi) // calculates tan(77777*pi)", "0"} ,
    {ParityBasicCase::Kind::Expected, "tan(-77777*pi) // calculates tan(-77777*pi)", "0"} ,
    {ParityBasicCase::Kind::Expected, "tan(77778*pi) // calculates tan(77778*pi)", "0"} ,
    {ParityBasicCase::Kind::Expected, "tan(-77778*pi) // calculates tan(-77778*pi)", "0"} ,
    {ParityBasicCase::Kind::Expected, "tan((-77778*pi, 77778*pi, -77777*pi, 77777*pi, -2*pi, 2*pi, -pi, pi, 0, -77777*pi/2, 77777*pi/2, -pi/2, pi/2))", "(0, 0, 0, 0, 0, 0, 0, 0, 0, -Inf, Inf, -Inf, Inf)"} ,
    {ParityBasicCase::Kind::Expected, "cos((pi/2, pi, -pi/2, -pi, 2*pi, 0, 77777*pi/2, -77777*pi/2))", "(0, -1, 0, -1, 1, 1, 0, 0)"} ,
    {ParityBasicCase::Kind::Expected, "tan((pi, pi/2, -pi, 2*pi, -pi/2, 0, 77777*pi/2, -77777*pi/2))", "(0, Inf, 0, 0, -Inf, 0, Inf, -Inf)"} ,
    {ParityBasicCase::Kind::ErrorContains, "[]", "unexpected token"} ,
    {ParityBasicCase::Kind::ErrorContains, "b=2; 2b", "unexpected token"} ,
    {ParityBasicCase::Kind::ErrorContains, "(2;2b;3)", "missing closing parenthesis"} ,
    {ParityBasicCase::Kind::ErrorContains, "(2,2b,3)", "unexpected token"} ,
    {ParityBasicCase::Kind::ErrorContains, "(2+3x)", "unexpected token"} ,
    {ParityBasicCase::Kind::ErrorContains, "(2+3t)", "unexpected token"} ,
    {ParityBasicCase::Kind::Expected, "hex(12)", "0xC"} ,
    {ParityBasicCase::Kind::Expected, "hex((12,255))", "(0xC,0xFF)"} ,
    {ParityBasicCase::Kind::Expected, "10 + hex(12) + 14", "36"} ,
    {ParityBasicCase::Kind::ErrorContains, "hex(12.5)", "hex() expects integer values"} ,
    {ParityBasicCase::Kind::ErrorContains, "hex", "function: hex(...)"} ,
    {ParityBasicCase::Kind::ErrorContains, "0x", "invalid hex literal"} ,
    {ParityBasicCase::Kind::ErrorContains, "0xG", "invalid hex literal"} ,
    {ParityBasicCase::Kind::Expected, "hex(0x7FFFFFFFFFFFFFFF)", "0x7FFFFFFFFFFFFFFF"} ,
    {ParityBasicCase::Kind::Expected, "hex(0xFFFFFFFFFFFFFFFF)", "0xFFFFFFFFFFFFFFFF"} ,
    {ParityBasicCase::Kind::Expected, "0b01110110011", "947"} ,
    {ParityBasicCase::Kind::Expected, "bin(13)", "0b1101"} ,
    {ParityBasicCase::Kind::Expected, "bin((1,2,5))", "(0b1,0b10,0b101)"} ,
    {ParityBasicCase::Kind::Expected, "10 + bin(12) + 14", "36"} ,
    {ParityBasicCase::Kind::ErrorContains, "0b", "invalid binary literal"} ,
    {ParityBasicCase::Kind::ErrorContains, "bin(12.5)", "bin() expects integer values"} ,
    {ParityBasicCase::Kind::ErrorContains, "bin", "function: bin(...)"} ,
    {ParityBasicCase::Kind::Expected, "9007199254740991+1", "9007199254740992"} ,
    {ParityBasicCase::Kind::Expected, "9007199254740992+1", "9007199254740993"} ,
    {ParityBasicCase::Kind::Expected, "3037000499*3037000499", "9223372030926249001"} ,
    {ParityBasicCase::Kind::Expected, "5/2", "2.5"} ,
    {ParityBasicCase::Kind::Expected, "2**10+1", "1025"} ,
    {ParityBasicCase::Kind::Expected, "2**-1", "0.5"} ,
    {ParityBasicCase::Kind::Expected, "9007199254740993&1", "1"} ,
    {ParityBasicCase::Kind::Expected, "9223372036854775807+1", "9223372036854775808"} ,
    {ParityBasicCase::Kind::Expected, "-9223372036854775808-1", "-9.223372036854778e+018"} ,
    {ParityBasicCase::Kind::Expected, "3037000500*3037000500", "9.223372037000249e+018"} ,
    {ParityBasicCase::Kind::Expected, "2**63", "9223372036854775808"} ,
    {ParityBasicCase::Kind::Expected, "2**64", "1.844674407370955e+019"} ,
    {ParityBasicCase::Kind::Expected, "9223372036854775807+0.5", "9.223372036854778e+018"} ,
    {ParityBasicCase::Kind::Expected, "hex(9223372036854775807+1)", "0x8000000000000000"} ,
    {ParityBasicCase::Kind::Expected, "(2**58, 2**58+123)", "(288230376151711744, 288230376151711867)"} ,
    //{ParityBasicCase::Kind::Expected, "(2**58, 2**58+123);hex", "(0x400000000000000, 0x40000000000007B)"} ,
    {ParityBasicCase::Kind::Expected, "(2**61, 2**61+123)", "(2305843009213693952, 2305843009213694075)"} ,
    //{ParityBasicCase::Kind::Expected, "(2**61, 2**61+123);hex", "(0x2000000000000000, 0x200000000000007B)"} ,
    {ParityBasicCase::Kind::Expected, "2**58+123", "288230376151711867"} ,
    {ParityBasicCase::Kind::Expected, "2**61+123", "2305843009213694075"} ,
    //{ParityBasicCase::Kind::Expected, "2**58+123;hex", "0x40000000000007B"} ,
    //{ParityBasicCase::Kind::Expected, "2**61+123;hex", "0x200000000000007B"} ,
    {ParityBasicCase::Kind::Expected, "9.0*10**18", "9000000000000000000"} ,
    {ParityBasicCase::Kind::Expected, "9.2e17", "920000000000000000"} ,
    {ParityBasicCase::Kind::Expected, "9e18", "9000000000000000000"} ,
    {ParityBasicCase::Kind::Expected, "90.123e15", "90123000000000000"} ,
    {ParityBasicCase::Kind::Expected, "1.23456789e18", "1234567890000000000"} ,
    {ParityBasicCase::Kind::Expected, "(9.0*10**18,9.2e17,9e18,90.123e15,1.23456789e18)", "(9000000000000000000, 920000000000000000, 9000000000000000000, 90123000000000000, 1234567890000000000)"} ,
    {ParityBasicCase::Kind::Expected, "(9.2233e18,9.2234e18,0.123,0.123e3,0.12345e4,0.123e5,0.012345678901234e18,1.234567890123456e18,222.0,0)", "(9223300000000000000, 9.2234e+18, 0.123, 123, 1234.5, 12300, 12345678901234000, 1234567890123456000, 222, 0)"} ,
    {ParityBasicCase::Kind::ErrorContains, "0x10000000000000000", "invalid hex literal"} ,
    {ParityBasicCase::Kind::ErrorContains, "0b10000000000000000000000000000000000000000000000000000000000000000", "invalid binary literal"} ,
    {ParityBasicCase::Kind::ErrorContains, "0o2000000000000000000000", "invalid octal literal"} ,
    {ParityBasicCase::Kind::ErrorContains, "1e", "unexpected token"} ,
    {ParityBasicCase::Kind::ErrorContains, "1e+", "unexpected token"} ,
    {ParityBasicCase::Kind::ErrorContains, "1e-", "unexpected token"} ,
    {ParityBasicCase::Kind::ErrorContains, ".", "unexpected token"} ,
    {ParityBasicCase::Kind::Expected, ".0", "0"} ,
    {ParityBasicCase::Kind::Expected, "0.", "0"} ,
    {ParityBasicCase::Kind::Expected, "0.0", "0"} ,
    {ParityBasicCase::Kind::Expected, "0xFFFFFFFFFFFFFFFF+2", "1.844674407370955e+019"} ,
    {ParityBasicCase::Kind::Expected, "0xFFFFFFFFFFFFFFFF-2", "18446744073709551613"} ,
    {ParityBasicCase::Kind::Expected, "0xFFFFFFFFFFFFFFFF*2", "3.689348814741911e+019"} ,
    {ParityBasicCase::Kind::Expected, "0xFFFFFFFFFFFFFFFF/2", "9.223372036854776e+018"} ,
    {ParityBasicCase::Kind::Expected, "0xFFFFFFFFFFFFFFFF%2", "1"} ,
    {ParityBasicCase::Kind::Expected, "0xFFFFFFFFFFFFFFFF>>1", "9223372036854775807"} ,
    {ParityBasicCase::Kind::Expected, "0xFFFFFFFFFFFFFFFF<<1", "3.689348814741911e+019"} ,
    {ParityBasicCase::Kind::Expected, "0xFFFFFFFFFFFFFFFF**2", "3.402823669209385e+038"} ,
    {ParityBasicCase::Kind::Expected, "0x7FFFFFFFFFFFFFFF**2", "8.507059173023462e+037"} ,
    {ParityBasicCase::Kind::Expected, "-0xFFFFFFFFFFFFFFFF+2", "-1.844674407370955e+019"} ,
    {ParityBasicCase::Kind::Expected, "-0xFFFFFFFFFFFFFFFF-2", "-1.844674407370955e+019"} ,
    {ParityBasicCase::Kind::ErrorContains, "(-0xFFFFFFFFFFFFFFFF)%2", "modulo operands must be integer values"} ,
    {ParityBasicCase::Kind::ErrorContains, "(-0xFFFFFFFFFFFFFFFF)>>1", "bitwise operands must be integer values"} ,
    {ParityBasicCase::Kind::Expected, "-0x7FFFFFFFFFFFFFFF+2", "-9223372036854775805"} ,
    {ParityBasicCase::Kind::Expected, "0xFFFFFFFFFFFFFFFF+0", "18446744073709551615"} ,
    {ParityBasicCase::Kind::Expected, "uhex(0xFFFFFFFFFFFFFFFF+0)", "0xFFFFFFFFFFFFFFFF"} ,
    {ParityBasicCase::Kind::Expected, "100000000000000000000", "1.0000000000000000e+20"} ,
    {ParityBasicCase::Kind::Expected, "0x7FFFFFFFFFFFFFFF+1", "9223372036854775808"} ,
    {ParityBasicCase::Kind::Expected, "0x7FFFFFFFFFFFFFFF+2", "9223372036854775809"} ,
    {ParityBasicCase::Kind::Expected, "0x7FFFFFFFFFFFFFFF-2", "9223372036854775805"} ,
    {ParityBasicCase::Kind::Expected, "0x7FFFFFFFFFFFFFFF<<1", "18446744073709551614"} ,
    {ParityBasicCase::Kind::Expected, "0x7FFFFFFFFFFFFFFF<<2", "3.68934881474191e+019"} ,
    {ParityBasicCase::Kind::Expected, "0x7FFFFFFFFFFFFFFF>>1", "4611686018427387903"} ,
    {ParityBasicCase::Kind::Expected, "a=(-0x7FFFFFFFFFFFFFFF,0x7FFFFFFFFFFFFFFF); a>>1", "(-4611686018427387904, 4611686018427387903)"} ,
    {ParityBasicCase::Kind::Expected, "a=(0x7FFFFFFFFFFFFFFF,0x7FFFFFFFFFFFFFFF); b=(2,2); a+b", "(9223372036854775809, 9223372036854775809)"} ,
    {ParityBasicCase::Kind::Expected, "-0x7FFFFFFFFFFFFFFF-2", "-9.223372036854776e+018"} ,
    {ParityBasicCase::Kind::Expected, "0x7FFFFFFFFFFFFFFF*2", "18446744073709551614"} ,
    {ParityBasicCase::Kind::Expected, "a=(-0xFFFFFFFFFFFFFFFF,-0x7FFFFFFFFFFFFFFF,0x7FFFFFFFFFFFFFFF); b=(2,2,2); a+b", "(-1.844674407370955e+019, -9223372036854775805, 9223372036854775809)"} ,
    {ParityBasicCase::Kind::Expected, "a=(-0xFFFFFFFFFFFFFFFF,-0x7FFFFFFFFFFFFFFF,0x7FFFFFFFFFFFFFFF); b=(2,2,2); a*b", "(-3.689348814741911e+019, -1.844674407370955e+019, 18446744073709551614)"} ,
    {ParityBasicCase::Kind::ErrorContains, "a=(-0xFFFFFFFFFFFFFFFF,-0x7FFFFFFFFFFFFFFF,0x7FFFFFFFFFFFFFFF); mod(a,2)", "mod() expects integer values"} ,
    {ParityBasicCase::Kind::ErrorContains, "a=(-0xFFFFFFFFFFFFFFFF,-0x7FFFFFFFFFFFFFFF,0x7FFFFFFFFFFFFFFF); a>>1", "bitwise operands must be integer values"} ,
    {ParityBasicCase::Kind::Expected, "a=(0xFFFFFFFFFFFFFFFF,0x7FFFFFFFFFFFFFFF,0x7FFFFFFFFFFFFFF); a>>1", "(9223372036854775807, 4611686018427387903, 288230376151711743)"} ,
    {ParityBasicCase::Kind::Expected, "a=(0xFFFFFFFFFFFFFFFF,0x7FFFFFFFFFFFFFFF,0x7FFFFFFFFFFFFFF); a<<1", "(3.68934881474191e+19, 18446744073709551614, 1152921504606846974)"} ,
    {ParityBasicCase::Kind::ErrorContains, "3.68934881474191e+19>>1", "bitwise operands must be integer values"} ,
    {ParityBasicCase::Kind::Expected, "3.68934881474191e+19/2", "1.844674407370955e+019"} ,
    {ParityBasicCase::Kind::Expected, "3.68934881474191e+19**0.5", "6074000999.952098"} ,
    {ParityBasicCase::Kind::ErrorContains, "3.68934881474191e+19%3", "modulo operands must be integer values"} ,
    {ParityBasicCase::Kind::Expected, "uhex(product(18446744073709551615,1))", "0xFFFFFFFFFFFFFFFF"} ,
    {ParityBasicCase::Kind::Expected, "prod(-9223372036854775807,1)&1", "1"} ,
    {ParityBasicCase::Kind::Expected, "uhex(product(-9223372036854775807-1,-1))", "0x8000000000000000"} ,
    {ParityBasicCase::Kind::Expected, "uhex(2**63)", "0x8000000000000000"} ,
    {ParityBasicCase::Kind::Expected, "uhex(pow(2,63))", "0x8000000000000000"} ,
    {ParityBasicCase::Kind::Expected, "uhex(pow((2,3),(63,2)))", "(0x8000000000000000,0x9)"} ,
    {ParityBasicCase::Kind::Expected, "uhex((-2)**63)", "0x8000000000000000"} ,
    {ParityBasicCase::Kind::Expected, "sort(0xFFFFFFFFFFFFFFFF,nan,0x7FFFFFFFFFFFFFFF);hex", "(nan,0x7FFFFFFFFFFFFFFF,0xFFFFFFFFFFFFFFFF)"} ,
    {ParityBasicCase::Kind::Expected, "sort(-5,nan,3,-inf,inf,0)", "(nan,-inf,-5,0,3,inf)"} ,
    {ParityBasicCase::Kind::Expected, "sort(nan,inf,2,-inf,nan,-2)", "(nan,nan,-inf,-2,2,inf)"} ,
    {ParityBasicCase::Kind::Expected, "sorted((nan,-3,4,inf,-inf))", "(nan,-inf,-3,4,inf)"} ,
    {ParityBasicCase::Kind::Expected, "x=0.5; f(x)=x<<2; f(3)", "12"} ,
    {ParityBasicCase::Kind::ErrorContains, "f(x)=x*2; f=10; f(3)", "unknown function: f"} ,
    {ParityBasicCase::Kind::Expected, "f(x)=x*2; f=10; f", "10"} ,
    {ParityBasicCase::Kind::ErrorContains, "f=10; f(x)=x*2; f", "user-defined function: f(x)"} ,
    {ParityBasicCase::Kind::ErrorContains, "f(x)=x+1; f", "user-defined function: f(x)"} ,
    {ParityBasicCase::Kind::Expected, "x=0; f(x)=1/x; f(2)", "0.5"} ,
    {ParityBasicCase::Kind::Expected, "f(x)=1%x; f(7)", "1"} ,
    {ParityBasicCase::Kind::Expected, "_", "1"} ,
    {ParityBasicCase::Kind::Expected, "_+5", "6"} ,
    {ParityBasicCase::Kind::Expected, "_=10; g(x)=1%(x-1); g(5)", "1"} ,
    {ParityBasicCase::Kind::Expected, "log(8,2)", "3"} ,
    {ParityBasicCase::Kind::Expected, "log(100,10)", "2"} ,
    {ParityBasicCase::Kind::ErrorContains, "log(8)", "log() expects 2 argument(s)"} ,
    {ParityBasicCase::Kind::ErrorContains, "ln", "function: ln(value)"} ,
    {ParityBasicCase::Kind::Expected, "log(e,e)", "1"} ,
    {ParityBasicCase::Kind::Expected, "a=2;sum(a,a)", "4"} ,
    {ParityBasicCase::Kind::ErrorContains, "f(a,a)=a", "duplicate parameter name"} ,
    {ParityBasicCase::Kind::Expected, "2+3;ans", "5"} ,
    {ParityBasicCase::Kind::Expected, "(1,2,3);ans", "(1,2,3)"} ,
    {ParityBasicCase::Kind::Expected, "hex(15);ans", "0xF"} ,
    {ParityBasicCase::Kind::Expected, "7; ans*2", "14"} ,
    {ParityBasicCase::Kind::Expected, "v=(10,20);sum(ans)", "30"} ,
    {ParityBasicCase::Kind::Expected, "atan2(1,1)", "0.7853981633974483"} ,
    {ParityBasicCase::Kind::Expected, "floor(2.9)", "2"} ,
    {ParityBasicCase::Kind::Expected, "ceil(2.1)", "3"} ,
    {ParityBasicCase::Kind::Expected, "ceil((1e14+0.5, 1e30, -1e30))", "(100000000000001, 1e+30, -1e+30)"},
    {ParityBasicCase::Kind::Expected, "trunc(-2.9)", "-2"} ,
    {ParityBasicCase::Kind::Expected, "round(2.5)", "3"} ,
    {ParityBasicCase::Kind::Expected, "sign(-123)", "-1"} ,
    {ParityBasicCase::Kind::Expected, "mod(17,5)", "2"} ,
    {ParityBasicCase::Kind::Expected, "avg(1,2,3,4)", "2.5"} ,
    {ParityBasicCase::Kind::Expected, "mean((1,2,3),9)", "3.75"} ,
    {ParityBasicCase::Kind::Expected, "clamp(15,0,10)", "10"} ,
    {ParityBasicCase::Kind::Expected, "deg(pi)", "180"} ,
    {ParityBasicCase::Kind::Expected, "rad(180)", "3.141592653589793"} ,
    {ParityBasicCase::Kind::Expected, "hypot(3,4)", "5"} ,
    {ParityBasicCase::Kind::Expected, "gcd(84,30)", "6"} ,
    {ParityBasicCase::Kind::Expected, "lcm(6,8)", "24"} ,
    {ParityBasicCase::Kind::Expected, "median(1,8,3)", "3"} ,
    {ParityBasicCase::Kind::Expected, "median((1,9),3,7)", "5"} ,
    {ParityBasicCase::Kind::Expected, "variance(1,2,3)", "0.6666666666666666"} ,
    {ParityBasicCase::Kind::Expected, "stddev(1,2,3)", "0.816496580927726"} ,
    {ParityBasicCase::Kind::Expected, "fact(5)", "120"} ,
    {ParityBasicCase::Kind::Expected, "factorial(10)", "3628800"} ,
    {ParityBasicCase::Kind::ErrorContains, "fact(2.5)", "fact() expects integer values"} ,
    {ParityBasicCase::Kind::Expected, "factorial(21)", "5.109094217170944e+019"} ,
    {ParityBasicCase::Kind::Expected, "fact(171)", "inf"} ,
    {ParityBasicCase::Kind::Expected, "6*(1/2)", "3"} ,
    {ParityBasicCase::Kind::Expected, "6*(1/3)", "2"} ,
    {ParityBasicCase::Kind::Expected, "5*(1/3)", "1.666666666666667"} ,
    {ParityBasicCase::Kind::Expected, "0xFFFFFFFFFFFFFFFF*(1/5)", "3689348814741910323"} ,
    {ParityBasicCase::Kind::Expected, "0x7FFFFFFFFFFFFFFF*(1/7)", "1317624576693539401"} ,
    {ParityBasicCase::Kind::Expected, "2**64*(1/2)", "9.223372036854776e+018"} ,
    {ParityBasicCase::Kind::Expected, "0x7FFFFFFFFFFFFFFF*(1/8)", "1.152921504606847e+018"} ,
    {ParityBasicCase::Kind::Expected, "2**52*(1/3)", "1.501199875790165e+015"} ,
    {ParityBasicCase::Kind::Expected, "-0x7FFFFFFFFFFFFFFF*(1/7)", "-1317624576693539401"} ,
    {ParityBasicCase::Kind::Expected, "0x7FFFFFFFFFFFFFFF*(-1/7)", "-1317624576693539401"} ,
    {ParityBasicCase::Kind::Expected, "-2**52*(1/2)", "-2251799813685248"} ,
    {ParityBasicCase::Kind::Expected, "2**52*(-1/2)", "-2251799813685248"} ,
    {ParityBasicCase::Kind::Expected, "1317624576693539401/(1/11)", "14493870343628933411"} ,
    {ParityBasicCase::Kind::Expected, "1317624576693539401/(1/7)", "9223372036854775807"} ,
    {ParityBasicCase::Kind::Expected, "-1317624576693539401/(1/7)", "-9223372036854775807"} ,
    {ParityBasicCase::Kind::Expected, "1317624576693539401/(-1/7)", "-9223372036854775807"} ,
    {ParityBasicCase::Kind::Expected, "1317624576693539401/(1/8)", "10540996613548315208"} ,
    {ParityBasicCase::Kind::Expected, "5/(1/3)", "15"} ,
    {ParityBasicCase::Kind::Expected, "-5/(1/3)", "-15"} ,
    {ParityBasicCase::Kind::Expected, "5/(-1/3)", "-15"} ,
    {ParityBasicCase::Kind::Expected, "0x7FFFFFFFFFFFFFFF/(1/8)", "7.37869762948382e+019"} ,
    {ParityBasicCase::Kind::Expected, "2**64/(1/2)", "3.68934881474191e+019"} ,
    {ParityBasicCase::Kind::Expected, "1317624576693539401*(1/11)", "1.197840524266854e+017"} ,
    {ParityBasicCase::Kind::Expected, "0x7FFFFFFFFFFFFFFF*(-1/8)", "-1.152921504606847e+018"} ,
    {ParityBasicCase::Kind::Expected, "6/(10,20,30)", "(0.6, 0.3, 0.2)"} ,
    {ParityBasicCase::Kind::Expected, "ratio(6/(10,20,30))", "(3/5, 3/10, 1/5)"} ,
    {ParityBasicCase::Kind::Expected, "sort(ratio(6/(10,20,30)))", "(1/5, 3/10, 3/5)"} ,
    {ParityBasicCase::Kind::Expected, "unique(abs((2,-1,2,3,-3,1)))", "(2, 1, 3)"} ,
#if SMARTMATH_FACTORINT
    {ParityBasicCase::Kind::Expected, "factorint(33)", "(3, 11)"} ,
    {ParityBasicCase::Kind::Expected, "factorint(12)", "(2**2, 3)"} ,
    {ParityBasicCase::Kind::Expected, "factorint(-33)", "(-3, 11)"} ,
    {ParityBasicCase::Kind::Expected, "factorint(2**52)", "(2**52)"} ,
    {ParityBasicCase::Kind::Expected, "factorint(90)", "(2, 3**2, 5)"} ,
    {ParityBasicCase::Kind::Expected, "factorint(9007)", "(9007)"} ,
    {ParityBasicCase::Kind::Expected, "factorint(900719)", "(900719)"} ,
    {ParityBasicCase::Kind::Expected, "factorint(90071992)", "(2**3, 11258999)"} ,
    {ParityBasicCase::Kind::Expected, "factorint(9007199254)", "(2, 89, 50602243)"} ,
    {ParityBasicCase::Kind::Expected, "factorint(900719925474)", "(2, 3, 12907, 11630897)"} ,
    {ParityBasicCase::Kind::Expected, "prod(factorint(12))", "12"} ,
    {ParityBasicCase::Kind::ErrorContains, "factorint((33,12))", "expects scalar values"} ,
    {ParityBasicCase::Kind::ErrorContains, "factorint(2**64)", "expects integer values"} ,
#endif
    {ParityBasicCase::Kind::ErrorContains, "rand", "function: rand()"} ,
    {ParityBasicCase::Kind::ErrorContains, "random", "function: random(min, max)"} ,
    {ParityBasicCase::Kind::ErrorContains, "median", "function: median(...)"} ,
    {ParityBasicCase::Kind::Expected, "sort((3,1,2))", "(1,2,3)"} ,
    {ParityBasicCase::Kind::Expected, "a=(5,2,9);sort(a)", "(2,5,9)"} ,
    {ParityBasicCase::Kind::Expected, "sort(5)", "(5)"} ,
    {ParityBasicCase::Kind::Expected, "sort(2,5,1)", "(1,2,5)"} ,
    {ParityBasicCase::Kind::ErrorContains, "sort", "function: sort(...)" } ,
    {ParityBasicCase::Kind::Expected, "unique((3,1,3,2,1,2))", "(3,1,2)"} ,
    {ParityBasicCase::Kind::Expected, "a=(5,2,5,9,2);unique(a)", "(5,2,9)"} ,
    {ParityBasicCase::Kind::Expected, "unique(5)", "(5)"} ,
    {ParityBasicCase::Kind::Expected, "unique(1,2,1,2,3)", "(1,2,3)"} ,
    {ParityBasicCase::Kind::ErrorContains, "unique", "function: unique(...)" } ,
    {ParityBasicCase::Kind::Expected, "(1,2)+(3)", "(4,5)"} ,
    {ParityBasicCase::Kind::ErrorContains, "1<<64", "incompatible operands"} ,
    {ParityBasicCase::Kind::ErrorContains, "1>>-1", "incompatible operands"} ,
    {ParityBasicCase::Kind::ErrorContains, "5%0", "incompatible operands"} ,
    {ParityBasicCase::Kind::ErrorContains, "1+", "unexpected token"} ,
    {ParityBasicCase::Kind::ErrorContains, "pow()", "expects 2 argument(s)"} ,
    {ParityBasicCase::Kind::ErrorContains, "pow(2,3,4)", "expects 2 argument(s)"} ,
    {ParityBasicCase::Kind::ErrorContains, "log()", "expects 2 argument(s)"} ,
    {ParityBasicCase::Kind::ErrorContains, "log(10,10,10)", "expects 2 argument(s)"} ,
    {ParityBasicCase::Kind::ErrorContains, "atan2()", "expects 2 argument(s)"} ,
    {ParityBasicCase::Kind::ErrorContains, "atan2(1,2,3)", "expects 2 argument(s)"} ,
    {ParityBasicCase::Kind::ErrorContains, "hypot()", "expects 2 argument(s)"} ,
    {ParityBasicCase::Kind::ErrorContains, "hypot(3,4,5)", "expects 2 argument(s)"} ,
    {ParityBasicCase::Kind::ErrorContains, "mod()", "expects 2 argument(s)"} ,
    {ParityBasicCase::Kind::ErrorContains, "mod(10,3,1)", "expects 2 argument(s)"} ,
    {ParityBasicCase::Kind::ErrorContains, "mod(10,0)", "numeric error in mod()"} ,
    {ParityBasicCase::Kind::ErrorContains, "gcd()", "expects 2 argument(s)"} ,
    {ParityBasicCase::Kind::ErrorContains, "gcd(6,8,10)", "expects 2 argument(s)"} ,
    {ParityBasicCase::Kind::ErrorContains, "gcd(6.5,3)", "expects integer values"} ,
    {ParityBasicCase::Kind::Expected, "lcm(6,0)", "0"} ,
    {ParityBasicCase::Kind::ErrorContains, "lcm(6.5,3)", "expects integer values"} ,
    {ParityBasicCase::Kind::ErrorContains, "hex()", "expects at least 1 argument"} ,
    {ParityBasicCase::Kind::Expected, "hex(1,2)", "(0x1,0x2)"} ,
    {ParityBasicCase::Kind::ErrorContains, "bin()", "expects at least 1 argument"} ,
    {ParityBasicCase::Kind::Expected, "bin(1,2)", "(0b1,0b10)"} ,
    {ParityBasicCase::Kind::ErrorContains, "clamp()", "expects 3 argument(s)"} ,
    {ParityBasicCase::Kind::ErrorContains, "clamp(1,2)", "expects 3 argument(s)"} ,
    {ParityBasicCase::Kind::ErrorContains, "clamp(1,2,3,4)", "expects 3 argument(s)"} ,
    {ParityBasicCase::Kind::ErrorContains, "clamp((1,2),(3,4),4)", "expects scalar min/max"} ,
    {ParityBasicCase::Kind::Expected, "rand()*0", "0"} ,
    {ParityBasicCase::Kind::ErrorContains, "random()", "expects 2 argument(s)"} ,
    {ParityBasicCase::Kind::ErrorContains, "random(1)", "expects 2 argument(s)"} ,
    {ParityBasicCase::Kind::ErrorContains, "random(1,2,3)", "expects 2 argument(s)"} ,
    {ParityBasicCase::Kind::ErrorContains, "random((1,2),3)", "expects scalar values"} ,
    {ParityBasicCase::Kind::ErrorContains, "sort()", "expects at least 1 argument"} ,
    {ParityBasicCase::Kind::Expected, "sort((2,5,1),4,3)", "(1,2,3,4,5)"} ,
    {ParityBasicCase::Kind::ErrorContains, "unique()", "expects at least 1 argument"} ,
    {ParityBasicCase::Kind::ErrorContains, "sum()", "expects at least 1 argument"} ,
    {ParityBasicCase::Kind::ErrorContains, "product()", "expects at least 1 argument"} ,
    {ParityBasicCase::Kind::ErrorContains, "min()", "expects at least 1 argument"} ,
    {ParityBasicCase::Kind::ErrorContains, "max()", "expects at least 1 argument"} ,
    {ParityBasicCase::Kind::ErrorContains, "avg()", "expects at least 1 argument"} ,
    {ParityBasicCase::Kind::ErrorContains, "median()", "expects at least 1 argument"} ,
    {ParityBasicCase::Kind::ErrorContains, "variance()", "expects at least 1 argument"} ,
    {ParityBasicCase::Kind::ErrorContains, "stddev()", "expects at least 1 argument"} ,
    {ParityBasicCase::Kind::ErrorContains, "sin(1,2)", "expects 1 argument(s)"} ,
    {ParityBasicCase::Kind::ErrorContains, "cos(1,2)", "expects 1 argument(s)"} ,
    {ParityBasicCase::Kind::ErrorContains, "ln(1,2)", "expects 1 argument(s)"} ,
    {ParityBasicCase::Kind::ErrorContains, "sqrt(1,2)", "expects 1 argument(s)"} ,
    {ParityBasicCase::Kind::ErrorContains, "fact()", "expects 1 argument(s)"} ,
    {ParityBasicCase::Kind::ErrorContains, "factorial(1,2)", "expects 1 argument(s)"} ,
    {ParityBasicCase::Kind::ErrorContains, "pow(2,)", "unexpected comma"} ,
    {ParityBasicCase::Kind::ErrorContains, "sum((1,2),)", "unexpected comma"} ,
    {ParityBasicCase::Kind::ErrorContains, "1/", "unexpected token"} ,
    {ParityBasicCase::Kind::ErrorContains, "1**", "unexpected token"} ,
    {ParityBasicCase::Kind::ErrorContains, "1<<", "unexpected token"} ,
    {ParityBasicCase::Kind::ErrorContains, "1>>", "unexpected token"} ,
    {ParityBasicCase::Kind::ErrorContains, "1&", "unexpected token"} ,
    {ParityBasicCase::Kind::ErrorContains, "1|", "unexpected token"} ,
    {ParityBasicCase::Kind::ErrorContains, "1^", "unexpected token"} ,
    {ParityBasicCase::Kind::Expected, "1%", "0.01"} ,
    {ParityBasicCase::Kind::ErrorContains, "pow(,2)", "unexpected"} ,
    {ParityBasicCase::Kind::ErrorContains, "atan2(,2)", "unexpected"} ,
    {ParityBasicCase::Kind::ErrorContains, "random(,2)", "unexpected"} ,
    {ParityBasicCase::Kind::ErrorContains, "clamp(1,,3)", "unexpected"} ,
    {ParityBasicCase::Kind::ErrorContains, "sum(,1)", "unexpected"} ,
    {ParityBasicCase::Kind::ErrorContains, "sin()", "expects 1 argument(s)"} ,
    {ParityBasicCase::Kind::ErrorContains, "tan()", "expects 1 argument(s)"} ,
    {ParityBasicCase::Kind::ErrorContains, "asin()", "expects 1 argument(s)"} ,
    {ParityBasicCase::Kind::ErrorContains, "acos()", "expects 1 argument(s)"} ,
    {ParityBasicCase::Kind::ErrorContains, "atan()", "expects 1 argument(s)"} ,
    {ParityBasicCase::Kind::ErrorContains, "sinh()", "expects 1 argument(s)"} ,
    {ParityBasicCase::Kind::ErrorContains, "cosh()", "expects 1 argument(s)"} ,
    {ParityBasicCase::Kind::ErrorContains, "tanh()", "expects 1 argument(s)"} ,
    {ParityBasicCase::Kind::ErrorContains, "acosh()", "expects 1 argument(s)"} ,
    {ParityBasicCase::Kind::ErrorContains, "asinh()", "expects 1 argument(s)"} ,
    {ParityBasicCase::Kind::ErrorContains, "atanh()", "expects 1 argument(s)"} ,
    {ParityBasicCase::Kind::Expected, "acosh(1)", "0"} ,
    {ParityBasicCase::Kind::Expected, "acosh(2)", "1.3169578969248166"} ,
    {ParityBasicCase::Kind::Expected, "acosh(0)", "nan"} ,
    {ParityBasicCase::Kind::Expected, "acosh(inf)", "inf"} ,
    {ParityBasicCase::Kind::Expected, "asinh(0)", "0"} ,
    {ParityBasicCase::Kind::Expected, "asinh(1)", "0.8813735870195431"} ,
    {ParityBasicCase::Kind::Expected, "asinh(-1)", "-0.8813735870195431"} ,
    {ParityBasicCase::Kind::Expected, "asinh((0,1))", "(0,0.8813735870195431)"} ,
    {ParityBasicCase::Kind::Expected, "atanh(0)", "0"} ,
    {ParityBasicCase::Kind::Expected, "atanh(0.5)", "0.5493061443340549"} ,
    {ParityBasicCase::Kind::Expected, "atanh(1)", "inf"} ,
    {ParityBasicCase::Kind::Expected, "atanh(-1)", "-inf"} ,
    {ParityBasicCase::Kind::Expected, "atanh(2)", "nan"} ,
    {ParityBasicCase::Kind::ErrorContains, "exp()", "expects 1 argument(s)"} ,
    {ParityBasicCase::Kind::ErrorContains, "log10()", "expects 1 argument(s)"} ,
    {ParityBasicCase::Kind::ErrorContains, "abs()", "expects 1 argument(s)"} ,
    {ParityBasicCase::Kind::ErrorContains, "floor()", "expects 1 argument(s)"} ,
    {ParityBasicCase::Kind::ErrorContains, "ceil()", "expects 1 argument(s)"} ,
    {ParityBasicCase::Kind::ErrorContains, "trunc()", "expects 1 argument(s)"} ,
    {ParityBasicCase::Kind::ErrorContains, "round()", "expects 1 argument(s)"} ,
    {ParityBasicCase::Kind::ErrorContains, "sign()", "expects 1 argument(s)"} ,
    {ParityBasicCase::Kind::ErrorContains, "deg()", "expects at least 1 argument"} ,
    {ParityBasicCase::Kind::ErrorContains, "rad()", "expects at least 1 argument"} ,
    {ParityBasicCase::Kind::Expected, "ln((1,2))", "(0,0.6931471805599453)"} ,
    {ParityBasicCase::Kind::Expected, "sqrt((1,2))", "(1,1.414213562373095)"} ,
    {ParityBasicCase::Kind::Expected, "abs((1,2))", "(1,2)"} ,
    {ParityBasicCase::Kind::Expected, "arcsin(1)", "1.570796326794897"} ,
    {ParityBasicCase::Kind::Expected, "arccos(1)", "0"} ,
    {ParityBasicCase::Kind::Expected, "arctan(1)", "0.7853981633974483"} ,
    {ParityBasicCase::Kind::ErrorContains, "prod()", "expects at least 1 argument"} ,
    {ParityBasicCase::Kind::ErrorContains, "mean()", "expects at least 1 argument"} ,
    {ParityBasicCase::Kind::ErrorContains, "variance(())", "expects at least 1 argument"} ,
    {ParityBasicCase::Kind::ErrorContains, "stddev(())", "expects at least 1 argument"} ,
    {ParityBasicCase::Kind::ErrorContains, "gcd(1)", "expects 2 argument(s)"} ,
    {ParityBasicCase::Kind::ErrorContains, "lcm(1)", "expects 2 argument(s)"} ,
    {ParityBasicCase::Kind::ErrorContains, "mod(1)", "expects 2 argument(s)"} ,
    {ParityBasicCase::Kind::ErrorContains, "hypot(1)", "expects 2 argument(s)"} ,
    {ParityBasicCase::Kind::ErrorContains, "atan2(1)", "expects 2 argument(s)"} ,
    {ParityBasicCase::Kind::Expected, "log((1,2),10)", "(0,0.3010299956639812)"} ,
    {ParityBasicCase::Kind::Expected, "log(10,(2,3))", "(3.321928094887362,2.095903274289384)"} ,
    {ParityBasicCase::Kind::Expected, "pow((2,3),2)", "(4,9)"} ,
    {ParityBasicCase::Kind::Expected, "pow(2,(2,3))", "(4,8)"} ,
    {ParityBasicCase::Kind::Expected, "hex((1,2,3),(4))", "(0x1,0x2,0x3,0x4)"} ,
    {ParityBasicCase::Kind::Expected, "bin((1,2,3),(4))", "(0b1,0b10,0b11,0b100)"} ,
    {ParityBasicCase::Kind::Expected, "random(10,10)", "10"} ,
    {ParityBasicCase::Kind::Expected, "random(1.5,1.5)", "1.5"} ,
    {ParityBasicCase::Kind::Expected, "random(3.5,3.5)", "3.5"} ,
    {ParityBasicCase::Kind::ErrorContains, "fact((1,2))", "expects a non-negative integer"} ,
    {ParityBasicCase::Kind::ErrorContains, "factorial((1,2))", "expects a non-negative integer"} ,
    {ParityBasicCase::Kind::ErrorContains, "clamp((1,2,3),(4,5),6)", "expects scalar min/max"} ,
    {ParityBasicCase::Kind::Expected, "sum((1,2),(3,4),5)", "15"} ,
    {ParityBasicCase::Kind::ErrorContains, "sort(())", "expects at least 1 argument"} ,
    {ParityBasicCase::Kind::ErrorContains, "unique(())", "expects at least 1 argument"} ,
    {ParityBasicCase::Kind::ErrorContains, "RestoreAnsFromCachedRender(g_cachedRenderText(i))", "unknown function"} ,
    {ParityBasicCase::Kind::Expected, "deg(pi/2,pi/4)", "(90,45)"} ,
    {ParityBasicCase::Kind::Expected, "rad(180,90)", "(3.141592653589793,1.570796326794897)"} ,
    {ParityBasicCase::Kind::ErrorContains, "mean", "function: mean(...)"} ,
    {ParityBasicCase::Kind::ErrorContains, "floor", "function: floor(value)"} ,
    {ParityBasicCase::Kind::ErrorContains, "ceil", "function: ceil(value)"} ,
    {ParityBasicCase::Kind::ErrorContains, "trunc", "function: trunc(value)"} ,
    {ParityBasicCase::Kind::ErrorContains, "round", "function: round(value)"} ,
    {ParityBasicCase::Kind::ErrorContains, "sign", "function: sign(value)"} ,
    {ParityBasicCase::Kind::ErrorContains, "deg", "function: deg(...)"} ,
    {ParityBasicCase::Kind::ErrorContains, "rad", "function: rad(...)"} ,
    {ParityBasicCase::Kind::Expected, "int(2.9)", "2"} ,
    {ParityBasicCase::Kind::Expected, "int(-2.9)", "-2"} ,
    {ParityBasicCase::Kind::Expected, "frac(2.9)", "0.8999999999999999"} ,
    {ParityBasicCase::Kind::Expected, "frac(-2.9)", "-0.8999999999999999"} ,
    {ParityBasicCase::Kind::Expected, "int((2.9,-2.9))", "(2,-2)"} ,
    {ParityBasicCase::Kind::Expected, "frac((2.9,-2.9))", "(0.8999999999999999,-0.8999999999999999)"} ,
    {ParityBasicCase::Kind::ErrorContains, "int", "function: int(value)"} ,
    {ParityBasicCase::Kind::ErrorContains, "frac", "function: frac(value)"} ,
    {ParityBasicCase::Kind::Expected, "fract(2.9)", "0.8999999999999999"} ,
    {ParityBasicCase::Kind::Expected, "fract((2.9,-2.9))", "(0.8999999999999999,-0.8999999999999999)"} ,
    {ParityBasicCase::Kind::ErrorContains, "fract", "function: frac(value)"} ,
    {ParityBasicCase::Kind::Expected, "oct(12)", "0o14"} ,
    {ParityBasicCase::Kind::Expected, "oct((12,255))", "(0o14,0o377)"} ,
    {ParityBasicCase::Kind::Expected, "10 + oct(12) + 14", "36"} ,
    {ParityBasicCase::Kind::ErrorContains, "oct(12.5)", "oct() expects integer values"} ,
    {ParityBasicCase::Kind::ErrorContains, "oct", "function: oct(...)"} ,
    {ParityBasicCase::Kind::ErrorContains, "oct()", "expects at least 1 argument"} ,
    {ParityBasicCase::Kind::Expected, "oct(1,2)", "(0o1,0o2)"} ,
    {ParityBasicCase::Kind::Expected, "oct((1,2,3),(4))", "(0o1,0o2,0o3,0o4)"} ,
    {ParityBasicCase::Kind::Expected, "oct(15);ans", "0o17"} ,
    {ParityBasicCase::Kind::Expected, "oct(9223372036854775807+1)", "0o1000000000000000000000"} ,
    {ParityBasicCase::Kind::Expected, "0O77", "63"} ,
    {ParityBasicCase::Kind::Expected, "0o123 + 1", "84"} ,
    {ParityBasicCase::Kind::Expected, "0o20 & 0xF", "0"} ,
    {ParityBasicCase::Kind::Expected, "oct((0o7,0o10))", "(0o7,0o10)"} ,
    {ParityBasicCase::Kind::Expected, "0o64", "52"} ,
    {ParityBasicCase::Kind::ErrorContains, "0o", "invalid octal literal"} ,
    {ParityBasicCase::Kind::Expected, "oct(0o17)", "0o17"} ,
    {ParityBasicCase::Kind::Expected, "0o10 + 8", "16"} ,
    {ParityBasicCase::Kind::Expected, "0b110011 & 0x37 | 0o64", "55"} ,
    {ParityBasicCase::Kind::ErrorContains, "0o8", "invalid octal literal"} ,
    {ParityBasicCase::Kind::Expected, "5=5", "1"} ,
    {ParityBasicCase::Kind::Expected, "5==4", "0"} ,
    {ParityBasicCase::Kind::Expected, "5<>4", "1"} ,
    {ParityBasicCase::Kind::Expected, "5!=5", "0"} ,
    {ParityBasicCase::Kind::Expected, "5>4", "1"} ,
    {ParityBasicCase::Kind::Expected, "5>=5", "1"} ,
    {ParityBasicCase::Kind::Expected, "4<5", "1"} ,
    {ParityBasicCase::Kind::Expected, "4<=4", "1"} ,
    {ParityBasicCase::Kind::Expected, "1|2=3", "1"} ,
    {ParityBasicCase::Kind::Expected, "1|2<2", "0"} ,
    {ParityBasicCase::Kind::Expected, "(1,2,3)=(1,2,3)", "1"} ,
    {ParityBasicCase::Kind::Expected, "(1,2,3)!=(1,2,4)", "1"} ,
    {ParityBasicCase::Kind::Expected, "(1,2)<(1,2,0)", "1"} ,
    {ParityBasicCase::Kind::Expected, "(1,2,9)>(1,2,3)", "1"} ,
    {ParityBasicCase::Kind::Expected, "(1,2,3)<=(1,2,3)", "1"} ,
    {ParityBasicCase::Kind::Expected, "(1)<(1,0)", "1"} ,
    {ParityBasicCase::Kind::Expected, "(1,0)>(1)", "1"} ,
    {ParityBasicCase::Kind::Expected, "5<(5,1)", "1"} ,
    {ParityBasicCase::Kind::Expected, "(5,1)>5", "1"} ,
    {ParityBasicCase::Kind::Expected, "2<3<4", "1"} ,
    {ParityBasicCase::Kind::Expected, "!0", "1"} ,
    {ParityBasicCase::Kind::Expected, "!5", "0"} ,
    {ParityBasicCase::Kind::Expected, "not 0", "1"} ,
    {ParityBasicCase::Kind::Expected, "not 2", "0"} ,
    {ParityBasicCase::Kind::Expected, "1&&1", "1"} ,
    {ParityBasicCase::Kind::Expected, "1&&0", "0"} ,
    {ParityBasicCase::Kind::Expected, "1||0", "1"} ,
    {ParityBasicCase::Kind::Expected, "0||0", "0"} ,
    {ParityBasicCase::Kind::Expected, "1 and 1", "1"} ,
    {ParityBasicCase::Kind::Expected, "1 and 0", "0"} ,
    {ParityBasicCase::Kind::Expected, "1 or 0", "1"} ,
    {ParityBasicCase::Kind::Expected, "0 or 0", "0"} ,
    {ParityBasicCase::Kind::Expected, "!1=0", "1"} ,
    {ParityBasicCase::Kind::Expected, "1=1 && 0=1", "0"} ,
    {ParityBasicCase::Kind::Expected, "1|2==3 && 5>3", "1"} ,
    {ParityBasicCase::Kind::Expected, "1 || 0 && 0", "1"} ,
    {ParityBasicCase::Kind::Expected, "(0,0) && 1", "1"} ,
    {ParityBasicCase::Kind::Expected, "(0,0) || 0", "1"} ,
    {ParityBasicCase::Kind::Expected, "not (0,0)", "0"} ,
    {ParityBasicCase::Kind::Expected, "0 or (0,0)", "1"} ,
    {ParityBasicCase::Kind::Expected, "not 1<0", "1"} ,
    {ParityBasicCase::Kind::Expected, "!1<0", "0"} ,
    {ParityBasicCase::Kind::Expected, "reverse((3,1,2))", "(2,1,3)"} ,
    {ParityBasicCase::Kind::Expected, "reverse(2,5,1)", "(1,5,2)"} ,
    {ParityBasicCase::Kind::Expected, "reverse((1,2,3),(4,5,6),(7,8,9))", "(9,8,7,6,5,4,3,2,1)"} ,
    {ParityBasicCase::Kind::Expected, "reverse(5)", "(5)"} ,
    {ParityBasicCase::Kind::ErrorContains, "reverse", "function: reverse(...)"} ,
    {ParityBasicCase::Kind::ErrorContains, "reverse()", "expects at least 1 argument"} ,
    {ParityBasicCase::Kind::Expected, "reverse((2,5,1),4,3)", "(3,4,1,5,2)"} ,
    {ParityBasicCase::Kind::ErrorContains, "reverse(())", "expects at least 1 argument"} ,
    {ParityBasicCase::Kind::Expected, "(10,20,30)[0]", "10"} ,
    {ParityBasicCase::Kind::Expected, "(10,20,30)[2]", "30"} ,
    {ParityBasicCase::Kind::Expected, "(10,20,30)[-1]", "30"} ,
    {ParityBasicCase::Kind::Expected, "(10,20,30)[-2]", "20"} ,
    {ParityBasicCase::Kind::Expected, "(10,20,30)[-3]", "10"} ,
    {ParityBasicCase::Kind::ErrorContains, "(10,20,30)[-4]", "array index is out of range"} ,
    {ParityBasicCase::Kind::ErrorContains, "(10,20,30)[3]", "array index is out of range"} ,
    {ParityBasicCase::Kind::Expected, "sort((3,1,2,4))[-1]", "4"} ,
    {ParityBasicCase::Kind::Expected, "reverse((1,2,3,4))[-1]", "1"} ,
    {ParityBasicCase::Kind::Expected, "reverse((1,2,3,4))[0]", "4"} ,
    {ParityBasicCase::Kind::Expected, "sorted((3,1,2))", "(1,2,3)"} ,
    {ParityBasicCase::Kind::Expected, "sorted(2,5,1)", "(1,2,5)"} ,
    {ParityBasicCase::Kind::ErrorContains, "sorted", "function: sort(...)"} ,
    {ParityBasicCase::Kind::Expected, "reversed((1,2,3),(4,5))", "(5,4,3,2,1)"} ,
    {ParityBasicCase::Kind::ErrorContains, "reversed", "function: reverse(...)"} ,
    {ParityBasicCase::Kind::Expected, "reversed((1,2,3,4))[-1]", "1"} ,
    {ParityBasicCase::Kind::Expected, "3 + not 4", "3"} ,
    {ParityBasicCase::Kind::Expected, "3 + not 4 + 5", "3"} ,
    {ParityBasicCase::Kind::Expected, "oct(80)", "0o120"} ,
    {ParityBasicCase::Kind::Expected, "oct(0)", "0o0"} ,
    {ParityBasicCase::Kind::Expected, "oct(-1)", "-0o1"} ,
    {ParityBasicCase::Kind::ErrorContains, "sin(x)=x", "reserved function name"} ,
    {ParityBasicCase::Kind::ErrorContains, "oct(x)=x", "reserved function name"} ,
    {ParityBasicCase::Kind::ErrorContains, "not(x)=x", "reserved function name"} ,
    {ParityBasicCase::Kind::ErrorContains, "ans(x)=x", "reserved built-in variable name"} ,
    {ParityBasicCase::Kind::ErrorContains, "_(x)=x", "reserved built-in variable name"} ,
    {ParityBasicCase::Kind::ErrorContains, "ANS(x)=x", "reserved built-in variable name"} ,
    {ParityBasicCase::Kind::Expected, "f(x,y)=x*y; a=(2,3); f(unpack(a))", "6"} ,
    {ParityBasicCase::Kind::Expected, "f(x,y,z)=x+y+z; f(unpack((1,2,3)))", "6"} ,
    {ParityBasicCase::Kind::Expected, "unpack((1,2,3))", "(1,2,3)"} ,
    {ParityBasicCase::Kind::Expected, "unpack(5)", "5"} ,
    {ParityBasicCase::Kind::ErrorContains, "unpack()", "expects at least 1 argument"} ,
    {ParityBasicCase::Kind::ErrorContains, "unpack", "function: unpack(...)"} ,
    {ParityBasicCase::Kind::Expected, "sum(unpack((1,2,3)))", "6"} ,
    {ParityBasicCase::Kind::ErrorContains, "f(x,y)=x*y; f(unpack((2,3,4)))", "expects 2 argument(s)"} ,
    {ParityBasicCase::Kind::Expected, "f(x,y,z)=x+y+z; f(unpack(1,2,3))", "6"} ,
    {ParityBasicCase::Kind::Expected, "f(a,b,c,d,t)=a+b+c+d+t; f(unpack((1,2),3,(4,5)))", "15"} ,
    {ParityBasicCase::Kind::Expected, "unpack((1,2),3,(4,5))", "(1,2,3,4,5)"} ,
    {ParityBasicCase::Kind::Expected, "repeat(0, 3)", "(0,0,0)"} ,
    {ParityBasicCase::Kind::Expected, "repeat((10), 5)", "(10,10,10,10,10)"} ,
    {ParityBasicCase::Kind::Expected, "repeat((1,2), 1)", "(1,2)"} ,
    {ParityBasicCase::Kind::Expected, "repeat(5, 2)", "(5,5)"} ,
    {ParityBasicCase::Kind::ErrorContains, "repeat(1, 0)", "expects a positive integer"} ,
    {ParityBasicCase::Kind::ErrorContains, "repeat(10, -3)", "expects a positive integer"} ,
    {ParityBasicCase::Kind::ErrorContains, "repeat((1,2), -1)", "expects a positive integer"} ,
    {ParityBasicCase::Kind::ErrorContains, "repeat((1,2), 1.23)", "expects integer values"} ,
    {ParityBasicCase::Kind::ErrorContains, "repeat", "function: repeat(x, n)"} ,
    {ParityBasicCase::Kind::ErrorContains, "repeat()", "expects 2 argument(s)"} ,
    {ParityBasicCase::Kind::ErrorContains, "repeat(1)", "expects 2 argument(s)"} ,
#if SMARTMATH_TIME_VALUES
    {ParityBasicCase::Kind::ErrorContains, "repeat((1,2), 1:00)", "expects integer values"} ,
#endif
    {ParityBasicCase::Kind::Expected, "a=repeat((18446744073709551615,2), 2); uhex(a[2])", "0xFFFFFFFFFFFFFFFF"} ,
    {ParityBasicCase::Kind::Expected, "range(1, 4)", "(1,2,3)"} ,
    {ParityBasicCase::Kind::Expected, "range(1, 10, 3)", "(1,4,7)"} ,
    {ParityBasicCase::Kind::Expected, "range(10, 0, -2)", "(10,8,6,4,2)"} ,
    {ParityBasicCase::Kind::Expected, "range(1, 2)", "(1)"} ,
    {ParityBasicCase::Kind::Expected, "range(-1, 2)", "(-1,0,1)"} ,
    {ParityBasicCase::Kind::Expected, "range(1, 10, 2)", "(1,3,5,7,9)"} ,
    {ParityBasicCase::Kind::Expected, "sum(range(1,4))", "6"} ,
    {ParityBasicCase::Kind::ErrorContains, "range(1, 10, 0)", "step must not be zero"} ,
    {ParityBasicCase::Kind::ErrorContains, "range(1, 10, -1)", "produces no values"} ,
    {ParityBasicCase::Kind::ErrorContains, "range(10, 0, 2)", "produces no values"} ,
    {ParityBasicCase::Kind::ErrorContains, "range(0, 0, 2)", "produces no values"} ,
    {ParityBasicCase::Kind::ErrorContains, "range(1, 1)", "produces no values"} ,
    {ParityBasicCase::Kind::ErrorContains, "range(5, 1)", "produces no values"} ,
    {ParityBasicCase::Kind::ErrorContains, "range(1, 2, 1.5)", "expects integer values"} ,
    {ParityBasicCase::Kind::ErrorContains, "range((1,2), 5)", "expects integer values"} ,
#if SMARTMATH_TIME_VALUES
    {ParityBasicCase::Kind::ErrorContains, "range(1, 1:00)", "expects integer values"} ,
#endif
    {ParityBasicCase::Kind::ErrorContains, "range", "function: range(start, stop [, step])"} ,
    {ParityBasicCase::Kind::ErrorContains, "range()", "expects at least 2 arguments"} ,
    {ParityBasicCase::Kind::ErrorContains, "range(1)", "expects at least 2 arguments"} ,
    {ParityBasicCase::Kind::ErrorContains, "range(1,2,3,4)", "expects 3 argument(s)"} ,
    {ParityBasicCase::Kind::Expected, "hex(range(1,4)[0])", "0x1"} ,
    {ParityBasicCase::Kind::Expected, "range(1, 1001)[999]", "1000"} ,
    {ParityBasicCase::Kind::ErrorContains, "range(1, 1002)", "produces too many values"} ,
    {ParityBasicCase::Kind::ErrorContains, "range(1, 2**53)", "produces too many values"} ,
    {ParityBasicCase::Kind::Expected, "repeat(0, 1000)[999]", "0"} ,
    {ParityBasicCase::Kind::ErrorContains, "repeat(0, 1001)", "produces too many values"} ,
    {ParityBasicCase::Kind::Expected, "repeat((1,2), 500)[999]", "2"} ,
    {ParityBasicCase::Kind::ErrorContains, "repeat((1,2), 501)", "produces too many values"} ,
    {ParityBasicCase::Kind::ErrorContains, "repeat(1, 2**53)", "produces too many values"} ,
    {ParityBasicCase::Kind::Expected, "len((10,20,30))", "3"} ,
#if SMARTMATH_TIME_VALUES
    {ParityBasicCase::Kind::Expected, "len(1:30)", "1"} ,
#endif
    {ParityBasicCase::Kind::Expected, "len(range(0,10))", "10"} ,
    {ParityBasicCase::Kind::Expected, "len(1,2,3)", "3"} ,
    {ParityBasicCase::Kind::Expected, "len()", "0"} ,
    {ParityBasicCase::Kind::Expected, "len(())", "0"} ,
    {ParityBasicCase::Kind::Expected, "len(1, (2, 3))", "3"} ,
    {ParityBasicCase::Kind::Expected, "len(42)", "1"} ,
    {ParityBasicCase::Kind::Expected, "len(nan)", "1"} ,
    {ParityBasicCase::Kind::Expected, "length((1,2))", "2"} ,
    {ParityBasicCase::Kind::ErrorContains, "len", "function: len(x)"} ,
    {ParityBasicCase::Kind::ErrorContains, "length", "function: len(x)"} ,
    {ParityBasicCase::Kind::Expected, "a=(1,2,3,4,5,6,7,8,9); a[2:7]", "(3, 4, 5, 6, 7)"} ,
    {ParityBasicCase::Kind::Expected, "a=(1,2,3,4,5,6,7,8,9); a[3:-3]", "(4, 5, 6)"} ,
    {ParityBasicCase::Kind::Expected, "a=(1,2,3,4,5,6,7,8,9); a[0:1]", "(1)"} ,
    {ParityBasicCase::Kind::Expected, "a=(1,2,3,4,5,6,7,8,9); a[-3:-2]", "(7)"} ,
    {ParityBasicCase::Kind::Expected, "a=(1,2,3,4,5,6,7,8,9); a[-2:-3]", "()"} ,
    {ParityBasicCase::Kind::Expected, "a=(1,2,3,4,5,6,7,8,9); a[3:1]", "()"} ,
    {ParityBasicCase::Kind::Expected, "a=(1,2,3,4,5,6,7,8,9); a[1:1]", "()"} ,
    {ParityBasicCase::Kind::Expected, "a=(1,2,3,4,5,6,7,8,9); a[11:11]", "()"} ,
    {ParityBasicCase::Kind::Expected, "a=(1,2,3,4,5,6,7,8,9); a[:3]", "(1, 2, 3)"} ,
    {ParityBasicCase::Kind::Expected, "a=(1,2,3,4,5,6,7,8,9); a[3:]", "(4, 5, 6, 7, 8, 9)"} ,
    {ParityBasicCase::Kind::Expected, "a=(1,2,3,4,5,6,7,8,9); a[:]", "(1, 2, 3, 4, 5, 6, 7, 8, 9)"} ,
    {ParityBasicCase::Kind::Expected, "a=(1,2,3,4,5,6,7,8,9); a[::2]", "(1, 3, 5, 7, 9)"} ,
    {ParityBasicCase::Kind::Expected, "a=(1,2,3,4,5,6,7,8,9); a[1:8:2]", "(2, 4, 6, 8)"} ,
    {ParityBasicCase::Kind::Expected, "a=(1,2,3,4,5,6,7,8,9); a[::-1]", "(9, 8, 7, 6, 5, 4, 3, 2, 1)"} ,
    {ParityBasicCase::Kind::Expected, "a=(1,2,3,4,5,6,7,8,9); a[8:1:-2]", "(9, 7, 5, 3)"} ,
    {ParityBasicCase::Kind::Expected, "a=(1,2,3,4,5,6,7,8,9); a[1:30]", "(2, 3, 4, 5, 6, 7, 8, 9)"} ,
    {ParityBasicCase::Kind::ErrorContains, "a=(1,2,3,4,5,6,7,8,9); a[::0]", "slice step must not be zero"} ,
    {ParityBasicCase::Kind::ErrorContains, "a=(1,2,3,4,5,6,7,8,9); a[1.5:3]", "array index must be an integer"} ,
    {ParityBasicCase::Kind::Expected, "(10,20,30)[1:2]", "(20)"} ,
    {ParityBasicCase::Kind::Expected, "a=(1,2,3,4,5); a[1:4][0]", "2"} ,
    {ParityBasicCase::Kind::Expected, "a=(1,2,3,4,5,6,7,8,9); a[-100:100]", "(1, 2, 3, 4, 5, 6, 7, 8, 9)"} ,
    {ParityBasicCase::Kind::Expected, "a=(1,2,3,4,5,6,7,8,9); a[100:-100:-1]", "(9, 8, 7, 6, 5, 4, 3, 2, 1)"} ,
    {ParityBasicCase::Kind::Expected, "_=(1,2,3,4,5,6); ff(a)=a[1:4]; ff(_)", "(2, 3, 4)"} ,
    {ParityBasicCase::Kind::Expected, "a=(1,2,3,4,5,6,7,8,9); i=2; j=7; a[i:j]", "(3, 4, 5, 6, 7)"} ,
    {ParityBasicCase::Kind::ErrorContains, "5[0:1]", "indexing requires an array value"} ,
#if SMARTMATH_TIME_VALUES
    {ParityBasicCase::Kind::ErrorContains, "a=range(1,101); a[1m3s]", "array index must be an integer"} ,
    {ParityBasicCase::Kind::ErrorContains, "a=range(1,101); a[1m:1m3s]", "array index must be an integer"} ,
#endif
    {ParityBasicCase::Kind::Expected, "a=5; a==5", "1"} ,
    {ParityBasicCase::Kind::Expected, "b=3; b==4", "0"} ,
    {ParityBasicCase::Kind::Expected, "7==7", "1"} ,
    {ParityBasicCase::Kind::Expected, "3==3 && 4==4 && 5==5", "1"} ,
    {ParityBasicCase::Kind::Expected, "(1==1)==(0==0)", "1"} ,
    {ParityBasicCase::Kind::Expected, "5==5==1", "1"} ,
    {ParityBasicCase::Kind::Expected, "u=2; v=2; u==v", "1"} ,
    {ParityBasicCase::Kind::Expected, "k=9; k==9 && k==3+6", "1"} ,
    {ParityBasicCase::Kind::Expected, "a=1; b=1; a==b==1", "1"} ,
    {ParityBasicCase::Kind::Expected, "int(7/3) << 60", "2305843009213693952"} ,
    {ParityBasicCase::Kind::Expected, "sqrt(9) << 2", "12"} ,
    {ParityBasicCase::Kind::Expected, "abs(-4) << 1", "8"} ,
    {ParityBasicCase::Kind::Expected, "sum(3,5) << 2", "32"} ,
    {ParityBasicCase::Kind::Expected, "(7/3.5) << 3", "16"} ,
    {ParityBasicCase::Kind::Expected, "x=7; x=x", "7"} ,
    {ParityBasicCase::Kind::Expected, "x=7; x==x", "1"} ,
    {ParityBasicCase::Kind::Expected, "x=7; (x)=(x)", "1"} ,
    {ParityBasicCase::Kind::Expected, "x=5; x=x+1", "6"} ,
    {ParityBasicCase::Kind::Expected, "z=0; (z)=(z)", "1"} ,
    {ParityBasicCase::Kind::Expected, "a=2; 1+a=3", "1"} ,
    {ParityBasicCase::Kind::Expected, "b=2; b+0=2", "1"} ,
    {ParityBasicCase::Kind::Expected, "c=1; c=2; c=c", "2"} ,
    {ParityBasicCase::Kind::Expected, "d=4; d = d", "4"} ,
    {ParityBasicCase::Kind::Expected, "t=3; (t=3)", "1"} ,
    {ParityBasicCase::Kind::Expected, "x=2;y=5;x+y=x", "0"} ,
    {ParityBasicCase::Kind::Expected, "x=2;y=3;x*y=x*y", "1"} ,
    {ParityBasicCase::Kind::ErrorContains, "e=1", "reserved constant name"} ,
    {ParityBasicCase::Kind::ErrorContains, "PI=2", "reserved constant name"} ,
    {ParityBasicCase::Kind::ErrorContains, "f(e)=e+1", "reserved constant name"} ,
    {ParityBasicCase::Kind::ErrorContains, "pi(x)=x", "reserved constant name"} ,
    {ParityBasicCase::Kind::Expected, "inf+1", "inf"} ,
    {ParityBasicCase::Kind::Expected, "INF+1", "inf"} ,
    {ParityBasicCase::Kind::Expected, "nan+1", "nan"} ,
    {ParityBasicCase::Kind::Expected, "NAN+1", "nan"} ,
    {ParityBasicCase::Kind::ErrorContains, "inf=1", "reserved constant name"} ,
    {ParityBasicCase::Kind::ErrorContains, "Inf=1", "reserved constant name"} ,
    {ParityBasicCase::Kind::ErrorContains, "f(inf)=inf+1", "reserved constant name"} ,
    {ParityBasicCase::Kind::ErrorContains, "inf(x)=x", "reserved constant name"} ,
    {ParityBasicCase::Kind::ErrorContains, "nan=1", "reserved constant name"} ,
    {ParityBasicCase::Kind::ErrorContains, "NaN=1", "reserved constant name"} ,
    {ParityBasicCase::Kind::ErrorContains, "f(nan)=nan+1", "reserved constant name"} ,
    {ParityBasicCase::Kind::ErrorContains, "nan(x)=x", "reserved constant name"} ,
    {ParityBasicCase::Kind::Expected, "-inf", "-inf"} ,
    {ParityBasicCase::Kind::Expected, "-Inf+1", "-inf"} ,
    {ParityBasicCase::Kind::Expected, "sum(-inf,2,3)", "-inf"} ,
    {ParityBasicCase::Kind::Expected, "max(-inf,2,3)", "3"} ,
    {ParityBasicCase::Kind::ErrorContains, "(-inf)&1", "bitwise operands must be integer values"} ,
    {ParityBasicCase::Kind::Expected, "(-inf)==(-inf)", "1"} ,
    {ParityBasicCase::Kind::Expected, "-inf + inf", "nan"} ,
    {ParityBasicCase::Kind::Expected, "inf/-inf", "nan"} ,
    {ParityBasicCase::Kind::Expected, "pow(inf,-inf)", "0"} ,
    {ParityBasicCase::Kind::Expected, "pow(-inf,3)", "-inf"} ,
    {ParityBasicCase::Kind::Expected, "pow(-inf,2)", "inf"} ,
    {ParityBasicCase::Kind::Expected, "inf-inf", "nan"} ,
    {ParityBasicCase::Kind::Expected, "(-inf)/(-inf)", "nan"} ,
    {ParityBasicCase::Kind::Expected, "inf*0", "nan"} ,
    {ParityBasicCase::Kind::Expected, "0*inf", "nan"} ,
    {ParityBasicCase::Kind::Expected, "pow(inf,-2)", "0"} ,
    {ParityBasicCase::Kind::Expected, "hex(~0x0D)", "-0xE"} ,
    {ParityBasicCase::Kind::Expected, "hex(-1)", "-0x1"} ,
    {ParityBasicCase::Kind::Expected, "uhex(~0x0D)", "0xFFFFFFFFFFFFFFF2"} ,
    {ParityBasicCase::Kind::Expected, "uhex(-1)", "0xFFFFFFFFFFFFFFFF"} ,
    {ParityBasicCase::Kind::Expected, "ubin(-1)", "0b1111111111111111111111111111111111111111111111111111111111111111"} ,
    {ParityBasicCase::Kind::Expected, "uoct(-1)", "0o1777777777777777777777"} ,
    {ParityBasicCase::Kind::ErrorContains, "uhex()", "expects at least 1 argument"} ,
    {ParityBasicCase::Kind::ErrorContains, "uhex", "function: uhex(...)"} ,
    {ParityBasicCase::Kind::Expected, "uhex(1,2)", "(0x1,0x2)"} ,
    {ParityBasicCase::Kind::Expected, "bin(-2)", "-0b10"} ,
    {ParityBasicCase::Kind::Expected, "hex(int((9007199254740992+2)/1))", "0x20000000000002"} ,
    {ParityBasicCase::Kind::Expected, "hex(int((9007199254740992+2)/2+0.0))", "0x10000000000001"} ,
    {ParityBasicCase::Kind::Expected, "a=int(((9007199254740992+2),(9007199254740992+6))/1); hex(a[0])", "0x20000000000002"} ,
    {ParityBasicCase::Kind::Expected, "a=int(((9007199254740992+2),(9007199254740992+6))/1); mod(a[1],4)", "2"} ,
    {ParityBasicCase::Kind::Expected, "a=int(((9007199254740992+2),(9007199254740992+6))/2); a[0]&1", "1"} ,
    {ParityBasicCase::Kind::Expected, "a=int(((9007199254740992+2),(9007199254740992+6))/2); hex(a[1])", "0x10000000000003"} ,
    {ParityBasicCase::Kind::Expected, "-1>>1", "-1"} ,
    {ParityBasicCase::Kind::Expected, "uhex(1<<63)", "0x8000000000000000"} ,
    {ParityBasicCase::Kind::Expected, "a=(-1,-2)>>1; a[0]", "-1"} ,
    {ParityBasicCase::Kind::Expected, "a=(-1,3)<<1; uhex(a[0])", "0xFFFFFFFFFFFFFFFE"} ,
    {ParityBasicCase::Kind::Expected, "a=(-1,-2)>>1; uhex(a[1])", "0xFFFFFFFFFFFFFFFF"} ,
    {ParityBasicCase::Kind::Expected, "(1+2)(3+4)", "21"} ,
    {ParityBasicCase::Kind::Expected, "2(1+pi)", "8.283185307179586"} ,
    {ParityBasicCase::Kind::Expected, "a=(3,4)+(5,6); hex(a[0])", "0x8"} ,
    {ParityBasicCase::Kind::Expected, "a=(3,4)*2; hex(a[1])", "0x8"} ,
    {ParityBasicCase::Kind::Expected, "a=2*(3,4); hex(a[1])", "0x8"} ,
    {ParityBasicCase::Kind::Expected, "a=(1,2)&3; hex(a[1])", "0x2"} ,
    {ParityBasicCase::Kind::Expected, "a=3&(1,2); hex(a[1])", "0x2"} ,
    {ParityBasicCase::Kind::Expected, "a=(1,2)&(3,4); hex(a[1])", "0x0"} ,
    {ParityBasicCase::Kind::Expected, "log((8,100),(2,10))", "(3,2)"} ,
    {ParityBasicCase::Kind::ErrorContains, "clamp((1,9),(0,10),(5,7))", "expects scalar min/max"} ,
    {ParityBasicCase::Kind::ErrorContains, "clamp(5,(1,6),(4,7))", "expects scalar min/max"} ,
    {ParityBasicCase::Kind::Expected, "gcd((84,30),6)", "(6,6)"} ,
    {ParityBasicCase::Kind::Expected, "gcd(6,(84,30))", "(6,6)"} ,
    {ParityBasicCase::Kind::Expected, "lcm((6,8),3)", "(6,24)"} ,
    {ParityBasicCase::Kind::Expected, "lcm((6,8),(3,5))", "(6,40)"} ,
    {ParityBasicCase::Kind::Expected, "16>>1>>2", "2"} ,
    {ParityBasicCase::Kind::Expected, "7&3|8", "11"} ,
    {ParityBasicCase::Kind::Expected, "7^3^1", "5"} ,
    {ParityBasicCase::Kind::Expected, "hex(-2)", "-0x2"} ,
    {ParityBasicCase::Kind::Expected, "uhex(-2)", "0xFFFFFFFFFFFFFFFE"} ,
    {ParityBasicCase::Kind::Expected, "hex((15,-2))", "(0xF,-0x2)"} ,
    {ParityBasicCase::Kind::ErrorContains, "hypot(1,2,3,4)", "expects 2 argument(s)"} ,
    {ParityBasicCase::Kind::Expected, "123", "123"} ,
    {ParityBasicCase::Kind::ErrorContains, "unknownFunc(1)", "unknown function"} ,
    {ParityBasicCase::Kind::Expected, "2+3; ans", "5"} ,
    {ParityBasicCase::Kind::Expected, "(e=3)", "0"} ,
    {ParityBasicCase::Kind::Expected, "(pi=3.141592653589793)", "1"} ,
    {ParityBasicCase::Kind::ErrorContains, "(1,2,3)[1.2]", "array index must be an integer"} ,
    {ParityBasicCase::Kind::ErrorContains, "(1,2,3)[(1,2)]", "array index must be a scalar integer"} ,
    {ParityBasicCase::Kind::Expected, "1<<0", "1"} ,
    {ParityBasicCase::Kind::Expected, "8>>0", "8"} ,
    {ParityBasicCase::Kind::Expected, "1<<63", "9223372036854775808"} ,
    {ParityBasicCase::Kind::Expected, "-1>>63", "-1"} ,
    {ParityBasicCase::Kind::Expected, "a=(1,2)<<0; a", "(1,2)"} ,
    {ParityBasicCase::Kind::Expected, "a=(8,9)>>0; a", "(8,9)"} ,
    {ParityBasicCase::Kind::Expected, "a=(1,2)<<63; uhex(a[0])", "0x8000000000000000"} ,
    {ParityBasicCase::Kind::Expected, "a=(-1,-2)>>63; a[1]", "-1"} ,
    {ParityBasicCase::Kind::ErrorContains, "0b102", "unexpected token"} ,
    {ParityBasicCase::Kind::ErrorContains, "0o89", "invalid octal literal"} ,
    {ParityBasicCase::Kind::ErrorContains, "0x1G", "unexpected token"} ,
    {ParityBasicCase::Kind::ErrorContains, "hex=1", "reserved function name"} ,
    {ParityBasicCase::Kind::ErrorContains, "HEX=1", "reserved function name"} ,
    {ParityBasicCase::Kind::ErrorContains, "random=1", "reserved function name"} ,
    {ParityBasicCase::Kind::Expected, "0xAA; hex", "0xAA"} ,
    {ParityBasicCase::Kind::Expected, "0xAA; hex()", "0xAA"} ,
    {ParityBasicCase::Kind::Expected, "(0x3C & 0x75, 0x01 | 0x30); hex", "(0x34,0x31)"} ,
    {ParityBasicCase::Kind::Expected, "(8,9); bin()", "(0b1000,0b1001)"} ,
    {ParityBasicCase::Kind::Expected, "15; uhex", "0xF"} ,
    {ParityBasicCase::Kind::Expected, "(45,60,90); rad", "(0.7853981633974483, 1.047197551196598, 1.570796326794897)"} ,
    {ParityBasicCase::Kind::Expected, "(pi/4,pi/3,pi/2); deg", "(45, 60, 90)"} ,
    {ParityBasicCase::Kind::ErrorContains, "0xAA; foo()", "unknown function"} ,
    {ParityBasicCase::Kind::ErrorContains, "y(a)=g(a)+y(a)+4", "recursive function call: y"} ,
#if SMARTMATH_TIME_VALUES
    {ParityBasicCase::Kind::Expected, "d(x)=(ratio(x), x-ratio(x)); d(seconds(20ms))", "(1/50, 0)" } ,
#endif
    {ParityBasicCase::Kind::Expected, "f(x)=x*(2,3,4); f(10)", "(20, 30, 40)" } ,
    {ParityBasicCase::Kind::Expected, "1+2;", "3"} ,
    {ParityBasicCase::Kind::Expected, "   1+2 ; ", "3"} ,
    {ParityBasicCase::Kind::ErrorContains, ";", "empty statement"} ,
    {ParityBasicCase::Kind::ErrorContains, " ; ", "empty statement"} ,
    {ParityBasicCase::Kind::ErrorContains, "1;;2", "empty statement"} ,
    {ParityBasicCase::Kind::Expected, "a=1; b=2; c=3; a+b+c;", "6"} ,
    {ParityBasicCase::Kind::Expected, "r=10; r;", "10"} ,
    {ParityBasicCase::Kind::Expected, "r()=5; r();", "5"} ,
    {ParityBasicCase::Kind::Expected, "1+2; 3+4", "7"} ,
    {ParityBasicCase::Kind::Expected, "1+2; 3+4;", "7"} ,
    {ParityBasicCase::Kind::Expected, "f(x)=x+1; f(3)", "4"} ,
    {ParityBasicCase::Kind::Expected, "f(x)=x+1; f(3);", "4"} ,
    {ParityBasicCase::Kind::Expected, "0xAA; hex", "0xAA"} ,
    {ParityBasicCase::Kind::Expected, "0xAA; hex;", "0xAA"} ,
    {ParityBasicCase::Kind::Expected, "(10,20,30); hex()", "(0xA,0x14,0x1E)"} ,
    {ParityBasicCase::Kind::Expected, "(10,20,30); hex;", "(0xA,0x14,0x1E)"} ,
    {ParityBasicCase::Kind::Expected, "sortby((2,1), polar)", "(1,2)"} ,
    {ParityBasicCase::Kind::Expected, "f(x)=x*(10,20); sortby((3,1,2), f)", "(1,2,3)"},
    {ParityBasicCase::Kind::Expected, "sum(1,Inf,2)", "inf"},
    {ParityBasicCase::Kind::Expected, "prod(1,Inf,2)", "inf"},
    {ParityBasicCase::Kind::Expected, "product(1,Inf,2)", "inf"},
    {ParityBasicCase::Kind::Expected, "avg(1,Inf,2)", "inf"},
    {ParityBasicCase::Kind::Expected, "mean(1,Inf,2)", "inf"},
    {ParityBasicCase::Kind::Expected, "min(1,Inf,2)", "1"},
    {ParityBasicCase::Kind::Expected, "max(1,Inf,2)", "inf"},
    {ParityBasicCase::Kind::Expected, "median(1,Inf,2)", "2"},
    {ParityBasicCase::Kind::Expected, "variance(1,Inf,2)", "nan"},
    {ParityBasicCase::Kind::Expected, "stddev(1,Inf,2)", "nan"},
    {ParityBasicCase::Kind::Expected, "sort(1,Inf,2)", "(1,2,inf)"},
    {ParityBasicCase::Kind::Expected, "sorted(1,Inf,2)", "(1,2,inf)"},
    {ParityBasicCase::Kind::Expected, "reverse(1,Inf,2)", "(2,inf,1)"},
    {ParityBasicCase::Kind::Expected, "reversed(1,Inf,2)", "(2,inf,1)"},
    {ParityBasicCase::Kind::Expected, "unique(1,Inf,2,Inf)", "(1,inf,2)"},
    {ParityBasicCase::Kind::Expected, "unpack(1,Inf,2)", "(1,inf,2)"},
    {ParityBasicCase::Kind::Expected, "sum(unpack(1,Inf,2))", "inf"},
    {ParityBasicCase::Kind::Expected, "sortby((1,Inf,2),abs)", "(1,2,inf)"}
};

static const ParityBasicCase kParityTimeFromSmokeCases[] = {
    {ParityBasicCase::Kind::Expected, "0:60", "01:00"} ,
    {ParityBasicCase::Kind::Expected, "1:30 + 2:45.111", "04:15.111"} ,
    {ParityBasicCase::Kind::Expected, "second + 5", "00:06"} ,
    {ParityBasicCase::Kind::Expected, "minute - second", "00:59"} ,
    {ParityBasicCase::Kind::Expected, "1:00 == 0:60", "1"} ,
    {ParityBasicCase::Kind::Expected, "0:30 > 20", "1"} ,
    {ParityBasicCase::Kind::Expected, "milliseconds(minute + 30*second)", "90000"} ,
    {ParityBasicCase::Kind::Expected, "seconds(2:00)", "120"} ,
    {ParityBasicCase::Kind::Expected, "minutes(0:45)", "0.75"} ,
    {ParityBasicCase::Kind::Expected, "hours(12:00:00)", "12"} ,
    {ParityBasicCase::Kind::Expected, "days(12:00:00)", "0.5"} ,
    {ParityBasicCase::Kind::Expected, "0:25 * 6", "02:30"} ,
    {ParityBasicCase::Kind::Expected, "1:30 / 0:30", "3"} ,
    {ParityBasicCase::Kind::ErrorContains, "1::0", "empty segment"} ,
    {ParityBasicCase::Kind::ErrorContains, "second=1", "reserved constant name"} ,
    {ParityBasicCase::Kind::ErrorContains, "minute=1", "reserved constant name"} ,
    {ParityBasicCase::Kind::ErrorContains, "hour=1", "reserved constant name"} ,
    {ParityBasicCase::Kind::ErrorContains, "day=1", "reserved constant name"} ,
    {ParityBasicCase::Kind::ErrorContains, "sin(minute)", "incompatible operands"} ,
    {ParityBasicCase::Kind::Expected, "milliseconds((0:01:00,0:02:00))", "(60000, 120000)"} ,
    {ParityBasicCase::Kind::Expected, "seconds((0:30,1:00))", "(30, 60)"} ,
    {ParityBasicCase::Kind::Expected, "hours((1:00:00,0:30:00))", "(1, 0.5)"} ,
    {ParityBasicCase::Kind::Expected, "minutes((0:30,1:00))", "(0.5, 1)"} ,
    {ParityBasicCase::Kind::ErrorContains, "milliseconds((5,6))", "milliseconds() expects a time value"} ,
    {ParityBasicCase::Kind::ErrorContains, "milliseconds(5)", "milliseconds() expects a time value"} ,
    {ParityBasicCase::Kind::Expected, "sum(0:30,1:00)", "01:30"} ,
    {ParityBasicCase::Kind::ErrorContains, "1:00*1:00", "incompatible operands"} ,
    {ParityBasicCase::Kind::ErrorContains, "product(1:00,0:30)", "incompatible operands"} ,
    {ParityBasicCase::Kind::ErrorContains, "prod(1:00)", "incompatible operands"} ,
    {ParityBasicCase::Kind::ErrorContains, "product(1:00)", "incompatible operands"} ,
    {ParityBasicCase::Kind::Expected, "1d2h3m4s5ms", "1:02:03:04.005"} ,
    {ParityBasicCase::Kind::Expected, "1d 2h 3m 4s 5ms", "1:02:03:04.005"} ,
    {ParityBasicCase::Kind::Expected, "1d3m + 2h5ms + 4s == 1:02:03:04.005", "1"} ,
    {ParityBasicCase::Kind::Expected, "5000ms", "00:05"} ,
    {ParityBasicCase::Kind::Expected, "2h90m", "03:30:00"} ,
    {ParityBasicCase::Kind::Expected, "23h3600s", "1:00:00:00"} ,
    {ParityBasicCase::Kind::Expected, "20m or 10h", "1"} ,
    {ParityBasicCase::Kind::Expected, "20m10s and 10h5m", "1"} ,
    {ParityBasicCase::Kind::Expected, "1h and not 0", "1"} ,
    {ParityBasicCase::Kind::Expected, "-1m 1s", "-01:01"} ,
    {ParityBasicCase::Kind::ErrorContains, "1d2d", "compact time literal: unit order or duplicate unit"} ,
    {ParityBasicCase::Kind::ErrorContains, "1h9d", "compact time literal: unit order or duplicate unit"} ,
    {ParityBasicCase::Kind::ErrorContains, "45s37m", "compact time literal: unit order or duplicate unit"} ,
    {ParityBasicCase::Kind::ErrorContains, "1day", "compact time literal: invalid suffix"} ,
    {ParityBasicCase::Kind::Expected, "1e2", "100"} ,
    {ParityBasicCase::Kind::Expected, "milliseconds(1.5*1ms)", "2"} ,
    {ParityBasicCase::Kind::Expected, "milliseconds(1.5*00:00.001)", "2"} ,
    {ParityBasicCase::Kind::Expected, "seconds(-1s-1m)", "-61"} ,
    {ParityBasicCase::Kind::Expected, "1d == 1:00:00:00", "1"} ,
    {ParityBasicCase::Kind::ErrorContains, "1d2", "compact time literal: expected unit suffix"} ,
    {ParityBasicCase::Kind::ErrorContains, "1ms1m", "compact time literal: unit order or duplicate unit"} ,
    {ParityBasicCase::Kind::Expected, "real(5)", "5"} ,
    {ParityBasicCase::Kind::Expected, "imag(-3.5)", "0"} ,
    {ParityBasicCase::Kind::Expected, "phase(42)", "0"} ,
    {ParityBasicCase::Kind::Expected, "polar(10)", "(10,0)"} ,
    {ParityBasicCase::Kind::Expected, "conj(-7)", "-7"} ,
    {ParityBasicCase::Kind::Expected, "cart(3.5)", "3.5"} ,
    {ParityBasicCase::Kind::Expected, "cart(polar(2))", "2"} ,
    {ParityBasicCase::Kind::Expected, "imag((5,10))", "(0,0)"} ,
    {ParityBasicCase::Kind::Expected, "cart((5,0))", "5"} ,
    {ParityBasicCase::Kind::ErrorContains, "cart((5,1))", "incompatible operands"} ,
#include "../tools/smoke_parity_append_generated.cpp.inc"
};

std::vector<TestCase> buildParityBasicFromSmokeCases() {
  std::vector<TestCase> t;
  constexpr std::size_t kCount = sizeof(kParityBasicFromSmokeCases) / sizeof(kParityBasicFromSmokeCases[0]);
  t.reserve(kCount);

  for (std::size_t i = 0; i < kCount; ++i) {
    const auto& c = kParityBasicFromSmokeCases[i];
    std::string name("parity/basic: "); name += c.expr;
    if (c.kind == ParityBasicCase::Kind::Expected) {
      t.push_back({name, [expr = c.expr, payload = c.payload](std::string& why) {
                     MathParser p;
                     return expectEval(p, expr, payload, why);
                   }});
    } else {
      t.push_back({name, [expr = c.expr, payload = c.payload](std::string& why) {
                     MathParser p;
                     return expectEvalErrorContains(p, expr, payload, why);
                   }});
    }
  }
  return t;
}

std::vector<TestCase> buildParityTimeFromSmokeCases() {
  std::vector<TestCase> t;
  constexpr std::size_t kCount = sizeof(kParityTimeFromSmokeCases) / sizeof(kParityTimeFromSmokeCases[0]);
  t.reserve(kCount);
  for (std::size_t i = 0; i < kCount; ++i) {
    const auto& c = kParityTimeFromSmokeCases[i];
    std::string name("parity/time: ");
    name += c.expr;
    if (c.kind == ParityBasicCase::Kind::Expected) {
      t.push_back({name, [expr = c.expr, payload = c.payload](std::string& why) {
                     MathParser p;
                     return expectEval(p, expr, payload, why);
                   }});
    } else {
      t.push_back({name, [expr = c.expr, payload = c.payload](std::string& why) {
                     MathParser p;
                     return expectEvalErrorContains(p, expr, payload, why);
                   }});
    }
  }
  return t;
}

std::vector<TestCase> buildTimeValuesSupportOptionCases() {
  std::vector<TestCase> t;
#if !SMARTMATH_TIME_VALUES
  return t;
#else
  t.push_back({"time-opt: flag and duration eval under enabled mode",
               [](std::string& why) {
                 MathParser p;
                 p.setSupportTimeValues(true);
                 if (!p.getSupportTimeValues()) {
                   why = "expected true after setSupportTimeValues(true)";
                   return false;
                 }
                 p.parseAndEvaluate("1:30 + 2:45.111");
                 if (!p.getError().empty()) {
                   why = p.getError();
                   return false;
                 }
                 if (p.getResult() != "04:15.111") {
                   why = std::string("got ") + p.getResult();
                   return false;
                 }
                 p.setSupportTimeValues(false);
                 if (p.getSupportTimeValues()) {
                   why = "expected false after setSupportTimeValues(false)";
                   return false;
                 }
                 return true;
               }});
  t.push_back({"time-opt: fresh parser defaults to on",
               [](std::string& why) {
                 MathParser p;
                 if (!p.getSupportTimeValues()) {
                   why = "expected default true";
                   return false;
                 }
                 return true;
               }});
  t.push_back({"time-opt: literals and converters disabled when support off",
               [](std::string& why) {
                 MathParser p;
                 p.setSupportTimeValues(false);
                 p.parseAndEvaluate("1:30 + 2:45.111");
                 if (p.getError().empty()) {
                   why = "expected evaluation failure for colon time literal with time support off";
                   return false;
                 }
                 std::string el = p.getError();
                 std::transform(el.begin(), el.end(), el.begin(), [](unsigned char c) { return static_cast<char>(std::tolower(c)); });
                 const bool okParse =
                     el.find("unexpected token") != std::string::npos ||
                     el.find("unexpected characters") != std::string::npos ||
                     el.find("invalid numeric") != std::string::npos ||
                     el.find("invalid segment") != std::string::npos;
                 if (!okParse) {
                   why = std::string("unexpected error for colon literal: ") + p.getError();
                   return false;
                 }
                 p.parseAndEvaluate("second");
                 if (p.getError().empty()) {
                   why = "expected failure for second with time support off";
                   return false;
                 }
                 el = p.getError();
                 std::transform(el.begin(), el.end(), el.begin(), [](unsigned char c) { return static_cast<char>(std::tolower(c)); });
                 if (el.find("unknown variable") == std::string::npos) {
                   why = std::string("unexpected error for second: ") + p.getError();
                   return false;
                 }
                 p.parseAndEvaluate("milliseconds(5)");
                 if (p.getError().empty()) {
                   why = "expected failure for milliseconds(5) with time support off";
                   return false;
                 }
                 el = p.getError();
                 std::transform(el.begin(), el.end(), el.begin(), [](unsigned char c) { return static_cast<char>(std::tolower(c)); });
                 if (el.find("incompatible operands") == std::string::npos) {
                   why = std::string("unexpected error for milliseconds(5): ") + p.getError();
                   return false;
                 }
                 return true;
               }});
  t.push_back({"time-opt: oversized plain decimal is numeric, not time literal",
               [](std::string& why) {
                 MathParser p;
                 p.setSupportTimeValues(true);
                 p.parseAndEvaluate("100000000000000000000");
                 if (!p.getError().empty()) {
                   why = p.getError();
                   return false;
                 }
                 if (!smokeScalarCloseEnough(p.getResult(), "1.0000000000000000e+20")) {
                   why = std::string("got ") + p.getResult();
                   return false;
                 }
                 return true;
               }});
  t.push_back({"time-opt: UDF accepts duration argument",
               [](std::string& why) {
                 MathParser p;
                 p.setSupportTimeValues(true);
                 p.parseAndEvaluate("_=0:00; rd(t)=ratio(days(t)); rd(1h)");
                 if (!p.getError().empty()) {
                   why = p.getError();
                   return false;
                 }
                 if (p.getResult() != "1/24") {
                   why = std::string("got ") + p.getResult();
                   return false;
                 }
                 return true;
               }});
  t.push_back({"time-opt: time converter arity enforced by table",
               [](std::string& why) {
                 MathParser p;
                 p.setSupportTimeValues(true);
                 if (!expectEvalErrorContains(p, "hours()", "hours() expects 1 argument(s)", why)) {
                   return false;
                 }
                 return expectEvalErrorContains(p, "days(1:00,2:00)", "days() expects 1 argument(s)", why);
               }});
  t.push_back({"time-opt: sin still rejects duration argument",
               [](std::string& why) {
                 MathParser p;
                 p.setSupportTimeValues(true);
                 p.parseAndEvaluate("sin(1:00)");
                 if (p.getError().empty()) {
                   why = std::string("expected error, got ") + p.getResult();
                   return false;
                 }
                 std::string el = p.getError();
                 std::transform(el.begin(), el.end(), el.begin(), [](unsigned char c) { return static_cast<char>(std::tolower(c)); });
                 if (el.find("incompatible operands") == std::string::npos) {
                   why = std::string("unexpected error: ") + p.getError();
                   return false;
                 }
                 return true;
               }});
  t.push_back({"time-opt: sum/avg/mean/min/max/ratio on duration values (parity with Basic timeAggOk)",
               [](std::string& why) {
                 MathParser p;
                 p.setSupportTimeValues(true);
                 static const struct {
                   const char* expr;
                   const char* expect;
                 } kRows[] = {
                     {"sum(0:30, 1:00)", "01:30"},
                     {"avg(0:30, 1:30)", "01:00"},
                     {"mean(0:20, 1:00)", "00:40"},
                     {"min(1:30, 0:45)", "00:45"},
                     {"max(1:30, 0:45)", "01:30"},
                     {"sum((0:15, 0:15, 0:30))", "01:00"},
                     {"reverse(0:30, Inf, 1:00)", "(01:00,inf,00:30)"},
                     {"unpack(0:30, Inf)", "(00:30,inf)"},
                     {"unique(0:30, Inf, 0:30)", "(00:30,inf)"},
                     {"ratio((hours(1:00), 15m/1h))", "(1/60,1/4)"},
                     {"repeat((1:00, 2:00), 2)", "(01:00,02:00,01:00,02:00)"}
                 };
                 for (const auto& row : kRows) {
                   p.parseAndEvaluate(row.expr);
                   if (!p.getError().empty()) {
                     why = std::string(row.expr) + ": " + p.getError();
                     return false;
                   }
                   if (p.getResult() != row.expect) {
                     why = std::string(row.expr) + " -> " + p.getResult() + " (want " + row.expect + ")";
                     return false;
                   }
                 }
                 return true;
               }});
  t.push_back({"time-opt: aggregates with Inf reject or mix (parity with Basic timeAggErr/timeAggOk Inf)",
               [](std::string& why) {
                 MathParser p;
                 p.setSupportTimeValues(true);
                 static const struct {
                   const char* expr;
                   const char* errSubstr;
                 } kErrRows[] = {
                     {"sum(0:30, Inf, 1:00)", "sum(): time values cannot be mixed with non-time values"},
                     {"avg(0:30, Inf, 1:30)", "avg(): time values cannot be mixed with non-time values"},
                     {"mean(0:20, Inf, 1:00)", "mean(): time values cannot be mixed with non-time values"},
                     {"min(0:45, Inf, 1:30)", "min(): time values cannot be mixed with non-time values"},
                     {"max(0:45, Inf, 1:30)", "max(): time values cannot be mixed with non-time values"},
                     {"median(0:30, Inf, 1:00)", "median(): time values cannot be mixed with non-time values"},
                     {"prod(0:30, Inf, 1:00)", "incompatible operands"},
                     {"variance(0:30, Inf, 1:00)", "incompatible operands"},
                     {"stddev(0:30, Inf, 1:00)", "incompatible operands"},
                     {"sort(0:30, Inf, 1:00)", "sort(): time values cannot be mixed with non-time values"},
                     {"sort(0:30, -Inf, 1:00)", "sort(): time values cannot be mixed with non-time values"},
                     {"sort(0:30, NaN, 1:00)", "sort(): time values cannot be mixed with non-time values"},
                     {"keyf(x)=x; sortby((1:30, 2), keyf)", "array literal: time values cannot be mixed with non-time values"}
                 };
                 for (const auto& row : kErrRows) {
                   p.parseAndEvaluate(row.expr);
                   if (p.getError().empty()) {
                     why = std::string(row.expr) + " expected error, got " + p.getResult();
                     return false;
                   }
                   std::string el = p.getError();
                   std::transform(el.begin(), el.end(), el.begin(),
                       [](unsigned char c) { return static_cast<char>(std::tolower(c)); });
                   std::string want = row.errSubstr;
                   std::transform(want.begin(), want.end(), want.begin(),
                       [](unsigned char c) { return static_cast<char>(std::tolower(c)); });
                   if (el.find(want) == std::string::npos) {
                     why = std::string(row.expr) + ": " + p.getError();
                     return false;
                   }
                 }
                 return true;
               }});
#endif
  return t;
}

std::vector<TestCase> buildLambdaFunctionsSupportOptionCases() {
  std::vector<TestCase> t;
#if !SMARTMATH_LAMBDA_FUNCTIONS
  return t;
#else
  t.push_back({"lambda-opt: flag and lambda eval under enabled mode",
               [](std::string& why) {
                 MathParser p;
                 p.setSupportLambdaFunctions(true);
                 if (!p.getSupportLambdaFunctions()) {
                   why = "expected true after setSupportLambdaFunctions(true)";
                   return false;
                 }
                 p.parseAndEvaluate("sortby((3,1,2), x:-x)");
                 if (!p.getError().empty()) {
                   why = p.getError();
                   return false;
                 }
                 if (p.getResult() != "(3,2,1)") {
                   why = std::string("got ") + p.getResult();
                   return false;
                 }
                 p.parseAndEvaluate("f=x:x+2; f(3)");
                 if (!p.getError().empty()) {
                   why = p.getError();
                   return false;
                 }
                 if (p.getResult() != "5") {
                   why = std::string("got ") + p.getResult();
                   return false;
                 }
                 p.setSupportLambdaFunctions(false);
                 if (p.getSupportLambdaFunctions()) {
                   why = "expected false after setSupportLambdaFunctions(false)";
                   return false;
                 }
                 return true;
               }});
  t.push_back({"lambda-opt: fresh parser defaults to on",
               [](std::string& why) {
                 MathParser p;
                 if (!p.getSupportLambdaFunctions()) {
                   why = "expected default true";
                   return false;
                 }
                 return true;
               }});
  t.push_back({"lambda-opt: lambda syntax disabled when support off",
               [](std::string& why) {
                 MathParser p;
                 p.setSupportLambdaFunctions(false);
                 p.parseAndEvaluate("sortby((-3,-1,2), abs)");
                 if (!p.getError().empty()) {
                   why = p.getError();
                   return false;
                 }
                 if (p.getResult() != "(-1,2,-3)" && p.getResult() != "(-1, 2, -3)") {
                   why = std::string("got ") + p.getResult();
                   return false;
                 }
                 const char* const rejects[] = {
                     "f=x:x+1",
                     "f=(x,y):(x+y)",
                     "sortby((3,1,2), x:-x)",
                     "sortby((1,2), (x):(1/x))"};
                 for (const char* expr : rejects) {
                   p.parseAndEvaluate(expr);
                   if (p.getError().empty()) {
                     why = std::string("expected failure for ") + expr;
                     return false;
                   }
                   std::string el = p.getError();
                   std::transform(el.begin(), el.end(), el.begin(),
                                  [](unsigned char c) { return static_cast<char>(std::tolower(c)); });
                   if (el.find("unexpected token") == std::string::npos) {
                     why = std::string("unexpected error for ") + expr + ": " + p.getError();
                     return false;
                   }
                 }
                 return true;
               }});
#endif
  return t;
}

struct NegBandRow {
  const char* name;
  const char* expr;
  const char* expect;
  const char* errPart;
};

static const NegBandRow kNegBandRealRows[] = {
#include "../tools/neg_band_cases_generated.cpp.inc"
};

#if SMARTMATH_COMPLEX_NUMBERS
static const NegBandRow kNegBandComplexRows[] = {
#include "../tools/neg_band_cases_generated_cx.cpp.inc"
};
#endif

static const NegBandRow kPosBandRealRows[] = {
#include "../tools/pos_band_cases_generated.cpp.inc"
};

#if SMARTMATH_COMPLEX_NUMBERS
static const NegBandRow kPosBandComplexRows[] = {
#include "../tools/pos_band_cases_generated_cx.cpp.inc"
};
#endif

static const NegBandRow kFloatMagnitudeRows[] = {
#include "../tools/float_magnitude_cases_generated.cpp.inc"
};

static const NegBandRow kComplexCoverageRows[] = {
#include "../tools/complex_coverage_generated.cpp.inc"
};

static bool runNegBandRowBatch(MathParser& p, const NegBandRow* rows, std::size_t count, std::string& why) {
  for (std::size_t i = 0; i < count; ++i) {
    const NegBandRow& row = rows[i];
    if (row.errPart != nullptr) {
      if (!expectEvalErrorContains(p, row.expr, row.errPart, why)) {
        why = std::string(row.name) + ": " + why;
        return false;
      }
    } else if (!expectEval(p, row.expr, row.expect, why)) {
      why = std::string(row.name) + ": " + why;
      return false;
    }
  }
  return true;
}

std::vector<TestCase> buildTrigAngleReductionCases() {
  static const char* kNonzero[] = {
      "sin(2**52)", "sin(2**51)", "sin(2**50)", "sin(2**49)", "sin(2**48)", "sin(2**47)",
      "sin(3**32)", "sin(7**18)", "cos(2**52)", "cos(2**51)", "tan(2**52)", "tan(3**32)"};
  static const char* kNearZero[] = {
      "sin(pi)", "sin(2*pi)", "cos(pi/2)", "cos(3*pi/2)", "tan(pi)", "tan(2*pi)"};
  struct ExactRow {
    const char* expr;
    const char* expect;
  };
  static const ExactRow kExactOne[] = {
      {"tan(pi/4)", "1"},
      {"tan(3*pi/4)", "-1"},
      {"tan(-pi/4)", "-1"},
      {"tan(-3*pi/4)", "1"},
  };
  std::vector<TestCase> t;
  for (const char* expr : kNonzero) {
    t.push_back({std::string("trig-reduction/nonzero/") + expr,
                 [expr](std::string& why) {
                   MathParser p;
                   p.parseAndEvaluate(expr);
                   if (!p.getError().empty()) {
                     why = p.getError();
                     return false;
                   }
                   double v = 0.0;
                   if (!tryParseFullDouble(p.getResult(), v)) {
                     why = "non-numeric result: " + p.getResult();
                     return false;
                   }
                   if (std::fabs(v) < 1e-6) {
                     why = "expected |result| > 1e-6, got " + p.getResult();
                     return false;
                   }
                   return true;
                 }});
  }
  for (const char* expr : kNearZero) {
    t.push_back({std::string("trig-reduction/nearzero/") + expr,
                 [expr](std::string& why) {
                   MathParser p;
                   p.parseAndEvaluate(expr);
                   if (!p.getError().empty()) {
                     why = p.getError();
                     return false;
                   }
                   double v = 0.0;
                   if (!tryParseFullDouble(p.getResult(), v)) {
                     why = "non-numeric result: " + p.getResult();
                     return false;
                   }
                   if (std::fabs(v) >= 1e-9) {
                     why = "expected |result| < 1e-9, got " + p.getResult();
                     return false;
                   }
                   return true;
                 }});
  }
  for (const ExactRow& row : kExactOne) {
    const std::string expr = row.expr;
    const std::string expected = row.expect;
    t.push_back({std::string("trig-reduction/exact/") + expr,
                 [expr, expected](std::string& why) {
                   MathParser p;
                   return expectEval(p, expr, expected, why);
                 }});
  }
  return t;
}

std::vector<TestCase> buildOperatorPrecedenceDocCases() {
  struct Row {
    const char* expr;
    const char* expected;
  };
  static const Row kRows[] = {
      {"-2**2", "-4"},
      {"-(2**2)", "-4"},
      {"2**-2", "0.25"},
      {"2**3**2", "512"},
      {"2**3*4", "32"},
      {"2*3**4", "162"},
      {"3*-2**2", "-12"},
      {"-2%", "-0.02"},
      {"2*3%", "0.06"},
      {"5+2*3%", "5.06"},
      {"2+3*4", "14"},
      {"2*3+4", "10"},
      {"2+3<<1", "10"},
      {"1|2^3&6<<1", "3"},
      {"!2==1", "0"},
      {"not 2==1", "1"},
      {"1||0&&0", "1"},
      {"0&&1||1", "1"},
      {"16**-0.5", "0.25"},
      {"2(3+4)**2", "98"},
      {"~2**2", "-5"},
      {"!!3", "1"},
      {"5+(2*3)%", "5.3"},
      {"200+15%", "230"},
  };
  std::vector<TestCase> t;
  for (const Row& row : kRows) {
    const std::string expr = row.expr;
    const std::string expected = row.expected;
    t.push_back({"precedence-doc/" + expr,
                 [expr, expected](std::string& why) {
                   MathParser p;
                   return expectEval(p, expr, expected, why);
                 }});
  }
  return t;
}

std::vector<TestCase> buildNegativeArgumentMagnitudeBandCases() {
  std::vector<TestCase> t;
  t.push_back({"neg-band/real: operators and builtins (3 negative magnitude ranges)",
               [](std::string& why) {
                 MathParser p;
                 const std::size_t count = sizeof(kNegBandRealRows) / sizeof(kNegBandRealRows[0]);
                 return runNegBandRowBatch(p, kNegBandRealRows, count, why);
               }});
#if SMARTMATH_COMPLEX_NUMBERS
  t.push_back({"neg-band/complex: operators and builtins (3 negative magnitude ranges)",
               [](std::string& why) {
                 MathParser p;
                 p.setSupportComplexNumbers(true);
                 const std::size_t count = sizeof(kNegBandComplexRows) / sizeof(kNegBandComplexRows[0]);
                 return runNegBandRowBatch(p, kNegBandComplexRows, count, why);
               }});
#endif
  return t;
}

std::vector<TestCase> buildFloatMagnitudeLiteralCases() {
  std::vector<TestCase> t;
  t.push_back({"float-magnitude/literals grid (Python reference)",
               [](std::string& why) {
                 MathParser p;
                 const std::size_t count = sizeof(kFloatMagnitudeRows) / sizeof(kFloatMagnitudeRows[0]);
                 return runNegBandRowBatch(p, kFloatMagnitudeRows, count, why);
               }});
  return t;
}

std::vector<TestCase> buildPositiveArgumentMagnitudeBandCases() {
  std::vector<TestCase> t;
  t.push_back({"pos-band/real: operators and builtins (3 positive magnitude ranges)",
               [](std::string& why) {
                 MathParser p;
                 const std::size_t count = sizeof(kPosBandRealRows) / sizeof(kPosBandRealRows[0]);
                 return runNegBandRowBatch(p, kPosBandRealRows, count, why);
               }});
#if SMARTMATH_COMPLEX_NUMBERS
  t.push_back({"pos-band/complex: operators and builtins (3 positive magnitude ranges)",
               [](std::string& why) {
                 MathParser p;
                 p.setSupportComplexNumbers(true);
                 const std::size_t count = sizeof(kPosBandComplexRows) / sizeof(kPosBandComplexRows[0]);
                 return runNegBandRowBatch(p, kPosBandComplexRows, count, why);
               }});
#endif
  return t;
}

std::vector<TestCase> buildComplexNumberSupportOptionCases() {
  std::vector<TestCase> t;
#if !SMARTMATH_COMPLEX_NUMBERS
  return t;
#endif
  t.push_back({"complex-opt: flag and basic eval under enabled mode",
               [](std::string& why) {
                 MathParser p;
                 p.setSupportComplexNumbers(true);
                 if (!p.getSupportComplexNumbers()) {
                   why = "expected true after setSupportComplexNumbers(true)";
                   return false;
                 }
                 p.parseAndEvaluate("1+1");
                 if (!p.getError().empty()) {
                   why = p.getError();
                   return false;
                 }
                 if (p.getResult() != "2") {
                   why = std::string("got ") + p.getResult();
                   return false;
                 }
                 p.setSupportComplexNumbers(false);
                 if (p.getSupportComplexNumbers()) {
                   why = "expected false after setSupportComplexNumbers(false)";
                   return false;
                 }
                 return true;
               }});
  t.push_back({"complex-opt: abs preserves exact integers when complex mode on",
               [](std::string& why) {
                 MathParser p;
                 p.setSupportComplexNumbers(true);
                 if (!expectEval(p, "abs(0x7FFFFFFFFFFFFFFF+20)", "9223372036854775827", why)) {
                   return false;
                 }
                 return expectEval(p, "hex(abs(0x7FFFFFFFFFFFFFFF+20))", "0x8000000000000013", why);
               }});
  t.push_back({"complex-opt: array index rejects complex bound",
               [](std::string& why) {
                 MathParser p;
                 p.setSupportComplexNumbers(true);
                 return expectEvalErrorContains(p, "a=range(1,101); a[7+3i]", "array index must be an integer", why);
               }});
  t.push_back({"complex-opt: real/imag/cart/conj preserve exact integers",
               [](std::string& why) {
                 MathParser p;
                 p.setSupportComplexNumbers(true);
                 static const struct {
                   const char* expr;
                   const char* expect;
                 } kRows[] = {
                     {"real(0x7FFFFFFFFFFFFFFF+0x7FFFFFFFFFFFFFFFi)", "9223372036854775807"},
                     {"imag(0x7FFFFFFFFFFFFFFF+0x7FFFFFFFFFFFFFFFi)", "9223372036854775807"},
                     {"cart(0x7FFFFFFFFFFFFFFF+0x7FFFFFFFFFFFFFFFi)",
                      "9223372036854775807+9223372036854775807i"},
                     {"conj(0x7FFFFFFFFFFFFFFF+0x7FFFFFFFFFFFFFFFi)",
                      "9223372036854775807-9223372036854775807i"},
                     {"hex(real(0x7FFFFFFFFFFFFFFF+0x7FFFFFFFFFFFFFFFi))", "0x7FFFFFFFFFFFFFFF"},
                     {"hex(imag(0x7FFFFFFFFFFFFFFF+0x7FFFFFFFFFFFFFFFi))", "0x7FFFFFFFFFFFFFFF"},
                     {"hex(cart(0x7FFFFFFFFFFFFFFF+0x7FFFFFFFFFFFFFFFi))",
                      "0x7FFFFFFFFFFFFFFF+0x7FFFFFFFFFFFFFFFi"},
                     {"hex(conj(0x7FFFFFFFFFFFFFFF+0x7FFFFFFFFFFFFFFFi))",
                      "0x7FFFFFFFFFFFFFFF-0x7FFFFFFFFFFFFFFFi"},
                 };
                 for (const auto& row : kRows) {
                   if (!expectEval(p, row.expr, row.expect, why)) {
                     return false;
                   }
                 }
                 return true;
               }});
  t.push_back({"complex-opt: unary components with negative exact integers",
               [](std::string& why) {
                 MathParser p;
                 p.setSupportComplexNumbers(true);
                 static const struct {
                   const char* expr;
                   const char* expect;
                 } kRows[] = {
                     {"real(0x7FFFFFFFFFFFFFFF-0x7FFFFFFFFFFFFFFFi)", "9223372036854775807"},
                     {"imag(0x7FFFFFFFFFFFFFFF-0x7FFFFFFFFFFFFFFFi)", "-9223372036854775807"},
                     {"0x7FFFFFFFFFFFFFFF-0x7FFFFFFFFFFFFFFFi",
                      "9223372036854775807-9223372036854775807i"},
                     {"cart(0x7FFFFFFFFFFFFFFF-0x7FFFFFFFFFFFFFFFi)",
                      "9223372036854775807-9223372036854775807i"},
                     {"conj(0x7FFFFFFFFFFFFFFF-0x7FFFFFFFFFFFFFFFi)",
                      "9223372036854775807+9223372036854775807i"},
                     {"-0x7FFFFFFFFFFFFFFF-0x7FFFFFFFFFFFFFFFi",
                      "-9223372036854775807-9223372036854775807i"},
                     {"conj(-0x7FFFFFFFFFFFFFFF-0x7FFFFFFFFFFFFFFFi)",
                      "-9223372036854775807+9223372036854775807i"},
                     {"real(-0x7FFFFFFFFFFFFFFF-0x7FFFFFFFFFFFFFFFi)", "-9223372036854775807"},
                     {"imag(-0x7FFFFFFFFFFFFFFF-0x7FFFFFFFFFFFFFFFi)", "-9223372036854775807"},
                     {"cart(-0x7FFFFFFFFFFFFFFF-0x7FFFFFFFFFFFFFFFi)",
                      "-9223372036854775807-9223372036854775807i"},
                 };
                 for (const auto& row : kRows) {
                   if (!expectEval(p, row.expr, row.expect, why)) {
                     return false;
                   }
                 }
                 return true;
               }});
  t.push_back({"complex-opt: fresh parser defaults to off",
               [](std::string& why) {
                 MathParser p;
                 if (p.getSupportComplexNumbers()) {
                   why = "expected default false";
                   return false;
                 }
                 return true;
               }});
  t.push_back({"complex-opt: literals and arithmetic (parity with Basic RunComplexNumberSupportOptionTests)",
               [](std::string& why) {
                 MathParser p;
                 p.setSupportComplexNumbers(true);
                 static const struct {
                   const char* expr;
                   const char* expect;
                 } kRows[] = {
                     {"10+5i", "10+5i"},
                     {"-1+3i", "-1+3i"},
                     {"2-3*i", "2-3i"},
                     {"-i+5", "5-i"},
                     {"(1+2i)*(3+4i)", "-5+10i"},
                     {"10+5i-10-5i", "0"},
                     {"(1+2i)/i", "2-i"},
                     {"inf+i*inf", "inf+inf*i"},
                     {"inf+i*nan", "nan"},
                     {"(1+2i)/0", "nan"},
                 };
                 for (const auto& row : kRows) {
                   p.parseAndEvaluate(row.expr);
                   if (!p.getError().empty()) {
                     why = std::string(row.expr) + ": " + p.getError();
                     return false;
                   }
                   if (p.getResult() != row.expect) {
                     why = std::string(row.expr) + " -> " + p.getResult() + " (want " + row.expect + ")";
                     return false;
                   }
                 }
                 p.setSupportComplexNumbers(false);
                 p.parseAndEvaluate("10+5i");
                 if (p.getError().empty()) {
                   why = "expected evaluation failure for 10+5i with complex support off";
                   return false;
                 }
                 std::string el = p.getError();
                 std::transform(el.begin(), el.end(), el.begin(), [](unsigned char c) { return static_cast<char>(std::tolower(c)); });
                 const bool okUnknown = el.find("unknown variable") != std::string::npos && el.find("i") != std::string::npos;
                 const bool okUnexpected = el.find("unexpected token") != std::string::npos ||
                                           el.find("unexpected characters") != std::string::npos;
                 if (!okUnknown && !okUnexpected) {
                   why = std::string("unexpected error text: ") + p.getError();
                   return false;
                 }
                 return true;
               }});
  t.push_back({"complex-opt: imaginary suffix precedence (parity Basic cxSuffix*)",
               [](std::string& why) {
                 MathParser p;
                 p.setSupportComplexNumbers(true);
                 static const struct {
                   const char* expr;
                   const char* expect;
                 } kRows[] = {
                     {"cart(3i)", "3i"},
                     {"cart(0x10i)", "16i"},
                     {"cart(0b101i)", "5i"},
                     {"cart(0o7i)", "7i"},
                     {"cart(-3i)", "-3i"},
                     {"cart(2i)", "2i"},
                     {"cart(2*i)", "2i"},
                     {"cart((2**3)i)", "8i"},
                     {"cart((2*3)i)", "6i"},
                     {"cart((2/3)i)", "0.6666666666666666i"},
                     {"cart((2%3)i)", "2i"},
                     {"cart((1+2)i)", "3i"},
                     {"cart((1+2)*i)", "3i"},
                     {"cart(2*3i)", "6i"},
                     {"cart(2/3i)", "-0.6666666666666666i"},
                     {"cart(1/2i)", "-0.5i"},
                     {"cart(1/2*i)", "0.5i"},
                     {"cart(1/(1+2)i)", "-0.3333333333333333i"},
                     {"cart(1/(1+2)*i)", "0.3333333333333333i"},
                     {"cart(1/2/3i)", "-0.1666666666666666i"},
                     {"ratio(1/2i)", "-1/2*i"},
                     {"cart(-2**57/3**30+2**55/7**20*i)",
                      "-699.9582090286702+0.4515324440664775i"},
                     {"cart(2**3i)", "-0.4869944179657813+0.8734050817748715i"},
                     {"cart(7**20i)", "0.3444991159598805+0.9387866419495224i"},
                     {"cart(7**2i)", "-0.7315336785874539-0.681805307321898i"},
                 };
                 for (const auto& row : kRows) {
                   if (!expectEval(p, row.expr, row.expect, why)) {
                     return false;
                   }
                 }
                 p.parseAndEvaluate("3 i");
                 if (p.getError().empty()) {
                   why = "expected error for \"3 i\"";
                   return false;
                 }
                 std::string el = p.getError();
                 std::transform(el.begin(), el.end(), el.begin(),
                                [](unsigned char c) { return static_cast<char>(std::tolower(c)); });
                 if (el.find("unexpected") == std::string::npos) {
                   why = std::string("3 i: ") + p.getError();
                   return false;
                 }
                 p.parseAndEvaluate("2%3i");
                 if (p.getError().empty()) {
                   why = "expected error for \"2%3i\"";
                   return false;
                 }
                 std::string em = p.getError();
                 std::transform(em.begin(), em.end(), em.begin(),
                                [](unsigned char c) { return static_cast<char>(std::tolower(c)); });
                 if (em.find("incompatible operands") == std::string::npos &&
                     em.find("integer operands") == std::string::npos) {
                   why = std::string("2%3i: ") + p.getError();
                   return false;
                 }
                 return true;
               }});
  t.push_back({"complex-opt: complex component unaries on reals with support off (parity Basic cxRealOffOk)",
               [](std::string& why) {
                 MathParser p;
                 p.setSupportComplexNumbers(true);
                 p.parseAndEvaluate("1+1");
                 p.setSupportComplexNumbers(false);
                 static const struct {
                   const char* expr;
                   const char* expect;
                 } kRows[] = {
                     {"real(5)", "5"},
                     {"cart(polar(2))", "2"},
                     {"cart((5,0))", "5"},
                 };
                 for (const auto& row : kRows) {
                   p.parseAndEvaluate(row.expr);
                   if (!p.getError().empty()) {
                     why = std::string(row.expr) + ": " + p.getError();
                     return false;
                   }
                   if (p.getResult() != row.expect) {
                     why = std::string(row.expr) + " -> " + p.getResult() + " (want " + row.expect + ")";
                     return false;
                   }
                 }
                 p.parseAndEvaluate("cart((5,1))");
                 if (p.getError().empty()) {
                   why = "expected error for cart((5,1)) with complex support off";
                   return false;
                 }
                 std::string elCart = p.getError();
                 std::transform(elCart.begin(), elCart.end(), elCart.begin(),
                     [](unsigned char c) { return static_cast<char>(std::tolower(c)); });
                 if (elCart.find("incompatible operands") == std::string::npos) {
                   why = std::string("cart((5,1)): ") + p.getError();
                   return false;
                 }
                 return true;
               }});
  t.push_back({"complex-opt: comparisons (= == <> !=), ordering and complex-vs-time errors",
               [](std::string& why) {
                 MathParser p;
                 p.setSupportComplexNumbers(true);
                 static const struct {
                   const char* expr;
                   const char* expect;
                 } kOk[] = {
                     {"1+2i == 1+2i", "1"},
                     {"1+2i == 3+4i", "0"},
                     {"3 <> 1+2i", "1"},
                     {"1+0*i == 1", "1"},
                     {"(1+1i, 2) == (1+1i, 2)", "1"},
                     {"(1+1i, 2) <> (1+1i, 3)", "1"},
                     {"(1+1i, 2) == (1, 2)", "0"},
                     {"(1, 2+1i) != (1, 2+1i)", "0"},
                     {"2 == 1+1i", "0"},
                     {"(2, 0) == (1+1i, 0)", "0"},
                 };
                 for (const auto& row : kOk) {
                   p.parseAndEvaluate(row.expr);
                   if (!p.getError().empty()) {
                     why = std::string(row.expr) + ": " + p.getError();
                     return false;
                   }
                   if (p.getResult() != row.expect) {
                     why = std::string(row.expr) + " -> " + p.getResult() + " (want " + row.expect + ")";
                     return false;
                   }
                 }
                 static const char* kErr[] = {
                     "1+2i > 0",
                     "1+2i >= 1+2i",
                     "(1, 2+1i) < (1, 3)",
#if SMARTMATH_TIME_VALUES
                     "1+2i == 1s",
                    "1+2i + 1s",
                    "1+2i - 1s",
                    "1+2i*1s",
                     "1+2i <> 0:01",
#endif
                 };
                 for (const char* expr : kErr) {
                   p.parseAndEvaluate(expr);
                   if (p.getError().empty()) {
                     why = std::string("expected error for: ") + expr;
                     return false;
                   }
                   std::string el = p.getError();
                   std::transform(el.begin(), el.end(), el.begin(), [](unsigned char c) { return static_cast<char>(std::tolower(c)); });
                   if (el.find("incompatible operands") == std::string::npos) {
                     why = std::string(expr) + ": want incompatible operands, got: " + p.getError();
                     return false;
                   }
                 }
                 return true;
               }});
  t.push_back({"complex-opt: logical not, bitwise complement, mod and bitwise ops on complex",
               [](std::string& why) {
                 MathParser p;
                 p.setSupportComplexNumbers(true);
                 static const struct {
                   const char* expr;
                   const char* expect;
                 } kOk[] = {
                     {"!(5+5*i)", "0"},
                     {"!(-3*i)", "0"},
                     {"!(0+0*i)", "1"},
                     {"not (1+1*i)", "0"},
                     {"not (0)", "1"},
                 };
                 for (const auto& row : kOk) {
                   p.parseAndEvaluate(row.expr);
                   if (!p.getError().empty()) {
                     why = std::string(row.expr) + ": " + p.getError();
                     return false;
                   }
                   if (p.getResult() != row.expect) {
                     why = std::string(row.expr) + " -> " + p.getResult() + " (want " + row.expect + ")";
                     return false;
                   }
                 }
                 static const char* kErr[] = {
                     "~(1+2*i)",
                     "~(1,2+1*i)",
                     "(1+2*i) % 2",
                     "mod(1+2*i, 3)",
                     "(1+2*i) & 1",
                     "(1+2*i) | 1",
                     "(1+2*i) ^ 1",
                     "(1+2*i) << 1",
                     "(1+2*i) >> 1",
                 };
                 for (const char* expr : kErr) {
                   p.parseAndEvaluate(expr);
                   if (p.getError().empty()) {
                     why = std::string("expected error for: ") + expr;
                     return false;
                   }
                   std::string el = p.getError();
                   std::transform(el.begin(), el.end(), el.begin(), [](unsigned char c) { return static_cast<char>(std::tolower(c)); });
                   const bool okErr = el.find("incompatible operands") != std::string::npos ||
                                        el.find("modulo operands must be integer values") != std::string::npos ||
                                        el.find("bitwise operands must be integer values") != std::string::npos;
                   if (!okErr) {
                     why = std::string(expr) + ": want int/mod/bit error, got: " + p.getError();
                     return false;
                   }
                 }
                 return true;
               }});
  t.push_back({"complex-opt: mixed real/complex array broadcast (+ - * /)",
               [](std::string& why) {
                 MathParser p;
                 p.setSupportComplexNumbers(true);
                 static const struct {
                   const char* expr;
                   const char* expect;
                 } kRows[] = {
                     {"(1,1+2*i)+10", "(11,11+2i)"},
                     {"10+(1,1+2*i)", "(11,11+2i)"},
                     {"(5,1+2*i)-(1,i)", "(4,1+i)"},
                     {"(1,2)-i", "(1-i,2-i)"},
                     {"(2,1+2*i)*3", "(6,3+6i)"},
                     {"3*(2,1+2*i)", "(6,3+6i)"},
                     {"(3,4)*(1,i)", "(3,4i)"},
                     {"(1,2)*(1+i,0)", "(1+i,0)"},
                     {"(8,6+8*i)/2", "(4,3+4i)"},
                     {"(8,6)/(2,3)", "(4,2)"},
                     {"(4+2*i,6)/(2,3)", "(2+i,2)"},
                     {"i+(1,2)", "(1+i,2+i)"},
                 };
                 for (const auto& row : kRows) {
                   p.parseAndEvaluate(row.expr);
                   if (!p.getError().empty()) {
                     why = std::string(row.expr) + ": " + p.getError();
                     return false;
                   }
                   if (p.getResult() != row.expect) {
                     why = std::string(row.expr) + " -> " + p.getResult() + " (want " + row.expect + ")";
                     return false;
                   }
                 }
                 return true;
               }});
  t.push_back({"complex-opt: ** / pow / sqrt / sqr / hypot on complex (parity with Basic powCases)",
               [](std::string& why) {
                 MathParser p;
                 p.setSupportComplexNumbers(true);
                 static const struct {
                   const char* expr;
                   const char* expect;
                 } kRows[] = {
                     {"(1+i)**2", "2i"},
                     {"(3+4*i)**0.5", "2+i"},
                     {"pow(1+i,2)", "2i"},
                     {"(1+i,2)**(2,2)", "(2i,4)"},
                     {"sqr(1+2*i)", "-3+4i"},
                     {"sqrt(3+4*i)", "2+i"},
                     {"hypot(3*i,4)", "5"},
                     {"hypot(1+i,1-i)", "2"},
                     {"sqrt(-4)", "2i"},
                     {"(-2)**(1/2)", "1.414213562373095i"},
                     {"(-7)**(3/2)", "-18.52025917745212i"},
                     {"(-5)**(1/3)", "-1.709975946676696"},
                     {"64**(1/3)", "4"},
                     {"pow(-27,1/3)", "-3"},
                     {"pow(2**64,1/2)", "4294967296"},
                     {"(-8)**(1/3)", "-2"},
                    {"pow(-1-0i,1/2)", "i"},
                     {"sqrt(81)", "9"},
                     {"sqrt(3+4*i)", "2+i"},
                     {"pow(2**64,1/2)", "4294967296"},
                 };
                 for (const auto& row : kRows) {
                   p.parseAndEvaluate(row.expr);
                   if (!p.getError().empty()) {
                     why = std::string(row.expr) + ": " + p.getError();
                     return false;
                   }
                   if (p.getResult() != row.expect) {
                     why = std::string(row.expr) + " -> " + p.getResult() + " (want " + row.expect + ")";
                     return false;
                   }
                 }
                 return true;
               }});
  t.push_back({"complex-opt: exp / ln / log10 / log(value,base) (principal branch; parity with Basic)",
               [](std::string& why) {
                 MathParser p;
                 p.setSupportComplexNumbers(true);
                 static const struct {
                   const char* expr;
                   const char* expect;
                 } kRows[] = {
                     {"ln(-1)", "3.141592653589793i"},
                     {"ln(-2)", "0.6931471805599452+3.141592653589793i"},
                     {"log10(-10)", "1+1.364376353841841i"},
                     {"exp(i)", "0.5403023058681397+0.8414709848078963i"},
                     {"exp(pi*i)", "-1"},
                     {"ln((-1,-4))", "(3.141592653589793i,1.38629436111989+3.141592653589793i)"},
                     {"exp((0,pi*i))", "(1,-1)"},
                     {"log10((-10,-100))", "(1+1.364376353841841i,2+1.364376353841841i)"},
                     {"ln(2)", "0.6931471805599452"},
                     {"ln(0)", "-inf"},
                    {"ln(-1-0i)", "3.141592653589793i"},
                     {"log(8,2)", "3"},
                     {"log(-9,7)", "1.129150068107159+1.614459257080781i"},
                     {"log((-9,16),(7,2))", "(1.129150068107159+1.614459257080781i,4)"},
                     {"log(1+i,2)", "0.5+1.133090035456798i"},
                     {"log((8,16),(2,4))", "(3,2)"},
                 };
                 for (const auto& row : kRows) {
                   p.parseAndEvaluate(row.expr);
                   if (!p.getError().empty()) {
                     why = std::string(row.expr) + ": " + p.getError();
                     return false;
                   }
                   if (p.getResult() != row.expect) {
                     why = std::string(row.expr) + " -> " + p.getResult() + " (want " + row.expect + ")";
                     return false;
                   }
                 }
                 return true;
               }});
  t.push_back({"complex-opt: sum/prod/avg/reverse/unique/unpack on complex (parity with Basic cxAggOk)",
               [](std::string& why) {
                 MathParser p;
                 p.setSupportComplexNumbers(true);
                 static const struct {
                   const char* expr;
                   const char* expect;
                 } kRows[] = {
                     {"sum(1+2i, 3)", "4+2i"},
                     {"sum((1+2i, 3+4i))", "4+6i"},
                     {"prod(1+i, 2)", "2+2i"},
                     {"product((2, 1+i))", "2+2i"},
                     {"avg(2+2i, 4+4i)", "3+3i"},
                     {"mean((1+2i, 3))", "2+i"},
                     {"reverse((1+2i, 3))", "(3,1+2i)"},
                     {"unique((1+2i, 1+2i, 3))", "(1+2i,3)"},
                     {"unique((1+2i, 1+3i))", "(1+2i,1+3i)"},
                     {"unpack((1+2i, 3))", "(1+2i,3)"},
                     {"sum(unpack((1+2i, 2+2i)))", "3+4i"},
                     {"f(x,y)=x*y; f(unpack((1+i, 2)))", "2+2i"},
                     {"sum(1i, 2i)", "3i"},
                     {"prod(2+i, 2-i)", "5"},
                     {"product(1+i, 1+i)", "2i"},
                     {"mean(1+2i, 3+4i, 5+6i)", "3+4i"},
                     {"avg(0, 2i)", "i"},
                     {"sum(2+i, -2+i)", "2i"},
                     {"sum(1, Inf+2i, 3)", "inf+2i"},
                     {"prod(1, Inf+2i, 3)", "inf+6i"},
                     {"prod(5, Inf+2i, 3)", "inf+30i"},
                     {"avg(1, Inf+2i, 3)", "inf+0.6666666666666666i"},
                     {"reverse(1, Inf+2i, 3)", "(3,inf+2i,1)"},
                     {"unique(1, Inf+2i, 3)", "(1,inf+2i,3)"},
                     {"unpack(Inf+2i, 3)", "(inf+2i,3)"},
                     {"sum(1, 2+Inf*i, 3)", "6+inf*i"},
                     {"prod(1, 2+Inf*i, 3)", "6+inf*i"},
                     {"prod(5, 2+Inf*i, 3)", "30+inf*i"},
                     {"avg(1, 2+Inf*i, 3)", "2+inf*i"},
                     {"reverse(1, 2+Inf*i, 3)", "(3,2+inf*i,1)"},
                     {"unique(1, 2+Inf*i, 3)", "(1,2+inf*i,3)"},
                     {"unpack(2+Inf*i, 3)", "(2+inf*i,3)"},
                     {"repeat((1, Inf, i), 2)", "(1,inf,i,1,inf,i)"},
                 };
                 for (const auto& row : kRows) {
                   p.parseAndEvaluate(row.expr);
                   if (!p.getError().empty()) {
                     why = std::string(row.expr) + ": " + p.getError();
                     return false;
                   }
                   if (p.getResult() != row.expect) {
                     why = std::string(row.expr) + " -> " + p.getResult() + " (want " + row.expect + ")";
                     return false;
                   }
                 }
                 p.parseAndEvaluate("repeat((1,2), i)");
                 if (p.getError().empty()) {
                   why = "repeat((1,2), i): expected error";
                   return false;
                 }
                 if (p.getError().find("repeat() expects integer values") == std::string::npos) {
                   why = std::string("repeat((1,2), i): ") + p.getError();
                   return false;
                 }
                 p.parseAndEvaluate("range(1, 5, i)");
                 if (p.getError().empty()) {
                   why = "range(1, 5, i): expected error";
                   return false;
                 }
                 if (p.getError().find("range() expects integer values") == std::string::npos) {
                   why = std::string("range(1, 5, i): ") + p.getError();
                   return false;
                 }
                 return true;
               }});
  t.push_back({"complex-opt: sortby by abs and ratio on complex (parity Basic 1046+)",
               [](std::string& why) {
                 MathParser p;
                 p.setSupportComplexNumbers(true);
                 static const struct {
                   const char* expr;
                   const char* expect;
                 } kRows[] = {
                     {"sortby((3+4i, 1+2i), abs)", "(1+2i,3+4i)"},
                     {"sortby((3+4i, 1+2i), polar)", "(1+2i,3+4i)"},
                     {"ratio(1+2i)", "1+2i"},
                     {"ratio(0.5+0.25i)", "1/2+1/4*i"},
                     {"ratio(0.25i)", "1/4*i"},
                     {"ratio(-0.5+0.25i)", "-1/2+1/4*i"},
                     {"ratio(3+0.25i)", "3+1/4*i"},
                     {"ratio(0.5+0.25i)+1", "1.5+0.25i"},
                     {"ratio(2+3i)", "2+3i"},
                     {"ratio(e+10i)", "14665106/5394991+10i"},
                     {"ratio(-2**57/3**30+2**55/7**20*i)", "-113792906/162571+1293463/2864607*i"},
                     {"ratio(-2**57/3**30+2**55/7**20i)", "12411888722130364-33823353366914148i"},
                     {"ratio((1+2i)/(3+4i))", "11/25+2/25*i"},
                     {"ratio((1.11+2.33i)*(3.37+4.71i))", "-4521/625+65401/5000*i"},
                     {"ratio((1+i)**(2+3i))", "-740749/4531935+329494/3432051*i"},
                     {"ratio(1/(1+i))", "1/2-1/2*i"},
                     {"ratio((1+3i)/2)", "1/2+3/2*i"},
                     {"ratio((1+2i)/2)", "1/2+i"},
                     {"ratio(2/(1+3i))", "1/5-3/5*i"},
                     {"ratio((2+3i)**2)", "-5+12i"},
                     {"ratio((1+i)/(1-i))", "i"},
                     {"ratio(3i/2)", "3/2*i"},
                     {"ratio(1/2+3i/4)", "1/2+3/4*i"},
                     {"3+ratio(0.2)i", "3+0.2i"}
                 };
                 for (const auto& row : kRows) {
                   p.parseAndEvaluate(row.expr);
                   if (!p.getError().empty()) {
                     why = std::string(row.expr) + ": " + p.getError();
                     return false;
                   }
                   if (p.getResult() != row.expect) {
                     why = std::string(row.expr) + " -> " + p.getResult() + " (want " + row.expect + ")";
                     return false;
                   }
                 }
                 return true;
               }});
  t.push_back({"complex-opt: getRawResult exposes complex and rational components",
               [](std::string& why) {
                 MathParser p;
                 p.setSupportComplexNumbers(true);

                 p.parseAndEvaluate("10+5i");
                 if (!p.getError().empty()) {
                   why = p.getError();
                   return false;
                 }
                 MathParser::RawResult r = p.getRawResult();
                 if (!r.isScalar() || !r.scalar.isComplex() || !r.scalar.real.isInt64() || r.scalar.real.intValue != 10LL ||
                     !r.scalar.imag.isInt64() || r.scalar.imag.intValue != 5LL) {
                   why = "expected complex int64 10+5i";
                   return false;
                 }

                 p.parseAndEvaluate("ratio(0.5+0.25i)");
                 if (!p.getError().empty()) {
                   why = p.getError();
                   return false;
                 }
                 r = p.getRawResult();
                 if (!r.isScalar() || !r.scalar.isComplex() || !r.scalar.real.isRational() ||
                     r.scalar.real.rational.numerator != 1LL || r.scalar.real.rational.denominator != 2ULL ||
                     !r.scalar.imag.isRational() || r.scalar.imag.rational.numerator != 1LL ||
                     r.scalar.imag.rational.denominator != 4ULL) {
                   why = "expected complex rational 1/2+1/4i";
                   return false;
                 }

                 p.parseAndEvaluate("ratio(e+10i)");
                 if (!p.getError().empty()) {
                   why = p.getError();
                   return false;
                 }
                 r = p.getRawResult();
                 if (!r.isScalar() || !r.scalar.isComplex() || !r.scalar.real.isRational() ||
                     r.scalar.real.rational.numerator != 14665106LL || r.scalar.real.rational.denominator != 5394991ULL) {
                   why = "expected complex rational real part for e";
                   return false;
                 }
                 const bool imagTenInt =
                     r.scalar.imag.isInt64() && r.scalar.imag.intValue == 10LL;
                 const bool imagTenRat = r.scalar.imag.isRational() && r.scalar.imag.rational.numerator == 10LL &&
                                         r.scalar.imag.rational.denominator == 1ULL;
                 if (!imagTenInt && !imagTenRat) {
                   why = "expected imag 10 as int64 or rational 10/1";
                   return false;
                 }

                 p.setSupportComplexNumbers(false);
                 return true;
               }});
  t.push_back({"complex-opt: min/max/sort/median/variance/stddev reject complex (parity with Basic cxAggErr)",
               [](std::string& why) {
                 MathParser p;
                 p.setSupportComplexNumbers(true);
                 static const char* kExprs[] = {
                     "min(1+2i, 3)",
                     "max((1+2i, 2))",
                     "sort(3+4i, 1+2i)",
                     "median(1+2i, 5)",
                     "variance(1+2i, 2)",
                     "stddev((1+2i, 2))",
                     "min(1, Inf+2i, 3)",
                     "max(1, 2+Inf*i, 3)",
                     "sort(Inf+2i, 1+2i)",
                     "median(Inf+2i, 1)",
                     "variance(2+Inf*i, 1)",
                     "stddev((Inf+2i, 1))",
                 };
                 for (const char* expr : kExprs) {
                   p.parseAndEvaluate(expr);
                   if (p.getError().empty()) {
                     why = std::string(expr) + " expected error, got " + p.getResult();
                     return false;
                   }
                   std::string el = p.getError();
                   std::transform(el.begin(), el.end(), el.begin(),
                       [](unsigned char c) { return static_cast<char>(std::tolower(c)); });
                   if (el.find("incompatible operands") == std::string::npos) {
                     why = std::string(expr) + ": " + p.getError();
                     return false;
                   }
                 }
                 return true;
               }});
  t.push_back({"complex-opt: unary rounding/abs/sign/fact and complex-only funcs (parity Basic cxUniOk)",
               [](std::string& why) {
                 MathParser p;
                 p.setSupportComplexNumbers(true);
                 static const struct {
                   const char* expr;
                   const char* expect;
                 } kRows[] = {
                     {"int(2.7+3.2i)", "2+3i"},
                     {"frac(2.5+0.5i)", "0.5+0.5i"},
                     {"round(2.4+3.6i)", "2+4i"},
                     {"floor(2.9+3.1i)", "2+3i"},
                     {"ceil(2.1+3.1i)", "3+4i"},
                     {"abs(3+4i)", "5"},
                     {"sign(3+4i)", "0.5999999999999999+0.8i"},
                     {"real(1+2i)", "1"},
                     {"imag(1+2i)", "2"},
                     {"conj(1+2i)", "1-2i"},
                     {"phase(1)", "0"},
                     {"polar(1+1i)", "(1.414213562373095,0.7853981633974482)"},
                     {"cart(polar(1+1i))", "1+i"},
                    {"phase(-1-0i)", "3.141592653589793"},
                    {"polar(-1-0i)", "(1,3.141592653589793)"},
                     {"fact(5)", "120"},
                     {"int((1+2.7i,4.2+5.8i))", "(1+2i,4+5i)"},
                     {"abs((3+4i,0))", "(5,0)"},
                 };
                 for (const auto& row : kRows) {
                   p.parseAndEvaluate(row.expr);
                   if (!p.getError().empty()) {
                     why = std::string(row.expr) + ": " + p.getError();
                     return false;
                   }
                   if (p.getResult() != row.expect) {
                     why = std::string(row.expr) + " -> " + p.getResult() + " (want " + row.expect + ")";
                     return false;
                   }
                 }
                 return true;
               }});
  t.push_back({"complex-opt: polar/cart arity and shape (parity Basic cxPolCartOk/cxPolCartErr)",
               [](std::string& why) {
                 MathParser p;
                 p.setSupportComplexNumbers(true);
                 static const struct {
                   const char* expr;
                   const char* expect;
                 } kOk[] = {
                     {"polar(3)", "(3,0)"},
                     {"polar(4i)", "(4,1.570796326794897)"},
                     {"polar(3+4i)", "(5,0.9272952180016122)"},
                     {"polar((3+4i))", "(5,0.9272952180016122)"},
                     {"cart(2)", "2"},
                     {"cart((2))", "2"},
                     {"cart((2,pi/3))", "1+1.732050807568877i"},
                     {"cart(2,pi/3)", "1+1.732050807568877i"},
                 };
                 for (const auto& row : kOk) {
                   p.parseAndEvaluate(row.expr);
                   if (!p.getError().empty()) {
                     why = std::string(row.expr) + ": " + p.getError();
                     return false;
                   }
                   if (!smokeResultCloseEnough(p.getResult(), row.expect)) {
                     why = std::string(row.expr) + " -> " + p.getResult() + " (want " + row.expect + ")";
                     return false;
                   }
                 }
                 static const struct {
                   const char* expr;
                   const char* errSubstr;
                 } kErr[] = {
                     {"polar(3+4i, 2)", "expects"},
                     {"polar((3+4i, 2))", "incompatible operands"},
                     {"cart((2,pi/3,1))", "incompatible operands"},
                     {"cart(2,pi/3,1)", "expects"},
                     {"cart((2,pi/3),(1))", "incompatible operands"},
                     {"polar(3+4i, 2i)", "expects"},
                     {"polar(3+4i, 1+2i)", "expects"},
                 };
                 for (const auto& row : kErr) {
                   p.parseAndEvaluate(row.expr);
                   if (p.getError().empty()) {
                     why = std::string(row.expr) + " expected error, got " + p.getResult();
                     return false;
                   }
                   const std::string el = p.getError();
                   if (el.find(row.errSubstr) == std::string::npos) {
                     why = std::string(row.expr) + ": " + el;
                     return false;
                   }
                 }
                 return true;
               }});
  t.push_back({"complex-opt: trig/atan2 auxiliary exact float parts (parity Basic cxTrigAuxOk)",
               [](std::string& why) {
                 MathParser p;
                 p.setSupportComplexNumbers(true);
                 static const struct {
                   const char* expr;
                   const char* expect;
                 } kRows[] = {
                     {"cart(6,0)", "6"},
                     {"cart(5,pi/2)", "5i"},
                     {"cart(2,pi)", "-2"},
                     {"cart(3,3*pi/2)", "-3i"},
                     {"real(cart(8,pi/2))", "0"},
                     {"imag(cart(8,pi/2))", "8"},
                     {"real(cart(7,0))", "7"},
                     {"imag(cart(7,0))", "0"},
                     {"real(cart(5,pi))", "-5"},
                     {"imag(cart(5,3*pi/2))", "-5"},
                     {"polar(9)", "(9,0)"},
                     {"polar(9i)", "(9,1.570796326794897)"},
                     {"polar(-9)", "(9,3.141592653589793)"},
                     {"polar(-9i)", "(9,-1.570796326794897)"},
                     {"phase(9)", "0"},
                     {"phase(9i)", "1.570796326794897"},
                     {"phase(-9)", "3.141592653589793"},
                     {"phase(-9i)", "-1.570796326794897"},
                     {"phase(0)", "0"},
                     {"sin(pi/2)", "1"},
                     {"cos(pi/2)", "0"},
                     {"tan(pi/2)", "inf"},
                     {"tan(-3*pi/2)", "-inf"},
                     {"sin(pi)", "0"},
                     {"tan(pi/4)", "1"},
                     {"tan(-pi/4)", "-1"},
                     {"real(sin(pi/2))", "1"},
                     {"imag(sin(pi/2))", "0"},
                     {"sinh(i*pi/2)", "i"},
                     {"cosh(i*pi/2)", "0"},
                     {"cosh(i*pi)", "-1"},
                     {"exp(pi/2*i)", "i"},
                     {"exp(2*pi*i)", "1"},
                     {"sin(3*pi/2)", "-1"},
                     {"real(ln(-1))", "0"},
                     {"imag(ln(-1))", "3.141592653589793"},
                     {"imag(ln(-i))", "-1.570796326794897"},
                     {"cart(polar(6i))", "6i"},
                     {"cart(polar(-4))", "-4"},
                 };
                 for (const auto& row : kRows) {
                   p.parseAndEvaluate(row.expr);
                   if (!p.getError().empty()) {
                     why = std::string(row.expr) + ": " + p.getError();
                     return false;
                   }
                   if (!smokeResultCloseEnough(p.getResult(), row.expect)) {
                     why = std::string(row.expr) + " -> " + p.getResult() + " (want " + row.expect + ")";
                     return false;
                   }
                 }
                 return true;
               }});
  t.push_back({"complex-opt: sin/cos/tan and inverses on complex (parity Basic cxTrigOk)",
               [](std::string& why) {
                 MathParser p;
                 p.setSupportComplexNumbers(true);
                 static const struct {
                   const char* expr;
                   const char* expect;
                 } kRows[] = {
                     {"sin(i)", "1.175201193643801i"},
                     {"cos(i)", "1.543080634815243"},
                     {"sin(1+i)", "1.298457581415977+0.634963914784736i"},
                     {"tan(i)", "0.7615941559557648i"},
                     {"sinh(i)", "0.8414709848078963i"},
                     {"cosh(1+i)", "0.8337300251311491+0.988897705762865i"},
                     {"asinh(1+i)", "1.061275061905035+0.6662394324925152i"},
                     {"atan(1+i)", "1.017221967897851+0.402359478108525i"},
                     {"sin((1+i,2))", "(1.298457581415977+0.634963914784736i,0.9092974268256815)"},
                     {"atanh(i)", "0.7853981633974482i"},
                     {"acos(1)", "0"},
                     {"asin(i)", "0.8813735870195428i"},
                 };
                 for (const auto& row : kRows) {
                   p.parseAndEvaluate(row.expr);
                   if (!p.getError().empty()) {
                     why = std::string(row.expr) + ": " + p.getError();
                     return false;
                   }
                   if (p.getResult() != row.expect) {
                     why = std::string(row.expr) + " -> " + p.getResult() + " (want " + row.expect + ")";
                     return false;
                   }
                 }
                 return true;
               }});
  t.push_back({"complex-opt: atan2/deg/rad reject complex (parity Basic cxTrigErr)",
               [](std::string& why) {
                 MathParser p;
                 p.setSupportComplexNumbers(true);
                 static const char* kExprs[] = {"atan2(1+2i,1)", "deg(1+2i)", "rad(1+i)"};
                 for (const char* expr : kExprs) {
                   p.parseAndEvaluate(expr);
                   if (p.getError().empty()) {
                     why = std::string(expr) + " expected error, got " + p.getResult();
                     return false;
                   }
                   std::string el = p.getError();
                   std::transform(el.begin(), el.end(), el.begin(), [](unsigned char c) { return static_cast<char>(std::tolower(c)); });
                   if (el.find("incompatible operands") == std::string::npos) {
                     why = std::string(expr) + ": " + p.getError();
                     return false;
                   }
                 }
                 return true;
               }});
  t.push_back({"complex-opt: hex/oct/bin/u* on integer complex literals (parity Basic cxFmtOk)",
               [](std::string& why) {
                 MathParser p;
                 p.setSupportComplexNumbers(true);
                 static const struct {
                   const char* expr;
                   const char* expect;
                 } kRows[] = {
                     {"hex(10+15i)", "0xA+0xFi"},
                     {"bin(1+i)", "0b1+i"},
                     {"oct(8i)", "0o10i"},
                     {"hex((1+2i,10+11i))", "(0x1+0x2i,0xA+0xBi)"},
                     {"uhex(-1+2i)", "0xFFFFFFFFFFFFFFFF+0x2i"},
                     {"uoct(-1+i)", "0o1777777777777777777777+i"},
                     {"ubin(5+10i)", "0b101+0b1010i"},
                 };
                 for (const auto& row : kRows) {
                   p.parseAndEvaluate(row.expr);
                   if (!p.getError().empty()) {
                     why = std::string(row.expr) + ": " + p.getError();
                     return false;
                   }
                   if (p.getResult() != row.expect) {
                     why = std::string(row.expr) + " -> " + p.getResult() + " (want " + row.expect + ")";
                     return false;
                   }
                 }
                 p.parseAndEvaluate("(0xFFFFFFFFFFFFFF + 0x7FFFFFFFFFFFFFi);uhex");
                 if (!p.getError().empty()) {
                   why = std::string("large hex complex uhex: ") + p.getError();
                   return false;
                 }
                 if (p.getResult() != "0xFFFFFFFFFFFFFF+0x7FFFFFFFFFFFFFi") {
                   why = std::string("large hex complex uhex -> ") + p.getResult();
                   return false;
                 }
                 p.parseAndEvaluate("(0xFFFFFFFFFFFFFF + 0x7FFFFFFFFFFFFFi)");
                 if (!p.getError().empty()) {
                   why = std::string("large hex complex sum: ") + p.getError();
                   return false;
                 }
                 if (p.getResult() != "72057594037927935+36028797018963967i") {
                   why = std::string("large hex complex sum -> ") + p.getResult();
                   return false;
                 }
                 p.parseAndEvaluate("sum(0xFFFFFFFFFFFFFF, 0x7FFFFFFFFFFFFFi)");
                 if (!p.getError().empty()) {
                   why = std::string("sum exact complex scalars: ") + p.getError();
                   return false;
                 }
                 if (p.getResult() != "72057594037927935+36028797018963967i") {
                   why = std::string("sum exact complex scalars -> ") + p.getResult();
                   return false;
                 }
                 p.parseAndEvaluate("conj(0xFFFFFFFFFFFFFF+0x7FFFFFFFFFFFFFi);uhex");
                 if (!p.getError().empty()) {
                   why = std::string("conj exact uhex: ") + p.getError();
                   return false;
                 }
                 if (p.getResult() != "0xFFFFFFFFFFFFFF-0x7FFFFFFFFFFFFFi") {
                   why = std::string("conj exact uhex -> ") + p.getResult();
                   return false;
                 }
                 p.parseAndEvaluate("hex(1+2.5i)");
                 if (p.getError().empty()) {
                   why = "hex(1+2.5i): expected error";
                   return false;
                 }
                 std::string eh = p.getError();
                 std::transform(eh.begin(), eh.end(), eh.begin(), [](unsigned char c) { return static_cast<char>(std::tolower(c)); });
                 if (eh.find("expects integer values") == std::string::npos) {
                   why = std::string("bad error: ") + p.getError();
                   return false;
                 }
                 return true;
               }});
  t.push_back({"complex-opt: clamp/gcd/mod/ncr/lcm reject complex (parity Basic cxUniErr)",
               [](std::string& why) {
                 MathParser p;
                 p.setSupportComplexNumbers(true);
                 static const char* kExprs[] = {
                     "clamp(1+2i, 0, 5)",
                     "gcd(1+2i, 2)",
                     "mod(1+2i, 2)",
                     "ncr(5+5i, 2)",
                     "lcm(2, 1+2i)",
                 };
                 for (const char* expr : kExprs) {
                   p.parseAndEvaluate(expr);
                   if (p.getError().empty()) {
                     why = std::string(expr) + " expected error, got " + p.getResult();
                     return false;
                   }
                   std::string el = p.getError();
                   std::transform(el.begin(), el.end(), el.begin(),
                       [](unsigned char c) { return static_cast<char>(std::tolower(c)); });
                   if (el.find("incompatible operands") == std::string::npos) {
                     why = std::string(expr) + ": " + p.getError();
                     return false;
                   }
                 }
                 return true;
               }});
  t.push_back({"complex-opt: coverage extension (Python reference, C3/C4/C5/lexer)",
               [](std::string& why) {
                 MathParser p;
                 p.setSupportComplexNumbers(true);
                 const std::size_t count = sizeof(kComplexCoverageRows) / sizeof(kComplexCoverageRows[0]);
                 return runNegBandRowBatch(p, kComplexCoverageRows, count, why);
               }});
  return t;
}

void runSuite(const std::string& title, const std::vector<TestCase>& cases, TestState& s) {
  std::cout << "=== " << title << " ===\n";
  std::unordered_set<std::string> seen;
  g_seen = &seen;
  g_duplicateCount = 0;
  clearParserConstDedupState();

  for (const auto& tc : cases) {
    g_skipDuplicate = false;
    std::string why;
    bool ok = false;
    try {
      ok = tc.run(why);
    } catch (const std::exception& ex) {
      ok = false;
      why = std::string("exception: ") + ex.what();
    } catch (...) {
      ok = false;
      why = "unknown exception";
    }

    if (g_skipDuplicate) {
      std::cout << "[DUPLICATE] " << tc.name << "\n";
      continue;  // duplicates are skipped (no count, no print).
    }

    ++s.total;
    if (ok) {
      ++s.passed;
      std::cout << "[PASS] " << tc.name << "\n";
    } else {
      ++s.failed;
      std::cout << "[FAIL] " << tc.name << " -> " << why << "\n";
    }
  }
  std::cout << "Detected duplicate cases: " << g_duplicateCount << "\n";
  std::cout << "\n";
  g_seen = nullptr;
  clearParserConstDedupState();
}

}  // namespace

int main(int argc, char** argv) {
  if (argc >= 3 && std::string(argv[1]) == "--eval") {
    MathParser p;
    const char* expr = argv[2];
    if (argc >= 4 && std::string(argv[2]) == "--complex") {
      p.setSupportComplexNumbers(true);
      expr = argv[3];
    }
    p.parseAndEvaluate(expr);
    if (!p.getError().empty()) {
      std::cout << "ERROR:" << p.getError() << "\n";
      return 1;
    }
    std::cout << p.getResult() << "\n";
    return 0;
  }

  TestState s;

  const auto unit = buildUnitCases();
  const auto edgeIntFloat = buildEdgeIntFloatCases();
  const auto nanInf = buildNanInfCases();
  const auto regression = buildRegressionCases();
  const auto parityFromSmoke = buildParityBasicFromSmokeCases();
  const auto parityTimeFromSmoke = buildParityTimeFromSmokeCases();
  const auto complexNumberSupportOption = buildComplexNumberSupportOptionCases();
  const auto trigAngleReduction = buildTrigAngleReductionCases();
  const auto operatorPrecedenceDoc = buildOperatorPrecedenceDocCases();
  const auto negativeArgumentMagnitudeBand = buildNegativeArgumentMagnitudeBandCases();
  const auto positiveArgumentMagnitudeBand = buildPositiveArgumentMagnitudeBandCases();
  const auto floatMagnitudeLiteral = buildFloatMagnitudeLiteralCases();
  const auto timeValuesSupportOption = buildTimeValuesSupportOptionCases();
  const auto lambdaFunctionsSupportOption = buildLambdaFunctionsSupportOptionCases();
  const auto sortbyRatio = buildSortbyRatioCases();
  const auto ratioInExpression = buildRatioInExpressionCases();
  const auto minMaxPreserveWinner = buildMinMaxPreserveWinnerCases();
  const auto clampPreserveIntegerMetadata = buildClampPreserveIntegerMetadataCases();
  const auto clampInvertedBounds = buildClampInvertedBoundsCases();
  const auto clampNanBound = buildClampNanBoundCases();

  runSuite("Unit", unit, s);
  runSuite("Edge/int-float", edgeIntFloat, s);
  runSuite("NaN/Inf", nanInf, s);
  runSuite("Regression", regression, s);
  runSuite("Parity/SmokeTest_MathParser (from Basic)", parityFromSmoke, s);
#if SMARTMATH_TIME_VALUES
  runSuite("Parity/Time (Smoke alignment)", parityTimeFromSmoke, s);
#endif
#if SMARTMATH_COMPLEX_NUMBERS
  runSuite("Complex number support (parser option)", complexNumberSupportOption, s);
#endif
  runSuite("Trig angle reduction (large integers)", trigAngleReduction, s);
  runSuite("Operator precedence (USAGE_AND_SYNTAX.md)", operatorPrecedenceDoc, s);
  runSuite("Negative argument magnitude bands", negativeArgumentMagnitudeBand, s);
  runSuite("Positive argument magnitude bands", positiveArgumentMagnitudeBand, s);
  runSuite("Float magnitude literals", floatMagnitudeLiteral, s);
#if SMARTMATH_TIME_VALUES
  runSuite("Time value support (parser option)", timeValuesSupportOption, s);
#endif
#if SMARTMATH_LAMBDA_FUNCTIONS
  runSuite("Lambda function support (parser option)", lambdaFunctionsSupportOption, s);
#endif
  runSuite("Sortby/Ratio", sortbyRatio, s);
  runSuite("Ratio in expressions", ratioInExpression, s);
  runSuite("Min/max preserve winner", minMaxPreserveWinner, s);
  runSuite("Clamp preserve integer metadata", clampPreserveIntegerMetadata, s);
  runSuite("Clamp inverted bounds", clampInvertedBounds, s);
  runSuite("Clamp NaN bounds", clampNanBound, s);

  std::cout << "TOTAL: " << s.total << ", PASSED: " << s.passed << ", FAILED: " << s.failed << "\n";
  return (s.failed == 0) ? 0 : 1;
}

