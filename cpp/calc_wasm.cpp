// WebAssembly calculator for the web page in tools/web/. Built by tools/build_web_calc.py with Emscripten.
#include "MathParser.hpp"

#include <emscripten/bind.h>

#include <cctype>
#include <memory>
#include <string>

namespace {

const bool kEnableTimeValues = true;
const bool kEnableComplexNumbers = true;
const bool kEnableLambdaFunctions = true;

std::string trimCopy(const std::string& text) {
  std::size_t begin = 0;
  while (begin < text.size() && std::isspace(static_cast<unsigned char>(text[begin])) != 0) {
    ++begin;
  }
  std::size_t end = text.size();
  while (end > begin && std::isspace(static_cast<unsigned char>(text[end - 1])) != 0) {
    --end;
  }
  return text.substr(begin, end - begin);
}

bool startsWith(const std::string& text, const char* prefix) {
  const std::size_t prefixSize = std::char_traits<char>::length(prefix);
  return text.compare(0, prefixSize, prefix) == 0;
}

bool isWholeLineComment(const std::string& trimmed) {
  return startsWith(trimmed, "#") || startsWith(trimmed, "//");
}

enum class LineStatus { Skipped, Result, Error };

struct LineOutput {
  LineStatus status;
  std::string text;  // result or error text
};

// One parser for the whole session, so variables, user functions, and ans carry over.
class WebCalc {
 public:
  WebCalc() { reset(); }

  void reset() {
    parser_.reset(new MathParser());
    parser_->setSupportTimeValues(kEnableTimeValues);
    parser_->setSupportComplexNumbers(kEnableComplexNumbers);
    parser_->setSupportLambdaFunctions(kEnableLambdaFunctions);
  }

  // Blank lines and whole-line # or // comments are skipped, as in calc.
  LineOutput evaluate(const std::string& line) {
    LineOutput out{LineStatus::Skipped, std::string()};
    const std::string trimmed = trimCopy(line);
    if (trimmed.empty() || isWholeLineComment(trimmed)) {
      return out;
    }
    parser_->parseAndEvaluate(trimmed);
    const std::string error = parser_->getError();
    if (error.empty()) {
      out.status = LineStatus::Result;
      out.text = parser_->getResult();
    } else {
      out.status = LineStatus::Error;
      out.text = error;
    }
    return out;
  }

 private:
  std::unique_ptr<MathParser> parser_;
};

}  // namespace

EMSCRIPTEN_BINDINGS(smartmath) {
  emscripten::enum_<LineStatus>("LineStatus")
      .value("Skipped", LineStatus::Skipped)
      .value("Result", LineStatus::Result)
      .value("Error", LineStatus::Error);
  emscripten::value_object<LineOutput>("LineOutput")
      .field("status", &LineOutput::status)
      .field("text", &LineOutput::text);
  emscripten::class_<WebCalc>("WebCalc")
      .constructor<>()
      .function("reset", &WebCalc::reset)
      .function("evaluate", &WebCalc::evaluate);
}
