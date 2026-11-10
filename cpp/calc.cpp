// Command-line calculator on the C++ MathParser. Keep in sync with tools/calc.bas.
#include "MathParser.hpp"

#include <cctype>
#include <cstdio>
#include <fstream>
#include <iostream>
#include <string>
#include <vector>

#if defined(_WIN32)
#include <io.h>
#else
#include <unistd.h>
#endif

namespace {

#if SMARTMATH_TIME_VALUES
const bool kEnableTimeValues = true;
#endif

#if SMARTMATH_COMPLEX_NUMBERS
const bool kEnableComplexNumbers = true;
#endif

#if SMARTMATH_LAMBDA_FUNCTIONS
const bool kEnableLambdaFunctions = true;
#endif

const bool kShowErrors = true;

typedef unsigned LineFlags;
const LineFlags kLinePrompt = 1u << 0;      // print "> " before reading each line
const LineFlags kLineStopOnQuit = 1u << 1;  // "quit" / "exit" ends the input

const char kFileOptionPrefix[] = "--file=";
const char kStdinFilePath[] = "-";
const char kStdinLabel[] = "<stdin>";

enum class StepKind { Expression, File };

struct Step {
  StepKind kind;
  std::string text;  // expression or file path
};

struct ProgramArgs {
  bool help = false;
  bool interactive = false;
  std::string filePath;
  std::vector<Step> steps;  // in command-line order
  std::string error;
};

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

bool isQuitCommand(std::string line) {
  for (char& c : line) {
    c = static_cast<char>(std::tolower(static_cast<unsigned char>(c)));
  }
  return line == "quit" || line == "exit";
}

bool isWholeLineComment(const std::string& trimmed) {
  return startsWith(trimmed, "#") || startsWith(trimmed, "//");
}

void stripLeadingUtf8Bom(std::string& text) {
  if (startsWith(text, "\xEF\xBB\xBF")) {
    text.erase(0, 3);
  }
}

// Rebinds stdin (and std::cin, which is synced with stdio) to the terminal.
bool reopenStdinFromTerminal() {
#if defined(_WIN32)
  FILE* reopened = nullptr;
  if (freopen_s(&reopened, "CONIN$", "r", stdin) != 0) {
    return false;
  }
#else
  if (std::freopen("/dev/tty", "r", stdin) == nullptr) {
    return false;
  }
#endif
  std::cin.clear();
  return true;
}

bool stdinIsConsole() {
#if defined(_WIN32)
  return _isatty(_fileno(stdin)) != 0;
#else
  return isatty(fileno(stdin)) != 0;
#endif
}

void printUsage(std::ostream& out) {
  out << "Usage: calc [expression] [-f <file>] [expression] [-i]\n"
      << "\n"
      << "Evaluates expressions and a file (one per line) in argument order on one\n"
      << "parser, so variables and ans carry over. With neither, reads standard\n"
      << "input (interactively in a terminal).\n"
      << "\n"
      << "  -f, --file <file>  Evaluate <file>. Use \"-f -\" to read standard input.\n"
      << "      --file=<file>\n"
      << "  -i, --interactive  When done, keep reading standard input (or the terminal,\n"
      << "                     if \"-f -\" used it up). Needs an expression or a file.\n"
      << "  -h, --help         Show this help (-h only as the first argument).\n"
      << "  --                 Treat the remaining arguments as expression text.\n"
      << "\n"
      << "Adjacent arguments form one expression; -f or -i ends it. Only the first\n"
      << "-i and -f are options; later ones are expression text. Blank lines and\n"
      << "whole-line # or // comments are skipped. quit or exit ends standard input\n"
      << "but is an expression in a file.\n"
      << "\n"
      << "  calc -f setup.txt x*2      file, then x*2\n"
      << "  calc x=5 -f test.txt       x=5, then file\n"
      << "  echo a=10 | calc -f - a+5  a=10 from standard input, then a+5\n"
      << "  calc -i 2+3 -i             2+3-i (imaginary unit), then interactive\n"
      << "  calc -- -i                 -i\n";
}

// "--name" cannot be math, unlike "-x", "--3" or "--(x)".
bool isLongOptionLike(const std::string& arg) {
  return arg.size() > 2 && startsWith(arg, "--") && std::isalpha(static_cast<unsigned char>(arg[2])) != 0;
}

ProgramArgs parseArgs(int argc, char** argv) {
  ProgramArgs args;
  if (argc > 1 && std::string(argv[1]) == "-h") {
    args.help = true;
    return args;
  }

  bool optionsOpen = true;
  bool expressionOpen = false;
  for (int i = 1; i < argc; ++i) {
    const std::string arg = argv[i];
    if (optionsOpen) {
      if (arg == "--") {
        optionsOpen = false;
        continue;
      }
      if (arg == "--help") {
        args.help = true;
        return args;
      }
      if (!args.interactive && (arg == "-i" || arg == "--interactive")) {
        args.interactive = true;
        expressionOpen = false;
        continue;
      }
      const bool inlineFile = startsWith(arg, kFileOptionPrefix);
      if (args.filePath.empty() && (inlineFile || arg == "-f" || arg == "--file")) {
        if (inlineFile) {
          args.filePath = arg.substr(sizeof(kFileOptionPrefix) - 1);
        } else if (i + 1 < argc) {
          args.filePath = argv[++i];
        } else {
          args.error = "missing file path after " + arg;
          return args;
        }
        if (args.filePath.empty()) {
          args.error = "missing file path";
          return args;
        }
        args.steps.push_back(Step{StepKind::File, args.filePath});
        expressionOpen = false;
        continue;
      }
      if (isLongOptionLike(arg)) {
        args.error = "unknown option: " + arg;
        return args;
      }
    }

    if (expressionOpen) {
      args.steps.back().text += ' ';
      args.steps.back().text += arg;
    } else {
      args.steps.push_back(Step{StepKind::Expression, arg});
      expressionOpen = true;
    }
  }

  if (args.interactive && args.steps.empty()) {
    args.error = "-i/--interactive requires an expression or -f/--file"
                 " (to evaluate -i, use: calc -- -i)";
  }
  return args;
}

// Prints the result or the error. A non-empty errorLabel prefixes errors as "label:line: ".
bool evaluateAndPrint(MathParser& parser,
                      const std::string& expression,
                      const std::string& errorLabel = std::string(),
                      std::size_t lineNumber = 0) {
  parser.parseAndEvaluate(expression);
  const std::string error = parser.getError();
  if (error.empty()) {
    std::cout << parser.getResult() << std::endl;
    return true;
  }
  if (kShowErrors) {
    if (!errorLabel.empty()) {
      std::cout << errorLabel << ':' << lineNumber << ": ";
    }
    std::cout << error << std::endl;
  }
  return false;
}

// Returns 0 when every evaluated line succeeded, 1 otherwise.
int evaluateLines(MathParser& parser, std::istream& input, LineFlags flags, const std::string& errorLabel) {
  int status = 0;
  std::string line;
  for (std::size_t lineNumber = 1;; ++lineNumber) {
    if ((flags & kLinePrompt) != 0) {
      std::cout << "> " << std::flush;
    }
    if (!std::getline(input, line)) {
      if ((flags & kLinePrompt) != 0) {
        std::cout << std::endl;
      }
      break;
    }
    if (lineNumber == 1) {
      stripLeadingUtf8Bom(line);
    }
    const std::string trimmed = trimCopy(line);
    if (trimmed.empty() || isWholeLineComment(trimmed)) {
      continue;
    }
    if ((flags & kLineStopOnQuit) != 0 && isQuitCommand(trimmed)) {
      break;
    }
    if (!evaluateAndPrint(parser, trimmed, errorLabel, lineNumber)) {
      status = 1;
    }
  }
  return status;
}

void printInteractiveBanner(const std::string& loadedFile) {
  if (!loadedFile.empty()) {
    std::cout << "Loaded " << loadedFile << '.' << std::endl;
  }
#if (SMARTMATH_TIME_VALUES && SMARTMATH_COMPLEX_NUMBERS && SMARTMATH_LAMBDA_FUNCTIONS)
  std::cout << "Time values, complex numbers, and lambdas are on." << std::endl
#else
  // assuming it is a minimal build, where everything is off
  std::cout << "Time values, complex numbers, and lambdas are off." << std::endl
#endif
            << "Waiting for a math expression or 'quit'/'exit'." << std::endl;
}

}  // namespace

int main(int argc, char** argv) {
  const ProgramArgs args = parseArgs(argc, argv);
  if (args.help) {
    printUsage(std::cout);
    return 0;
  }
  if (!args.error.empty()) {
    std::cerr << args.error << "\n\n";
    printUsage(std::cerr);
    return 1;
  }

  MathParser parser;

#if SMARTMATH_TIME_VALUES
  parser.setSupportTimeValues(kEnableTimeValues);
#endif

#if SMARTMATH_COMPLEX_NUMBERS
  parser.setSupportComplexNumbers(kEnableComplexNumbers);
#endif

#if SMARTMATH_LAMBDA_FUNCTIONS
  parser.setSupportLambdaFunctions(kEnableLambdaFunctions);
#endif

  const bool fileIsStdin = args.filePath == kStdinFilePath;
  std::ifstream file;
  if (!args.filePath.empty() && !fileIsStdin) {
    file.open(args.filePath.c_str());
    if (!file) {
      std::cerr << "cannot open file: " << args.filePath << '\n';
      return 1;
    }
  }

  int status = 0;
  for (const Step& step : args.steps) {
    if (step.kind == StepKind::File) {
      status |= fileIsStdin ? evaluateLines(parser, std::cin, 0, kStdinLabel)
                            : evaluateLines(parser, file, 0, step.text);
    } else if (!evaluateAndPrint(parser, step.text)) {
      status = 1;
    }
  }
  if (!args.steps.empty() && !args.interactive) {
    return status;
  }
  if (fileIsStdin && !reopenStdinFromTerminal()) {
    std::cerr << "-i/--interactive after -f -: no terminal to read from\n";
    return 1;
  }

  const bool interactive = stdinIsConsole();
  if (interactive) {
    printInteractiveBanner(fileIsStdin ? kStdinLabel : args.filePath);
  }
  const int stdinStatus =
      evaluateLines(parser, std::cin, kLineStopOnQuit | (interactive ? kLinePrompt : 0), std::string());
  return interactive ? 0 : (status | stdinStatus);
}
