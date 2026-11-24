#include "MathParser.hpp"

#include <chrono>
#include <cstdint>
#include <cstdlib>
#include <iostream>
#include <algorithm>
#include <string>

namespace {

long long rawCartesianToInt64(const MathParser::RawResult::CartesianScalar& s) {
  switch (s.kind) {
    case MathParser::RawResult::ScalarKind::Time:
    case MathParser::RawResult::ScalarKind::Int64:
      return s.intValue;
    case MathParser::RawResult::ScalarKind::UInt64:
      return static_cast<long long>(s.uintValue);
    case MathParser::RawResult::ScalarKind::Rational:
      return static_cast<long long>(static_cast<double>(s.rational.numerator) /
                                    static_cast<double>(s.rational.denominator));
#if SMARTMATH_FACTORINT
    case MathParser::RawResult::ScalarKind::IntPower:
      return s.rational.numerator;
#endif
    case MathParser::RawResult::ScalarKind::FloatingPoint:
    default:
      return static_cast<long long>(s.floatingPoint);
  }
}

long long rawScalarToInt64(const MathParser::RawResult::Scalar& s) {
  return rawCartesianToInt64(s.real);
}

long long benchResultChecksum(const MathParser::RawResult& r) {
  if (!r.hasValue()) {
    return 0;
  }
  if (r.isScalar()) {
    return rawScalarToInt64(r.scalar);
  }
  long long sum = 0;
  for (const auto& item : r.array) {
    sum += rawScalarToInt64(item);
  }
  return sum;
}

double benchParseAndEvaluateRaw(MathParser& parser, const std::string& expr, std::uint64_t& errors) {
  const MathParser::RawResult r = parser.parseAndEvaluateRaw(expr);
  if (!r.hasValue()) {
    ++errors;
    return 0.0;
  }
  return static_cast<double>(benchResultChecksum(r));
}

void runLambdaFlagStressBench(std::uint64_t iterations) {
  std::cout << "=== lambda-stress (sortby/assign paths that differ by lambda flag) ===\n";
  std::cout << "Default hot-path expressions (myconst1 +/- i, etc.) never hit assignment or sortby key parsing,\n";
  std::cout << "so toggling setSupportLambdaFunctions() there usually shows no ops/sec difference.\n\n";

  for (const bool lambdaOn : {true, false}) {
    MathParser parser;
    parser.setSupportTimeValues(false);
    parser.setSupportComplexNumbers(false);
    parser.setSupportLambdaFunctions(lambdaOn);
    parser.addConst("myconst1", 123LL);

    const char* const exprsLambdaOn[] = {
        "sortby((3,1,2), x:-x)",
        "f=x:x+2; f(3)",
        "sortby((1,2), (x):(1/x))",
    };
    const char* const exprsLambdaOff[] = {
        "sortby((-3,-1,2), abs)",
        "myconst1 + 3",
        "sortby((3,1,2), abs)",
    };

    std::uint64_t errors = 0;
    long long checksum = 0;
    const auto t0 = std::chrono::steady_clock::now();
    for (std::uint64_t i = 0; i < iterations; ++i) {
      const char* expr = lambdaOn ? exprsLambdaOn[i % 3] : exprsLambdaOff[i % 3];
      checksum += static_cast<long long>(benchParseAndEvaluateRaw(parser, expr, errors));
    }
    const auto t1 = std::chrono::steady_clock::now();
    const auto elapsedNs = std::chrono::duration_cast<std::chrono::nanoseconds>(t1 - t0).count();
    const double elapsedSec = static_cast<double>(elapsedNs) / 1e9;
    const double opsPerSec = (elapsedSec > 0.0) ? (static_cast<double>(iterations) / elapsedSec) : 0.0;

    std::cout << "lambda_support=" << (lambdaOn ? "true" : "false") << "\n";
    std::cout << "elapsed_sec: " << elapsedSec << "\n";
    std::cout << "ops_per_sec: " << opsPerSec << "\n";
    std::cout << "checksum: " << checksum << "\n";
    std::cout << "errors: " << errors << "\n\n";
  }
}

void runPhaseProfileBench(std::uint64_t iterations) {
  std::cout << "=== phase-profile (compile / evaluate / raw export) ===\n";
  MathParser parser;
  parser.setSupportTimeValues(false);
  parser.setSupportLambdaFunctions(false);
  parser.setSupportComplexNumbers(false);
  parser.addConst("myconst1", 123LL);
  parser.addConst("myconst2", 345LL);

  std::uint64_t errors = 0;
  long long checksum = 0;
  std::int64_t nsExprBuild = 0;
  std::int64_t nsCompile = 0;
  std::int64_t nsEvaluate = 0;
  std::int64_t nsRawExport = 0;

  for (std::uint64_t i = 0; i < iterations; ++i) {
    const auto tBuild0 = std::chrono::steady_clock::now();
    std::string expr;
    if ((i % 10) == 0) {
      expr = "myconst1 / (" + std::to_string(i) + " + 0.5)";
    } else if ((i % 3) == 0) {
      expr = "myconst1 + " + std::to_string(i);
    } else if ((i % 4) == 0) {
      expr = "myconst1 - " + std::to_string(i);
    } else if ((i % 5) == 0) {
      expr = "myconst2 | " + std::to_string(i);
    } else if ((i % 7) == 0) {
      expr = "myconst2 & " + std::to_string(i);
    } else if ((i % 11) == 0) {
      expr = "myconst1 << (" + std::to_string(i) + "%63)";
    } else {
      expr = "((myconst1 & 0xFFFF) + (myconst2 | 0x07)) << (2 + 3)";
    }
    const auto tBuild1 = std::chrono::steady_clock::now();
    nsExprBuild += std::chrono::duration_cast<std::chrono::nanoseconds>(tBuild1 - tBuild0).count();

    const auto tCompile0 = std::chrono::steady_clock::now();
    const bool compiled = parser.compile(expr);
    const auto tCompile1 = std::chrono::steady_clock::now();
    nsCompile += std::chrono::duration_cast<std::chrono::nanoseconds>(tCompile1 - tCompile0).count();

    const auto tEval0 = std::chrono::steady_clock::now();
    if (compiled) {
      parser.evaluate();
    }
    const auto tEval1 = std::chrono::steady_clock::now();
    nsEvaluate += std::chrono::duration_cast<std::chrono::nanoseconds>(tEval1 - tEval0).count();

    const auto tRaw0 = std::chrono::steady_clock::now();
    const MathParser::RawResult r = parser.getRawResult();
    const auto tRaw1 = std::chrono::steady_clock::now();
    nsRawExport += std::chrono::duration_cast<std::chrono::nanoseconds>(tRaw1 - tRaw0).count();

    if (!compiled || !r.hasValue() || !r.isScalar()) {
      ++errors;
    } else {
      checksum += rawScalarToInt64(r.scalar);
    }
  }

  const auto totalNs = nsExprBuild + nsCompile + nsEvaluate + nsRawExport;
  const double totalSec = static_cast<double>(totalNs) / 1e9;
  const double opsPerSec = (totalSec > 0.0) ? (static_cast<double>(iterations) / totalSec) : 0.0;
  const auto pct = [&](std::int64_t partNs) -> double {
    return (totalNs > 0) ? (100.0 * static_cast<double>(partNs) / static_cast<double>(totalNs)) : 0.0;
  };

  std::cout << "iterations: " << iterations << "\n";
  std::cout << "ops_per_sec: " << opsPerSec << "\n";
  std::cout << "checksum: " << checksum << "\n";
  std::cout << "errors: " << errors << "\n";
  std::cout << "expr_build_pct: " << pct(nsExprBuild) << "\n";
  std::cout << "compile_pct: " << pct(nsCompile) << "\n";
  std::cout << "evaluate_pct: " << pct(nsEvaluate) << "\n";
  std::cout << "raw_export_pct: " << pct(nsRawExport) << "\n";
  std::cout << "expr_build_ms_per_op: " << (static_cast<double>(nsExprBuild) / static_cast<double>(iterations) / 1e6) << "\n";
  std::cout << "compile_ms_per_op: " << (static_cast<double>(nsCompile) / static_cast<double>(iterations) / 1e6) << "\n";
  std::cout << "evaluate_ms_per_op: " << (static_cast<double>(nsEvaluate) / static_cast<double>(iterations) / 1e6) << "\n";
  std::cout << "raw_export_ms_per_op: " << (static_cast<double>(nsRawExport) / static_cast<double>(iterations) / 1e6) << "\n\n";
}

} // namespace

int main(int argc, char** argv) {
  std::uint64_t iterations = 200000;
  std::uint64_t repeats = 1;
  std::uint64_t warmupRuns = 0;
  bool lambdaStress = false;
  bool phaseProfile = false;
  if (argc > 1) {
    const std::uint64_t parsed = static_cast<std::uint64_t>(_strtoui64(argv[1], nullptr, 10));
    if (parsed > 0) {
      iterations = parsed;
    }
  }
  if (argc > 2) {
    const std::uint64_t parsed = static_cast<std::uint64_t>(_strtoui64(argv[2], nullptr, 10));
    if (parsed > 0) {
      repeats = parsed;
    }
  }
  if (argc > 3) {
    warmupRuns = static_cast<std::uint64_t>(_strtoui64(argv[3], nullptr, 10));
  }
  if (argc > 4) {
    const std::string mode = argv[4];
    if (mode == "lambda-stress") {
      lambdaStress = true;
    } else if (mode == "profile") {
      phaseProfile = true;
    }
  }

  if (lambdaStress) {
    runLambdaFlagStressBench(iterations);
    return 0;
  }

  if (phaseProfile) {
    runPhaseProfileBench(iterations);
    return 0;
  }

  double totalElapsedSec = 0.0;
  double totalOpsPerSec = 0.0;
  std::uint64_t totalErrors = 0;
  long long totalChecksum = 0;
  std::uint64_t totalFloatPathCount = 0;
  std::uint64_t totalIntLikePathCount = 0;
  std::vector<double> measuredOpsPerSec;
  measuredOpsPerSec.reserve(static_cast<std::size_t>(repeats));

  const std::uint64_t totalRuns = warmupRuns + repeats;

  for (std::uint64_t run = 0; run < totalRuns; ++run) {
    const bool isWarmup = (run < warmupRuns);
    MathParser parser;
    // Note:
    // On my machine, I get around 395_000 ops_per_sec with SupportTimeValues enabled.
    // With them disabled, I get around 400_000 ops_per_sec.
    // Toggling setSupportLambdaFunctions() here usually does NOT change this number: the
    // expressions below are plain arithmetic (no assignment RHS, no sortby lambda keys).
    // Use: PerfHotPath_MathParser.exe <iterations> <repeats> <warmup> lambda-stress
    parser.setSupportTimeValues(false);
    parser.setSupportLambdaFunctions(false);
    parser.setSupportComplexNumbers(false);
    parser.addConst("myconst1", 123LL);
    parser.addConst("myconst2", 345LL);

    std::uint64_t errors = 0;
    long long checksum = 0;

    const auto t0 = std::chrono::steady_clock::now();
    for (std::uint64_t i = 0; i < iterations; ++i) {
      std::string expr;
      if ((i % 10) == 0) {
        ++totalFloatPathCount;
        expr = "myconst1 / (" + std::to_string(i) + " + 0.5)";
      } else if ((i % 3) == 0) {
        ++totalIntLikePathCount;
        expr = "myconst1 + " + std::to_string(i);
      } else if ((i % 4) == 0) {
        ++totalIntLikePathCount;
        expr = "myconst1 - " + std::to_string(i);
      } else if ((i % 5) == 0) {
        ++totalIntLikePathCount;
        expr = "myconst2 | " + std::to_string(i);
      } else if ((i % 7) == 0) {
        ++totalIntLikePathCount;
        expr = "myconst2 & " + std::to_string(i);
      } else if ((i % 11) == 0) {
        ++totalIntLikePathCount;
        expr = "myconst1 << (" + std::to_string(i) + "%63)";
      } else {
        ++totalIntLikePathCount;
        expr = "((myconst1 & 0xFFFF) + (myconst2 | 0x07)) << (2 + 3)";
      }

      // Keep this inline: the benchmark measures parser throughput, not helper wrappers.
      // benchParseAndEvaluateRaw() adds extra calls/branches per iteration and skews ops/sec.
      const MathParser::RawResult r = parser.parseAndEvaluateRaw(expr);
      if (!r.hasValue() || !r.isScalar()) {
        ++errors;
      } else {
        checksum += rawScalarToInt64(r.scalar);
      }
    }
    const auto t1 = std::chrono::steady_clock::now();

    const auto elapsedNs = std::chrono::duration_cast<std::chrono::nanoseconds>(t1 - t0).count();
    const double elapsedSec = static_cast<double>(elapsedNs) / 1e9;
    const double opsPerSec = (elapsedSec > 0.0) ? (static_cast<double>(iterations) / elapsedSec) : 0.0;

    std::cout << "run: " << (run + 1) << "/" << totalRuns;
    if (isWarmup) {
      std::cout << " (warmup)";
    }
    std::cout << "\n";
    std::cout << "elapsed_sec: " << elapsedSec << "\n";
    std::cout << "ops_per_sec: " << opsPerSec << "\n";
    std::cout << "checksum: " << checksum << "\n";
    std::cout << "errors: " << errors << "\n";

    if (!isWarmup) {
      totalElapsedSec += elapsedSec;
      totalOpsPerSec += opsPerSec;
      totalChecksum += checksum;
      totalErrors += errors;
      measuredOpsPerSec.push_back(opsPerSec);
    }
  }

  std::cout << "iterations: " << iterations << "\n";
  std::cout << "repeats: " << repeats << "\n";
  std::cout << "warmup_runs: " << warmupRuns << "\n";
  std::cout << "avg_elapsed_sec: " << (totalElapsedSec / static_cast<double>(repeats)) << "\n";
  std::cout << "avg_ops_per_sec: " << (totalOpsPerSec / static_cast<double>(repeats)) << "\n";
  if (!measuredOpsPerSec.empty()) {
    std::sort(measuredOpsPerSec.begin(), measuredOpsPerSec.end());
    const std::size_t n = measuredOpsPerSec.size();
    const double medianOpsPerSec = (n % 2 == 1)
        ? measuredOpsPerSec[n / 2]
        : 0.5 * (measuredOpsPerSec[(n / 2) - 1] + measuredOpsPerSec[n / 2]);
    std::cout << "median_ops_per_sec: " << medianOpsPerSec << "\n";
    std::cout << "min_ops_per_sec: " << measuredOpsPerSec.front() << "\n";
    std::cout << "max_ops_per_sec: " << measuredOpsPerSec.back() << "\n";
  }
  std::cout << "total_checksum: " << totalChecksum << "\n";
  std::cout << "total_errors: " << totalErrors << "\n";
  const std::uint64_t totalOps = iterations * (repeats + warmupRuns);
  const double floatShare = (totalOps > 0) ? (100.0 * static_cast<double>(totalFloatPathCount) / static_cast<double>(totalOps)) : 0.0;
  const double intLikeShare = (totalOps > 0) ? (100.0 * static_cast<double>(totalIntLikePathCount) / static_cast<double>(totalOps)) : 0.0;
  std::cout << "float_path_count: " << totalFloatPathCount << "\n";
  std::cout << "int_like_path_count: " << totalIntLikePathCount << "\n";
  std::cout << "float_path_percent: " << floatShare << "\n";
  std::cout << "int_like_path_percent: " << intLikeShare << "\n";
  return 0;
}
