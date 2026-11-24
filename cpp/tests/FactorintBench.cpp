#include "MathParser.hpp"
#include <chrono>
#include <cstdio>

static double benchExpr(MathParser& p, const char* expr) {
  const auto t0 = std::chrono::steady_clock::now();
  p.parseAndEvaluate(expr);
  const auto t1 = std::chrono::steady_clock::now();
  const double ms =
      std::chrono::duration<double, std::milli>(t1 - t0).count();
  std::printf("%8.2f ms  %s\n", ms, expr);
  return ms;
}

int main() {
  static const char* exprs[] = {
      "factorint(13)",
      "factorint(677599634708959)",
      "factorint(76568758722112367)",
      "factorint(18446744073709551615)",
      "factorint(2**63-1)",
  };
  MathParser p;
  double total = 0.0;
  for (const char* expr : exprs) {
    total += benchExpr(p, expr);
  }
  std::printf("TOTAL: %.2f ms\n", total);
  return 0;
}
