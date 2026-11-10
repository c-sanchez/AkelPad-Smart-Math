#include "MathParser.hpp"

#include <cmath>

void MathParser::RawResult::CartesianScalar::clear() {
  kind = ScalarKind::FloatingPoint;
  floatingPoint = 0.0;
}

bool MathParser::RawResult::CartesianScalar::isZero() const {
  switch (kind) {
  case ScalarKind::Int64:
  case ScalarKind::Time:
    return intValue == 0;
  case ScalarKind::UInt64:
    return uintValue == 0;
  case ScalarKind::Rational:
#if SMARTMATH_FACTORINT
  case ScalarKind::IntPower:
#endif
    return rational.numerator == 0;
  case ScalarKind::FloatingPoint:
  case ScalarKind::Complex:
    break;
  }
  if (std::isnan(floatingPoint)) {
    return false;
  }
  return floatingPoint == 0.0;
}

void MathParser::RawResult::clear() {
  kind = Kind::None;
  scalar.kind = ScalarKind::FloatingPoint;
  scalar.real.clear();
  scalar.imag.clear();
  array.clear();
}
