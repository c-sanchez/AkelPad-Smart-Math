#include <cstdint>
#include <cstdio>
#include <limits>

static std::uint64_t mulModAddDouble(std::uint64_t a, std::uint64_t b, std::uint64_t modN) {
  if (modN <= 1u) return 0u;
  std::uint64_t res = 0;
  a %= modN;
  while (b > 0u) {
    if ((b & 1u) != 0u) res = (res + a) % modN;
    b >>= 1;
    if (b == 0u) break;
    a = (a + a) % modN;
  }
  return res;
}

static std::uint64_t mulMod(std::uint64_t a, std::uint64_t b, std::uint64_t modN) {
  if (modN <= 1u) return 0u;
  a %= modN;
  b %= modN;
  if (a == 0u || b == 0u) return 0u;
  if (b <= UINT64_MAX / a) return (a * b) % modN;
  return mulModAddDouble(a, b, modN);
}

static std::uint64_t powMod(std::uint64_t baseVal, std::uint64_t exp, std::uint64_t modN) {
  volatile std::uint64_t res = 1;
  volatile std::uint64_t b = baseVal % modN;
  volatile std::uint64_t e = exp;
  while (e > 0u) {
    if ((e & 1u) != 0u) res = mulMod(static_cast<std::uint64_t>(res), static_cast<std::uint64_t>(b), modN);
    b = mulMod(static_cast<std::uint64_t>(b), static_cast<std::uint64_t>(b), modN);
    e >>= 1;
  }
  return static_cast<std::uint64_t>(res);
}

static bool isPrime(std::uint64_t n) {
  static const std::uint64_t mrBases[] = {2, 32544231, 2567547226ull, 4118087717ull,
                                        6700417, 12917328, 1297059741};
  std::uint64_t d = n - 1;
  int s = 0;
  while ((d & 1u) == 0u) {
    d >>= 1;
    ++s;
  }
  for (std::uint64_t a : mrBases) {
    if (n <= a) break;
    std::uint64_t x = powMod(a, d, n);
    if (x == 1u || x == n - 1u) continue;
    bool composite = true;
    for (int r = 1; r < s; ++r) {
      x = mulMod(x, x, n);
      if (x == n - 1u) {
        composite = false;
        break;
      }
    }
    if (composite) return false;
  }
  return true;
}

int main() {
  const std::uint64_t n = 677599634708959ull;
  std::printf("powMod(2,d,n)=%llu isPrime=%d\n", powMod(2, (n - 1u) / 2u, n), isPrime(n) ? 1 : 0);
  return isPrime(n) ? 0 : 1;
}
