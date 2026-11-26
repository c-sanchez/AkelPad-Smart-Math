# Reports hard-coded numeric types and 64-bit / binary64 constants that must live only in the
# type-config layer (cpp/MathParserTypes.hpp, Inc/MathParserTypes.bi).
# Usage (from anywhere):  powershell -File tools\check_type_literals.ps1 [-Strict] [-Detail]
#   default : prints per-file / per-pattern counts, exit code 0 (report only)
#   -Strict : exit code 1 when any non-allowlisted hit remains
#   -Detail : also prints every hit as file:line:text
param([switch]$Strict, [switch]$Detail)

$root = Split-Path -Parent $PSScriptRoot

$cppFiles = @(
  'cpp\MathParser.cpp', 'cpp\MathParser.hpp', 'cpp\MathParserInternal.hpp',
  'cpp\MathParserFactorInt.cpp', 'cpp\MathParserRawResult.cpp',
  'cpp\calc.cpp', 'cpp\calc_wasm.cpp')
$basFiles = @(
  'MathParser.bas', 'MathParserFactorInt.bas', 'MathParserRawResult.bas',
  'Inc\MathParser.bi', 'Inc\MathParserRawResult.bi',
  'SmartMath_Format.bas', 'SmartMath.bas', 'SmartMath_Globals.bi', 'tools\calc.bas')

$cppPatterns = [ordered]@{
  'double'        = '\bdouble\b'
  'long double'   = '\blong\s+double\b'
  'long long'     = '\blong\s+long\b'
  'u/int64_t'     = '\bstd::u?int64_t\b|\bu?int64_t\b'
  '__int64'       = '__int64'
  '1ull<<N'       = '\b1(ull|ULL)\s*<<'
  'ldexp(1,N)'    = 'ldexp\(\s*1\.0\s*,'
  '2^53 literal'  = '9007199254740992'
  '2^63/2^64 lit' = '922337203685477|1844674407370955'
}
$basPatterns = [ordered]@{
  'Double'        = '\bDouble\b'
  'LongInt'       = '\bLongInt\b'
  'ULongInt'      = '\bULongInt\b'
  'CLngInt'       = '\bC(U)?LngInt\b'
  '2^53 literal'  = '9007199254740992'
  '2^63/2^64 lit' = '922337203685477|1844674407370955|&h(7|8|F)FFFFFFFFFFFFFFF'
}

$total = 0
function Scan($files, $patterns, $commentPrefix) {
  foreach ($rel in $files) {
    $path = Join-Path $root $rel
    if (-not (Test-Path $path)) { continue }
    $lines = Get-Content $path
    $counts = [ordered]@{}
    $inFixed = $false
    for ($i = 0; $i -lt $lines.Count; $i++) {
      $t = $lines[$i].TrimStart()
      # Regions/lines marked SM_FIXED_WIDTH use a width on purpose (algorithm bound to binary64 / 64-bit).
      if ($lines[$i] -match "SM_FIXED_WIDTH_BEGIN") { $inFixed = $true; continue }
      if ($lines[$i] -match "SM_FIXED_WIDTH_END") { $inFixed = $false; continue }
      if ($inFixed -or $lines[$i] -match "SM_FIXED_WIDTH") { continue }
      if ($t.StartsWith($commentPrefix)) { continue }
      foreach ($k in $patterns.Keys) {
        if ($lines[$i] -match $patterns[$k]) {
          if (-not $counts.Contains($k)) { $counts[$k] = 0 }
          $counts[$k]++
          $script:total++
          if ($Detail) { "{0}:{1}: [{2}] {3}" -f $rel, ($i + 1), $k, $lines[$i].Trim() }
        }
      }
    }
    if ($counts.Count -gt 0) {
      $s = ($counts.GetEnumerator() | ForEach-Object { "$($_.Key)=$($_.Value)" }) -join ', '
      "{0,-34} {1}" -f $rel, $s
    }
  }
}

Scan $cppFiles $cppPatterns '//'
Scan $basFiles $basPatterns "'"
"TOTAL hits outside the config layer: $total"
if ($Strict -and $total -gt 0) { exit 1 }
exit 0
