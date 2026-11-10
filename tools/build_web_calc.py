"""Build the SmartMath web calculator (GitHub Pages site).

Output folder (default: tools/_site):
  index.html, app.js, style.css   copied from tools/web/
  builtin-functions.js            function names from AkelFiles/Plugs/Coder/smartmath.coder
  smartmath.js, smartmath.wasm    cpp/calc_wasm.cpp + the C++ MathParser, built with Emscripten (emcc)
  help.html                       USAGE_AND_SYNTAX.md converted with pandoc
                                  (the AkelPad-only "SmartMath Options" section is omitted)

Usage:
  python tools/build_web_calc.py [--out DIR] [--no-help] [--tests]

--tests also builds cpp/Tests_MathParser.cpp with Emscripten into tools/_wasm-tests/
and runs it with Node.js, so the WebAssembly build of the parser is checked against
the same tests as the MSVC build.
em++, pandoc and node are looked up on PATH. The tools this run will
actually call are checked first; any that are missing are listed, with a
link to where to get them, and the build stops before it writes anything.
"""

import argparse
import shutil
import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
TOOLS = Path(__file__).resolve().parent
CPP = ROOT / "cpp"
WEB = TOOLS / "web"
WASM_TESTS = TOOLS / "_wasm-tests"

PARSER_SOURCES = ["MathParser.cpp", "MathParserFactorInt.cpp", "MathParserRawResult.cpp"]
STATIC_FILES = ["index.html", "app.js", "style.css"]

# The parser recurses per nesting level; Emscripten's default 64 KB stack is far below MSVC's 1 MB.
COMMON_EMCC_FLAGS = ["-O2", "-sALLOW_MEMORY_GROWTH=1", "-sSTACK_SIZE=4MB"]

# Printed when the tool is not on PATH. Keyed by the executable name.
TOOL_HINTS = {
    "em++": "Emscripten SDK (activate it so em++ is on PATH): https://emscripten.org/docs/getting_started/downloads.html",
    "pandoc": "Pandoc (or rerun with --no-help): https://pandoc.org/installing.html",
    "node": "Node.js: https://nodejs.org/",
}


def tools_needed(no_help, tests):
    names = ["em++"]
    if not no_help:
        names.append("pandoc")
    if tests:
        names.append("node")
    return names


def resolve_tools(names):
    resolved = {}
    missing = []
    for name in names:
        path = shutil.which(name)
        if path is None:
            missing.append(name)
        else:
            resolved[name] = path
    if missing:
        width = max(len(name) for name in missing)
        lines = ["error: missing tools (not on PATH):"]
        for name in missing:
            lines.append(f"  {name:<{width}}  {TOOL_HINTS[name]}")
        sys.exit("\n".join(lines))
    return resolved


def run(args):
    print("+ " + " ".join(str(a) for a in args), flush=True)
    subprocess.run([str(a) for a in args], check=True)


def coder_function_names():
    path = ROOT / "AkelFiles" / "Plugs" / "Coder" / "smartmath.coder"
    names = []
    in_functions = False
    for line in path.read_text(encoding="utf-16").splitlines():
        stripped = line.strip()
        if not in_functions:
            if stripped == "; functions:":
                in_functions = True
            continue
        if not stripped or stripped.startswith(";"):
            break
        names.append(stripped.split()[-1])
    if not names:
        sys.exit(f"error: no function names found in {path}")
    return names


def write_builtin_functions(out_dir):
    names = coder_function_names()
    lines = [
        "// Generated from AkelFiles/Plugs/Coder/smartmath.coder (the ; functions: section).",
        "export default new Set([",
    ]
    lines.extend(f'  "{name}",' for name in names)
    lines.append("]);\n")
    (out_dir / "builtin-functions.js").write_text("\n".join(lines), encoding="utf-8")
    print(f"+ wrote {len(names)} builtin function names")


def build_wasm(out_dir, emxx):
    sources = [CPP / "calc_wasm.cpp"] + [CPP / s for s in PARSER_SOURCES]
    run([emxx, "-std=c++17", *COMMON_EMCC_FLAGS, "-lembind",
         "-sMODULARIZE=1", "-sEXPORT_ES6=1", "-sEXPORT_NAME=createSmartMath",
         "-sENVIRONMENT=web",
         *sources, "-o", out_dir / "smartmath.js"])


def markdown_for_web_help():
    text = (ROOT / "USAGE_AND_SYNTAX.md").read_text(encoding="utf-8")
    marker = "\n## SmartMath Options\n"
    cut = text.find(marker)
    if cut < 0:
        sys.exit("error: USAGE_AND_SYNTAX.md is missing the '## SmartMath Options' section")
    return text[:cut].rstrip() + "\n"


def build_help(out_dir, pandoc):
    args = [pandoc, "--from=gfm", "--to=html5", "--standalone",
            "--toc", "--toc-depth=3",
            "--template", WEB / "help-template.html",
            "--metadata", "pagetitle=SmartMath Usage and Syntax",
            "--output", out_dir / "help.html"]
    print("+ " + " ".join(str(a) for a in args), flush=True)
    subprocess.run([str(a) for a in args], input=markdown_for_web_help(),
                   text=True, encoding="utf-8", check=True)


def run_tests(build_dir, emxx, node):
    build_dir.mkdir(parents=True, exist_ok=True)
    test_js = build_dir / "Tests_MathParser.js"
    sources = [CPP / "Tests_MathParser.cpp"] + [CPP / s for s in PARSER_SOURCES]
    run([emxx, "-std=c++14", *COMMON_EMCC_FLAGS, "-sENVIRONMENT=node",
         *sources, "-o", test_js])
    run([node, test_js])


def main():
    parser = argparse.ArgumentParser(description="Build the SmartMath web calculator.")
    parser.add_argument("--out", default=str(TOOLS / "_site"),
                        help="output folder (default: tools/_site)")
    parser.add_argument("--no-help", action="store_true", help="skip help.html (no pandoc needed)")
    parser.add_argument("--tests", action="store_true", help="also build and run the C++ tests under Node.js")
    args = parser.parse_args()
    tools = resolve_tools(tools_needed(args.no_help, args.tests))

    out_dir = Path(args.out).resolve()
    out_dir.mkdir(parents=True, exist_ok=True)
    for name in STATIC_FILES:
        shutil.copy2(WEB / name, out_dir / name)
    write_builtin_functions(out_dir)
    (out_dir / ".nojekyll").write_text("", encoding="utf-8")
    build_wasm(out_dir, tools["em++"])
    if not args.no_help:
        build_help(out_dir, tools["pandoc"])
    if args.tests:
        run_tests(WASM_TESTS, tools["em++"], tools["node"])
    print(f"Done: {out_dir}")


if __name__ == "__main__":
    main()
