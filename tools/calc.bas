'' Command-line calculator on MathParser.bas. Keep in sync with cpp/calc.cpp.
#include once "crt/stdio.bi"
#ifdef __FB_WIN32__
#include once "crt/io.bi"
#else
extern "C"
declare function isatty (byval fd as long) as long
declare function fileno (byval stream as FILE ptr) as long
end extern
#endif
#include once "..\Inc\MathParser.bi"

#ifdef __FB_WIN32__
'' Keeps the MinGW runtime from expanding "*" in arguments such as: calc 2 * 3
extern dowildcard alias "_dowildcard" as long
dim shared dowildcard as long = 0
#endif

const ENABLE_TIME_VALUES as Boolean = TRUE
const ENABLE_COMPLEX_NUMBERS as Boolean = TRUE
const ENABLE_LAMBDA_FUNCTIONS as Boolean = TRUE
const SHOW_ERRORS as Boolean = TRUE

const LINE_PROMPT as Long = 1        '' print "> " before reading each line
const LINE_STOP_ON_QUIT as Long = 2  '' "quit" / "exit" ends the input

const FILE_OPTION_PREFIX as String = "--file="
const STDIN_FILE_PATH as String = "-"
const STDIN_LABEL as String = "<stdin>"
const TRIM_CHARS as String = !" \t\r\n\v\f"

const USAGE_TEXT as String = _
  !"Usage: calc [expression] [-f <file>] [expression] [-i]\n" & _
  !"\n" & _
  !"Evaluates expressions and a file (one per line) in argument order on one\n" & _
  !"parser, so variables and ans carry over. With neither, reads standard\n" & _
  !"input (interactively in a terminal).\n" & _
  !"\n" & _
  !"  -f, --file <file>  Evaluate <file>. Use \"-f -\" to read standard input.\n" & _
  !"      --file=<file>\n" & _
  !"  -i, --interactive  When done, keep reading standard input (or the terminal,\n" & _
  !"                     if \"-f -\" used it up). Needs an expression or a file.\n" & _
  !"  -h, --help         Show this help (-h only as the first argument).\n" & _
  !"  --                 Treat the remaining arguments as expression text.\n" & _
  !"\n" & _
  !"Adjacent arguments form one expression; -f or -i ends it. Only the first\n" & _
  !"-i and -f are options; later ones are expression text. Blank lines and\n" & _
  !"whole-line # or // comments are skipped. quit or exit ends standard input\n" & _
  !"but is an expression in a file.\n" & _
  !"\n" & _
  !"  calc -f setup.txt x*2      file, then x*2\n" & _
  !"  calc x=5 -f test.txt       x=5, then file\n" & _
  !"  echo a=10 | calc -f - a+5  a=10 from standard input, then a+5\n" & _
  !"  calc -i 2+3 -i             2+3-i (imaginary unit), then interactive\n" & _
  !"  calc -- -i                 -i\n"

enum CalcStepKind
  STEP_EXPRESSION
  STEP_FILE
end enum

type CalcStep
  kind as CalcStepKind
  text as String  '' expression or file path
end type

type ProgramArgs
  help as Boolean
  interactive as Boolean
  filePath as String
  steps(any) as CalcStep  '' in command-line order
  errorText as String
end type

private function StartsWith(byref text as String, byref prefix as String) as Boolean
  return left(text, len(prefix)) = prefix
end function

private function IsQuitCommand(byref text as String) as Boolean
  dim lowered as String = lcase(text)
  return lowered = "quit" orelse lowered = "exit"
end function

private function IsWholeLineComment(byref trimmed as String) as Boolean
  return StartsWith(trimmed, "#") orelse StartsWith(trimmed, "//")
end function

private sub StripLeadingUtf8Bom(byref text as String)
  if StartsWith(text, chr(&hEF, &hBB, &hBF)) then
    text = mid(text, 4)
  end if
end sub

'' "--name" cannot be math, unlike "-x", "--3" or "--(x)".
private function IsLongOptionLike(byref arg as String) as Boolean
  if len(arg) <= 2 orelse StartsWith(arg, "--") = FALSE then return FALSE
  dim c as Integer = asc(lcase(mid(arg, 3, 1)))
  return c >= asc("a") andalso c <= asc("z")
end function

private function StdinIsConsole() as Boolean
#ifdef __FB_WIN32__
  return _isatty(_fileno(stdin)) <> 0
#else
  return isatty(fileno(stdin)) <> 0
#endif
end function

'' Rebinds stdin to the terminal.
private function ReopenStdinFromTerminal() as Boolean
#ifdef __FB_WIN32__
  return freopen("CONIN$", "r", stdin) <> NULL
#else
  return freopen("/dev/tty", "r", stdin) <> NULL
#endif
end function

'' Returns FALSE at end of input.
private function ReadLine(byval stream as FILE ptr, byref lineText as String) as Boolean
  dim buffer as zstring * 4096
  lineText = ""
  do
    if fgets(@buffer, sizeof(buffer), stream) = NULL then
      return len(lineText) > 0
    end if
    lineText &= buffer
  loop until right(lineText, 1) = !"\n"
  return TRUE
end function

private sub PrintToStderr(byref text as String)
  dim f as Long = freefile
  open err for output as #f
  print #f, text;
  close #f
end sub

private sub AddStep(byref args as ProgramArgs, byval kind as CalcStepKind, byref text as String)
  dim n as Integer = ubound(args.steps) + 1
  redim preserve args.steps(0 to n)
  args.steps(n).kind = kind
  args.steps(n).text = text
end sub

private sub ParseArgs(byval argc as Integer, byref args as ProgramArgs)
  if argc > 1 andalso command(1) = "-h" then
    args.help = TRUE
    exit sub
  end if

  dim optionsOpen as Boolean = TRUE
  dim expressionOpen as Boolean = FALSE
  dim i as Integer = 1
  while i < argc
    dim arg as String = command(i)
    i += 1
    if optionsOpen then
      if arg = "--" then
        optionsOpen = FALSE
        continue while
      end if
      if arg = "--help" then
        args.help = TRUE
        exit sub
      end if
      if args.interactive = FALSE andalso cbool(arg = "-i" orelse arg = "--interactive") then
        args.interactive = TRUE
        expressionOpen = FALSE
        continue while
      end if
      dim inlineFile as Boolean = StartsWith(arg, FILE_OPTION_PREFIX)
      if len(args.filePath) = 0 andalso (inlineFile orelse cbool(arg = "-f" orelse arg = "--file")) then
        if inlineFile then
          args.filePath = mid(arg, len(FILE_OPTION_PREFIX) + 1)
        elseif i < argc then
          args.filePath = command(i)
          i += 1
        else
          args.errorText = "missing file path after " & arg
          exit sub
        end if
        if len(args.filePath) = 0 then
          args.errorText = "missing file path"
          exit sub
        end if
        AddStep(args, STEP_FILE, args.filePath)
        expressionOpen = FALSE
        continue while
      end if
      if IsLongOptionLike(arg) then
        args.errorText = "unknown option: " & arg
        exit sub
      end if
    end if

    if expressionOpen then
      args.steps(ubound(args.steps)).text &= " " & arg
    else
      AddStep(args, STEP_EXPRESSION, arg)
      expressionOpen = TRUE
    end if
  wend

  if args.interactive andalso ubound(args.steps) < 0 then
    args.errorText = "-i/--interactive requires an expression or -f/--file" & _
                     " (to evaluate -i, use: calc -- -i)"
  end if
end sub

'' Prints the result or the error. A non-empty errorLabel prefixes errors as "label:line: ".
private function EvaluateAndPrint(byref expression as String, _
                                  byref errorLabel as const String = "", _
                                  byval lineNumber as Integer = 0) as Boolean
  dim value as Double
  dim resultText as String
  dim isArray as Boolean
  if Parser_TryEvaluateEx(expression, value, resultText, isArray) then
    print resultText
    return TRUE
  end if
  if SHOW_ERRORS then
    if len(errorLabel) > 0 then
      print errorLabel & ":" & lineNumber & ": ";
    end if
    print Parser_GetLastError()
  end if
  return FALSE
end function

'' Returns 0 when every evaluated line succeeded, 1 otherwise.
private function EvaluateLines(byval stream as FILE ptr, byval flags as Long, byref errorLabel as const String) as Long
  dim status as Long = 0
  dim lineText as String
  dim lineNumber as Integer = 0
  do
    if (flags and LINE_PROMPT) <> 0 then
      print "> ";
    end if
    if ReadLine(stream, lineText) = FALSE then
      if (flags and LINE_PROMPT) <> 0 then
        print
      end if
      exit do
    end if
    lineNumber += 1
    if lineNumber = 1 then
      StripLeadingUtf8Bom(lineText)
    end if
    dim trimmed as String = trim(lineText, any TRIM_CHARS)
    if len(trimmed) = 0 orelse IsWholeLineComment(trimmed) then
      continue do
    end if
    if (flags and LINE_STOP_ON_QUIT) <> 0 andalso IsQuitCommand(trimmed) then
      exit do
    end if
    if EvaluateAndPrint(trimmed, errorLabel, lineNumber) = FALSE then
      status = 1
    end if
  loop
  return status
end function

private sub PrintInteractiveBanner(byref loadedFile as const String)
  if len(loadedFile) > 0 then
    print "Loaded " & loadedFile & "."
  end if
  print "Time values, complex numbers, and lambdas are on."
  print "Waiting for a math expression or 'quit'/'exit'."
end sub

private function Main(byval argc as Integer) as Long
  dim args as ProgramArgs
  ParseArgs(argc, args)
  if args.help then
    print USAGE_TEXT;
    return 0
  end if
  if len(args.errorText) > 0 then
    PrintToStderr(args.errorText & !"\n" & USAGE_TEXT)
    return 1
  end if

  Parser_SetSupportTimeValues(ENABLE_TIME_VALUES)
  Parser_SetSupportComplexNumbers(ENABLE_COMPLEX_NUMBERS)
  Parser_SetSupportLambdaFunctions(ENABLE_LAMBDA_FUNCTIONS)

  dim fileIsStdin as Boolean = (args.filePath = STDIN_FILE_PATH)
  dim fileStream as FILE ptr = NULL
  if fileIsStdin then
    fileStream = stdin
  elseif len(args.filePath) > 0 then
    fileStream = fopen(args.filePath, "r")
    if fileStream = NULL then
      PrintToStderr("cannot open file: " & args.filePath & !"\n")
      return 1
    end if
  end if
  dim fileLabel as String = iif(fileIsStdin, STDIN_LABEL, args.filePath)

  dim status as Long = 0
  for i as Integer = 0 to ubound(args.steps)
    if args.steps(i).kind = STEP_FILE then
      status or= EvaluateLines(fileStream, 0, fileLabel)
    elseif EvaluateAndPrint(args.steps(i).text) = FALSE then
      status = 1
    end if
  next
  if fileStream <> NULL andalso fileIsStdin = FALSE then
    fclose(fileStream)
  end if
  if ubound(args.steps) >= 0 andalso args.interactive = FALSE then
    return status
  end if
  if fileIsStdin andalso ReopenStdinFromTerminal() = FALSE then
    PrintToStderr(!"-i/--interactive after -f -: no terminal to read from\n")
    return 1
  end if

  dim interactive as Boolean = StdinIsConsole()
  if interactive then
    PrintInteractiveBanner(fileLabel)
  end if
  dim stdinStatus as Long = EvaluateLines(stdin, LINE_STOP_ON_QUIT or iif(interactive, LINE_PROMPT, 0), "")
  return iif(interactive, 0, status or stdinStatus)
end function

end Main(__FB_ARGC__)
