# SmartMath Plugin for AkelPad

**SmartMath** is a lightweight, powerful plugin that turns AkelPad into a dynamic real-time calculator, inspired by tools like *Soulver*, *Numara* and *SpeQ Mathematics*. It allows you to perform calculations naturally in plain text, showing results instantly on the screen without altering your actual file content.

![image](https://i.ibb.co.com/Pvw762qh/akelpad-smartmath-01.png)

![image](https://i.ibb.co.com/7tX48Y1g/akelpad-smartmath-02.png)

![image](https://i.ibb.co.com/NgRWmW9M/akelpad-smartmath-03.png)

### ✨ Features
*   **Real-Time Evaluation:** Instantly evaluates mathematical expressions as you type.
*   **Non-Invasive Rendering:** Results are drawn directly onto the screen margin using Windows GDI. Your actual text document remains 100% untouched and clean.
*   **Variable Support:** Define variables (e.g., `base = 100`) and reference them in subsequent lines (e.g., `base * 1.15`).
*   **Arrays:** Use array arguments and variables (e.g. `(1, 2, 3)*10` and `a = (pi/4, pi/2); sin(a)`).
*   **Smart Percentage Logic:** Intuitive handling of the `%` character relative to the left operand (e.g., `500 + 21%` correctly yields `605`).
*   **Advanced Number Formatting:** 
    *   Customizable precision from 0 up to 14 decimal places.
    *   Optional **Thousands Separator** with `.` as decimal separator and `'` as thousands separator (e.g., `1'000'500.50`).
*   **Hexadecimal, Octal and Binary Numbers:** SmartMath supports prefixes for hexadecimal numbers (`0x7F`), octal numbers (`0o15`) and binary numbers (`0b01001`). The output formatting functions `hex`, `oct` and `bin` are available.
*   **Built-in Functions:** SmartMath supports most of the functions you usually find in calculators (e.g. `abs`, `sin`, `log`, `max` etc.).
*   **User-Defined Functions:** Define your own functions (e.g. `f(x,y) = x**2 + y**2; f(5,6)`).
*   **Customizable Aesthetics:** Choose from 6 different text colors (Green, Blue, Red, Yellow, White, Black) to match your AkelPad theme. It also perfectly integrates with AkelPad's "Active Line" background highlight.
*   **Native Autoload:** Integrates directly with AkelPad's plugin manager to remember its active state across sessions.
*   **Copy the Result to the Clipboard:** Double-click the result to copy it to the clipboard.

### 🛠 Technical Details
*   **Language:** 100% FreeBASIC (Modular architecture).
*   **Interface:** Built on top of the AkelPad Plugin API (`AkelDLL.bi`) and Windows API.
*   **Rendering Engine:** Uses **Window Subclassing** to intercept `WM_PAINT` and `WM_SIZE` messages, dynamically expanding the right margin (`EM_SETMARGINS`) and drawing numbers via `TextOut`. It does not rely on text annotations like Scintilla.
*   **Parsing:** Features a custom, highly optimized Recursive Descent Parser. Zero dependencies on external scripting engines (like VBScript or JScript).
*   **Architecture:** Compiles to a highly optimized **x86** (32-bit) and **x64** (64-bit) native Windows DLL.
*   **Storage:** Settings are saved either in a `SmartMath.ini` file within the AkelPad plugins configuration folder or in the Registry according to AkelPad's settings.

### ⚙️ Compilation
To compile the plugin yourself:
1.  Ensure you have the **FreeBASIC** compiler installed and added to your PATH (the MinGW toolchain bundled with FB should include **`windres`**, which compiles `SmartMath.rc` into version metadata embedded in the DLL).
2.  Open a terminal in the project's root folder.
3.  Run the `Compile.bat` file. This will link all the modules (`SmartMath.bas`, `SmartMath_Config.bas`, etc.) and the Windows resource script **`SmartMath.rc`** (file description, version, copyright) into a single DLL.
4.  Use `Compile32.bat` to get a 32-bit (x86) binary and `Compile64.bat` to get a 64-bit (x86_64) binary. Note: path to `fbc` should either be in PATH or explicitly specified as `FB_HOME` in these `.bat` files.

### 📘 Language Reference
The full expression language reference (functions, operators, precedence, arrays, variables, comments, and usage tips) is documented in:
- `USAGE_AND_SYNTAX.md`

### 📦 Installation
1.  Locate your AkelPad installation directory.
2.  Copy the compiled DLL (`SmartMath.dll`) into the `AkelFiles\Plugs\` folder.
    *   Typical path: `C:\Program Files (x86)\AkelPad\AkelFiles\Plugs\`
3.  Restart AkelPad.
4.  Go to `Options -> Plugins` (or press `Alt+P`), find `SmartMath::ToggleSmartMath`, and check it to enable it and set it to Autoload.

### Development
1. It is highly recommended to use an AI-assistant for any further development and bug fixing.
2. Use the `add-mathparser-function` skill (under ".cursor\skills\add-mathparser-function") to add new functions or operators. The prompt to AI will be similar to "using the `add-mathparser-function` skill, implement functionality which ... (description of the functionality)". In such way, the new code will automatically follow the existing code structure and paths, reuse the existing helpers, will be covered with new tests and will be reflected in the documentation ("USAGE_AND_SYNTAX.md"). To have more clarity and control of what AI is going to do, add the following to the prompt: "If you have any questions or concerns, ask explicitly. Don't make assumptions".
3. Use the `parser-reusability-cleanup` skill (under ".cursor\skills\parser-reusability-cleanup") for code efficiency, maintenance, refactoring and cleanup. The prompt to AI may be similar to "using the `parser-reusability-cleanup` skill, what would you refactor, simplify or optimize in the parser to either improve the performance or to reduce the code base?".

### 🙏 Special Thanks
I would like to express my deepest gratitude to **Mysoft**, whose persistence and insights were instrumental in convincing me to embrace FreeBASIC for plugin development. While the Notepad++ version of this tool was originally a port from FreePascal, this AkelPad version was built from the ground up using **FreeBASIC**.

I also want to give a special mention to **Jepalza**, who provided essential help in overcoming a critical technical roadblock at the very beginning of this project. Thanks to his expertise, I was able to move forward with the development of SmartMath.

A huge thank you to **Vitaliy "DV" Dovgan**. Without him, SmartMath wouldn't truly be "Smart", but rather just a plugin with basic mathematical evaluations and little else. He has done an excellent job adding numerous advanced features, various improvements, and optimizations throughout the entire code.