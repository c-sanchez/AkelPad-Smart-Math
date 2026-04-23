# SmartMath Plugin for AkelPad

**SmartMath** is a lightweight, powerful plugin that turns AkelPad into a dynamic real-time calculator, inspired by tools like *Soulver* or *Numara*. It allows you to perform calculations naturally in plain text, showing results instantly on the screen without altering your actual file content.

![image](https://i.ibb.co.com/Pvw762qh/akelpad-smartmath-01.png)

![image](https://i.ibb.co.com/7tX48Y1g/akelpad-smartmath-02.png)

![image](https://i.ibb.co.com/NgRWmW9M/akelpad-smartmath-03.png)

### ✨ Features
*   **Real-Time Evaluation:** Instantly evaluates addition, subtraction, multiplication, and division as you type.
*   **Non-Invasive Rendering:** Results are drawn directly onto the screen margin using Windows GDI. Your actual text document remains 100% untouched and clean.
*   **Variable Support:** Define variables (e.g., `base = 100`) and reference them in subsequent lines (e.g., `base * 1.15`).
*   **Smart Percentage Logic:** Intuitive handling of the `%` character relative to the left operand (e.g., `500 + 21%` correctly yields `605`).
*   **Advanced Number Formatting:** 
    *   Customizable precision from 0 up to 14 decimal places.
    *   Optional **Thousands Separator** with smart locale handling (uses `.` for thousands and `,` for decimals, e.g., `1.000.500,50`).
*   **Customizable Aesthetics:** Choose from 6 different text colors (Green, Blue, Red, Yellow, White, Black) to match your AkelPad theme. It also perfectly integrates with AkelPad's "Active Line" background highlight.
*   **Native Autoload:** Integrates directly with AkelPad's plugin manager to remember its active state across sessions.

### 🛠 Technical Details
*   **Language:** 100% FreeBASIC (Modular architecture).
*   **Interface:** Built on top of the AkelPad Plugin API (`AkelDLL.bi`) and Windows API.
*   **Rendering Engine:** Uses **Window Subclassing** to intercept `WM_PAINT` and `WM_SIZE` messages, dynamically expanding the right margin (`EM_SETMARGINS`) and drawing numbers via `TextOut`. It does not rely on text annotations like Scintilla.
*   **Parsing:** Features a custom, highly optimized Recursive Descent Parser. Zero dependencies on external scripting engines (like VBScript or JScript).
*   **Architecture:** Compiles to a highly optimized **x86** (32-bit) and **x64** (64-bit) native Windows DLL.
*   **Storage:** Settings are saved in a `SmartMath.ini` file seamlessly within the AkelPad plugins configuration folder.

### ⚙️ Compilation
To compile the plugin yourself:
1.  Ensure you have the **FreeBASIC** compiler installed and added to your PATH.
2.  Open a terminal in the project's root folder.
3.  Run the `Compile.bat` file. This will link all the modules (`SmartMath.bas`, `SmartMath_Config.bas`, etc.) into a single DLL.

### 📦 Installation
1.  Locate your AkelPad installation directory.
2.  Copy the compiled DLL (`SmartMath.dll`) into the `AkelFiles\Plugs\` folder.
    *   Typical path: `C:\Program Files (x86)\AkelPad\AkelFiles\Plugs\`
3.  Restart AkelPad.
4.  Go to `Options -> Plugins` (or press `Alt+P`), find `SmartMath::ToggleSmartMath`, and check it to enable it and set it to Autoload.

### 🙏 Special Thanks
I would like to express my deepest gratitude to **Mysoft**, whose persistence and insights were instrumental in convincing me to embrace FreeBASIC for plugin development. While the Notepad++ version of this tool was originally a port from FreePascal, this AkelPad version was built from the ground up using **FreeBASIC**.

I also want to give a special mention to **Jepalza**, who provided essential help in overcoming a critical technical roadblock at the very beginning of this project. Thanks to his expertise, I was able to move forward with the development of SmartMath.