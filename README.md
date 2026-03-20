## How to Use – LaTeX to Readable Formula

### Step 1 – Save your Word document first
Before running anything, save your `.docx` file so you can recover it if needed.

---

### Step 2 – Open the VBA Editor
Press **Alt + F11** inside Microsoft Word.

---

### Step 3 – Create a new Module
Click **Insert** in the top menu → Click **Module**

---

### Step 4 – Paste the code
Open the `ConvertLatex.txt` file, copy everything, and paste it into the module window.

---

### Step 5 – Run the macro
Press **F5** or click the green **▶ Run** button.

---

### What it converts
| LaTeX in text | Result in Word |
|---|---|
| `$C = P^e \mod n$` | C = Pᵉ mod n |
| `$\phi(n)$` | φ(n) |
| `$y^2 = x^3 + ax + b$` | y² = x³ + ax + b |
| `$k_a = y^a \mod P$` | kₐ = yᵃ mod P |
| `$\lambda$`, `$\alpha$` etc. | λ, α etc. |

---

### Requirements
- Microsoft Word (any version with VBA support)
- Works on Windows only (VBA not supported on Mac Word)
