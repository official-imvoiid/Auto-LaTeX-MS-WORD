# Auto-LaTeX MS Word

Converts `$formula$` latex text in your word doc to proper readable 
math with superscripts subscripts and greek symbols automatically.

---

## How to Use

1. **Save your Word doc first** (just in case)
2. Press `Alt + F11` to open VBA editor
3. Click **Insert → Module**
4. Open `ConvertLatex.bas` copy everything paste it in
5. Press `F5` to run
6. Done — all formulas converted instantly

---

## What it converts

| Before | After |
|---|---|
| `$C = P^e \mod n$` | C = Pᵉ mod n |
| `$\phi(n)$` | φ(n) |
| `$y^2 = x^3 + ax + b$` | y² = x³ + ax + b |
| `$k_a = y^a \mod P$` | kₐ = yᵃ mod P |

---
> ⚠️ **Note:** Please check your document after running. Depending on 
> how the LaTeX was written in your text, some formulas may not convert 
> perfectly and will need to be fixed manually.

## Requirements
- Microsoft Word on **Windows only**
- VBA must be enabled (it is by default)
