# LaTeX_renderer_for_Excel
Render inline LaTeX equations in Excel
# 📊 Excel LaTeX Renderer Add-in

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![Excel Version](https://img.shields.io/badge/Excel-2016%2B-green.svg)](https://www.microsoft.com/en-us/microsoft-365/excel)
![Platform](https://img.shields.io/badge/Platform-Windows%20%7C%20Mac-blue.svg)

Transform your Excel spreadsheets with beautiful mathematical notation! This Excel add-in automatically converts LaTeX mathematical expressions into properly formatted Unicode symbols, making your data analysis and scientific documentation more professional and readable.

## ✨ Features

- 🔤 **Greek Letters**: α, β, γ, δ, Φ, Ω, and more
- 🔢 **Mathematical Operators**: ±, ×, ÷, ≤, ≥, ≈, ∞
- 📐 **Set Theory**: ∈, ⊂, ∪, ∩, ∅
- 🔱 **Logic Symbols**: ∀, ∃, ∧, ∨, ¬
- ➡️ **Arrows**: →, ⇒, ↔, ⇔
- 📝 **Mathematical Alphabets**: 𝔸, 𝕭, ℂ, 𝒟 (mathbb, mathcal, mathfrak)
- 🔺 **Superscripts/Subscripts**: x², H₂O, automatic formatting
- ➗ **Fractions**: Converts `\frac{a}{b}` to formatted a/b
- 🎯 **Real-time Rendering**: Converts as you type
- 🔄 **Toggle Support**: Enable/disable rendering on demand

##  Installation

To avoid security warnings when installing the add-in, please follow these steps:

1. **Download** the latest ZIP release from the [Releases](https://github.com/trumpetkern27/LaTeX_renderer_for_Excel/releases).
2. **Extract** all files from the ZIP to a folder on your computer (e.g., `Documents\ExcelAddins`).
3. **Add the folder as a Trusted Location in Excel**:
   - Go to **File > Options > Trust Center > Trust Center Settings > Trusted Locations**.
   - Click **Add new location...** and select the folder where you extracted the add-in.
4. **Load the Add-in**:
   - Open Excel and go to **File > Options > Add-ins**.
   - At the bottom, select **Excel Add-ins** from the dropdown and click **Go**.
   - Click **Browse**, navigate to the extracted folder, and select `LaTeXRenderer.xlam`.
5. **If you see a security warning** when opening the add-in:
   - Right-click the `LaTeXRenderer.xlam` file in File Explorer.
   - Click **Properties**.
   - If there is an **Unblock** checkbox, check it and click **OK**.
   - See the README.txt file if issues persist
6. Restart Excel and enjoy the add-in!

---

If you encounter any issues, please check the [GitHub Issues](https://github.com/trumpetkern27/LaTeX_renderer_for_Excel/issues) page or reach out.

### Basic Usage

Simply type LaTeX expressions between dollar signs and press Enter:

```
$E = mc^2$          →  E = mc²
$\alpha + \beta$    →  α + β  
$x \in \mathbb{R}$  →  x ∈ ℝ
$\sum_{i=1}^n x_i$  →  ∑ᵢ₌₁ⁿ xᵢ
```

## 📖 Some of the many Supported LaTeX Commands


| LaTeX | Unicode | LaTeX | Unicode |
|-------|---------|-------|---------|
| `\alpha` | α | `\Alpha` | Α |
| `\beta` | β | `\Beta` | Β |
| `\gamma` | γ | `\Gamma` | Γ |
| `\delta` | δ | `\Delta` | Δ |
| `\epsilon` | ε | `\Epsilon` | Ε |
| `\lambda` | λ | `\Lambda` | Λ |
| `\mu` | μ | `\Mu` | Μ |
| `\pi` | π | `\Pi` | Π |
| `\sigma` | σ | `\Sigma` | Σ |
| `\phi` | φ | `\Phi` | Φ |
| `\omega` | ω | `\Omega` | Ω |
| `\pm` | ± | Plus-minus |
| `\times` | × | Multiplication |
| `\div` | ÷ | Division |
| `\cdot` | · | Center dot |
| `\leq` | ≤ | Less than or equal |
| `\geq` | ≥ | Greater than or equal |
| `\neq` | ≠ | Not equal |
| `\approx` | ≈ | Approximately equal |
| `\equiv` | ≡ | Equivalent |
| `\propto` | ∝ | Proportional to |
| `\in` | ∈ | Element of |
| `\notin` | ∉ | Not element of |
| `\subset` | ⊂ | Subset |
| `\supset` | ⊃ | Superset |
| `\cup` | ∪ | Union |
| `\cap` | ∩ | Intersection |
| `\emptyset` | ∅ | Empty set |
| `\forall` | ∀ | For all |
| `\exists` | ∃ | There exists |
| `\neg` | ¬ | Negation |
...


##  Examples

### Scientific Notation
```
$6.022 \times 10^{23}$ → 6.022 × 10²³
$\Delta H = -285.8 \text{ kJ/mol}$ → ΔH = -285.8 kJ/mol
```

### Mathematical Formulas
```
$\int_0^\infty e^{-x} dx = 1$ → ∫₀^∞ e⁻ˣ dx = 1
$\sum_{n=1}^\infty \frac{1}{n^2} = \frac{\pi^2}{6}$ → ∑ₙ₌₁^∞ 1/n² = π²/6
```

### Set Theory
```
$\mathbb{N} \subset \mathbb{Z} \subset \mathbb{Q} \subset \mathbb{R} \subset \mathbb{C}$
→ ℕ ⊂ ℤ ⊂ ℚ ⊂ ℝ ⊂ ℂ
```

## 🛠️ Development

### Prerequisites
- Excel 2016 or later
- VBA development environment
- Git for version control

### Building from Source

1. **Clone the repository**:
```bash
git clone https://github.com/trumpetkern27/LaTeX_renderer_for_Excel.git
cd LaTeX_renderer_for_excel
```

2. **Open Excel** and import the VBA modules from `src/`
3. **Test the functionality**
4. **Export as .xlam** add-in file

### Project Structure
```
├── src/
│   ├── LaTeXRenderer.bas      # Core rendering logic
│   ├── ThisWorkbook.cls       # Add-in initialization
|   ├── Ribbon.bas             # Add-in Ribbon
│   └── Global.bas             # Global variables/settings
├── installation.md            # Setup guide
└── LaTeXRenderer.xlam         # Compiled add-in
```

## 🤝 Contributing

Contributions are welcome! Here's how you can help:

1. **🐛 Report Bugs**: Open an issue with detailed reproduction steps
2. **💡 Feature Requests**: Suggest new LaTeX symbols or functionality
3. **📝 Documentation**: Improve installation guides or examples
4. **🔧 Code Contributions**: 
   - Fork the repository
   - Create a feature branch
   - Submit a pull request

### Development Guidelines
- Follow VBA naming conventions
- Add comments for complex logic
- Include test cases for new features
- Update documentation for API changes

## 📋 Roadmap

- [ ] **Extended Symbol Support**: More mathematical operators and symbols
- [ ] **Custom Fonts**: Integration with mathematical font packages
- [ ] **Performance Optimization**: Faster rendering for large ranges
- [ ] **Formula Recognition**: Auto-detect mathematical expressions

## ⚠️ Known Limitations

- **Font Dependency**: Requires Unicode-compatible fonts
- **Complex Expressions**: Some advanced LaTeX features not supported, such as \frac and nested functions
- **Performance**: May slow down with large worksheets
- **Compatibility**: Designed for Excel 2016+

## 📄 License

This project is licensed under the MIT License - see the LICENSE file for details.

## 🙏 Acknowledgments

- I thank my Creator for enabling me to create such a thing
- I often will say that Excel is the pinnacle of man's creation
- It is shocking how there has not been a way to render LaTeX code in Excel up until now

## 📞 Support

- 📧 **Email**: trumpetkern@gmail.com
- 🐛 **Issues**: [GitHub Issues](../../issues)
- 💬 **Discussions**: [GitHub Discussions](../../discussions)

---

**⭐ Star this repository if you find it useful!**
