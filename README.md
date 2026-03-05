<div align="center">
  
# [Investments Recovery Path.](https://github.com/BrenoFariasdaSilva/Investments-Recovery-Path) <img src="https://github.com/BrenoFariasdaSilva/Investments-Recovery-Path/blob/main/.assets/Icons/Investment.png"  width="4%" height="4%">

</div>

<div align="center">
  
---

Analyzes cryptocurrency investment portfolios and calculates optimal recovery strategies for assets with negative returns through proportional budget allocation.
  
---

</div>

<div align="center">

![GitHub Code Size in Bytes](https://img.shields.io/github/languages/code-size/BrenoFariasdaSilva/Investments-Recovery-Path)
![GitHub Commits](https://img.shields.io/github/commit-activity/t/BrenoFariasDaSilva/Investments-Recovery-Path/main)
![GitHub Last Commit](https://img.shields.io/github/last-commit/BrenoFariasdaSilva/Investments-Recovery-Path)
![GitHub Forks](https://img.shields.io/github/forks/BrenoFariasDaSilva/Investments-Recovery-Path)
![GitHub Language Count](https://img.shields.io/github/languages/count/BrenoFariasDaSilva/Investments-Recovery-Path)
![GitHub License](https://img.shields.io/github/license/BrenoFariasdaSilva/Investments-Recovery-Path)
![GitHub Stars](https://img.shields.io/github/stars/BrenoFariasdaSilva/Investments-Recovery-Path)
![GitHub Contributors](https://img.shields.io/github/contributors/BrenoFariasdaSilva/Investments-Recovery-Path)
![GitHub Created At](https://img.shields.io/github/created-at/BrenoFariasdaSilva/Investments-Recovery-Path)
![wakatime](https://wakatime.com/badge/github/BrenoFariasdaSilva/Investments-Recovery-Path.svg)

</div>

<div align="center">
  
![RepoBeats Statistics](https://repobeats.axiom.co/api/embed/7aab7ac179c13d6489e877d918cd86023ba65c7d.svg "Repobeats analytics image")

</div>

## Table of Contents
- [Investments Recovery Path. ](#investments-recovery-path-)
  - [Table of Contents](#table-of-contents)
  - [Introduction](#introduction)
  - [Requirements](#requirements)
  - [Setup](#setup)
    - [Clone the repository](#clone-the-repository)
    - [Python, Pip and Venv](#python-pip-and-venv)
      - [Linux](#linux)
      - [MacOS](#macos)
      - [Windows](#windows)
  - [Run Python Code:](#run-python-code)
    - [Dependencies](#dependencies)
  - [Usage](#usage)
  - [Results](#results)
  - [Contributing](#contributing)
  - [Collaborators](#collaborators)
  - [License](#license)
    - [Apache License 2.0](#apache-license-20)

## Introduction

This Investments Recovery Path Calculator is a Python-based tool that analyzes cryptocurrency investment portfolios from Excel files and calculates optimal recovery strategies for assets with negative returns. The script performs proportional allocation of available budget based on current losses to minimize overall portfolio loss percentage.

**Key Features:**
- Automatic Excel data loading and preprocessing with data cleaning
- Proportional loss-based budget allocation across losing assets
- New loss percentage calculation after hypothetical investment
- Improvement metrics showing expected recovery in percentage points
- Comprehensive output table with investment recommendations
- Detailed logging with timestamps for execution history

## Requirements

- Python >= 3.7
- pandas >= 2.0.0
- numpy >= 1.24.0
- openpyxl >= 3.1.0 (for Excel file reading)
- colorama == 0.4.6 (for terminal coloring)
- Excel file with proper format containing columns: Data, Total Spent - R$, Current Amount - R$, Profit - R$, Profit - %

## Setup

### Clone the repository

1. Clone the repository with the following command:

   ```bash
   git clone https://github.com/BrenoFariasdaSilva/Investments-Recovery-Path.git
   cd Investments-Recovery-Path
   ```

### Python, Pip and Venv

In order to run the scripts, you must have python3, pip and venv installed in your machine. If you don't have it installed, you can use the following commands to install it:

#### Linux

In order to install python3, pip and venv in Linux, you can use the following commands:

```bash
sudo apt install python3 python3-pip python3-venv -y
```

#### MacOS

In order to install python3 and pip in MacOS, you can use the following commands:

```bash
brew install python3
```

#### Windows

In order to install python3 and pip in Windows, you can use the following commands in case you have `choco` installed:

```bash
choco install python3
```

Or just download the installer from the [official website](https://www.python.org/downloads/).

Great, you now have python3 and pip installed. Now, we need to install the additional project requirements. 

## Run Python Code:

```bash
# Run directly with Python:
python3 main.py # Mac and Linux users can use python3
python main.py # Windows users can use python

# Or using Makefile:
make run
```

### Dependencies

1. Install the project dependencies with the following command:

   ```bash
   # Using Makefile:
   make dependencies
   
   # Or manually with pip:
   pip install -r requirements.txt
   ```

## Usage

1. **Configure the script**: Edit the configuration constants in `main.py`:
   - `INPUT_FILE`: Path to your Excel file (default: "./Input/Invested Money.xlsx")
   - `SHEET_NAME`: Name of the Excel sheet to read (default: "CryptoCurrencies")
   - `AVAILABLE_BUDGET`: Available budget for investment recovery in R$ (default: 500.00)
   - `EXCLUDED_CRYPTOS`: List of cryptocurrencies to exclude from calculation (default: ["Bitcoin", "Ethereum", "USDC", "USDT", "Ripple"])
   - `EXCLUDE_POSITIVE_CRYPTOCURRENCIES`: Set to True to exclude cryptocurrencies with positive profit (default: True)

2. **Prepare your Excel file**: Ensure your Excel file has the following columns:
   - Data (CryptoCurrency names)
   - Total Spent - R$
   - Current Amount - R$
   - Profit - R$
   - Profit - %

3. **Run the project**:
   ```bash
   make run
   # or
   python main.py
   ```

4. **View results**: Check the terminal output for investment recommendations and review `Logs/main.log` for detailed execution history.

## Results

The script generates comprehensive investment recovery recommendations including:

- **Filtered Portfolio Analysis**: Identifies all cryptocurrencies with losses (excluding specified coins)
- **Proportional Budget Allocation**: Distributes the available budget proportionally based on loss magnitudes
- **Recovery Projections**: Calculates new loss percentages after hypothetical investment
- **Improvement Metrics**: Shows expected improvement in percentage points for each asset
- **Summary Table**: Displays CryptoCurrency name, current loss (R$), recommended investment amount, old % loss, new % loss, and improvement percentage
- **Total Calculations**: Provides total current losses and total investment allocation
- **Execution Logs**: Detailed logs saved to `Logs/main.log` with timestamps for audit and debugging purposes

Generated output files
- **Excel**: `Output/main_Results.xlsx`
  - Sheet contains columns (left-to-right): `#`, `CryptoCurrency`, `Current Loss (R$)`, `Investments`, `Old % Loss`, `New % Loss`, `Improvement %`.
  - The leading `#` column is a 1-based index (starts at 1).
  - Empty/NA cells are written as a dash (`-`) in the spreadsheet (for clarity on totals or missing percentages).

- **CSV**: `Output/main_Results.csv`
  - CSV containing header (left-to-right): `#`, `CryptoCurrency`, `Current Loss (R$)`, `Investments`, `Old % Loss`, `New % Loss`, `Improvement %`.
  - Standard CSV format: comma (`,`) delimiter, dot (`.`) decimal separator.
  - Encoded as UTF-8 with BOM to improve compatibility with Excel on Windows.
  - Floats are formatted with two decimal places (e.g. `-1295.39`).
  - The first column is `#` (index starting at 1) and the subsequent columns match the Excel sheet.

  Example CSV (excerpt):

  ```csv
  #,CryptoCurrency,Current Loss (R$),Investments,Old % Loss,New % Loss,Improvement %
  1,BTC,-1295.39,250.00,-12.95,-10.52,2.43
  2,XRP,-502.10,250.00,-25.10,-20.00,5.10
  3,TOTAL,-1797.49,500.00,-,-,-
  ```
Notes about table contents
- The table includes a final `TOTAL` row where `CryptoCurrency` is `TOTAL` and numeric totals appear under `Current Loss (R$)` and `Investments`; percentage columns for the total are represented with a dash (`-`).

All monetary values are displayed in Brazilian Real (R$) with proper formatting, and the output is color-coded for easy reading in the terminal.

## Contributing

Contributions are what make the open-source community such an amazing place to learn, inspire, and create. Any contributions you make are **greatly appreciated**. If you have suggestions for improving the code, your insights will be highly welcome.
In order to contribute to this project, please follow the guidelines below or read the [CONTRIBUTING.md](CONTRIBUTING.md) file for more details on how to contribute to this project, as it contains information about the commit standards and the entire pull request process.
Please follow these guidelines to make your contributions smooth and effective:

1. **Set Up Your Environment**: Ensure you've followed the setup instructions in the [Setup](#setup) section to prepare your development environment.

2. **Make Your Changes**:
   - **Create a Branch**: `git checkout -b feature/YourFeatureName`
   - **Implement Your Changes**: Make sure to test your changes thoroughly.
   - **Commit Your Changes**: Use clear commit messages, for example:
     - For new features: `git commit -m "FEAT: Add some AmazingFeature"`
     - For bug fixes: `git commit -m "FIX: Resolve Issue #123"`
     - For documentation: `git commit -m "DOCS: Update README with new instructions"`
     - For refactorings: `git commit -m "REFACTOR: Enhance component for better aspect"`
     - For snapshots: `git commit -m "SNAPSHOT: Temporary commit to save the current state for later reference"`
   - See more about crafting commit messages in the [CONTRIBUTING.md](CONTRIBUTING.md) file.

3. **Submit Your Contribution**:
   - **Push Your Changes**: `git push origin feature/YourFeatureName`
   - **Open a Pull Request (PR)**: Navigate to the repository on GitHub and open a PR with a detailed description of your changes.

4. **Stay Engaged**: Respond to any feedback from the project maintainers and make necessary adjustments to your PR.

5. **Celebrate**: Once your PR is merged, celebrate your contribution to the project!

## Collaborators

We thank the following people who contributed to this project:

<table>
  <tr>
    <td align="center">
      <a href="#" title="defina o titulo do link">
        <img src="https://github.com/BrenoFariasdaSilva.png" width="100px;" alt="My Profile Picture"/><br>
        <sub>
          <b>Breno Farias da Silva</b>
        </sub>
      </a>
    </td>
  </tr>
</table>

## License

### Apache License 2.0

This project is licensed under the [Apache License 2.0](LICENSE). This license permits use, modification, distribution, and sublicense of the code for both private and commercial purposes, provided that the original copyright notice and a disclaimer of warranty are included in all copies or substantial portions of the software. It also requires a clear attribution back to the original author(s) of the repository. For more details, see the [LICENSE](LICENSE) file in this repository.
