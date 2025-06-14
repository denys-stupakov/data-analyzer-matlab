# 📊 MATLAB Data Toolkit – SVP Menu System

This project is a multi-functional MATLAB application developed for academic purposes. It allows users to process statistical data, perform matrix operations, and generate plots based on user input. All outputs are saved to Excel or `.txt` files.

## 📂 Structure

- `SVP_Menu.m`: Main GUI menu for navigation
- `Štatistika`: Loads Excel data and analyzes cities & villages based on population, temperature, and altitude
- `Matice`: Generates random matrix A, computes B = A·Aᵗ, performs operations like determinant, rank, inverse
- `Grafy`: Plots a quadratic function and geometric series, calculates sequence sums

## 📁 Files & Folders

- `DataInput/`
  - `SVP-Statistika.xlsx`
  - `Matice.txt`
- `DataOutput/`
  - `MaticeVysledky.txt`
  - `VystupPostupnisti.txt`
- `SVP_Menu.m`
- `Other .m files` (functions used by the modules)

## 🔧 Technologies

- MATLAB
- Excel integration
- File I/O
- GUI (menu-based)
- Data visualization

## 📷 Screenshots

<img src="screenshots/menu.png" width="500"/>
<img src="screenshots/statistics.png" width="500"/>
<img src="screenshots/graph.png" width="500"/>

## 💡 Features

- Interactive menu with 3 major data modules
- Excel file read/write and formatting
- Matrix algebra and linear algebra tools
- Graphical visualization of functions and sequences
- Statistical analysis with summary output

## 🚀 How to Run

1. Open MATLAB
2. Run the `SVP_Menu.m` file
3. Select from the menu: Štatistika, Matice, or Grafy
