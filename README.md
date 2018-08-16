## Declaration
This is the Developer Manual for BOM Comparison Tool.

- NO COMMERCIAL USE
- NO ACCURACY GUARANTEE

*BOMComparer.exe is developed by Chen Wang, 2018 summer intern in CRDC Cisco. If you want to modify or reuse the source code, first please contact me via chenw.pop@gmail.com. Warm welcome to perfecting the tool.*

## Environment
- **Language:** Python 3.5.5
- **Package Management Tool:** Anaconda 4.5.8, pip 18.0
- **Packages:**
                numpy 1.15.0
                pandas 0.23.3
                xlrd 1.1.0
                openpyxl 2.5.4
                PyQt5 5.11.2
                PyQt5_sip 4.19.12
                PyInstaller 3.4dev0+
- **Operation System:** Virtual Environment on win10, OS X

## Source Code
Source code is slightly different on two OS, in terms of file naming rules, App background image, and position of menu bar.
- **WINDOWS:** WIN/BOMComparer.py
- **OS X:** OSX/BOMComparer.py

## from Python Script to Executable
1. Install PMT Anaconda(https://conda.io/docs/user-guide/install/index.html)
2. Create a virtual environment for Python 3.5.5 with Anaconda(let's name it bom).
3. Activate the Virtual Environment bom.
4. Install PMT pip 18.0. We must manage packages with pip since Anaconda 4.5.8 misunderstand PyQt5's path.
5. Install packages listed in **Environment**->**Packages** with pip.
6. Notice that we will install the latest PyInstaller via command
          pip install https://github.com/pyinstaller/pyinstaller/archive/develop.zip
7. make sure your App run smoothly before packaging it.
8. We must run the PyInstaller command from bom's root path, you can check the root path via
          conda env list
   Enter the dir of source script, for win10 OS, run the following command
          $VENVPATH\Scripts\pyinstaller BOMComparer.py --onefile ^
          --windowed --add-data $IMAGEPATH;. ^
          --path=$VENVPATH
   For OS X, run
          $VENVPATH/bin/pyinstaller BOMComparer.py --onefile \
          --windowed --add-data $IMAGEPATH:. \
          --path=$VENVPATH
9. Double-click to open the Executable App in dist folder. If an error occurs, open it in cmd to debug, see reference in https://pyinstaller.readthedocs.io/en/stable/when-things-go-wrong.html
10. Reach to me(chenw.pop@gmail.com) for help if necessary.
