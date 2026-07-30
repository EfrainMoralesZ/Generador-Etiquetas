@echo off
setlocal

REM Build del ejecutable de Windows para Generador de Etiquetas
python -m PyInstaller --version >nul 2>&1
if errorlevel 1 (
  echo PyInstaller no esta instalado en este entorno. Instalando...
  python -m pip install pyinstaller
  if errorlevel 1 (
    echo.
    echo No se pudo instalar PyInstaller.
    exit /b 1
  )
)

REM Para agregar un icono mas adelante: coloca el .ico en img\icono.ico y
REM agrega la linea "  --icon "img\icono.ico" ^" a la llamada de abajo.

python -m PyInstaller ^
  --noconfirm ^
  --clean ^
  --windowed ^
  --onefile ^
  --name GeneradorEtiquetas ^
  --icon "img\icono.ico" ^
  --hidden-import openpyxl ^
  --hidden-import xlrd ^
  --hidden-import PIL._tkinter_finder ^
  --collect-all customtkinter ^
  --collect-all tkinterdnd2 ^
  --collect-all reportlab ^
  --collect-all pillow ^
  --collect-all pymupdf ^
  app.py

if errorlevel 1 (
  echo.
  echo Build failed.
  exit /b 1
)

REM app.py / armadoEtiqueta.py leen y crean "data\config_etiquetas.json"
REM (y su historial en data\estado_app.json, data\lotes, data\etiquetas)
REM con rutas relativas a la carpeta donde se ejecuta el .exe, asi que la
REM configuracion base debe quedar junto al ejecutable generado.
if not exist "dist\data" mkdir "dist\data"
copy /Y "data\config_etiquetas.json" "dist\data\config_etiquetas.json" >nul

echo.
echo Build completado. Ejecutable en:
echo dist\GeneradorEtiquetas.exe
echo (junto con dist\data\config_etiquetas.json, necesario para que arranque)
