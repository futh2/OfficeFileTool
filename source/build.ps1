$ErrorActionPreference = 'Stop'
$buildRoot = $PSScriptRoot
python -m pip install pyinstaller psutil pywin32
if ($LASTEXITCODE -ne 0) { throw 'Dependency installation failed.' }
python -m unittest discover -s $buildRoot -p 'test_office_tool.py' -v
if ($LASTEXITCODE -ne 0) { throw 'Tests failed.' }
python -m PyInstaller --noconfirm --onefile --windowed --name OfficeFileTool --distpath (Join-Path $buildRoot '..') --workpath (Join-Path $buildRoot 'build') --specpath $buildRoot (Join-Path $buildRoot 'office_tool.py')
if ($LASTEXITCODE -ne 0) { throw 'Build failed.' }
