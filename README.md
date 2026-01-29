# gslidestoppt

Folder structure (example)
```
gslidestoppt/
├── venv/
├── slides.py
└── pdfs/
    ├── deck1.pdf
    ├── deck2.pdf
    └── deck3.pdf
```

What this script does

Loops through all .pdf files in a folder

For each PDF:

Converts pages → images

Creates <pdf_name>_locked.pptx


How to run
1. Go to your project folder

```cd /Users/<user>/gslidestoppt```

2. Create a virtual environment (using your Homebrew Python)

```python3 -m venv venv```




3. Activate the virtual environment

```source venv/bin/activate```


4. Install required packages inside the venv
   
```pip install pdf2image python-pptx pillow```

5. Install Poppler (system dependency) macos
```brew install poppler```
Verify:
```pdftoppm -h```

Windows

Download Poppler from:
```https://github.com/oschwartz10612/poppler-windows/releases/```

Extract it (e.g. C:\poppler)

Add C:\poppler\Library\bin to PATH

Restart terminal

7.Run your script (inside venv)
```python3 slides.py```




Output
```
output_ppt/
├── deck1_locked.pptx
├── deck2_locked.pptx
└── deck3_locked.pptx
```

Each PPT:
✔️ Opens cleanly in PowerPoint
✔️ Zero layout drift
✔️ Effectively read-only
