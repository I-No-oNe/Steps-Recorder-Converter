# 🚀 Steps Recorder Converter

Convert your Windows Steps Recorder output into a fully editable PowerPoint presentation!

---

## 📂 Project Structure

```yml
StepsRecorderConverter/
├── LICENSE                  # Project license file (MIT)
├── .gitignore               # Specifies intentionally untracked files to ignore
├── converter.py             # Main Python script for conversion
├── README.md                # Project documentation file
├── requirements.txt         # Python dependencies list
├── runRecorder.bat          # Batch file to quickly launch Windows Steps Recorder
````

---

## ✨ Features

- Automatically finds your **latest** Steps Recorder `.zip` file in the Downloads folder  
- Extracts the embedded `.mht` file and all screenshots inside  
- Generates a PowerPoint presentation for every image in the `.mht` file. 

---

## 🎬 How to Use

### 1. Record Your Steps

* Run `runRecorder.bat` to open Windows Steps Recorder quickly.
* Click **Start Record**, perform your actions, then **Stop Record**.
* Save the recording as a `.zip` **IN THE DOWNLOADS FOLDER**.

### 2. Convert to PowerPoint Presentation

* Run the converter script:

```bash
python converter.py
```

* The script will find your latest `.zip` file, extract images, and create a PowerPoint file with timestamped filename.

### 3. Open and Edit

* Find the generated `.pptx` in your project folder.
* Open it with PowerPoint or compatible software and edit/present!
---
*Happy documenting your workflow with ease!* 🎉
