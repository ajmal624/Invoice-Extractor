# Invoice-Extractor

Unzip the File
use internet for all installation

Install Python 3.14, then verify the installation:
    Press Windows + R, type `cmd`, and press Enter.
    In Command Prompt, run:
        python --version

Open Visual Studio Code and install the required Python extensions.

Go to File → New Window → Open Folder, then select the `Invoice Extractor` folder.

Open the terminal and run:
    cd Main
    py -m venv venv
    venv\Scripts\Activate
    pip install -r requirements.txt

Run the application:
    streamlit run app2.py

On the output page, Select PDF using Upload Files, Select Excel File using Upload Excel, and Generate the Output.

## ⚠️ Important Files – Security Limitation

Some important configuration files used by this project
could not be uploaded to this GitHub repository because they contain
sensitive API credentials.

These files are required for the complete project and are kept separately.

### Important files not included in this repository

- `.env`

The `.env` file contains the Gemini API key required by the application
to access the Gemini API for invoice data extraction.

The file is **not deleted or unnecessary**; it is excluded from GitHub
only for security reasons to prevent exposing sensitive API credentials.

The application code (`app2.py`) and dependency file (`requirements.txt`)
are included in this repository.

To run the complete project, create a `.env` file locally and add your
own Gemini API key.

> **Note:** Never share or upload your API key publicly.
