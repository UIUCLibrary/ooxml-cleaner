# ooxml-cleaner

Massage, normalize, and clean Office Open XML (MS Word `.docx`) files
for import and publication

## Prerequisites

A Bourne shell (such as `bash`) and `xsltproc` from LibXSLT are
prerequisites.

Pull requests to make the XSLT processor more generic are welcome, as
is scripting and packaging for Windows and OS X systems.

## Usage

Run the `cleanup.sh` script with a Word `.docx` file as input:
```
cleanup.sh example.docx
```
This will create `example_out.docx`.  The content should be
indistinguishable, but will be considerably cleaner for systems that
look at the raw XML data to import it.

The `-d` switch leaves the intermediary files in place for debugging
purposes.
