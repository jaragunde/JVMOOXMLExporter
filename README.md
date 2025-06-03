# OOXML document exporter

This project can export OOXML documents into PDF or HTML documents. It's built using Java technologies.

It currently supports:
* PowerPoint (pptx) to PDF.
* Word (docx) to HTML.

There is experimental support for:
* Word to PDF.
* Excel to HTML.

## Usage

Convert a document from the command line, writing the output to the designated location:
```
./gradlew run --args="/path/to/input/document.pptx /path/to/output/document.pdf"
```
If the output document is ommitted, it will output into the same location with the `pdf` or `html` suffix:
```
./gradlew run --args="/path/to/input/document.pptx"
```

## Structure

It's split in a command-line application (`app`) and a library (`lib`).
