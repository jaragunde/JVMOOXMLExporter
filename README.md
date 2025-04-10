# OOXML document exporter to PDF

This project can export OOXML documents into PDF documents. It's built using Java technologies.

It currently only supports PowerPoint (pptx) documents.

## Usage

Convert a document from the command line, writing the output to the designated location:
```
./gradlew run --args="/path/to/input/document.pptx /path/to/output/document.pdf"
```
If the output document is ommitted, it will output into the same location with the `pdf` suffix:
```
./gradlew run --args="/path/to/input/document.pptx"
```

## Structure

It's split in a command-line application (`app`) and a library (`lib`).
