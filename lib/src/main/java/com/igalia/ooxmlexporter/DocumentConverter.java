package com.igalia.ooxmlexporter;

public interface DocumentConverter {
    public void convert(String outputFile);

    public String getDefaultExtension();
}
