import com.igalia.ooxmlexporter.DocumentConverter;
import com.igalia.ooxmlexporter.DocumentConverterFactory;
import com.igalia.ooxmlexporter.PptxConverter;

public class Main {

    public static DocumentConverterFactory.DocumentFormat getDefaultFormat(String inputFile, String outputFile) {
        if (outputFile.toLowerCase().endsWith("html")) {
            return DocumentConverterFactory.DocumentFormat.HTML;
        } else if (outputFile.toLowerCase().endsWith("pdf")) {
            return DocumentConverterFactory.DocumentFormat.PDF;
        } else if (inputFile.toLowerCase().endsWith(".pptx")) {
            // For PowerPoint we only have a PDF converter so we default to it
            return DocumentConverterFactory.DocumentFormat.PDF;
        } else {
            // Otherwise, default to HTML which is generally more mature
            return DocumentConverterFactory.DocumentFormat.HTML;
        }
    }

    public static void main (String[] args) {
        if (args.length == 0) {
            System.out.println("""
                    Missing input file.
                    Arguments:
                      input-file   mandatory, path to file to be converted.
                      output-file  optional, defaults to input file with the corresponding output extension (\".pdf\" or \".html\").""");
            System.exit(0);
        }
        String inputFile = args[0];
        String outputFile = inputFile; // extension will be added by the DocumentConverter
        if (args.length > 1) {
            outputFile = args[1];
        }

        DocumentConverterFactory.DocumentFormat outputFormat = getDefaultFormat(inputFile, outputFile);
        DocumentConverter converter = DocumentConverterFactory.getConverterForDocument(inputFile, outputFormat);
        if (converter != null) {
            converter.convert(outputFile);
        } else {
            System.out.println("File type unsupported.");
            System.exit(0);
        }
    }
}
