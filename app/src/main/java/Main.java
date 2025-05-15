import com.igalia.ooxmlexporter.DocumentConverter;
import com.igalia.ooxmlexporter.DocumentConverterFactory;

public class Main {
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

        DocumentConverter converter = DocumentConverterFactory.getConverterForDocument(inputFile, outputFile);
        if (converter != null) {
            converter.convert();
        } else {
            System.out.println("File type unsupported.");
            System.exit(0);
        }
    }
}
