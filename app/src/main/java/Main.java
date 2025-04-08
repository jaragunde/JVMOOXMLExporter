import com.igalia.ooxmlexporter.DocumentConverter;
import com.igalia.ooxmlexporter.DocumentConverterFactory;

public class Main {
    public static void main (String[] args) {
        if (args.length == 0) {
            System.out.println("Input file required as argument.");
            System.exit(0);
        }
        String inputFile = args[0];

        DocumentConverter converter = DocumentConverterFactory.getConverterForDocument(inputFile, inputFile + ".pdf");
        if (converter != null) {
            converter.convert();
        } else {
            System.out.println("File type unsupported.");
            System.exit(0);
        }
    }
}
