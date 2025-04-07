import java.io.IOException;

public class Main {
    public static void main (String[] args) {
        if (args.length == 0) {
            System.out.println("Input file required as argument.");
            System.exit(0);
        }
        String inputFile = args[0];

        if (inputFile.endsWith(".pptx")) {
            DocumentConverter converter = new PptxConverter(inputFile, inputFile + ".pdf");
            converter.convert();
        } else if (inputFile.endsWith(".docx")) {
            DocumentConverter converter = new DocxConverter(inputFile, inputFile + ".pdf");
            converter.convert();
        } else {
            System.out.println("File type unsupported.");
            System.exit(0);
        }
    }
}
