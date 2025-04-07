import java.io.IOException;

public class Main {
    public static void main (String[] args) throws IOException {
        if (args.length == 0) {
            System.out.println("Input file required as argument.");
            System.exit(0);
        }
        String inputFile = args[0];

        if (inputFile.endsWith(".pptx")) {
            PptxConverter.convertPptx(inputFile, inputFile + ".pdf");
        } else if (inputFile.endsWith(".docx")) {
            DocxConverter.convertDocx(inputFile, inputFile + ".pdf");
        } else {
            System.out.println("File type unsupported.");
            System.exit(0);
        }
    }
}
