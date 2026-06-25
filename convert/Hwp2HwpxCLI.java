import kr.dogfoot.hwplib.object.HWPFile;
import kr.dogfoot.hwplib.reader.HWPReader;
import kr.dogfoot.hwpxlib.object.HWPXFile;
import kr.dogfoot.hwpxlib.writer.HWPXWriter;
import kr.dogfoot.hwp2hwpx.Hwp2Hwpx;

public class Hwp2HwpxCLI {
    public static void main(String[] args) throws Exception {
        if (args.length < 1) {
            System.out.println("Usage: Hwp2HwpxCLI <input.hwp> [output.hwpx]");
            System.exit(1);
        }
        String input = args[0];
        String output = args.length > 1 ? args[1] : input.replaceAll("\\.hwp$", ".hwpx");

        System.out.println("Input:  " + input);
        System.out.println("Output: " + output);

        HWPFile hwpFile = HWPReader.fromFile(input);
        HWPXFile hwpxFile = Hwp2Hwpx.toHWPX(hwpFile);
        HWPXWriter.toFilepath(hwpxFile, output);

        System.out.println("Done!");
    }
}
