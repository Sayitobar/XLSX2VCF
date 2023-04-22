package sayitobar;

import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.IOException;
import java.io.PrintStream;
import java.nio.charset.StandardCharsets;
import java.nio.file.Paths;
import org.apache.poi.openxml4j.exceptions.OpenXML4JException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.openxml4j.opc.PackageAccess;
import org.xml.sax.SAXException;


public class Example {
    public static String convert_XLSX_to_VCF(String XLSX_path, String newFilePath, boolean temporary) throws IOException, OpenXML4JException, SAXException {
        long startTime = System.nanoTime();

        // create new VCF file
        File f;
        if (temporary) {
            f = File.createTempFile(
                    "TEMP_" + Paths.get(newFilePath).getFileName().toString().split("\\.")[0] + "_",
                    "." + Paths.get(newFilePath).getFileName().toString().split("\\.")[1]
            );
            f.deleteOnExit();  // işi bitince otomatik silinir
        }
        else
            f = new File(newFilePath);

        
        // The package open is instantaneous, as it should be.
        OPCPackage p = OPCPackage.open(XLSX_path, PackageAccess.READ);

        final ByteArrayOutputStream baos = new ByteArrayOutputStream();

        // read xlsx
        try (PrintStream ps = new PrintStream(baos, true, StandardCharsets.UTF_8.name())) {
            XLSX2VCF xlsx2vcf = new XLSX2VCF(p, ps);
            xlsx2vcf.process(0);
        } finally {
            p.revert();
        }

        // get string data
        String data = baos.toString(StandardCharsets.UTF_8.name());
        baos.close();
        data = data.replaceAll("\\s+$", "");  // := stripTrailing() - just remove redundant \n's

        // And then, save
        PrintStream ps = new PrintStream(f.getPath());
        ps.print(data);
        ps.close();

        ps.close();
        System.out.println("XLSX -> VCF conversion accomplished:\t" + XLSX_path + "    -->    " + f.getPath() + " \t(in " + (System.nanoTime()-startTime)/1000000000.0 + " secs)");

        return f.getPath();
    }
    
    public static void main(String[] args) throws IOException, OpenXML4JException, SAXException {
        convert_XLSX_to_VCF("C:\\Users\\USERNAME\\MyFile.xlsx", "C:\\Users\\USERNAME\\MyNewFile.vcf", false);
        // xlsx file location, desired vcf file location, not temporary file
    }
}
