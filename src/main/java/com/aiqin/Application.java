package com.aiqin;

import com.github.liaochong.myexcel.core.HtmlToExcelFactory;
import org.apache.poi.ss.usermodel.Workbook;

import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;

/**
 * @author 南来
 * @version V1.0
 * @Description
 * @date 2019-01-13 15:25
 */
public class Application {

    public static void main(String[] args) {
        String htmlFilePath = "in.html";
        String outFilePath = "out.xlsx";
        int argsLen = args.length;

        try {

            if (argsLen > 0) {
                if (args[0].contains("--h") || args[0].contains("--help")) {
                    System.out.println("How to use it :");
                    System.out.println("    ExcelReport.jar [html_file:in.html] [out_file:out.xlsx]");
                    System.out.println("example： ");
                    System.out.println("    java -jar ExcelReport.jar in.html out.xlsx");
                    return;
                }

                if (argsLen >= 2) {
                    htmlFilePath = args[0];
                    outFilePath = args[1];
                }
            }

            Path htmlPath = Paths.get(htmlFilePath);

            Workbook workbook = HtmlToExcelFactory.readHtml(htmlPath.toFile()).build();

            Path outPath = Paths.get(outFilePath);

            workbook.write(Files.newOutputStream(outPath));
            System.out.println("export complete！");
            System.out.println(outPath.toAbsolutePath());
        } catch (Exception e) {
            if (!(e instanceof ArrayIndexOutOfBoundsException)) {
                e.printStackTrace();
            }
            System.out.println("you can use --help(--h) to get more information!");
        }
    }
}
