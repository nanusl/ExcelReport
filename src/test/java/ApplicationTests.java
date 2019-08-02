import com.github.liaochong.myexcel.core.ExcelBuilder;
import com.github.liaochong.myexcel.core.FreemarkerExcelBuilder;
import com.github.liaochong.myexcel.core.HtmlToExcelFactory;
import org.apache.poi.ss.usermodel.Workbook;
import org.junit.Test;

import java.io.File;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.HashMap;

/**
 * @author 南来
 * @version V1.0
 * @Description
 * @date 2019-01-13 15:26
 */
public class ApplicationTests {

    @Test
    public void HtmlExportTest() throws Exception {

        File htmlFile = Paths.get("test_20190802.html").toFile();

        Workbook workbook = HtmlToExcelFactory.readHtml(htmlFile).useDefaultStyle().build();

        Path outPath = Paths.get("excel.xlsx");

        workbook.write(Files.newOutputStream(outPath));
    }

    @Test
    public void freeMarketExportTest() throws Exception {

        ExcelBuilder excelBuilder = new FreemarkerExcelBuilder();

        Workbook workbook = excelBuilder.template("/templates/freemarker_template.ftl").build(new HashMap<>());

        Path outPath = Paths.get("freeMarketExport.xlsx");

        workbook.write(Files.newOutputStream(outPath));
    }
}