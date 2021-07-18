import static org.junit.Assert.assertEquals;

import java.io.File;
import java.io.IOException;
import java.util.List;
import java.util.Map;

import org.junit.Before;
import org.junit.Test;
import org.junit.runner.RunWith;
import org.mockito.junit.MockitoJUnitRunner;

import br.com.local.helper.JExcelHelper;
import jxl.read.biff.BiffException;
import jxl.write.WriteException;

@RunWith(MockitoJUnitRunner.class)
public class JExcelIntegrationTest {
	
	private JExcelHelper jexcelHelper;
	private static String FILE_NAME = "temp.xls";
	private String fileLocation;
	
	@Before
    public void generateExcelFile() throws IOException, WriteException {

        File currDir = new File(".");
        String path = currDir.getAbsolutePath();
        fileLocation = path.substring(0, path.length() - 1) + FILE_NAME;

        jexcelHelper = new JExcelHelper();
        jexcelHelper.writeJExcel();

    }
	
	@Test
	public void whenParsingJExelFile_thenCorrect() throws IOException, BiffException {
		
		Map<Integer, List<String>> data = jexcelHelper.readExcel(fileLocation);
		
		assertEquals("Name", data.get(0).get(0));
		assertEquals("Age", data.get(0).get(1));
		assertEquals("John Smith", data.get(1).get(0));
		assertEquals("20", data.get(1).get(1));
	}
	
	public void cleanUp() {
		File file = new File(fileLocation);
		if(file.exists()) {
			file.delete();
		}
	}
 
}
