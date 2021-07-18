package br.com.local.helper;

import java.io.File;
import java.util.Date;
import java.util.Locale;

import jxl.Workbook;
import jxl.WorkbookSettings;
import jxl.write.DateFormats;
import jxl.write.DateTime;
import jxl.write.Label;
import jxl.write.WritableCellFormat;
import jxl.write.WritableFont;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;

public class CreateExcelFile {

	public static void main(String[] args) {

		try {
			
			String filename = "entrada.xls";
			
			WorkbookSettings ws = new WorkbookSettings();
			ws.setLocale(new Locale("en", "EN"));
			WritableWorkbook workbook = Workbook.createWorkbook(new File(filename), ws);
			WritableSheet s = workbook.createSheet("Folha1", 0);
			WritableSheet s1 = workbook.createSheet("Folha1", 0);
			
			writeDataSheet(s);
			
			
		} catch (Exception e) {
			// TODO: handle exception
		}

	}

	private static void writeDataSheet(WritableSheet s) throws WriteException {
		
		WritableFont wf = new WritableFont(WritableFont.ARIAL, 10, WritableFont.BOLD);
		WritableCellFormat cf = new WritableCellFormat(wf);
		
		cf.setWrap(true);
		
		Label l = new Label(0, 0, "Data", cf);
		s.addCell(l);
		
		WritableCellFormat cf1 = new WritableCellFormat(DateFormats.FORMAT9);
		DateTime dt = new DateTime(0, 1, new Date(), cf1, DateTime.GMT);
		
		s.addCell(dt);
		
		/* Cria um label e escreve um float numver em uma célula da folha*/
		
		
	}

}
