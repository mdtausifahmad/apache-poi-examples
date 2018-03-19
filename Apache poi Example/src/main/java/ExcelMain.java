import java.io.IOException;

import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;

public class ExcelMain {

	public static void main(String[] args) throws IOException {

		/*
		 * //Initialize xlsx and create sheet and row FileOutputStream
		 * fileOutputStream = new FileOutputStream(new File("header.xlsx"));
		 * XSSFWorkbook workbook = new XSSFWorkbook(); XSSFSheet xssfSheet =
		 * workbook.createSheet("sheet0"); XSSFRow xssfRow =
		 * xssfSheet.createRow(0); XSSFCell cell = xssfRow.createCell(0);
		 * cell.setCellValue("Test");
		 * 
		 * //Set the Page margins xssfSheet.setMargin(Sheet.LeftMargin, 0.25);
		 * xssfSheet.setMargin(Sheet.RightMargin, 0.25);
		 * xssfSheet.setMargin(Sheet.TopMargin, 0.75);
		 * xssfSheet.setMargin(Sheet.BottomMargin, 0.75);
		 * xssfSheet.setAutobreaks(true); xssfSheet.setFitToPage(true);
		 * xssfSheet.setPrintGridlines(true);
		 * 
		 * //Set the Header and Footer Margins
		 * xssfSheet.setMargin(Sheet.HeaderMargin, 0.25);
		 * xssfSheet.setMargin(Sheet.FooterMargin, 0.25);
		 * 
		 * //Setup print layout settings XSSFPrintSetup layout =
		 * xssfSheet.getPrintSetup(); layout.setLandscape(true);
		 * layout.setFitWidth((short) 1); layout.setFitHeight((short) 0);
		 * layout.setPaperSize(PrintSetup.A4_PAPERSIZE);
		 * layout.setFooterMargin(0.25);
		 * 
		 * XSSFHeaderFooter header = (XSSFHeaderFooter) xssfSheet.getHeader();
		 * header.setCenter(HSSFHeader.font("Calibri", "Bold") +
		 * HSSFHeader.fontSize((short) 35) + "This is Title of page");
		 * header.setRight(HSSFHeader.font("Stencil-Normal", "Bold") +
		 * HSSFHeader.fontSize((short) 15) + "Page Number: " +
		 * HeaderFooter.page());
		 * 
		 * workbook.write(fileOutputStream);
		 */
  
		 Workbook workbook = new HSSFWorkbook() ;
		
		HSSFSheet sheet = (HSSFSheet) workbook.createSheet("ProductInflow Sheet ");

		CellStyle s1 = workbook.createCellStyle();
		Font f1 = workbook.createFont();
		f1.setFontHeightInPoints((short) 18);
		f1.setFontName(HSSFFont.FONT_ARIAL);
		//f1.setBoldweight(HSSFFont.COLOR_NORMAL);
		f1.setBold(true);
		f1.setColor(HSSFColor.DARK_BLUE.index);
		s1.setFont(f1);
		s1.setAlignment(HorizontalAlignment.LEFT);
		// Add these lines
		s1.setFillForegroundColor(IndexedColors.LIGHT_TURQUOISE.getIndex());

		Row r1 = sheet.createRow((short) 1);
		r1.setRowStyle(s1);
		Cell c1 = r1.createCell((short) 1);
		c1.setCellStyle(s1);
		c1.setCellValue(" ADAT WATER SERVICES LIMITED");

		sheet.addMergedRegion(new CellRangeAddress(1, // first row (0-based)
				1, // last row (0-based)
				1, // first column (0-based)
				11 // last column (0-based)
		));

		CellStyle s2 = workbook.createCellStyle();
		Font f2 = workbook.createFont();
		f2.setFontHeightInPoints((short) 8);
		f2.setFontName(HSSFFont.FONT_ARIAL);
		//f2.setBoldweight(HSSFFont.COLOR_NORMAL);
		f2.setBold(false);
		f2.setColor(HSSFColor.BLACK.index);
		s2.setFont(f2);
		s2.setAlignment(HorizontalAlignment.LEFT);

		Row r2 = sheet.createRow((short) 2);
		r2.setRowStyle(s2);
		Cell c2 = r2.createCell((short) 2);
		c2.setCellStyle(s2);
		c2.setCellValue(" BOX AD145. ADABRAKA, ACCRA");

		sheet.addMergedRegion(new CellRangeAddress(2, // first row (0-based)
				2, // last row (0-based)
				2, // first column (0-based)
				9 // last column (0-based)
		));

		CellStyle s3 = workbook.createCellStyle();
		Font f3 = workbook.createFont();
		f3.setFontHeightInPoints((short) 8);
		f3.setFontName(HSSFFont.FONT_ARIAL);
		//f3.setBoldweight(HSSFFont.COLOR_NORMAL);
		f3.setBold(false);
		f3.setColor(HSSFColor.BLACK.index);
		s3.setFont(f3);
		s3.setAlignment(HorizontalAlignment.LEFT);

		Row r3 = sheet.createRow((short) 3);
		r3.setRowStyle(s3);
		Cell c3 = r3.createCell((short) 3);
		c3.setCellStyle(s3);
		c3.setCellValue(" TEL: (024) 3220605, MOBILE : (020) 8130285 ");

		sheet.addMergedRegion(new CellRangeAddress(3, // first row (0-based)
				3, // last row (0-based)
				3, // first column (0-based)
				9 // last column (0-based)
		));

	}

}
