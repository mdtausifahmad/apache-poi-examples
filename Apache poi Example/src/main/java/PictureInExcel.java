import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;

import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Drawing;
import org.apache.poi.ss.usermodel.Picture;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.util.IOUtils;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class PictureInExcel {

	public static void main(String[] args) throws IOException {
		
		 //create a new workbook
	    Workbook wb = new XSSFWorkbook(); //or new HSSFWorkbook();

	    //add picture data to this workbook.
	    InputStream is = new FileInputStream("tausif.jpg");
	    byte[] bytes = IOUtils.toByteArray(is);
	    int pictureIdx = wb.addPicture(bytes, Workbook.PICTURE_TYPE_JPEG);
	    is.close();

	    CreationHelper helper = wb.getCreationHelper();

	    //create sheet
	    Sheet sheet = wb.createSheet();

	    // Create the drawing patriarch.  This is the top level container for all shapes. 
	    Drawing drawing = sheet.createDrawingPatriarch();

	    //add a picture shape
	    ClientAnchor anchor = helper.createClientAnchor();
	    anchor.setAnchorType(ClientAnchor.AnchorType.MOVE_AND_RESIZE);
	    //set top-left corner of the picture,
	    //subsequent call of Picture#resize() will operate relative to it
	    anchor.setCol1(4);
	    anchor.setRow1(4);
	    
	    anchor.setCol2(8);
	    anchor.setRow2(8);
	    
	    Picture pict = drawing.createPicture(anchor, pictureIdx);

	    //auto-size picture relative to its top-left corner
	    pict.resize();

	    //save workbook
	    String file = "picture.xls";
	    if(wb instanceof XSSFWorkbook) file += "x";
	    FileOutputStream fileOut = new FileOutputStream(file);
	    wb.write(fileOut);
	    fileOut.close();
	        

	}

}
