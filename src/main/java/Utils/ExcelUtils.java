package Utils;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelUtils {
	String filePath="/excelFiles/TestData.xlsx";
	 String projectPath=System.getProperty("user.dir");
	 XSSFWorkbook workBook;
	 XSSFSheet sheet;
	
	public ExcelUtils(String fileName)
	{
		try {
		workBook = new XSSFWorkbook(projectPath+filePath);
		sheet=workBook.getSheet(fileName);	
		} catch (Exception e) {
			// TODO Auto-generated catch block
			System.out.println(e.getMessage());
			System.out.println(e.getCause());
			e.printStackTrace();
		}
	}
	 
	 
	 public  int getRowCount() {
		int rowCount = 0;	
		
		try {
			
			rowCount=sheet.getPhysicalNumberOfRows();
			
		} catch (Exception e) {
			// TODO Auto-generated catch block
			System.out.println(e.getMessage());
			System.out.println(e.getCause());
			e.printStackTrace();
		}
		return rowCount; 
	}
	 public  String getCellDataString(int rowNumber,int colNum) {
			 String strValue=null;
			
			try {
				
				strValue=getCellData(sheet.getRow(rowNumber).getCell(colNum));
				
			} catch (Exception e) {
				// TODO Auto-generated catch block
				System.out.println(e.getMessage());
				System.out.println(e.getCause());
				e.printStackTrace();
			}
		 return strValue;
		}
	 
	public String getCellData(Cell cell) {
		CellType cellTypeStr=cell.getCellType();
		 
		
		switch(cellTypeStr) {
		case BOOLEAN:
			return ""+cell.getBooleanCellValue();	
			 
		case STRING:
			return cell.getStringCellValue();	
			 
		case NUMERIC:
			return ""+cell.getNumericCellValue();
		case FORMULA:
            if (DateUtil.isCellDateFormatted(cell)) {
            	return ""+cell.getDateCellValue();
            } else {
            	return ""+cell.getNumericCellValue();
            }
            
        case BLANK:
            System.out.print("");
            break;
        default:
            System.out.print("");
		}
		
		return "";
		
		
	}
	 
	 
	
}
