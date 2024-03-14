package ExcelGenerator;

// -----( IS Java Code Template v1.2

import com.wm.data.*;
import com.wm.util.Values;
import com.wm.app.b2b.server.Service;
import com.wm.app.b2b.server.ServiceException;
// --- <<IS-START-IMPORTS>> ---
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.w3c.dom.NodeList;
import org.xml.sax.InputSource;
import org.apache.poi.hssf.*;
import javax.xml.parsers.DocumentBuilderFactory;
import java.io.FileOutputStream;
import java.io.StringReader;
import java.io.File;
// --- <<IS-END-IMPORTS>> ---

public final class javaService

{
	// ---( internal utility methods )---

	final static javaService _instance = new javaService();

	static javaService _newInstance() { return new javaService(); }

	static javaService _cast(Object o) { return (javaService)o; }

	// ---( server methods )---




	public static final void demoExcel (IData pipeline)
        throws ServiceException
	{
		// --- <<IS-START(demoExcel)>> ---
		// @sigtype java 3.5
		// [i] recref:0:required TestCaseDocument HG_Excel_Demo.v2.docTypes:TestCaseDocument
		// [i] field:0:required filePath
		// [i] field:0:required fileName
		// [o] field:0:required Exception
		try {
			
			
			
		/*	// Create a new Excel workbook
			HSSFWorkbook workbook=new HSSFWorkbook();
		
			
			
			 * HSSFCellStyle style = workbook.createCellStyle();
			 * style.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex()); //Set
			 * the desired color style.setFillPattern(FillPatternType.THICK_BACKWARD_DIAG);
			 
			 
			
			
					
			//HSSFWorkbook workbook = new HSSFWorkbook();
		    HSSFSheet sheet = workbook.createSheet("Sheet2");
		    
		    // Add a header row
		    HSSFRow headerRow = sheet.createRow(0);
		    headerRow.createCell(0).setCellValue("TestCaseID");
		    headerRow.createCell(1).setCellValue("ExecutionDate");
		    headerRow.createCell(2).setCellValue("TestedBy");
		    headerRow.createCell(3).setCellValue("TestScenario");
		    headerRow.createCell(4).setCellValue("TestCase");
		    headerRow.createCell(5).setCellValue("TestSteps");
		    headerRow.createCell(6).setCellValue("ExpectedResult");
		    headerRow.createCell(7).setCellValue("ActualResult");
		    headerRow.createCell(8).setCellValue("Comments");
		    headerRow.createCell(9).setCellValue("Request");
		    headerRow.createCell(10).setCellValue("Response");
		    headerRow.createCell(11).setCellValue("MessageId");*/
			
			
			
			// Create a new Excel workbook
			HSSFWorkbook workbook = new HSSFWorkbook();
		
			// Create a cell style with your desired color and fill pattern
			HSSFCellStyle style = workbook.createCellStyle();
			style.setFillForegroundColor(IndexedColors.LIGHT_YELLOW.getIndex());
			style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		
			HSSFSheet sheet = workbook.createSheet("Sheet2");
		
			// Add a header row
			HSSFRow headerRow = sheet.createRow(0);
		
			// Create an array to hold the headers
			String[] headers = {"TestCaseID", "ExecutionDate", "TestedBy", "TestScenario", "TestCase", 
			                    "TestSteps", "ExpectedResult", "ActualResult", "Comments", "Request", 
			                    "Response", "MessageId"};
		
			// Add each header to the header row
			for (int i = 0; i < headers.length; i++) {
			    HSSFCell cell = headerRow.createCell(i);
			    cell.setCellValue(headers[i]);
			    cell.setCellStyle(style); // Apply the cell style to each cell
			}
			
		
			// pipeline
			IDataCursor pipelineCursor = pipeline.getCursor();
		
				// TestCaseDocument
				IData	TestCaseDocument = IDataUtil.getIData( pipelineCursor, "TestCaseDocument" );
				if ( TestCaseDocument != null)
				{
					IDataCursor TestCaseDocumentCursor = TestCaseDocument.getCursor();
						String	__version = IDataUtil.getString( TestCaseDocumentCursor, "@version" );
		
						// i.TestCase
						IData	TestCase = IDataUtil.getIData( TestCaseDocumentCursor, "TestCase" );
						if ( TestCase != null)
						{
							IDataCursor TestCaseCursor = TestCase.getCursor();
		
								// i.TestCaseDocument
								IData[]	TestCaseDocument_1 = IDataUtil.getIDataArray( TestCaseCursor, "TestCaseDocument" );
								if ( TestCaseDocument_1 != null)
								{
									for ( int i = 0; i < TestCaseDocument_1.length; i++ )
									{
										IDataCursor TestCaseDocument_1Cursor = TestCaseDocument_1[i].getCursor();
											String	TestCaseID = IDataUtil.getString( TestCaseDocument_1Cursor, "TestCaseID" );
											String	ExecutionDate = IDataUtil.getString( TestCaseDocument_1Cursor, "ExecutionDate" );
											String	TestedBy = IDataUtil.getString( TestCaseDocument_1Cursor, "TestedBy" );
											String	TestScenario = IDataUtil.getString( TestCaseDocument_1Cursor, "TestScenario" );
											String	TestCase_1 = IDataUtil.getString( TestCaseDocument_1Cursor, "TestCase" );
											String	TestSteps = IDataUtil.getString( TestCaseDocument_1Cursor, "TestSteps" );
											String	ExpectedResult = IDataUtil.getString( TestCaseDocument_1Cursor, "ExpectedResult" );
											String	ActualResult = IDataUtil.getString( TestCaseDocument_1Cursor, "ActualResult" );
											String	Comments = IDataUtil.getString( TestCaseDocument_1Cursor, "Comments" );
											String	Request = IDataUtil.getString( TestCaseDocument_1Cursor, "Request" );
											String	Response = IDataUtil.getString( TestCaseDocument_1Cursor, "Response" );
											String	MessageId = IDataUtil.getString( TestCaseDocument_1Cursor, "MessageId" );
										TestCaseDocument_1Cursor.destroy();
										
										//Create rows with run time data
			 							HSSFRow row = sheet.createRow(i + 1);
			 							row.createCell(0).setCellValue(TestCaseID);
								        row.createCell(1).setCellValue(ExecutionDate);
								        row.createCell(2).setCellValue(TestedBy);
								        row.createCell(3).setCellValue(TestScenario);	 
								       row.createCell(4).setCellValue(TestCase_1);	
								      row.createCell(5).setCellValue(TestSteps);	
								     row.createCell(6).setCellValue(ExpectedResult);	
								    row.createCell(7).setCellValue(ActualResult);	
								   row.createCell(8).setCellValue(Comments);	
								   row.createCell(9).setCellValue(Request);
								   row.createCell(10).setCellValue(Response);
								  row.createCell(11).setCellValue(MessageId);
										
										
									}
								}
							TestCaseCursor.destroy();
						}
					TestCaseDocumentCursor.destroy();
				}
				String	filePath = IDataUtil.getString( pipelineCursor, "filePath" );
				String	fileName = IDataUtil.getString( pipelineCursor, "fileName" );
			pipelineCursor.destroy();
		
			// pipeline
			// Create the directory if it doesn't exist
		    File directory = new File(filePath);
		    if (!directory.exists()) {
		        directory.mkdirs();
		    }
		 // Save the Excel file to the specified directory
		    try (FileOutputStream outputStream = new FileOutputStream(filePath + "\\"+fileName+".xls")) {
		        workbook.write(outputStream);
		        workbook.close();
		        System.out.println("Excel file generated successfully.");
		    }
					
			
		} catch (Exception e) {
			e.printStackTrace();
			IDataCursor pipelineCursor_1 = pipeline.getCursor();
			IDataUtil.put( pipelineCursor_1, "Exception", e );
			pipelineCursor_1.destroy();
		}
		// --- <<IS-END>> ---

                
	}
}

