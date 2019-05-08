package framework;

import org.testng.annotations.Test;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
//import org.apache.poi.hslf.model.Sheet;
//import org.apache.poi.hssf.model.Workbook;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.testng.annotations.DataProvider;
import org.testng.annotations.NoInjection;
import org.testng.annotations.Parameters;
import org.testng.annotations.Test;

//import com.sun.corba.se.spi.orbutil.threadpool.Work;

import org.testng.Assert;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.testng.annotations.AfterTest;
import org.testng.annotations.BeforeClass;
import java.awt.AWTException;
import java.awt.Robot;
import java.awt.event.KeyEvent;
import java.util.ArrayList;
import java.util.concurrent.TimeUnit;

public class ExcelToDataProvider {

    public static HSSFWorkbook workbook;
    public static HSSFSheet worksheet;
    public static DataFormatter formatter= new DataFormatter();
    public static String file_location = System.getProperty("user.dir")+"/Desktop";
    static String SheetName= "Sheet1";
    public  String Res;
   // Write obj1=new Write();
    public int DataSet=-1;

    public String ColName="RESULT";
    public int col_num;
 
	
	@DataProvider
    public static Object[][] ReadVariant() throws IOException
    {
    FileInputStream fileInputStream= new FileInputStream("C:\\Users\\Crest\\Desktop\\testdata.xls"); //Excel sheet file location get mentioned here
        workbook = new HSSFWorkbook (fileInputStream); //get my workbook 
        worksheet=workbook.getSheet(SheetName);// get my sheet from workbook
        HSSFRow Row=worksheet.getRow(0);     //get my Row which start from 0   
     
        int RowNum = worksheet.getPhysicalNumberOfRows();// count my number of Rows
        int ColNum= Row.getLastCellNum(); // get last ColNum 
         
        Object Data[][]= new Object[RowNum-1][ColNum]; // pass my  count data in array
         
            for(int i=0; i<RowNum-1; i++) //Loop work for Rows
            {  
                HSSFRow row= worksheet.getRow(i+1);
                 
                for (int j=0; j<ColNum; j++) //Loop work for colNum
                {
                    if(row==null)
                        Data[i][j]= "";
                    else
                    {
                        HSSFCell cell= row.getCell(j);
                        if(cell==null)
                            Data[i][j]= ""; //if it get Null value it pass no data 
                        else
                        {
                            String value=formatter.formatCellValue(cell);
                            Data[i][j]=value; //This formatter get my all values as string i.e integer, float all type data value
                        }
                    }
                }
            }
 
        return Data;
    }
	//*******************************REading from excel ******************************//
	@Test  //Test method
	(dataProvider="ReadVariant",priority=1) //It get values from ReadVariant function method
	 
	//Here my all parameters from excel sheet are mentioned.
	public void AddVariants(String NAME, String DESCRIPTION, String WEIGHT, String PRICE, String MODEL, String RS) throws Exception
	{
	//Data will set in Excel sheet once one parameter will get from excel sheet to Respective locator position.
	DataSet++;
	System.out.println("NAme of product available are:" +NAME);
	System.out.println("Weight for products are:" +DESCRIPTION);
	System.out.println("Volume of product are:" +WEIGHT);
	System.out.println("Description quotation are:" +PRICE);
	System.out.println("Description picklings are:" +MODEL);
	 
	}
	
	//*************************************Write in result column *****************************//
//Check other classs - Write.java
	
	@Parameters({ "Ress", "DR" })
	@Test //Test method
public void WriteResult(String Ress, int DR) throws Exception
{
	 System.out.println("5245545435543");
//	FileOutputStream out = new FileOutputStream(new File("C:\\Users\\Crest\\Desktop\\testdata.xls"));
    FileInputStream file_input_stream= new FileInputStream("C:\\Users\\Crest\\Desktop\\testdata.xls");
    workbook=new HSSFWorkbook(file_input_stream);
    worksheet=workbook.getSheet(ExcelToDataProvider.SheetName);
    HSSFRow Row=worksheet.getRow(0);

    int sheetIndex=workbook.getSheetIndex(ExcelToDataProvider.SheetName);
    DataFormatter formatter = new DataFormatter();
    if(sheetIndex==-1)
    {
        System.out.println("No such sheet in file exists");
    } else      {
        col_num=-1;
            for(int i=0;i<Row.getLastCellNum();i++)
            {
                HSSFCell cols=Row.getCell(i);
                String colsval=formatter.formatCellValue(cols);
                if(colsval.trim().equalsIgnoreCase(ColName.trim()))
                {
                    col_num=i;
                    break;
                }
            }
//          
            Row= worksheet.getRow(DR);
            try
                {
                //get my Row which is equal to Data  Result and that colNum
                    HSSFCell cell=worksheet.getRow(DR).getCell(col_num);
                    // if no cell found then it create cell
                    if(cell==null) {
                        cell=Row.createCell(col_num);                           
                    }
                    //Set Result is pass in that cell number
                    cell.setCellValue(Ress);
                                     
                     
                }
            catch (Exception e) //add exception in result colummn in failed 
            
            {
                System.out.println(e.getMessage()); 
                HSSFCell cell=worksheet.getRow(DR).getCell(col_num);
                // if no cell found then it create cell
                if(cell==null) {
                    cell=Row.createCell(col_num);                           
                }
                //Set Result is pass in that cell number
                cell.setCellValue(e.getMessage());
            } 

    }
        FileOutputStream file_output_stream=new FileOutputStream("C:\\Users\\Crest\\Desktop\\testdata.xls");
        workbook.write(file_output_stream);
        file_output_stream.close();
        if(col_num==-1) {
            System.out.println("Column you are searching for does not exist");
        }
}


}