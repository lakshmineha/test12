package test;

public class sub {


import java.io.FileInputStream;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
    public class ExcelUtils {
                private static XSSFSheet ExcelWSheet;
                private static XSSFWorkbook ExcelWBook;
                private static XSSFCell Cell;

            //This method is to set the File path and to open the Excel file
            //Pass Excel Path and SheetName as Arguments to this method
            public static void setExcelFile(String Path,String SheetName) throws Exception {
	                FileInputStream ExcelFile = new FileInputStream(Path);
	                ExcelWBook = new XSSFWorkbook(ExcelFile);
	                ExcelWSheet = ExcelWBook.getSheet(SheetName);
                   }

            //This method is to read the test data from the Excel cell
            //In this we are passing parameters/arguments as Row Num and Col Num
            public static String getCellData(int RowNum, int ColNum) throws Exception{
            	  Cell = ExcelWSheet.getRow(RowNum).getCell(ColNum);
                  String CellData = Cell.getStringCellValue();
                  return CellData;
	}

}

    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    Now Modify your main 'DriverScript'. With the help of Excel Utility, open the Excel file, read Action Keywords one by one and on each Action Keyword perform the required step.

    package executionEngine;

    import config.ActionKeywords;
    import utility.ExcelUtils;

    public class DriverScript {

        public static void main(String[] args) throws Exception {
        	// Declaring the path of the Excel file with the name of the Excel file
        	String sPath = "D://Tools QA Projects//trunk//Hybrid Keyword Driven//src//dataEngine//DataEngine.xlsx";

        	// Here we are passing the Excel path and SheetName as arguments to connect with Excel file 
        	ExcelUtils.setExcelFile(sPath, "Test Steps");

        	//Hard coded values are used for Excel row & columns for now
        	//In later chapters we will replace these hard coded values with varibales
        	//This is the loop for reading the values of the column 3 (Action Keyword) row by row
        	for (int iRow=1;iRow<=9;iRow++){
    		    //Storing the value of excel cell in sActionKeyword string variable
        		String sActionKeyword = ExcelUtils.getCellData(iRow, 3);

        		//Comparing the value of Excel cell with all the project keywords
        		if(sActionKeyword.equals("openBrowser")){
                            //This will execute if the excel cell value is 'openBrowser'
        			//Action Keyword is called here to perform action
        			ActionKeywords.openBrowser();}
        		else if(sActionKeyword.equals("navigate")){
        			ActionKeywords.navigate();}
        		else if(sActionKeyword.equals("click_MyAccount")){
        			ActionKeywords.click_MyAccount();}
        		else if(sActionKeyword.equals("input_Username")){
        			ActionKeywords.input_Username();}
        		else if(sActionKeyword.equals("input_Password")){
        			ActionKeywords.input_Password();}
        		else if(sActionKeyword.equals("click_Login")){
        			ActionKeywords.click_Login();}
        		else if(sActionKeyword.equals("waitFor")){
        			ActionKeywords.waitFor();}
        		else if(sActionKeyword.equals("click_Logout")){
        			ActionKeywords.click_Logout();}
        		else if(sActionKeyword.equals("closeBrowser")){
        			ActionKeywords.closeBrowser();}

        		}
        	}
     }
    
    
    
    
    
    Note: There is no change in the Action Keyword class, it is the same as in the last chapter.

    Driver Script Class:
    package executionEngine;

    import java.lang.reflect.Method;
    import config.ActionKeywords;
    import utility.ExcelUtils;

    public class DriverScript {
    	//This is a class object, declared as 'public static'
    	//So that it can be used outside the scope of main[] method
    	public static ActionKeywords actionKeywords;
    	public static String sActionKeyword;
    	//This is reflection class object, declared as 'public static'
    	//So that it can be used outside the scope of main[] method
    	public static Method method[];

    	//Here we are instantiating a new object of class 'ActionKeywords'
    	public DriverScript() throws NoSuchMethodException, SecurityException{
    		actionKeywords = new ActionKeywords();
    		//This will load all the methods of the class 'ActionKeywords' in it.
                    //It will be like array of method, use the break point here and do the watch
    		method = actionKeywords.getClass().getMethods();
    	}

        public static void main(String[] args) throws Exception {

        	//Declaring the path of the Excel file with the name of the Excel file
        	String sPath = "D://Tools QA Projects//trunk//Hybrid Keyword Driven//src//dataEngine//DataEngine.xlsx";

        	//Here we are passing the Excel path and SheetName to connect with the Excel file
            //This method was created in the last chapter of 'Set up Data Engine' 		
        	ExcelUtils.setExcelFile(sPath, "Test Steps");

        	//Hard coded values are used for Excel row & columns for now
        	//In later chapters we will use these hard coded value much efficiently
        	//This is the loop for reading the values of the column 3 (Action Keyword) row by row
    		//It means this loop will execute all the steps mentioned for the test case in Test Steps sheet
        	for (int iRow = 1;iRow <= 9;iRow++){
    		    //This to get the value of column Action Keyword from the excel
        		sActionKeyword = ExcelUtils.getCellData(iRow, 3);
                //A new separate method is created with the name 'execute_Actions'
    			//You will find this method below of the this test
    			//So this statement is doing nothing but calling that piece of code to execute
        		execute_Actions();
        		}
        	}
    	
    	//This method contains the code to perform some action
    	//As it is completely different set of logic, which revolves around the action only,
    	//It makes sense to keep it separate from the main driver script
    	//This is to execute test step (Action)
        private static void execute_Actions() throws Exception {
    		//This is a loop which will run for the number of actions in the Action Keyword class 
    		//method variable contain all the method and method.length returns the total number of methods
    		for(int i = 0;i < method.length;i++){
    			//This is now comparing the method name with the ActionKeyword value got from excel
    			if(method[i].getName().equals(sActionKeyword)){
    				//In case of match found, it will execute the matched method
    				method[i].invoke(actionKeywords);
    				//Once any method is executed, this break statement will take the flow outside of for loop
    				break;
    				}
    			}
    		}
     }

    The above code is now much more clear and simple. If there is any addition of method in Action Keyword class, driver script will not have any effect of it. It will automatically consider the newly created method.