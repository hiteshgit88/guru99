package parameters;

import java.lang.reflect.Method;
import java.util.concurrent.TimeUnit;
import org.openqa.selenium.By;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.testng.Assert;
import org.testng.annotations.BeforeTest;
import org.testng.annotations.DataProvider;
import org.testng.annotations.Test;

public class ParameterByMethodInDataprovider{

	ChromeDriver driver=null;
     
    @BeforeTest
    public void setup(){
    	System.setProperty("webdriver.chrome.driver", "E:\\eclipse\\bin\\chromedriver.exe");        
        driver = new ChromeDriver();
        driver.manage().timeouts().implicitlyWait(10, TimeUnit.SECONDS);
        driver.get("https://google.com");
    }

    @Test(dataProvider="SearchProvider")
    public void testMethodA(String author,String searchKey) throws InterruptedException{
        
    	WebElement searchText = driver.findElement(By.name("q"));
        //Search text in search box
        searchText.sendKeys(searchKey);
        //Print author and search string
        System.out.println("Welcome ->"+author+" Your search key is->"+searchKey);
        Thread.sleep(3000);
        String testValue = searchText.getAttribute("value");
        System.out.println(testValue +"::::"+searchKey);
        searchText.clear();
        //Verify if google text box is showing correct value
        Assert.assertTrue(testValue.equalsIgnoreCase(searchKey));
    }
      
    @Test(dataProvider="SearchProvider")
    public void testMethodB(String searchKey) throws InterruptedException{
        {
        	WebElement searchText = driver.findElement(By.name("q"));
            //Search text in search box
            searchText.sendKeys(searchKey);
            //Print only search string
            System.out.println("Welcome ->Unknown user Your search key is->"+searchKey);
            Thread.sleep(3000);
            String testValue = searchText.getAttribute("value");
            System.out.println(testValue +"::::"+searchKey);
            searchText.clear();
            //Verify if google text box is showing correct value
            Assert.assertTrue(testValue.equalsIgnoreCase(searchKey));
        }
    }    
    /**
     * Here DataProvider returning value on the basis of test method name
     * @param m
     * @return
     **/

    @DataProvider(name="SearchProvider")
    public Object[][] getDataFromDataprovider(Method m){
        if(m.getName().equalsIgnoreCase("testMethodA")){
        return new Object[][] {
                { "Guru99", "India" },
                { "Krishna", "UK" },
                { "Bhupesh", "USA" }
            };}
        else{
        return new Object[][] {
                { "Canada" },
                { "Russia" },
                { "Japan" }
            };}       
    }
}