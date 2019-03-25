package PerformOperation;

import java.awt.Robot;
import java.awt.Toolkit;
import java.awt.datatransfer.StringSelection;
import java.awt.event.KeyEvent;
import java.io.BufferedWriter;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.FileWriter;
import java.io.IOException;
import java.io.PrintWriter;
import java.net.Authenticator;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.ResultSet;
import java.sql.Statement;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.time.temporal.ChronoField;
import java.time.temporal.TemporalAccessor;
import java.util.Calendar;
import java.util.Date;
import java.util.Dictionary;
import java.util.HashMap;
import java.util.Hashtable;
import java.util.Iterator;
import java.util.List;
import java.util.Locale;
import java.util.Map;
import java.util.Properties;
import java.util.Set;
import java.util.concurrent.TimeUnit;
import javax.crypto.KeyGenerator;
import javax.crypto.SecretKey;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.Alert;
import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.Keys;
import org.openqa.selenium.NoAlertPresentException;
import org.openqa.selenium.UnexpectedAlertBehaviour;
import org.openqa.selenium.UnhandledAlertException;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.openqa.selenium.ie.InternetExplorerDriver;
import org.openqa.selenium.interactions.Actions;
import org.openqa.selenium.remote.CapabilityType;
import org.openqa.selenium.remote.DesiredCapabilities;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.Select;
import org.openqa.selenium.support.ui.WebDriverWait;
//import org.sikuli.script.Key;
//import org.sikuli.script.Pattern;
//import org.sikuli.script.Screen;
import org.testng.annotations.Listeners;
import org.testng.annotations.Test;

import MainExecution.MainClass;
//import junit.framework.Assert;



public class OperationPerform
{
	WebDriver driver;
	//{ 
		//System.setProperty("atu.reporter.config", "C:\\Users\\MGHQ1353\\git\\WebDriverTest_Maven\\WebDriverTest_Maven\\atu.properties"); }
	//Screen s=new Screen();
	//Pattern p=new Pattern();	
	static Hashtable dict= new Hashtable();
		public OperationPerform()
		{
       // this.driver = driver;
		}
	//public void perform(String sBrowser, String sObjectName, String sObjectType, String sOperation, String sValue,String cellVal)throws Exception
	//@SuppressWarnings("deprecation")
		
		@Test
	public void perform(String sBrowser, String sObjectName, String sObjectType, String sOperation, String sValue)throws Exception
	{		
		switch (sOperation)
			{
			case "setdata":
            {       
            	
                  if(sValue.isEmpty()) 
                  {
                         break;
                  }
                  else
                  {
                 // driver.findElement(fetchObject(sObjectName)).clear();  
                  driver.findElement(fetchObject(sObjectName)).sendKeys(sValue);
                  Thread.sleep(3000);
                  }
            	
                	  break;  
                  }
       
			case "set_config_file":
			{             
				String filePath = System.getProperty("user.dir");
				System.out.println(filePath);
            	FileInputStream fis=new FileInputStream(sValue);
            	Workbook wb=WorkbookFactory.create(fis);
            	Sheet sheet =wb.getSheet("Output");         
            	int rowsnum = sheet.getLastRowNum();
            	Row rownum = sheet.getRow(rowsnum);
            	String sValue1=rownum.getCell(0).toString();       
            	String fValue=filePath + "\\" +sValue1;
            	System.out.println(fValue);
            	driver.findElement(fetchObject(sObjectName)).sendKeys(fValue); 
            	Thread.sleep(3000);
            	break;
            }  
            
			case "geturl":
			{	
				driver.get(sValue);
				break;
			}
		
			case "launchbrowser":
			{	
			
				if(sBrowser.equals("IE"))
				{
					System.setProperty("webdriver.ie.driver", "C:\\Selenium\\IEDriverServer.exe");
					driver = new InternetExplorerDriver();
					driver.get(sValue);
					Thread.sleep(3000);
				}
				else if(sBrowser.equals("CHROME"))
				{
					System.setProperty("webdriver.chrome.driver", "C:\\Selenium\\chromedriver.exe");
					driver = new ChromeDriver();
					driver.get(sValue);
				}
				else if(sBrowser.equals("FIREFOX"))
				{
					//System.setProperty("webdriver.firefox.marionette", "C:\\Users\\mghq1353\\workspace\\lib\\geckodriver.exe");
					System.setProperty("webdriver.firefox.marionette", "C:\\Selenium\\geckodriver.exe");
					DesiredCapabilities dc = new DesiredCapabilities();
					dc.setCapability(CapabilityType.UNEXPECTED_ALERT_BEHAVIOUR, UnexpectedAlertBehaviour.IGNORE);
					driver = new FirefoxDriver(dc);
					try{
					driver.get(sValue);
					Thread.sleep(3000);
				}
					 catch (Exception e) {
						 
                         System.out.println("Exception is"+e );
                         }     
					break;
				}
				else
				{
					//by default execution on Firefox browser
					System.setProperty("webdriver.firefox.marionette", "C:\\Selenium\\geckodriver.exe");
					//driver = new FirefoxDriver();
				
					DesiredCapabilities dc = new DesiredCapabilities();
					dc.setCapability(CapabilityType.UNEXPECTED_ALERT_BEHAVIOUR, UnexpectedAlertBehaviour.IGNORE);
					driver = new FirefoxDriver(dc);
					driver.get(sValue);
				}
				break;
				}	
		
			case "end_test":
			{
				driver.quit();
	          	
				break;
			}
			case "clean_up":
			{
				Runtime.getRuntime().exec("C:\\Users\\mghq1353\\Desktop\\Cleanup.exe");
				
				break;
			}
			case "click":
			{
				//driver.manage().timeouts().implicitlyWait(5, TimeUnit.SECONDS);
				try
				{
					driver.findElement(fetchObject(sObjectName)).click();
                	Thread.sleep(3000);
                	
                	break;
				
				}
				catch(Exception e)
				{
                	
					break;
				}
			}
			case "keyboard_event":
			{
				Robot robot = new Robot();
				robot.keyPress(KeyEvent.VK_DOWN);
				robot.keyPress(KeyEvent.VK_ENTER);
				break;
			}
//			case "Encrypt_password":
//			{
//			try {
//			    // Generate a temporary key. In practice, you would save this key.
//			    // See also Encrypting with DES Using a Pass Phrase.
//			    SecretKey key = KeyGenerator.getInstance("DES").generateKey();
//
//			    // Create encrypter/decrypter class
//			    DesEncrypter encrypter = new DesEncrypter(key);
//
//			    // Encrypt
//			    String encrypted = encrypter.encrypt("Don't tell anybody!");
//
//			    // Decrypt
//			    String decrypted = encrypter.decrypt(encrypted);
//			} catch (Exception e)
//			{
//				
				
				
//			}
//			break;
//			}
//			
			case "click_anywhere":
			{
				Actions actions = new Actions(driver);

				Robot robot = new Robot();

				robot.mouseMove(50,50);

				actions.click().build().perform();

//				robot.mouseMove(200,70);
//
//				actions.click().build().perform();
				break;
			}
		
			case "getdata":
			{
				
				String data = driver.findElement(fetchObject(sObjectName)).getText();
				Thread.sleep(3000);
				break;
				}
			case "getattribute":
			{
				String data = driver.findElement(fetchObject(sObjectName)).getAttribute(sValue);
				Thread.sleep(3000);
				break;
				}
			
			case "select_checkbox":
			{
				//List<WebElement> checkboxes=driver.findElements(By.xpath(""));
				WebElement checkbox1=driver.findElement(fetchObject(sObjectName));
				if(checkbox1.isSelected())
				{
					checkbox1.click();
					break;
				}
				
				break;
			}
			
			case "switchframe":
			{		
//				try{
				driver.switchTo().frame(sObjectName);
				System.out.println("Frame switched");
				Thread.sleep(3000);
//				}
//				catch(Exception e)
//				{
					break;
				}
//			}
//			case "switch_multiple_frame":
//			{
//				try{
//					driver.switchTo().frame("PegaGadget0Ifr");
//					try{
//						driver.switchTo().frame("PegaGadget1Ifr");
//						try{
//							driver.switchTo().frame("PegaGadget2Ifr");
//							try{
//								driver.switchTo().frame("PegaGadget3Ifr");
//								try{
//									driver.switchTo().frame("PegaGadget4Ifr");
//									try{
//										driver.switchTo().frame("PegaGadget5Ifr");
//										}
//									catch(Exception ex)
//									{
//											System.out.println("Pega frame 5 not found");	
//									}
//								}
//									catch(Exception ex)
//									{
//											System.out.println("Pega frame 4 not found");	
//									}
//								}
//									catch(Exception ex)
//									{
//											System.out.println("Pega frame 3 not found");	
//									}
//						}
//						catch(Exception ex)
//						{
//						System.out.println("Pega frame 2 not found");	
//						}
//					}
//						catch(Exception ex)
//						{
//							System.out.println("Pega frame 1 not found");
//						}
//				}
//						catch(Exception ex)
//						{
//							System.out.println("Pega frame 0 not found");
//						}
//				break;
//					}
//			//	}
//			//}
//			
			
			case"switch_multiple_frames":
			{
				for(int f=0; f<=5 ; f++)
				{
					try{
					 driver.switchTo().frame("PegaGadget"+f+"Ifr");
					 System.out.println("PegaGadget"+f+"Ifr");
					
					 Thread.sleep(5000);
					 
						 driver.findElement(fetchObject(sObjectName)).click();
						 System.out.println("Clicked");
						 //break;
					 }
					 catch(Exception ex)
					 {
						 
					 }
					continue;
					}
				//break;
			}

			case "switchframe1":
			{		
			int total=driver.findElements(By.tagName("iframe")).size();
				for(int i=0;i<=total;i++)
					{
						driver.switchTo().frame(i);
						System.out.println(total);
			          //also locate for element
			        }
			}	
			case "switchtochildframe":
			{
				System.out.println(driver.switchTo().parentFrame());
				driver.switchTo().parentFrame().switchTo().frame(sObjectName);
				Thread.sleep(3000);
			}
			case "createfile":
			{
				String filePath = System.getProperty("user.dir");
				System.out.println(filePath);
				String filename=(driver.findElement(fetchObject(sObjectName)).getText());
				System.out.println(filename);
				File file = new File(filePath + "\\" + filename);
				if (!file.exists())
				{
	                file.createNewFile();
	               // FileWriter FW = new FileWriter(file);
	               // BufferedWriter BW = new BufferedWriter(FW);
	                //BW.write("test"); //Writing In To File.
	                FileOutputStream fos = new FileOutputStream(file);
	        		PrintWriter pw = new PrintWriter(fos);       			
	        				pw.write("test");	
	        				//pw.append("a");
	        				pw.close();
	        				fos.close();
	               
	                System.out.println("File is created");
	            } else {
	                System.out.println("File already exist");
	            }
					break;
			}
			case "switchframebyid":
			{		
				driver.switchTo().frame(0);
				System.out.println("Frame switched");
				Thread.sleep(3000);
				break;
			}	
			case "switchwindow":
			{
				for(String Child_window : driver.getWindowHandles()) 
				{
					driver.switchTo().window(Child_window);   
					System.out.println("Window got switched to"+Child_window);
					Thread.sleep(3000);
					
				}
				break;
			}
			case "Switch_Multiple_Windows":
			{
				Set < String > s = driver.getWindowHandles();   
		           Iterator < String > ite = s.iterator();
		           int i = 1;
		           while (ite.hasNext() && i < 10)
		           {
		              String popupHandle = ite.next().toString();
		              Thread.sleep(2000);
		               driver.switchTo().window(popupHandle);
		               Thread.sleep(2000);
		               System.out.println("Window title is : "+driver.getTitle());
		               int NoOfWin=Integer.parseInt(sValue);
		               
		               if (i == NoOfWin)
		            	   break;
		               i++;
		               if(driver.getTitle().equalsIgnoreCase(sObjectName))
		               {
		                     Thread.sleep(2000);
		                     driver.switchTo().window(popupHandle).getTitle().equalsIgnoreCase(sObjectName);
		                     Thread.sleep(2000);
		               }
		               
		           }
		           
		           Thread.sleep(3000);
		              
		              driver.manage().window().maximize();
		              break;
			}
			
			
			case "action_dblclick":
			{	
				try{
				WebElement element=driver.findElement(fetchObject(sObjectName));
				Actions action = new Actions(driver).doubleClick(element);
				action.build().perform();
				break;

				}
				catch(Exception e)
				{
				break;
				}
			}
			case "action_click":
			{
				//try{
				WebElement element=driver.findElement(fetchObject(sObjectName));
				Actions action = new Actions(driver).click(element);
				action.build().perform();
				Thread.sleep(3000);
				 //  }
				//catch(Exception e)
				//{
					break;
				//}
				
			}
			case "action_mouseover":
			{
			
				WebElement element=driver.findElement(fetchObject(sObjectName));
				Actions action = new Actions(driver);
				action.moveToElement(element);
				action.click(element);
				System.out.println("Actions click");			
	        	//action.clickAndHold(element);
				Thread.sleep(3000);
				WebElement SubMenu=driver.findElement(fetchObject(sValue));
				action.moveToElement(SubMenu);
				action.click().build().perform();
				
			
				break;
			
			}
			case "javascript_click":
			{
				WebElement elementToClick = driver.findElement(fetchObject(sObjectName));
				((JavascriptExecutor)driver).executeScript("window.scrollTo(0,"+elementToClick.getLocation().y+")");
				elementToClick.click();
			}
			case "clear":
			{
				
				driver.findElement(fetchObject(sObjectName)).clear();		
				break;
			}
			case "wait":
			{	
				int num = Integer.parseInt(sValue);
				Thread.sleep(num);
				break;
			}	
			case "js_mouseevent":
			{
				WebElement element = driver.findElement(fetchObject(sObjectName));
				JavascriptExecutor js = (JavascriptExecutor) driver;
				js.executeScript("var evt = document.createEvent('MouseEvents');" + "evt.initMouseEvent('click',true, true, window, 0, 0, 0, 0, 0, false, false, false, false, 0,null);" + "arguments[0].dispatchEvent(evt);", (element));
			}
			case "maximize_browser":
			{
				driver.manage().window().maximize();
				break;
			}
			case "defaultwindow":
			{
				driver.switchTo().defaultContent();
				break;
			}
			case "selecttext":
			{
				try{
					Select element= new Select(driver.findElement(fetchObject(sObjectName)));
					element.selectByVisibleText(sValue);
					Thread.sleep(3000);
					}
				catch(Exception e)
				{
					break;
				}
			}
			case "selectbyvalue":
			{
				driver.manage().timeouts().implicitlyWait(3,TimeUnit.SECONDS);
				try{
					Select element= new Select(driver.findElement(fetchObject(sObjectName)));
					element.selectByValue(sValue);
					Thread.sleep(3000);
				}
				catch(Exception e)
				{
					break;
				}
			}
			case "selectbyindex":
			{
				driver.manage().timeouts().implicitlyWait(3,TimeUnit.SECONDS);
				try{
	               int ind = Integer.parseInt(sValue);
	               Select element= new Select(driver.findElement(fetchObject(sObjectName)));
	               element.selectByIndex(ind);
	               Thread.sleep(3000);
					}
				catch(Exception e)
				{
					break;
				}
			}
        
			case "modaldialog":
			{
				((JavascriptExecutor) driver).executeScript("window.showModalDialog = window.open;");
			}
			case "js_scroll":
			{
				JavascriptExecutor jse = (JavascriptExecutor)driver;
				jse.executeScript("window.scrollBy(0,250)", "");
				break;
			}
			case "loading_wait":
			{
				By loadingElement = fetchObject(sObjectName);
				WebDriverWait wait = new WebDriverWait(driver, 30);
				wait.until(ExpectedConditions.visibilityOfElementLocated(loadingElement));
				break;
			}
			case "wait_3":
			{      
	              WebElement element2=driver.findElement(fetchObject(sObjectName));
	              new WebDriverWait(driver,10000).until(ExpectedConditions.visibilityOf(element2));
	              break;
			}
		//case "input_data":
		//{
  		//	driver.findElement(fetchObject(sObjectName)).sendKeys(cellVal); 
  		//	break;
		//}
			case "js_click":
			{
				
				JavascriptExecutor jse = (JavascriptExecutor)driver;
				WebElement element = driver.findElement(fetchObject(sObjectName));
				jse.executeScript("arguments[0].click();" , element);
				break;
			}
		
			case "js_input":
			{
				JavascriptExecutor jse = (JavascriptExecutor)driver;
				WebElement element = driver.findElement(fetchObject(sObjectName));
				jse.executeScript("arguments[0].setAttribute('parminder.kaur@orange.com', element)");
				break;
			}
			
			case "js_scrollandclick":
			{
				JavascriptExecutor jse = (JavascriptExecutor)driver;
				WebElement element = driver.findElement(fetchObject(sObjectName));
				jse.executeScript("arguments[0].scrollIntoView(true);", element);

				break;
			}
			case "close_alert":
			{
				JavascriptExecutor jse = (JavascriptExecutor)driver;
				jse.executeScript("closeSplashScreen();reloadScrollBars();");
				break;
			} 
			case "set_date":
			{		
				driver.findElement(fetchObject(sObjectName)).click();
				break;
			} 
			case "todaydate":
			{
				DateFormat dateFormat = new SimpleDateFormat("MM/dd/yyyy");
				Date date = new Date();
				driver.findElement(fetchObject(sObjectName)).sendKeys(dateFormat.format(date));
				break;
			}
			case "set_calendar_date":
			{
				DateFormat dateFormat = new SimpleDateFormat("M/d/yyyy");
				Date date = new Date();
				//String date1=dateFormat.toString();
				String date1=dateFormat.format(date);
				System.out.println(date1);
				String date2[]=(date1.split(" ")[0]).split("/");
				String fdate=date2[1];
				//System.out.println(date2[0]);
				System.out.println(date2[1]);
				//System.out.println(date2[2]);
				List<WebElement> ls = driver.findElements(By.xpath("//table//tbody//tr//td"));
				System.out.println(ls);
				for(WebElement datepick: ls)
				{
					if(datepick.getText().equals(fdate))
					{
						datepick.click();
					}
				}
				break;
			}		
			case "set_before_date":
			{
				String date=driver.findElement(fetchObject(sObjectName)).getText();
				DateFormat dateFormat = new SimpleDateFormat("MM/dd/yyyy");
				String date1=dateFormat.format(date);
				System.out.println(date1);
				String date2[]=(date1.split(" ")[0]).split("/");
				System.out.println(date2[0]);
				System.out.println(date2[1]);
				System.out.println(date2[2]);
				List<WebElement> ls = driver.findElements(By.xpath("//table//tbody//tr//td"));
				ls.get(Integer.parseInt(date2[0])-1);
				for(WebElement datepick: ls)
				{
					if(datepick.getText().equals(ls.get(Integer.parseInt(date2[0])-1)))
					{
						datepick.click();
					}
				}
					break;
			}
			case "set_tos_date":
			{
				//driver.findElement(fetchObject(sObjectName)).getText();
				DateFormat dateFormat = new SimpleDateFormat("MM/dd/yyyy");
				Date date = new Date();
				String date1=dateFormat.format(date);
				System.out.println(date1);
				String date2[]=(date1.split(" ")[0]).split("/");
				String fdate=date2[1];
				System.out.println(date2[0]);
				System.out.println(date2[1]);
				System.out.println(date2[2]);
				List<WebElement> ls = driver.findElements(By.xpath("//table//tbody//tr//td"));
				System.out.println(ls);
				for(WebElement datepick: ls)
				{
					if(datepick.getText().equals(fdate))
					{
						datepick.click();
					}
				}
					break;
			}
			case "read_data":
			{
				//for reading data from excel
				//send sheetname as sObjectType
				//send file path as sValue
				FileInputStream fis=new FileInputStream(sValue);
				Workbook wb=WorkbookFactory.create(fis);
				Sheet sheet =wb.getSheet(sObjectType ); 
				int rowsnum = sheet.getLastRowNum();
				Row rownum = sheet.getRow(rowsnum);
				String sValue1=rownum.getCell(0).toString();
				driver.findElement(fetchObject(sObjectName)).sendKeys(sValue1);
				Thread.sleep(3000);
				fis.close();
				break;
			}
			case "capture_data":
			{
	         //for capturing data and writing in excel
		  	//send sheetname as sObjectType
		  	//send file path as sValue
				try{
				String A = (driver.findElement(fetchObject(sObjectName)).getText());
				System.out.println(A);
				FileInputStream fis=new FileInputStream(sValue);
				Workbook wb=WorkbookFactory.create(fis);
				Sheet sheet =wb.getSheet(sObjectType); 
				int rowsnum = sheet.getLastRowNum();
				System.out.println(rowsnum);
				Row row1 = sheet.createRow(rowsnum+1);
				Cell cell1 = row1.createCell(0);
				cell1.setCellValue(A);
				FileOutputStream fout = new FileOutputStream(sValue); 
				wb.write(fout);         
				fout.close();
				
				break;
			}
				catch(Exception ex)
				{
                	
                	break;
                	}
        		}
			case "capture_scriptname":
			{
				try{
					FileInputStream fis=new FileInputStream(sValue);
					Workbook wb=WorkbookFactory.create(fis);
					Sheet sheet =wb.getSheet(sObjectType); 
					int rowsnum = sheet.getLastRowNum();
					System.out.println(rowsnum);
					Row row1 = sheet.createRow(rowsnum+1);
					Cell cell1 = row1.createCell(0);
					cell1.setCellValue(sObjectName);
					FileOutputStream fout = new FileOutputStream(sValue); 
					wb.write(fout);         
					fout.close();
					}
				catch(Exception e)
				{
					break;
				}
				
			}
			case "capture_data_mlan":
			{
				WebElement A = driver.findElement(fetchObject(sObjectName));
				String text = A.getText();
				System.out.println(text);
				String s1 = text.substring(0, 11);
				FileInputStream fis=new FileInputStream(sValue);
				Workbook wb=WorkbookFactory.create(fis);
				Sheet sheet =wb.getSheet(sObjectType); 
				int rowsnum = sheet.getLastRowNum();
				System.out.println(rowsnum);
				Row row1 = sheet.createRow(rowsnum+1);
				Cell cell1 = row1.createCell(0);
				cell1.setCellValue(s1);
				FileOutputStream fout = new FileOutputStream(sValue); 
				wb.write(fout);         
				fout.close();
				break;
			}
			case "acceptalert":
			{
				
				    try {
				        Alert alert = driver.switchTo().alert();
				        String alertText = alert.getText();
				        System.out.println("Alert data: " + alertText);
				        alert.accept();
				    	}
				    	catch (NoAlertPresentException e)
				    {
				        e.printStackTrace();
				    }
				
               break;
			}
			case "dismissalert":
			{
               try{
                     Alert al=driver.switchTo().alert();
                     al.dismiss();
                     Thread.sleep(3000);
                     break;
               }
               catch(Exception e){
                     break;
               }
			}
			case "getdata_gold":
			{
              String A = (driver.findElement(fetchObject(sObjectName)).getText());
              String result5 = A.replaceAll("[^a-zA-Z0-9-]", "");
              System.out.println("order:"+result5);
              System.out.println("order:"+A);
              FileInputStream fis=new FileInputStream(sValue);
              Workbook wb=WorkbookFactory.create(fis);
              Sheet sheet=wb.getSheet("Order_Ref");      
              int rowsnum = sheet.getLastRowNum();
              Row row1 = sheet.createRow(rowsnum+1);
              Cell cell1 = row1.createCell(0);
              cell1.setCellValue(result5);
              FileOutputStream fout = new FileOutputStream(sValue); 
              wb.write(fout);         
              fout.close();
              break; 
			}
			case "gatewaystatus":
			{ 
				Class.forName("oracle.jdbc.driver.OracleDriver");
                Connection con=DriverManager.getConnection("jdbc:"+"oracle:thin:@10.237.59.108:1521:cisora01","CISOBS","OBS173CIS"); 
                Statement stat=con.createStatement();
                String query="select * from ODS_GLOBAL.CE_GTW_TXN where GTW_ID='GTW024'";
                ResultSet set=stat.executeQuery(query);
                while(set.next())
                {
                	System.out.println(set.getString("STATUS"));
                }
                break;
			}
			case "LOIS_CIS":
			{ 
				Class.forName("oracle.jdbc.driver.OracleDriver");
				Connection con=DriverManager.getConnection("jdbc:"+"oracle:thin:@10.237.59.108:1521:cisora01","CISOBS","OBS173CIS");       
				FileInputStream fis=new FileInputStream(sValue);
				Workbook wb=WorkbookFactory.create(fis);
				Sheet sheet =wb.getSheet(sObjectType );         
				int rowsnum = sheet.getLastRowNum();
				Row rownum = sheet.getRow(rowsnum);    
				String sValue1=rownum.getCell(0).toString();
				System.out.println(sValue1);
				fis.close();
                Statement stat=con.createStatement();
                String query="select * from ods_global.lois_orders_input where ORDERREFERENCE='" +sValue1 + "'";
                ResultSet set=stat.executeQuery(query);
                while(set.next())
                 {
                     System.out.println(set.getString("ORDERREFERENCE"));
                 }
                break;
              
			}           


			case "set_order":
			{                                 
 
 //     System.out.println("order:"+result_order);
        
				FileInputStream fis=new FileInputStream(sValue);
				Workbook wb=WorkbookFactory.create(fis);
				Sheet sheet =wb.getSheet("Order_Ref");         
				int rowsnum = sheet.getLastRowNum();
				Row rownum = sheet.getRow(rowsnum);
				String sValue1=rownum.getCell(0).toString();       
				String result_order = sValue1.replaceAll("[^\\d.-]", ""); 
				System.out.println("order:"+result_order);
				driver.findElement(fetchObject(sObjectName)).sendKeys(result_order);
				Thread.sleep(3000);
				break;

			}  
			case "navigate":
			{
				driver.navigate().to(sValue);
				Thread.sleep(3000);
				break;
			}
			case "device_check":
			{
				List<WebElement> rows=driver.findElements(By.xpath("//table[@id='bodyTbl_right']/tbody/tr"));
				int rowCount=rows.size();
				for(int i=2; i<rowCount;i++)
				{
					//String Model_Name=driver.findElement(By.xpath("//table/tbody/tr[i]/td[5]")).getAttribute("value");
					String Model_Name=driver.findElement(By.xpath("//table[@id='bodyTbl_right']/tbody/tr[i]/td[5]/div/input[@class='autocomplete_input ac_']")).getAttribute("value");
					if (Model_Name.isEmpty()) 
					{
						System.out.println("Model Name is Empty");
						String Device_Name = driver.findElement(By.xpath("//table/tbody/tr[i]/td[2]")).getText();
						System.out.println(Device_Name);
						FileInputStream fis=new FileInputStream("C:\\Selenium\\Test Cases\\Salto Regression\\MLAN\\MLAN Cessation.xlsx");
						Workbook wb=WorkbookFactory.create(fis);
						Sheet sheet =wb.getSheet("Output"); 
						int rowsnum = sheet.getLastRowNum();
						System.out.println(rowsnum);
						Row row1 = sheet.createRow(rowsnum+1);
						Cell cell1 = row1.createCell(0);
						cell1.setCellValue(Device_Name);
						FileOutputStream fout = new FileOutputStream("C:\\Selenium\\Test Cases\\Salto Regression\\MLAN\\MLAN Cessation.xlsx"); 
						wb.write(fout);         
						fout.close();
						
					}
				}
				break;
			}
			case "click_alert_link":
			{
				FileInputStream fis=new FileInputStream(sValue);
				Workbook wb=WorkbookFactory.create(fis);
				Sheet sheet =wb.getSheet(sObjectType); 
				int rowsnum = sheet.getLastRowNum();
				Row rownum = sheet.getRow(rowsnum);
				// System.out.println(rowsnum);
				String sValue1=rownum.getCell(0).toString(); 
				WebElement table =driver.findElement(fetchObject(sObjectName));
				List<WebElement> trs = table.findElements(By.tagName("tr"));
				for( WebElement tr : trs)
				{
					List<WebElement> tds = tr.findElements(By.tagName("td"));
					for( WebElement td : tds)
                     {
                        if( td.getText().equals(sValue1))
                        {
                        	td.findElement(By.tagName("a")).click();
                            break;
                        }else
                             {
                                 System.out.println("link not present");
                                 break;
                             }
                     }
				}

			}
			case "click_link":
			{
				FileInputStream fis=new FileInputStream(sValue);
				Workbook wb=WorkbookFactory.create(fis);
				Sheet sheet =wb.getSheet(sObjectType);       
				int rowsnum = sheet.getLastRowNum();
				Row rownum = sheet.getRow(rowsnum);
				// System.out.println(rowsnum);
				String sValue1=rownum.getCell(0).toString();
				driver.findElement(By.linkText(sValue1)).click();
				Thread.sleep(3000);
                break;
			}
			case "delete_button":
			{
				FileInputStream fis=new FileInputStream(sValue);
				Workbook wb=WorkbookFactory.create(fis);
				Sheet sheet =wb.getSheet(sObjectType);     
				int rowsnum = sheet.getLastRowNum();
				Row rownum = sheet.getRow(rowsnum);
				// System.out.println(rowsnum);
				String sValue1=rownum.getCell(0).toString();
      
				WebElement table =driver.findElement(fetchObject(sObjectName));
				List<WebElement> trs = table.findElements(By.tagName("tr"));

				for( WebElement tr : trs)
				{
					List<WebElement> tds = tr.findElements(By.tagName("td"));
					for( WebElement td : tds)
                         {
                             if( td.getText().equals(sValue1))
                             {
                            	 td.findElement(By.xpath((".//*[@value='Delete']")));
                            	 break;
                             }
                             else
                             {
                                 System.out.println("link not present");
                                 break;
                             }
                         }
				}
			}
			case "is_link_present":
			{
				FileInputStream fis=new FileInputStream(sValue);
				Workbook wb=WorkbookFactory.create(fis);
				Sheet sheet =wb.getSheet(sObjectType); 
				int rowsnum = sheet.getLastRowNum();
				Row rownum = sheet.getRow(rowsnum);
				String sValue1=rownum.getCell(0).toString();
				if(driver.findElement(By.linkText(sValue1)) != null)
				{
					System.out.println("User Created. Test Pass");
				}
				else
				{
					System.out.println("User not created. Test Fail."); 
				}              
					break;
			}
			case "read_select_data":
			{
				FileInputStream fis=new FileInputStream(sValue);
				Workbook wb=WorkbookFactory.create(fis);
				String Locator = sObjectType;
				String sheetName = Locator.split(",")[0];
				String cellNum=Locator.split(",")[1];
				int num = Integer.parseInt(cellNum);
				Sheet sheet =wb.getSheet(sheetName); 
				int rowsnum = sheet.getLastRowNum();
				Row rownum = sheet.getRow(rowsnum);
				// System.out.println(rowsnum);
				String sValue1=rownum.getCell(num).toString();
				System.out.println(sValue1);
				Select element= new Select(driver.findElement(fetchObject(sObjectName)));
				element.selectByVisibleText(sValue1); 
				break;
			}

			case "read_form_data":
			{
				FileInputStream fis=new FileInputStream(sValue);
				Workbook wb=WorkbookFactory.create(fis);    
				String Locator = sObjectType;
				String sheetName = Locator.split(",")[0];
				String cellNum=Locator.split(",")[1];
				int num = Integer.parseInt(cellNum);
				Sheet sheet =wb.getSheet(sheetName); 
				int rowsnum = sheet.getLastRowNum();
				Row rownum = sheet.getRow(rowsnum);
				// System.out.println(rowsnum);
				String sValue1=rownum.getCell(num).toString();
				System.out.println(sValue1);
				driver.findElement(fetchObject(sObjectName)).sendKeys(sValue1); 
                break;
			}
			
			case "selectlastindex":
			{
				try
				{
					Select element= new Select(driver.findElement(fetchObject(sObjectName)));
					element.selectByIndex(element.getOptions().size()-1);
					Thread.sleep(3000);       
					break;
				}
				catch(Exception e){
					break;                     
				}
			}
			case "closeBrowser":
			{
				driver.close();
				break;
			}
			case "capture_time":
			{
//capture time value from popup and save in excel
//send sheetname in sObjectType cell and filepath in sValue cell

				try{
					String textPop=driver.findElement(fetchObject(sObjectName)).getText();
					System.out.println(textPop);
 					String mins=textPop.substring(18,20 );
 					String secs=textPop.substring(21,23 );
       
 					System.out.println(mins+":"+secs);
 					FileInputStream fis=new FileInputStream(sValue);
 					Workbook wb=WorkbookFactory.create(fis);
 					Sheet sheet =wb.getSheet(sObjectType); 
      
 					int rowsnum = sheet.getLastRowNum();
 					System.out.println(rowsnum);
 					Row row1 = sheet.createRow(rowsnum+1);
 					fis.close();
 					Cell cell1 = row1.createCell(0);
 					Cell cell2 = row1.createCell(1);
 					cell1.setCellValue(mins);
 					cell2.setCellValue(secs);
 					FileOutputStream fout = new FileOutputStream(sValue); 
 					wb.write(fout);         
 					fout.close();
 					// fis.close();
       
 					break;
				}
				catch(Exception e){
					break;
				}
       
			}

			case "gethr":
			{
				try{
//Read hour value from excel
//send sheetname in sObjectType cell and filepath in sValue cell

					FileInputStream fis=new FileInputStream(sValue);
					Workbook wb=WorkbookFactory.create(fis);
					Sheet sheet =wb.getSheet(sObjectType); 


					int rowsnum = sheet.getLastRowNum();
					Row rownum = sheet.getRow(rowsnum);
					// System.out.println(rowsnum);
					String sValue1=rownum.getCell(0).toString();
					System.out.println(sValue1);
					driver.findElement(fetchObject(sObjectName)).sendKeys(sValue1); 
					fis.close();

					break;
				}
				catch(Exception e){
					break;
					}
			}
			case "getMin":
			{
				try{
//Read Minute value from excel
//send sheetname in sObjectType cell and filepath in sValue cell

					FileInputStream fis=new FileInputStream(sValue);
					Workbook wb=WorkbookFactory.create(fis);
					Sheet sheet =wb.getSheet(sObjectType); 
					int rowsnum = sheet.getLastRowNum();
					Row rownum = sheet.getRow(rowsnum);
					String sValue1=rownum.getCell(1).toString();
					System.out.println(sValue1);
					driver.findElement(fetchObject(sObjectName)).sendKeys(sValue1); 
					fis.close();
					break;
				}
				catch(Exception e){
					break;

				}
			}
			
			case "Next_Date":
			{
				try{
				Calendar cal = Calendar.getInstance();
				Date date = cal.getTime() ;
				System.out.println(date);
				System.out.println(sValue);
				int ind = Integer.parseInt(sValue);
				cal.add(Calendar.DAY_OF_YEAR, ind);
				date = cal.getTime() ;
				System.out.println(date);
				SimpleDateFormat nDateFormat  = new SimpleDateFormat("d");
				String nextDay = nDateFormat.format(date);
				System.out.println(nextDay);    

				//int result_final=Integer.parseInt(nextDay);
				//System.out.println(result_final);

				driver.switchTo().defaultContent();
				driver.switchTo().frame(sObjectName);
				Thread.sleep(1000);
             			List<WebElement> columns = driver.findElements(By.xpath("//*[@id='controlCalBody']/tr/td/a"));
				//WebElement dateWidgetFrom = driver.findElement(By.xpath(".//*[@id='controlCalBody']/tr/td/a"));
				//List<WebElement> columns = dateWidgetFrom.findElements(By.tagName("td"));
				

				

				for (WebElement cell: columns) 
				{
					if (cell.getText().equals(nextDay))
					{
						cell.click();
						break;
					}	
				}
       
				
				break;
				}
				catch(Exception e)
				{
                		break;
				}
			}
			
			case "Next_Date_NEW":
			{	
				try{	 
				SimpleDateFormat  dateFormat = new SimpleDateFormat("dd");
				Calendar cal = Calendar.getInstance();
				Date date = cal.getTime();
				System.out.println("Calendar date format is"+date);
				String date1=date.toString();
				System.out.println("date extracted is"+date1);
				//date extracted isTue Jul 31 10:16:23 IST 2018
				String[] date_sp = date1.split(" ");
				String on_date = date_sp[2];
				System.out.println("Date is"+on_date);
				String monthName = date_sp[1];
				System.out.println("Month is" +monthName);
				String Year = date_sp[5];   
				System.out.println("Year is" +Year);
				
				DateTimeFormatter parser = DateTimeFormatter.ofPattern("mm").withLocale(Locale.ENGLISH);
				//TemporalAccessor accessor = parser.parse(on_date);
				TemporalAccessor accessor = parser.parse(monthName);
				int monthNumber = accessor.get(ChronoField.MONTH_OF_YEAR);
				Date myDate = dateFormat.parse(monthName);
				
				//Date date = cal.getTime() ;
				cal.setTime(myDate);			
				int ind = Integer.parseInt(sValue);
				//cal.add(Calendar.DAY_OF_YEAR, ind);
				cal.add(Calendar.DAY_OF_YEAR, ind);
				Date afterDate = cal.getTime();
				System.out.println(afterDate);
				String result = dateFormat.format(afterDate);
				int result_final=Integer.parseInt(result);
				System.out.println(result_final);
				List<WebElement> columns = driver.findElements(By.xpath("//*[@id='controlCalBody']/tr/td/a"));
				driver.switchTo().defaultContent();
				driver.switchTo().frame(sObjectName);
				if(result_final > 31)
					{
						driver.findElement(By.xpath("//*[@id='nextMonth']/a")).click();
					}
				else if(result_final<=31)
				{
					driver.findElement(By.xpath("//*[@id='previousMonth']/a")).click();
				}

				for (WebElement cell: columns) 
				{
					if (cell.getText().equals(result_final))
					{
						cell.click();
						break;
					}	
				}
       
				//driver.findElement(By.linkText(nextDay)).click();
				break;
				}
				catch(Exception e)
				{
                		break;
				}
			}
			
			case "Enter_Date":
            {
            	DateFormat dateformat = new SimpleDateFormat("MM/dd/yyyy");
                Date date = new Date();
                Calendar cal = Calendar.getInstance();
                cal.setTime(date);
                int d1 =Integer.parseInt(sValue);
                cal.add(Calendar.DATE, d1); //minus number would decrement the days
                Date after = cal.getTime();
                String date1 = dateformat.format(after);
                System.out.println(date1);
                driver.findElement(fetchObject(sObjectName)).sendKeys(date1);
                break;
            }

			case "SCM_Date":
			{
	
				//String get_scm_date= driver.findElement(By.xpath("//*[@id='CV']")).getText();
				//System.out.println(get_scm_date);
//				String date_dd_MM_yyyy[] = get_scm_date.split(" ");
//				String day = date_dd_MM_yyyy[0];
//				System.out.println(day);
//				String month = date_dd_MM_yyyy[1];
//				System.out.println(month);
//				String year = date_dd_MM_yyyy[2];
//				System.out.println(year);
				WebElement nextLink= driver.findElement(By.xpath("//*[@id='nextYear']/a"));
				nextLink.click();   
				Calendar cal = Calendar.getInstance();
				cal.get(Calendar.DAY_OF_YEAR);
				Date date = cal.getTime() ;
      
				SimpleDateFormat nDateFormat  = new SimpleDateFormat("d");
				String fDay = nDateFormat.format(date);     
         
				// driver.switchTo().frame(sObjectName);
				List<WebElement> columns = driver.findElements(By.xpath(".//*[@id='controlCalBody']/tr/td/a"));
				//WebElement dateWidgetFrom = driver.findElement(By.xpath(".//*[@id='controlCalBody']/tr/td/a"));
				//List<WebElement> columns = dateWidgetFrom.findElements(By.tagName("td"));

				for (WebElement cell: columns) 
				{
					if (cell.getText().equals(fDay))
					{
						cell.click();
						break;
					}
				}
       
				break;
                
			}
			case "ExpediteDate_ReqProv":
			{
			       //capture Date      
			         String date_cp = driver.findElement(By.xpath("//*[@id='CV']")).getText();
			         String[] date_sp = date_cp.split(" ");
			         String on_date = date_sp[0];    
			         String monthName = date_sp[1];
			         String Year = date_sp[2];   
			         System.out.println("MonthName:"+monthName);
			         //changing the month from Jan to 01		
			         DateTimeFormatter parser = DateTimeFormatter.ofPattern("MMM").withLocale(Locale.ENGLISH);
			         TemporalAccessor accessor = parser.parse(monthName);
			         int monthNumber = accessor.get(ChronoField.MONTH_OF_YEAR);
			         System.out.println("MonthNumber:"+monthNumber); 
			               
			         
			         DateFormat dateFormat = new SimpleDateFormat("d");
			      Date myDate = dateFormat.parse(on_date);

			       // Use the Calendar class to subtract one day
			       Calendar calendar = Calendar.getInstance();
			       calendar.setTime(myDate);
			       calendar.add(Calendar.DAY_OF_YEAR, -1);
			       
			       // Use the date formatter to produce a formatted date string
			       Date previousDate = calendar.getTime();
			       String result = dateFormat.format(previousDate);
			       System.out.println("result::"+result);
			       
			       if(result.equals("31"))
			       { 
			          if(date_cp.equals("01 Jan 2018"))
			           {
			       System.out.println("New date:"+result); 
			        
			       int Month_diff = monthNumber-1;
			        System.out.println(Month_diff);
			      
			         String new_month = new Integer(Month_diff).toString();       
			         String new_day = new Integer(result).toString();
			           if(new_month.equals("0"))
			           {
			                 new_month = "12";
			           }

			           int new_year = Integer.parseInt(Year);
			           int Year_diff = new_year-1;
			          // System.out.println(Year_diff);
			          
			           String new_date = new_day + new_month + Year_diff;
			           //System.out.println(new_date);
			                   
			           //get current date
			           Date date =  new Date();
			           SimpleDateFormat nDateFormat = new SimpleDateFormat("dd-MM-YYYY");
			           String todayDate = nDateFormat.format(date);
			          // System.out.println("Current Date:" +todayDate);
			           
			           driver.findElement(fetchObject(sObjectName)).click();
			           Thread.sleep(2000);
			           driver.switchTo().defaultContent();
			           driver.switchTo().frame(sValue);

			           String[] curr_date = todayDate.split("-");
			           String curr_mon = curr_date[1];
			           //System.out.println(curr_mon);
			           int curr_mon1 = Integer.parseInt(curr_mon);
			           
			           int target_mon = Integer.parseInt(new_month);
			          // System.out.println("Target_month:"+target_mon);
			           int mov_month =  target_mon - curr_mon1; 
			           //System.out.println("mov_month::"+mov_month);
			           
			           
			           for(int k=0; k<mov_month; k++)
			           {         
			           driver.findElement(By.xpath(".//*[@id='nextMonth']/a")).click();         
			           }
			           
			           Thread.sleep(2000);
			           List<WebElement> columns = driver.findElements(By.xpath("//*[@id='controlCalBody']/tr/td/a"));
			         
			           //List<WebElement> columns = dateWidgetFrom.findElements(By.tagName("td"));

			           for (WebElement cell: columns) 
			           {
			               if (cell.getText().equals(result))
			               {
			                   cell.click();
			                   break;
			               }
			           }
			           
			      }
			       }
			       else
			       {         
			           //Click and open the calendar
			    	   driver.findElement(fetchObject(sObjectName)).click();
			           Thread.sleep(2000);
			           driver.switchTo().defaultContent();
			           driver.switchTo().frame(sValue);
			                  
			          //get current date
			        Date date =  new Date();
			        SimpleDateFormat nDateFormat = new SimpleDateFormat("dd-MM-YYYY");
			        String todayDate = nDateFormat.format(date);
			        System.out.println("Current Date:" +todayDate);      
			        String[] curr_date = todayDate.split("-");
			        String curr_mon = curr_date[1];
			        String curr_year = curr_date[2];
			        System.out.println(curr_mon);
			        //get month diff
			        int curr_mon1 = Integer.parseInt(curr_mon);
			           int mon_diff = curr_mon1 - monthNumber;
			           System.out.println(mon_diff);
			           int month_diff=Math.abs(mon_diff);
			           System.out.println(month_diff);
			           //get year diff
			           int curr_year1 = Integer.parseInt(curr_year);
			           int target_year = Integer.parseInt(Year);
			           int Year_diff = target_year - curr_year1;
			           System.out.println(Year_diff);
			           //move the calendar to target year
			           for(int i=0; i<Year_diff; i++)
			              {
			                     System.out.println("i count:" +i);
			                     driver.findElement(By.xpath(".//*[@id='nextYear']/a")).click();
			              }
			         
			           Thread.sleep(2000);
			           //move the calendar to target month
			           for(int j=0;j<month_diff;j++)
			              {
			              System.out.println("j count:" +j);
			              
			              driver.findElement(By.xpath(".//*[@id='nextMonth']/a")).click();
			              Thread.sleep(1000);
			              System.out.println("J loop ended");
			              }
			         
			           Thread.sleep(2000);
			           
			           System.out.println("Click correct date"+result);
			           List<WebElement> columns = driver.findElements(By.xpath("//*[@id='controlCalBody']/tr/td/a"));
			          // List<WebElement> columns = dateWidgetFrom.findElements(By.tagName("td"));
			       
			           for (WebElement cell: columns) 
			           {
			        	//   System.out.println("For loop started for date");
			               if (cell.getText().equals(result))
			               {
			                   cell.click();
			              //     System.out.println("Click date done");
			                   break;
			               }
			           }              

			       }
			       break;
			}

			case "staging_shipping_date":
			{
				WebElement nextLink= driver.findElement(By.xpath("//*[@id='nextMonth']/a"));
				nextLink.click();   
				Calendar cal = Calendar.getInstance();
				cal.get(Calendar.DAY_OF_YEAR);
				Date date = cal.getTime() ;
      
				SimpleDateFormat nDateFormat  = new SimpleDateFormat("d");
				String fDay = nDateFormat.format(date);
				System.out.println(fDay);
       
         
				// driver.switchTo().frame(sObjectName);
				WebElement dateWidgetFrom = driver.findElement(By.xpath("//*[@id='controlCalBody']/tr/td/a"));
				List<WebElement> columns = dateWidgetFrom.findElements(By.tagName("td"));

				for (WebElement cell: columns) 
				{
					if (cell.getText().equals(fDay))
					{
						cell.click();
						break;
					}
				}
       
				break;
                
			}
			
			
			case "refresh_Saltotask":
			{
	
				int num = Integer.parseInt(sValue);
				for(int i =1; i<=num;i++)
				{
					driver.findElement(By.xpath("//*[@id='RULE_KEY']/table[1]/tbody/tr/td[11]/nobr/span/a")).click();
					Thread.sleep(60000);
				}
				break;
			}

			case "mlan_srf1":
			{              
				String service_usid=driver.findElement(By.xpath("//*[@id='BNAZZZO5UGKEZ0UK5UZOJ23MJLNCVHYZ__0__']/tr[2]/td[6]")).getText();
				System.out.println(service_usid);
				String order=driver.findElement(By.xpath("html/body/form[2]/table/tbody/tr[4]/td/table/tbody/tr[2]/td[2]/table/tbody/tr/td/table[1]/tbody/tr[2]/td[1]/span")).getText();
				System.out.println(order);
				FileInputStream fis=new FileInputStream(sValue);
				Workbook wb=WorkbookFactory.create(fis);
				Sheet sheet =wb.getSheet("Order Info"); 
				System.out.println("Sheet name is"+sheet);
				fis.close();
				Row usid = sheet.getRow(6);
				System.out.println("Row selected in sheet is"+usid);

	
				Cell cell =usid.getCell(8);	
				cell.setCellValue(service_usid);
	
				//	usid.getCell(7).setCellValue(service_usid);
	
				Row gold_order = sheet.getRow(12);
				System.out.println("Row selected in sheet is"+gold_order);
				Cell cell1 =gold_order.getCell(8);
				cell1.setCellValue(order);
				FileOutputStream fos = new FileOutputStream(sValue);
				wb.write(fos);
					
				System.out.println("SRF1 file updated");
	
				fos.close();
				Thread.sleep(3000);
				break;

			}  
			/*
			case "mlan_srf1_test":
			{              
				String service_usid=driver.findElement(By.xpath("//*[@id='BNAZZZO5UGKEZ0UK5UZOJ23MJLNCVHYZ__0__']/tr[2]/td[6]")).getText();
				System.out.println(service_usid);
				String order=driver.findElement(By.xpath("html/body/form[2]/table/tbody/tr[4]/td/table/tbody/tr[2]/td[2]/table/tbody/tr/td/table[1]/tbody/tr[2]/td[1]/span")).getText();
				System.out.println(order);
				FileInputStream fis=new FileInputStream(sValue);
				Workbook wb=WorkbookFactory.create(fis);
				Sheet sheet =wb.getSheet("Order Info"); 
				System.out.println("Sheet name is"+sheet);
				fis.close();
				
				Row usid = sheet.getRow(6);
				System.out.println("Row selected in sheet is"+usid);

	
				Cell cell =usid.getCell(8);
				System.out.println(cell);
				//cell.setCellValue(service_usid);
				Robot rb = new Robot();
				rb.keyPress(KeyEvent.VK_F2);
				rb.keyRelease(KeyEvent.VK_F2);
				System.out.println("F2 Pressed and Released");				rb.keyPress(KeyEvent.VK_BACK_SPACE);
				rb.keyRelease(KeyEvent.VK_BACK_SPACE);
				System.out.println("Backspace Pressed and Released");
				StringSelection sel = new StringSelection(service_usid);
				Toolkit.getDefaultToolkit().getSystemClipboard().setContents(sel, null);
				rb.keyPress(KeyEvent.VK_ENTER);

				rb.keyRelease(KeyEvent.VK_ENTER);
				System.out.println("Data entered");	
				//	usid.getCell(7).setCellValue(service_usid);
	
				Row gold_order = sheet.getRow(12);
				System.out.println("Row selected in sheet is"+gold_order);
				Cell cell1 =gold_order.getCell(8);
				System.out.println(cell1);
			//	cell1.setCellValue(order);
				
				
				rb.keyPress(KeyEvent.VK_F2);
				rb.keyRelease(KeyEvent.VK_F2);				rb.keyPress(KeyEvent.VK_BACK_SPACE);
				rb.keyRelease(KeyEvent.VK_BACK_SPACE);
				StringSelection sel1 = new StringSelection(order);
				Toolkit.getDefaultToolkit().getSystemClipboard().setContents(sel1, null);
				rb.keyPress(KeyEvent.VK_ENTER);

				rb.keyRelease(KeyEvent.VK_ENTER);
				
				FileOutputStream fos = new FileOutputStream(sValue);
				wb.write(fos);
					
				System.out.println("SRF1 file updated");

			
				
				fos.close();
				Thread.sleep(3000);
				break;

			}  
			
			*/
			case "acceptlvo":
			{
				driver.findElement(fetchObject(sObjectName)).click();
	
				for(String accept_btn_window : driver.getWindowHandles())
				{
					driver.switchTo().window(accept_btn_window  );
					//driver.manage().timeouts().implicitlyWait(20, TimeUnit.SECONDS);
				}
				System.out.println("child window-accept");
				driver.manage().timeouts().implicitlyWait(20, TimeUnit.SECONDS);
				driver.findElement(By.id("YES_BTN")).click();
   	
				for(String Parent_window : driver.getWindowHandles())
				{
					driver.switchTo().window(Parent_window);
				}
				System.out.println("parent accept");
	
				break;
			}
			
//			case "screen_click":
//			{
//				try
//				{
//				s.click(sValue);
//				break;
//				}
//				catch(Exception e)
//				{
//					break;
//				}
//			}
			
//			case "screen_dblClick":
//			{
//				try
//				{
//					s.doubleClick(sValue);
//					break;
//				}
//				catch(Exception e)
//				{
//					break;
//				}
//			}
//			case "screen_type":
//			{
//				try
//				{
//					s.type(sValue,sObjectType);
//					break;
//				}
//				catch(Exception e)
//				{
//					break;
//				}
//			}
			
			case "screen_Select_Down":
			{
				//s.type(sValue.sObjectType,KeyEvent.VK_DOWN)
			}
//			case "sikuli_date_type":
//			{
//				Calendar cal = Calendar.getInstance();
//				Date date = cal.getTime() ;
//				//System.out.println(date);
//				//System.out.println(sValue);
//				int ind = Integer.parseInt(sObjectType);
//				cal.add(Calendar.DAY_OF_YEAR, ind);
//				date = cal.getTime() ;
//				System.out.println(date);
//				SimpleDateFormat nDateFormat  = new SimpleDateFormat("M/dd/yyyy");
//				String fDay = nDateFormat.format(date);
//				try
//				{
//					s.type(sValue,fDay);
//					break;
//				}
//				catch(Exception e)
//				{
//					break;
//				}
//			}
			
			case "File_Upload":
			{
				//Upload file using Auto IT
				Runtime.getRuntime().exec(sValue); //path of au3 file
//				System.out.println("File is attached");
				break;
				
				
				
//				StringSelection sel = new StringSelection("C:\\Selenium\\Test Cases\\MLAN_SRF1_NEW_Complex_v4.0.xlsm");
//				// Copy to clipboard
//				Toolkit.getDefaultToolkit().getSystemClipboard().setContents(sel,null);
//				 System.out.println("selection" +sel);
//				 // Create object of Robot class
//				 Robot robot = new Robot();
//				 Thread.sleep(1000);
//				      
//				  // Press Enter
//				 robot.keyPress(KeyEvent.VK_ENTER);
//				 
//				// Release Enter
//				 robot.keyRelease(KeyEvent.VK_ENTER);
//				 
//				  // Press CTRL+V
//				 robot.keyPress(KeyEvent.VK_CONTROL);
//				 robot.keyPress(KeyEvent.VK_V);
//				 
//				// Release CTRL+V
//				 robot.keyRelease(KeyEvent.VK_CONTROL);
//				 robot.keyRelease(KeyEvent.VK_V);
//				 Thread.sleep(1000);
//				        
//				         
//				 robot.keyPress(KeyEvent.VK_ENTER);
//				 robot.keyRelease(KeyEvent.VK_ENTER);
	//			 break;
			}
			
			case "compare_text":
			{
				String text = driver.findElement(fetchObject(sObjectName)).getText();
				
				if (text.equalsIgnoreCase(sValue))
				{
					System.out.println("text captured is expected");
				}
				else
				{
					System.out.println("text captured is not expected");
				}
				break;
			}
			case "check_field":
			{
				String value = driver.findElement(fetchObject(sObjectName)).getAttribute("value");
				if (value.isEmpty())
				{
					System.out.println("CSD does not flow from Gold");
					break;
				}
				else
				{
					break;
				}
				
			}
//			case "compare_status":
//			{
//				String text = driver.findElement(fetchObject(sObjectName)).getText();
//				Assert.assertEquals(sValue, text);
//				System.out.println("Staus is compared");
//				break;
//			}

			default:
			break;
		}
	}

	private By fetchObject(String sObjectName)throws Exception
	{
			String Locator = sObjectName;
			
			String LocatorType = Locator.split(":",2)[0];
			String LocatorValue=Locator.split(":",2)[1];	
		//	String Locator2 = sObjectType;
		//	String LocatorType2 = Locator2.split(":")[0];
		//	String LocatorValue2=Locator2.split(":")[1];	
		if (LocatorType.equals("name"))
			{
				return By.name(LocatorValue);
			} 
		else if (LocatorType.equals("xpath"))
			{
				return By.xpath(LocatorValue);
			}
		 else if (LocatorType.equals("cssselector"))
		 	{
			 	return By.cssSelector(LocatorValue);
		 	}
		 else if (LocatorType.equals("id"))
		 	{
			 	return By.id(LocatorValue);
		 	}
		 else if (LocatorType.equals("linktext"))
			{
				return By.linkText(LocatorValue);
			}
		 else if (LocatorType.equals("partiallinktext"))
		 	{
			 	return By.partialLinkText(LocatorValue);
		 	}
		 else if (LocatorType.equals("tagname"))
         {
                return By.tagName(LocatorValue);
         }
		 else if (LocatorType.equals("classname"))
		 {
			 	return By.className(LocatorValue);
		 }

		 else
		 	{
			 	return null;
			
		 	}
		}
}
