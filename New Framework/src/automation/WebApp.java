package automation;
import java.awt.Rectangle;
import java.awt.image.BufferedImage;
import java.io.File;
import java.io.IOException;
import java.io.StringReader;
import java.io.StringWriter;
import java.lang.reflect.Field;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.util.HashMap;
import java.util.List;

import javax.imageio.ImageIO;
import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.transform.Transformer;
import javax.xml.transform.TransformerFactory;
import javax.xml.transform.stream.StreamResult;
import javax.xml.transform.stream.StreamSource;
import javax.xml.xpath.XPath;
import javax.xml.xpath.XPathExpression;
import javax.xml.xpath.XPathFactory;

import org.apache.commons.io.FileUtils;
import org.apache.log4j.Logger;
import org.apache.poi.ss.usermodel.Row;

import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.Keys;
import org.openqa.selenium.OutputType;
import org.openqa.selenium.Point;
import org.openqa.selenium.TakesScreenshot;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.openqa.selenium.interactions.Actions;
import org.openqa.selenium.support.ui.ExpectedCondition;
import org.openqa.selenium.support.ui.Select;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.w3c.dom.Document;
import org.xml.sax.InputSource;

import com.cisco.expressions.ExpressionEvaluator;

public class WebApp {
	
	WebDriver driver;
	
	ExcelDocument excelDocument;
	
	Logger logger = Logger.getRootLogger();
	
	String codeSnippetClass = "";
	
	XPath xPath = XPathFactory.newInstance().newXPath();
	
	ExpressionEvaluator expressionEvalutor = new ExpressionEvaluator();
	
	String winHandleBefore = "";
	
	public WebApp() {
		initApp();
	}
	
	public String getCodeSnippetClass() {
		return codeSnippetClass;
	}

	public void setCodeSnippetClass(String codeSnippetClass) {
		this.codeSnippetClass = codeSnippetClass;
	}
	
	public void initApp() {
//		System.setProperty("webdriver.firefox.bin", "C:\\FF_NEW\\firefox.exe");
		driver = new FirefoxDriver();
		driver.manage().window().maximize();
		excelDocument = new ExcelDocument();
		excelDocument.setDirecotyPath("");
		excelDocument.setWorkingFileName("Automation.xlsx");
		excelDocument.initiateFileInstance();
		excelDocument.selectSheet("Methods");
	}
	
	/**
     * Parsing object recursively
     * @param rootObject
     */
    @SuppressWarnings({ })
	public String parseObject(String pathString,Object rootObject) {
    	String stringInformation = null;
    	@SuppressWarnings("rawtypes")
    	Class cls = rootObject.getClass();
		Field[] fieldList = cls.getDeclaredFields();
		if (cls.getSimpleName().contains("String")) {

		} else {
			for (Field field : fieldList) {
				try {
					
					String objectType = field.getType().toString();
					if(objectType.contains("boolean")) {
						if(field.getName().equalsIgnoreCase(pathString)) {
							boolean booleanValue = (Boolean) field.get(rootObject);
							stringInformation = String.valueOf(booleanValue);
						}
					} else if(objectType.contains("int")) {
						if(field.getName().equalsIgnoreCase(pathString)) {
							int intVlaue = (Integer) field.get(rootObject);
							stringInformation = String.valueOf(intVlaue);
						}
					} else if(objectType.contains("float")) {
						if(field.getName().equalsIgnoreCase(pathString)) {
							float floatValue = (Integer) field.get(rootObject);
							stringInformation = String.valueOf(floatValue);
						}
					}else if(objectType.contains("Optional")) {

						Object optionalObject = field.get(rootObject);
						if(field.getName().equalsIgnoreCase(pathString)) {
							// Apply reflection on this!
							// See the list of methods!
							Method[] methods = optionalObject.getClass().getDeclaredMethods();
							
							boolean isPresentFlag = false;
							for(int count = 0; count<methods.length; count++){
								Method method = methods[count];
								if(isPresentFlag) {
									if(method.getName().contains("getValue")){
										try {
											method.setAccessible(true);
											Object object = (Object) method.invoke(optionalObject, null);
											if(object!=null){
												stringInformation = parseObject(pathString, object);
											}
											break;
										} catch (InvocationTargetException e) {
											// TODO Auto-generated catch block
											e.printStackTrace();
										}
									}
								} else {
									if(method.getName().contains("ispresent")){
										try {
											method.setAccessible(true);
											isPresentFlag = (Boolean) method.invoke(optionalObject, null);
											if(isPresentFlag){
												count = 0;
											} else{
												break;
											}
										} catch (InvocationTargetException e) {
											// TODO Auto-generated catch block
											e.printStackTrace();
										}
									}
								}
								count++;
							}
						}
						
					} else if (objectType.contains("[Ljava.lang.String")) {
	
						String[] stringObjects = (String[]) field.get(rootObject);
						if(stringObjects != null) {
							int i = 1;
							for(String string:stringObjects) {
								stringInformation = string;
								i++;
							}
						}

					} else if (objectType.contains("[L")) {
						
						Object[] arrayObjects = (Object[]) field
								.get(rootObject);
						if(arrayObjects != null) {
							for (Object object : arrayObjects) {
								stringInformation = parseObject(pathString, object);
							}
						}
						
					} else if (objectType.contains("java.lang.String")) {
						if(field.getName().equalsIgnoreCase(pathString)) {
							String value = (String) field.get(rootObject);
							if(value != null){
								stringInformation = value;
							}
						}
					} else {
						Object object = field.get(rootObject);
						stringInformation = parseObject(pathString,object);
					}
					
				} catch (IllegalArgumentException e) {
					// TODO Auto-generated catch block
					e.printStackTrace();
				} catch (SecurityException e) {
					// TODO Auto-generated catch block
					e.printStackTrace();
				} catch (IllegalAccessException e) {
					// TODO Auto-generated catch block
					e.printStackTrace();
				}
			}
		}
		return stringInformation;
    }
	
	public void runApplication(String methodName, Object genericObject) {
		
		HashMap<String, String> hashMap = new HashMap<String, String>();
		try {
			Class codeSnippetCls = Class.forName(codeSnippetClass);
			Object codeSnippetObject = codeSnippetCls.newInstance();
			
			boolean isMethodActivated = false;
			for(Row row:excelDocument.getSheet()) {
				String cellValue = excelDocument.getValueFromExcel(row.getRowNum(), 0);
				if (isMethodActivated && cellValue.length() > 0) {
					break;
				} else if(isMethodActivated) {
					
					String cellActionString = excelDocument.getValueFromExcel(row.getRowNum(), 1);
					if(cellActionString.length() > 0) {
						
						if(cellActionString.equalsIgnoreCase("code")) {
							String codemethodName = excelDocument.getValueFromExcel(row.getRowNum(), 4);
							Method method = codeSnippetCls.getDeclaredMethod(codemethodName, null);
							method.invoke(codeSnippetObject, null);
						} else if(cellActionString.equalsIgnoreCase("movewindow")) {
							//Store the current window handle
							winHandleBefore = driver.getWindowHandle();

							//Perform the click operation that opens new window

							//Switch to new window opened
							for(String winHandle : driver.getWindowHandles()){
								System.out.println("window Handle"+winHandle);
							    driver.switchTo().window(winHandle);
							}

//							// Perform the actions on new window
//
//							//Close the new window, if that window no more required
//							driver.close();

							//Switch back to original browser (first window)

//							driver.switchTo().window(winHandleBefore);

							//continue with original browser (first window)
						} else if(cellActionString.equalsIgnoreCase("defaultwindow")) {
							driver.switchTo().window(winHandleBefore);
						} else if(cellActionString.equalsIgnoreCase("alert-ok")) {
							driver.switchTo().alert().accept();
						} else if(cellActionString.startsWith("navigate")) {
							String opType = cellActionString.split("-")[1];
							if(opType.equalsIgnoreCase("back")) {
								driver.navigate().back();
							}
						} else if(cellActionString.equalsIgnoreCase("rallyReplace")) {
//							System.out.println("value"+excelDocument.getValueFromExcel(row.getRowNum(), 3));
//							System.out.println(hashMap.get(excelDocument.getValueFromExcel(row.getRowNum(), 3)));
							String value = hashMap.remove(excelDocument.getValueFromExcel(row.getRowNum(), 3));
//							System.out.println("output    "+value);
							value = value.replace("(Copy of)", "May 2014");
							System.out.println(value);
							hashMap.put(excelDocument.getValueFromExcel(row.getRowNum(), 3), value);
						} else {
							
							String varName = excelDocument.getValueFromExcel(row.getRowNum(), 3);
							String elementIdentifier = parseObject(excelDocument.getValueFromExcel(row.getRowNum(), 3), genericObject);
//							System.out.println(elementIdentifier);
							if(elementIdentifier == null) {
								elementIdentifier = hashMap.get(excelDocument.getValueFromExcel(row.getRowNum(), 3));
								if(elementIdentifier == null) {
									elementIdentifier = excelDocument.getValueFromExcel(row.getRowNum(), 3);
								}
							}
							System.out.println(elementIdentifier);
							String elementValue = excelDocument.getValueFromExcel(row.getRowNum(), 4);
							if(elementValue.length() == 0) {
								elementValue = elementIdentifier;
							}
							
							if(elementValue.contains("function:")) {
								System.out.println(elementValue);
								elementValue = elementValue.replace("function:", "");
								elementValue = expressionEvalutor.evaluateExpression(elementValue);
//								elementValue
							}
							
							hashMap.put(varName, elementValue);
							String xpathString = excelDocument.getValueFromExcel(row.getRowNum(), 2);
				
							WebElement webElement = null;
							if(xpathString.length() > 0) {
								if(xpathString.contains("@ID@")) {
									xpathString = xpathString.replace("@ID@", elementIdentifier);
								}
								System.out.println(xpathString);
								if(cellActionString.equalsIgnoreCase("displayclick")) {
//									System.out.println(driver.getPageSource());
									List<WebElement> webElements = driver.findElements(By.xpath(xpathString));
									System.out.println(webElements.size());
									for(WebElement webElement2:webElements){
										if(webElement2.isDisplayed()){
											webElement = webElement2;
											break;
										}
									}
								} else if(cellActionString.equalsIgnoreCase("evaluate")) {
//									Document doc = DocumentBuilderFactory.newInstance().newDocumentBuilder().parse(driver.getPageSource());
									XPathExpression xPathExpression = xPath.compile(xpathString);
									System.out.println(xPathExpression);
									String evlString = (String) xPathExpression.evaluate(convertToXML(driver.getPageSource()));
									System.out.println("eval "+evlString);
									hashMap.put(elementIdentifier, evlString);
//									String expressionValue = String.valueOf(xPath.evaluate(xpathString, driver.getPageSource()));
//									System.out.println(expressionValue);
//									hashMap.put(elementIdentifier, expressionValue);
								} else {
									webElement = returnWebElement(xpathString, "180");
								}
							}
											
							if(cellActionString.startsWith("code")) {
								
							} else if(cellActionString.equalsIgnoreCase("click") || cellActionString.equalsIgnoreCase("displayclick")) {
//								takeScreenShot(webElement);
								webElement.click();
							} else if(cellActionString.equalsIgnoreCase("verify")) {
								String verString = webElement.getText();
								if(verString.equalsIgnoreCase(elementValue)) {
									System.out.println("passed");
								} else {
									System.out.println("exepected string not present in mail");
								}
							} else if (cellActionString.equalsIgnoreCase("get")) {
								System.out.println("Text::"+webElement.getText());
								hashMap.put(elementIdentifier, webElement.getText());
//								takeScreenShot(webElement);
							} else if(cellActionString.equalsIgnoreCase("url")) {
								driver.get(elementValue);
							} else if(cellActionString.equalsIgnoreCase("send")) {
//								takeScreenShot(webElement);
								System.out.println("value "+elementValue);
								if(webElement.getText().length() > 0) {
									webElement.clear();
								} else if(webElement.getAttribute("value") != null && webElement.getAttribute("value").length() > 0) {
									JavascriptExecutor js = (JavascriptExecutor) driver;
									js.executeScript("arguments[0].value = '';", webElement, 10);
								}
								webElement.sendKeys(elementValue);
							} else if(cellActionString.startsWith("mouse-move")) {
//								String opType = cellActionString.split("-")[1];
								Actions actions = new Actions(driver);
								actions.moveToElement(webElement);
								actions.build().perform();
							} else if(cellActionString.startsWith("key")) {
								String keyType = cellActionString.split("-")[1];
								Keys[] keys = Keys.values();
								int keyCount = 0;
								while(keyCount < keys.length) {
									String keyName = keys[keyCount].name();
									if(keyName.matches(keyType)) {
										webElement.sendKeys(keys[keyCount]);
										break;
									}
									keyCount++;
								}
							} else if(cellActionString.equalsIgnoreCase("sleep")) {
								Thread.sleep(10000);
							} else if(cellActionString.equalsIgnoreCase("rightclick")) {
								Actions actions = new Actions(driver);
								actions.contextClick(webElement);
								actions.perform();
							} else if(cellActionString.equalsIgnoreCase("select")) {
								Select select = new Select(webElement);
								boolean optionStatus = false;
								while(!optionStatus) {
									List<WebElement> options = select.getOptions();
									for(WebElement option:options) {
										String optionText = option.getText();
										if(optionText.equalsIgnoreCase(elementValue)) {
											select.selectByVisibleText(optionText);
											optionStatus = true;
											break;
										}
									}
								}
							} else if(cellActionString.startsWith("attribute")) {
								String attributeValue = webElement.getAttribute(cellActionString.split("-")[1]);
								System.out.println(elementIdentifier);
								System.out.println("attributevalue: "+attributeValue);
								if(hashMap.containsKey(elementIdentifier)) {
									hashMap.remove(elementIdentifier);
								}
								if(!hashMap.containsKey(elementIdentifier)) {
									System.out.println("false");
								}
								hashMap.put(elementIdentifier, attributeValue);
//								System.out.println("check "+hashMap.get(elementIdentifier));
							} else {
								
							}
						}
						
					}
					
				} else if(cellValue.equalsIgnoreCase(methodName)) {
					isMethodActivated = true;
				} 
			}
		} catch (InstantiationException e1) {
			// TODO Auto-generated catch block
			e1.printStackTrace();
		} catch (IllegalAccessException e1) {
			// TODO Auto-generated catch block
			e1.printStackTrace();
		} catch (ClassNotFoundException e1) {
			// TODO Auto-generated catch block
			e1.printStackTrace();
		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		
	}
	
	public WebElement returnWebElement(String xpath, String waitTime) {
		final String xpathString = xpath;
		WebElement selectedElement = null;
		try {
			selectedElement = (new WebDriverWait(driver, Integer.parseInt(waitTime))
					.until(new ExpectedCondition<WebElement>() {
						public WebElement apply(WebDriver d) {
							return d.findElement(By.xpath(xpathString));
						}
					}));
			// selectedElement = (new WebDriverWait(driver, waitTime))
			// .until(new ExpectedCondition<WebElement>(){
			// @Override
			// public WebElement apply(WebDriver d) {
			// return d.findElement(by);
			// }});
		} catch (Exception e) {
			e.printStackTrace();
		}
		return selectedElement;
	}
	
	public void takeScreenShot(WebElement element) {
		try {
			File scrFile = ((TakesScreenshot)driver).getScreenshotAs(OutputType.FILE);
			
			if(element != null) {
				Point p = element.getLocation();
				int width = element.getSize().getWidth();
				int height = element.getSize().getHeight();
				Rectangle rectangle = new Rectangle(width, height);
				BufferedImage img = null;
				img = ImageIO.read(scrFile);
				BufferedImage dest = img.getSubimage(p.getX(), p.getY(), rectangle.width, rectangle.height);
				ImageIO.write(dest, "png", scrFile);
			}
			FileUtils.copyFile(scrFile, new File("ss/"+scrFile.lastModified()+".png"));
		} catch (IOException e1) {
			// TODO Auto-generated catch block
			e1.printStackTrace();
		}
	}
	
	public void closeApplication() {
		driver.close();
	}
	
	public Document convertToXML(String inputContent) {
		Document doc = null;
		try 
		{
			System.out.println(inputContent);
			DocumentBuilderFactory factory = DocumentBuilderFactory.newInstance();
	        DocumentBuilder builder = factory.newDocumentBuilder();
	        InputSource is = new InputSource(new StringReader(inputContent));
	        System.out.println(is);
	        doc = builder.parse(is);
	        System.out.println(doc);
//		    TransformerFactory tFactory = TransformerFactory.newInstance();
//		    Transformer transformer = tFactory.newTransformer();
//		    StringWriter strWriter = new StringWriter();
//		    transformer.transform(new StreamSource(new StringReader(inputContent)), new StreamResult(strWriter));
//		    String xmlString = strWriter.toString();
		}
		catch (Exception e)
		{
		    e.printStackTrace();
		}
		return doc;
	}

	/**
	 * @param args
	 */
	public static void main(String[] args) {
		// TODO Auto-generated method stub
		WebApp webApp = new WebApp();
		GenericObject genericObject = new GenericObject();
//		System.out.println(HMPUITesting.class.getName());
//		webApp.setCodeSnippetClass(HMPUITesting.class.getName());
		webApp.setCodeSnippetClass(ItemFoundationTesting.class.getName());
//		webApp.setCodeSnippetClass("HMPUITesting");
//		webApp.runApplication("Dashboard",genericObject);
//		webApp.runApplication("Login", genericObject);
//		webApp.runApplication("Dashboard", genericObject);
//		webApp.runApplication("CreateChildNode", genericObject);
//		webApp.runApplication("Workflowapproval", genericObject);
//		webApp.runApplication("BGNodeCreate", genericObject);
//		webApp.runApplication("itemCreate", genericObject);
//		webApp.runApplication("TestcaseUploading", genericObject);
//		webApp.runApplication("selectsourcetf", genericObject);
//		for(int tc = 4226; tc<= 4234; tc++) {
//			String tcid = "TC"+String.valueOf(tc);
//			genericObject.setTcid(tcid);
//			webApp.runApplication("selectdesttfcopy", genericObject);
//		}
//		webApp.runApplication("selectsourcetfdelete", genericObject);
//		for(int tc = 6017; tc<=6097; tc++) {
//			String tcid = "TC"+String.valueOf(tc);
//			genericObject.setTcid(tcid);
//			webApp.runApplication("selectdelete", genericObject);
////			webApp.runApplication("selectdesttfcopy", genericObject);
//		}
	}

}
