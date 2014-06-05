package automation;
import java.io.IOException;
import java.util.HashMap;
import java.util.Iterator;
import java.util.Set;

import org.apache.log4j.Logger;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.firefox.FirefoxDriver;

public class HMPUITesting {

	
	Logger consoleLogger = Logger.getRootLogger();
	
	WebDriver driver;
	
	public HMPUITesting() {
		
	}
	
	public void login() {
		try {
			String[] dialog = new String[] { "loginFF.exe",
//					"Authentication Required", "bteki", "werty123_", "ok" };
					"Authentication Required", "ckatta", "tmon$123", "ok" };
			Process pp1 = Runtime.getRuntime().exec(dialog);
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}
	
	public void sleep(long sleep) {
		try {
			Thread.sleep(sleep);
		} catch (InterruptedException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}
	
	public void nodeCreation() {
		
		String nodeSchema = "START nt=node:HMP_META_INDEX('*:*') " +
				"MATCH  nt-[rel1:hasPropertyGroup]-pg-[rel2:containsProperty]->p WHERE nt.nodeName = \"Technical Support\" " +
				"RETURN pg.nodeName,p.nodeName,p.isMandatory, p.isReadOnly, p.propertyRequired, p.propertyType, p.sequenceNumber;";
		driver = new FirefoxDriver();
		login();
		String[] keyBuilderStrings = {"pg.nodeName", "p.sequenceNumber"};
		HashMap<String, HashMap<String, String>> completeDataMap = (HashMap<String, HashMap<String, String>>) getCypherResult(nodeSchema, keyBuilderStrings);
		driver.close();
		Set<String> keys = completeDataMap.keySet();
		Iterator<String> keyIterator = keys.iterator();
		while(keyIterator.hasNext()) {
			String keyString = keyIterator.next();
			System.out.println(keyString+"--"+completeDataMap.get(keyString));
		}
		
		Set<String> completeDataKeys = completeDataMap.keySet();
		Iterator<String> completeDataIterator = completeDataKeys.iterator();
		while(completeDataIterator.hasNext()) {
			System.out.println(completeDataIterator.next());
		}
		
	}
	
	public void exceldatacreation() {
		
	}	
	
	public Object getCypherResult(String cypherQuery, String[] mapBasedOnKeys) {
		 driver.get("http://hmp-ermo1-dev-03:8080/hmp/ui");
		 driver.get("http://hmp-ermo1-dev-03:8080/hmp/ui/PrintCypherQueryNew.jsp");
		 driver.findElement(By.xpath(".//*[@id='cypherQuery']")).clear();
		 driver.findElement(By.xpath(".//*[@id='cypherQuery']")).sendKeys(cypherQuery);;;
		 driver.findElement(By.xpath(".//textarea[@id='cypherQuery']//following::input[@type='submit']")).click();
		 sleep(5000);
		 String cypherString = driver.findElement(By.xpath("html/body/form/pre")).getText();
//		cypherString = "+---------------------------------------------------------------------------------------------------------------------------+"+"\n"+
//		"| node.nodeName! | rel.start_date! | rel.end_date! | r.nodeName! |"+"\n"+
//		"+---------------------------------------------------------------------------------------------------------------------------+"+"\n"+
//		"| \"Mid-Range/Low-End Routing\" | 2013-04-11 15:00:01:000 | 2013-09-25 08:27:58:118 | \"1000\" |"+"\n"+
//		"| \"Optics - Enterprise Networking Allocation\" | 2013-09-25 21:02:15:000 | 2013-11-10 21:45:18:212 | \"1000\" |"+"\n"+
//		"| \"Mid-Range/Low-End Routing\" | 2013-09-25 08:27:59:000 | 2013-09-25 08:27:59:000 | \"1000\" |"+"\n"+
//		"| \"Optics - Enterprise Networking Allocation\" | 2013-11-10 21:45:19:000 | 2013-11-10 21:45:19:000 | \"1000\" |"+"\n"+
//		"| <null> | 2014-10-08 22:08:09:000 | 2099-12-31 23:59:59:000 | \"1000\" |"+"\n"+
//		"+---------------------------------------------------------------------------------------------------------------------------+"+"\n"+
//		"5 rows"+"\n"+
//		"25 ms";
		consoleLogger.info("Cypher Result");
		consoleLogger.info("---------------");
		consoleLogger.info(cypherString);
		String[] cypherResultLines = cypherString.split("\n");
		consoleLogger.info(cypherResultLines.length);

		/**
		* Header Information
		*/
		HashMap<String, String> headerMap = new HashMap<String, String>();
		HashMap<String, String> dataMap = new HashMap<String, String>();
		HashMap<String, HashMap<String, String>> completeDataMap = new HashMap<String, HashMap<String,String>>();

		for(int i=0; i<cypherResultLines.length; i++) {
			if(i == 0 || i==cypherResultLines.length-1 || i==cypherResultLines.length-2 || i==cypherResultLines.length-3 || i == 2) {
	
			} else if(i == 1) {
				consoleLogger.info(cypherResultLines[i]);
				String[] headerStrings = cypherResultLines[i].split("\\|");
				int k = 0;
				for(String headerString: headerStrings) {
					if(headerString.length() != 0) {
					headerString = headerString.trim();
					headerString = headerString.replace(" ", "");
					headerMap.put(String.valueOf(k+1), headerString);
					consoleLogger.info(headerString);
					k++;
					}
				}
			} else {
				dataMap = new HashMap<String, String>();
				consoleLogger.info(cypherResultLines[i]);
				String[] dataStrings = cypherResultLines[i].split("\\|");
				int k = 0;
				consoleLogger.info("Size:: "+headerMap.size());
				String keyValue = "";
				for(String dataString:dataStrings) {
					if(dataString.length() != 0) {
						dataString = dataString.trim();
						String headerValue = headerMap.get(String.valueOf(k+1));
//						consoleLogger.info(headerValue);
						System.out.println(dataString);
						dataString = dataString.replace("\"", "");
						dataMap.put(headerValue, dataString);
						for(String keyString: mapBasedOnKeys) {
							if(headerValue.equalsIgnoreCase(keyString)) {
								keyValue = keyValue+"@"+dataString;
							}
						}
						consoleLogger.info(dataString);
						k++;
					}
				}
				completeDataMap.put(keyValue, dataMap);
			}
		// consoleLogger.info(i+"--> "+cypherResultLines[i]);
		}
		return completeDataMap;
	}
	
	public static void main(String[] args) {
		HMPUITesting hmpuiTesting = new HMPUITesting();
		hmpuiTesting.nodeCreation();
	}
	
}
