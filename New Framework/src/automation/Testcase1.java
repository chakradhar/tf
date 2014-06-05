package automation;

import java.io.BufferedReader;
import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileReader;
import java.io.IOException;
import java.util.Iterator;
import java.util.Properties;
import java.util.Set;

import org.junit.Test;

import com.sun.tools.attach.AgentInitializationException;
import com.sun.tools.attach.AgentLoadException;
import com.sun.tools.attach.AttachNotSupportedException;
import com.sun.tools.attach.VirtualMachine;

public class Testcase1 {
	
	@Test
	public void google(Object one) {
		String className = Thread.currentThread().getStackTrace()[1].getClassName();
		String methodName = Thread.currentThread().getStackTrace()[1].getMethodName();
		String fileName= Thread.currentThread().getStackTrace()[1].getFileName();
		System.out.println(className);
		System.out.println(methodName);
		System.out.println(fileName);
		
		try {
			File file= new File("src/"+fileName);
			FileReader fr = new FileReader(file);
			BufferedReader br = new BufferedReader(fr);
			String stringInformation = null;
			while((stringInformation = br.readLine())!=null) {
				if(stringInformation.contains(methodName+"(")){
					System.out.println(stringInformation);
				}
			}
			
			// attach to target VM
		      VirtualMachine vm = VirtualMachine.attach("2820");
		      
		      // get system properties in target VM
		      Properties props = vm.getSystemProperties();

		      // construct path to management agent
		      String home = props.getProperty("java.home");
		      String agent = home + File.separator + "lib" + File.separator
		          + "management-agent.jar";

		      Set keys = props.keySet();
		      
		      Iterator<Object> keyIterator = keys.iterator();
		      
		      while(keyIterator.hasNext()) {
		    	  System.out.println(keyIterator.next());
		      }
		      
//		      this.
		      
		      // load agent into target VM
//		      vm.loadAgent(agent, "com.sun.management.jmxremote.port=5000");

		      // detach
		      vm.detach();
		      
		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (AttachNotSupportedException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
//		} catch (AgentLoadException e) {
//			// TODO Auto-generated catch block
//			e.printStackTrace();
//		} catch (AgentInitializationException e) {
//			// TODO Auto-generated catch block
//			e.printStackTrace();
		}
		
//		Runtime runtime = Runtime.getRuntime();
//		GenericObject genericObject = new GenericObject();
//		Properties properties = new Properties();
//		properties.put("url", "http://www.google.com");
//		properties.put("text", "hi");
//		genericObject.setProperties(properties);
//		WebApp webApp = new WebApp();d
//		webApp.runApplication(methodName, genericObject);
//		System.out.println("google");
	}

}
