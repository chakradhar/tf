package automation;
import java.io.IOException;
import java.util.Map;

import com.sun.jdi.Bootstrap;
import com.sun.jdi.VirtualMachine;
import com.sun.jdi.VirtualMachineManager;
import com.sun.jdi.connect.Connector.IntegerArgument;
import com.sun.jdi.connect.IllegalConnectorArgumentsException;
import com.sun.jdi.connect.ListeningConnector;


public class VirtualProgram {

	public static void main(String[] args){
		
		try {
			VirtualMachineManager mgr = Bootstrap.virtualMachineManager();
			System.out.println(mgr);
//			System.out.println(mgr.allConnectors().size());
//			ListeningConnector lc = (ListeningConnector)mgr.listeningConnectors().get(0);
//			if (lc == null) {
//			throw new RuntimeException("No com.sun.jdi.SocketListen type found");
//			}
//			
//			Map map = lc.defaultArguments();
//			IntegerArgument arg = (IntegerArgument) map.get("port");
//			arg.setValue(4000);
//			arg = (IntegerArgument) map.get("timeout");
//			arg.setValue(0);
//			lc.startListening(map);
//			VirtualMachine vm = lc.accept(map);
//		} catch (IOException e) {
//			// TODO Auto-generated catch block
//			e.printStackTrace();
//		} catch (IllegalConnectorArgumentsException e) {
//			// TODO Auto-generated catch block
//			e.printStackTrace();
		} catch (NullPointerException ne) {
			ne.printStackTrace();
		}
		
		
	}
	
}
