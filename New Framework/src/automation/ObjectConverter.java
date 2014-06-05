package automation;
import java.lang.reflect.Method;


public class ObjectConverter {

	public ObjectConverter() {		
	}
	
	public Object convertObject(Object object) {
		Method[] methods = object.getClass().getDeclaredMethods();
		
		for(Method method: methods) {
			
			
			
		}
		
		return null;
	}
	
}
