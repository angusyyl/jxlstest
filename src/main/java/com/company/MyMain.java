package main.java.com.company;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.Arrays;
import java.util.List;

import org.jxls.common.Context;
import org.jxls.util.JxlsHelper;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.slf4j.profiler.Profiler;
import org.slf4j.profiler.TimeInstrument;

public class MyMain {
	static Logger logger = LoggerFactory.getLogger(MyMain.class); 
			
	public static void main(String[] args) throws FileNotFoundException, IOException {
		MyMain mymain = new MyMain();
		
		Profiler profiler = new Profiler("MyMainProfiler");
		profiler.setLogger(logger);
		profiler.start("genJXLS task");
		
		mymain.genJXLS();
	    
	    TimeInstrument tm = profiler.stop();
	    tm.log();
//	    tm.print();
	}
	
	private void genJXLS() throws FileNotFoundException, IOException {
		logger.info("Running JXLS demo");
		
	    List<Employee> employees = generateSampleEmployeeData();
	    try(InputStream is = new FileInputStream("object_collection_template.xls")) {
	        try (OutputStream os = new FileOutputStream("target/object_collection_output.xls")) {
	            Context context = new Context();
	            context.putVar("employees", employees);
	            JxlsHelper.getInstance().processTemplateAtCell(is, os, context, "Result!A1");
	        }
	    }
	}

	private List<Employee> generateSampleEmployeeData() {
		return Arrays.asList(new Employee("John"), new Employee("Sam"));
	}
}
