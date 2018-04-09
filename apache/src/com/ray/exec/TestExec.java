package com.ray.exec;

import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.HashMap;

import org.apache.commons.exec.CommandLine;
import org.apache.commons.exec.DefaultExecuteResultHandler;
import org.apache.commons.exec.DefaultExecutor;
import org.apache.commons.exec.ExecuteStreamHandler;
import org.apache.commons.exec.ExecuteWatchdog;
import org.apache.commons.exec.Executor;
import org.apache.commons.exec.PumpStreamHandler;

public class TestExec {
	
	private static final String SHFILEPATH = "/home/rayweihao/test.sh";

	public static void main(String[] args) {
		// TODO Auto-generated method stub

		CommandLine cmdLine = new CommandLine(SHFILEPATH);
		
		cmdLine.addArgument("-u=infodba");
		
		cmdLine.addArgument("-p=infodba");
		
		cmdLine.addArgument("-g=dba");
		
		cmdLine.addArgument("-file=${file}");
		
		HashMap map = new HashMap();
		map.put("file", new File(SHFILEPATH));
		cmdLine.setSubstitutionMap(map);
		
		DefaultExecuteResultHandler resultHandler = new DefaultExecuteResultHandler();
		ByteArrayOutputStream baos = new ByteArrayOutputStream();
		ExecuteStreamHandler streamHandler = new PumpStreamHandler(baos);
		
		ExecuteWatchdog watchdog = new ExecuteWatchdog(5000);
		Executor executor = new DefaultExecutor();
		executor.setExitValue(1);
		executor.setWatchdog(watchdog);
		executor.setStreamHandler(streamHandler);
		
		try {
			executor.execute(cmdLine, resultHandler);
			
			resultHandler.waitFor();
			
			System.out.println(resultHandler.getExitValue());
			
			String result = baos.toString().trim();
			
			System.out.println(result);
			
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (InterruptedException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}

}
