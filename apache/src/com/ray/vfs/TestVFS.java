package com.ray.vfs;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;

import org.apache.commons.vfs2.FileContent;
import org.apache.commons.vfs2.FileObject;
import org.apache.commons.vfs2.FileSystemException;
import org.apache.commons.vfs2.FileSystemManager;
import org.apache.commons.vfs2.VFS;

public class TestVFS {

	public static void main(String[] args) {
		// TODO Auto-generated method stub

		try {
			FileSystemManager fsManager = VFS.getManager();
			FileObject jarFile = fsManager.resolveFile( "jar:///home/rayweihao/POI/commons-logging-1.2.jar" );
			
			getAllChildren(jarFile);
		} catch (FileSystemException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}
	
	public static void getAllChildren(FileObject fileObject)
	{
		if(null != fileObject) {
			try {
				if(fileObject.isFolder()) {
					FileObject[] children = fileObject.getChildren();
					if(null != children && children.length > 0) {
						for ( int i = 0; i < children.length; i++ )
						{
						    System.out.println( children[ i ].getName());
						    if(children[i].isFile()) {
						    	FileContent content = children[ i ].getContent();
							    if(null != content)
							    {
							    	File file = new File("/home/rayweihao/mytemp/" + children[ i ].getName().getBaseName());
							    	FileOutputStream output = new FileOutputStream(file);
							    	content.write(output);
							    	output.close();
							    }
						    }
						    if(null != children[i]) {
						    	getAllChildren(children[i]);
						    }
						}
					}
				}
			} catch (FileSystemException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			} catch (FileNotFoundException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			} catch (IOException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
		}
	}

}
