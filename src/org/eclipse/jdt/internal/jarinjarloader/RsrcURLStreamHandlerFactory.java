package org.eclipse.jdt.internal.jarinjarloader;

import java.net.URLStreamHandler;
import java.net.URLStreamHandlerFactory;

public class RsrcURLStreamHandlerFactory implements URLStreamHandlerFactory{

	private ClassLoader classLoader;
	private URLStreamHandlerFactory chainfac;
	
	public RsrcURLStreamHandlerFactory(ClassLoader classLoader) {
		this.classLoader = classLoader;
	}
	
	@Override
	public URLStreamHandler createURLStreamHandler(String protocol) {
		if("rsrc".equals(protocol)){
			return new RsrcURLStreamHandler(this.classLoader);
		}
		if(this.chainfac != null){
			return this.chainfac.createURLStreamHandler(protocol);
		}
		return null;
	}

	public void setURLStreamHandlerFactory(URLStreamHandlerFactory fac){
		this.chainfac = fac;
	}
}
