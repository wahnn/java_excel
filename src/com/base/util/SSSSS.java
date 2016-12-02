package com.base.util;

import java.io.UnsupportedEncodingException;

import com.sun.org.apache.bcel.internal.generic.NEW;

public class SSSSS {
	public static void main(String[] args) throws UnsupportedEncodingException {
		String[] columnArray = {"A","B","C","D","E","F","G","H","I","J","K","L","M","N","O","P","Q","R","S","T","U","V","W","X","Y","Z"};
		for (int i = 0; i < columnArray.length; i++) {
			String string = columnArray[i];
			System.out.println(string);
		}
		
		
		String str = "æ¯ä»å®";
		System.out.println(new String(str.getBytes("ISO-8859-1"),"UTF-8"));
	}
}
