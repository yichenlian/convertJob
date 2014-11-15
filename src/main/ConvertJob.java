package main;

import java.io.BufferedReader;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.io.OutputStream;
import java.util.ArrayList;

import com.artofsolving.jodconverter.DefaultDocumentFormatRegistry;
import com.artofsolving.jodconverter.DocumentConverter;
import com.artofsolving.jodconverter.DocumentFormat;
import com.artofsolving.jodconverter.openoffice.connection.OpenOfficeConnection;
import com.artofsolving.jodconverter.openoffice.connection.SocketOpenOfficeConnection;
import com.artofsolving.jodconverter.openoffice.converter.OpenOfficeDocumentConverter;
import com.lowagie.text.Document;
import com.lowagie.text.Font;
import com.lowagie.text.Paragraph;
import com.lowagie.text.pdf.BaseFont;
import com.lowagie.text.pdf.PdfWriter;

import util.ConfigManager;


public class ConvertJob {
	/**
	 * source file
	 */
	private static final String SOURCE_PATH = ConfigManager.instance().getProperty("respath", "source_path");
	/**
	 * openoffice服务ip
	 */
	private static final String DEFAULT_HOST = ConfigManager.instance().getProperty("convert", "doc_convert_ip");
	/**
	 * openoffice服务端口
	 */
	private static final int DEFAULT_PORT = ConfigManager.instance().getIntProperty("convert", "doc_convert_port");
	/**
	 * swftools工具安装路径
	 */
	private static final String SWFTOOLS_PATH = ConfigManager.instance().getProperty("convert", "swftools_path");
	/**
	 * 字体路径
	 */
	private static final String FONT_PATH = ConfigManager.instance().getProperty("convert", "font_path");
	/**
	 * 系统类型：windows或者linux
	 */
	private static final String SYSTEM_KIND = ConfigManager.instance().getProperty("convert", "system_kind");
	/**
	 * pdf临时文件夹
	 */
	private static final String TEMP_PDF_DIR = ConfigManager.instance().getProperty("convert", "temp_folder_pdf");
	/**
	 * swf临时文件夹
	 */
	private static final String TEMP_SWF_DIR = ConfigManager.instance().getProperty("convert", "temp_folder_swf");
	/**
	 * 图片临时文件夹
	 */
	private static final String TEMP_IMAGE_DIR = ConfigManager.instance().getProperty("convert", "temp_folder_img");
	
	
	private static DocumentConverter converter; 

	private static OpenOfficeConnection connection =null;
	
	static{
		/**
		 * 创建相关文件夹
		 */
		makeDirs(new File(TEMP_PDF_DIR));
		makeDirs(new File(TEMP_SWF_DIR));
		makeDirs(new File(TEMP_IMAGE_DIR));
	}

	private static void makeDirs(File dirFile){
		if(!dirFile.exists()){
			dirFile.mkdirs();
		}
	}
	/**
	 * 转码成功后将图片和pdf删除
	 * @param file
	 */
	private static void delFile(File file){
		if(file.exists()){
			file.delete();
		}
	}
	/**
	 * 得到异常的堆栈信息
	 * @param e
	 * @return
	 */
	public static String getExceptionStack(Exception e) {
		StackTraceElement[] stackTraceElements = e.getStackTrace();
		String result = e.toString() + "\n";
		for (int index = stackTraceElements.length - 1; index >= 0; --index) {
			result += "        at " + stackTraceElements[index].getClassName() + ".";
			result += stackTraceElements[index].getMethodName() + "(";
			result += stackTraceElements[index].getFileName() + ".";
			result += stackTraceElements[index].getLineNumber() + ")\n";
		}
		return result;
	}
	/**
	 * 为pdf和gif使用
	 * @param filePath
	 * @return
	 */
	static int getPdfPageCount(String filePath) {
		ProcessBuilder processBuilder = null;
		if("windows".equalsIgnoreCase(SYSTEM_KIND)){
			processBuilder = new ProcessBuilder(SWFTOOLS_PATH+"pdf2swf", filePath, "-I");
		}else{
			String arg = String.format(SWFTOOLS_PATH+"pdf2swf %s  -I  |grep page= | wc -l",
					filePath);
			processBuilder = new ProcessBuilder("sh", "-c", arg);
		}
		processBuilder.redirectErrorStream(true);
		System.out.println(processBuilder.command().toString());
		try {
			Process process = processBuilder.start();
			BufferedReader reader = new BufferedReader(new InputStreamReader(
					process.getInputStream()));
			String lastLine = "";
			String line = null;
			int winPages = 0;
			while ((line = reader.readLine()) != null) {
				lastLine = line;
				if("windows".equalsIgnoreCase(SYSTEM_KIND) && lastLine.contains("page")){
					winPages+=1;
				}
			}
			process.waitFor();			
			if("windows".equalsIgnoreCase(SYSTEM_KIND)){
				return winPages;
			}
			return intValueOf(lastLine);

		} catch (Exception e) {
			System.out.println(getExceptionStack(e));
		}

		return 0;
	}
	static int intValueOf(Object object) {
		try {
			return Integer.valueOf(String.valueOf(object));
		} catch (Exception e) {
			System.out.println(getExceptionStack(e));
		}

		return 0;
	}



	/**
	 * 将文档转换成pdf
	 * @param docJson
	 */
	public static boolean  docsToPdf(File srcFile){

		InputStream inputStream = null;
		try{
			inputStream = new FileInputStream(srcFile);
		}catch(Exception e){
			System.out.println(e.getMessage()+"\r\n"+getExceptionStack(e));
		}
		if(inputStream==null){
			return false;
		}
		OutputStream outputStream =null;
		String fileName = srcFile.getName();
		// 获得文件格式
		DefaultDocumentFormatRegistry formatReg = new DefaultDocumentFormatRegistry();
		DocumentFormat pdfFormat = formatReg.getFormatByFileExtension("pdf");
		DocumentFormat xlsFormat = null;
		String ext = fileName.substring(fileName.lastIndexOf(".")+1);
		String pdfFile = TEMP_PDF_DIR+fileName.replace(ext, "pdf");
		String swfFile = TEMP_SWF_DIR+fileName.replace(ext, "swf");
		String swfFileName = fileName.replace(ext, "swf");
		File targetFile = new File(pdfFile);
		if("pdf".equals(ext)){
			try{
				outPutFile(inputStream, targetFile);
				if(!targetFile.exists()){
					return false;
				}
				return pdfToSwf(pdfFile, swfFile, swfFileName);//将pdf转码成swf
			}catch(Exception e){
				System.out.println(getExceptionStack(e));
			}

		}else if("jpg".equalsIgnoreCase(ext)||"jpeg".equalsIgnoreCase(ext)||"gif".equalsIgnoreCase(ext)||"png".equalsIgnoreCase(ext)){
			targetFile = new File(TEMP_IMAGE_DIR+fileName);
			try{
				outPutFile(inputStream, targetFile);
				if(!targetFile.exists()){
					return false;
				}
				return imageToSwf(TEMP_IMAGE_DIR+fileName, swfFile, swfFileName,ext);//图片转码

			}catch(Exception e){
				System.out.println(getExceptionStack(e));
			}
			return true;

		}else if("txt".equals(ext)){
			try{
				txtToPdf(inputStream, targetFile);
			}catch(Exception e){
				System.out.println(getExceptionStack(e));
			}
			return pdfToSwf(pdfFile, swfFile, swfFileName);//将pdf转码成swf
		}else{
			xlsFormat = formatReg.getFormatByFileExtension(ext);
		} // stream 流的形式
		try {
			outputStream = new FileOutputStream(targetFile);
			connection = new SocketOpenOfficeConnection(DEFAULT_HOST,DEFAULT_PORT);
			connection.connect();
			converter = new OpenOfficeDocumentConverter(connection);

			converter.convert(inputStream, xlsFormat, outputStream, pdfFormat);
			if(!targetFile.exists()){
				return false;
			}
			return pdfToSwf(pdfFile, swfFile, swfFileName);//将pdf转码成swf

		} catch (Exception e) {
			System.out.println(e.getMessage()+"\r\n"+getExceptionStack(e));
		} finally {
			if (connection != null) {
				connection.disconnect();
				connection = null;
			}
		}
		return true;
	}
	
	/**
	 *
	 * @param iStream
	 * @param outFile
	 */
	public static void outPutFile(InputStream fis,File outFile){
		try{
			OutputStream os = new FileOutputStream(outFile);
			byte[] bytes = new byte[1024];
			int len = 0;
			while ((len = fis.read(bytes)) > 0) {
				os.write(bytes, 0, len);
				os.flush();
			}
			os.close();
			fis.close();
		}catch(Exception e){
			System.out.println(e.getMessage()+"\r\n"+getExceptionStack(e));
		}
	}

	public static void txtToPdf(InputStream is, File fileOut) throws Exception{
		Document doc = new Document();
		FileOutputStream out = new FileOutputStream(fileOut);
		PdfWriter.getInstance(doc, out);
		BaseFont bfHei = BaseFont.createFont(FONT_PATH, BaseFont.IDENTITY_H,
				BaseFont.NOT_EMBEDDED);
		Font font = new Font(bfHei, 12);
		BufferedReader reader = new BufferedReader(new InputStreamReader(is));
		String content = null;
		StringBuffer bu = new StringBuffer();
		doc.open();
		while ((content = reader.readLine()) != null) {
			bu.append(content);
		}
		Paragraph text = new Paragraph(bu.toString(), font);
		doc.add(text);
		doc.close();
		reader.close();
		if (is != null)
			is.close();
	}
	

	private static String bytesToHexString(byte[] src){     
		StringBuilder stringBuilder = new StringBuilder();     
		if (src == null || src.length <= 0) {     
			return null;     
		}     
		for (int i = 0; i < src.length; i++) {     
			int v = src[i] & 0xFF;     
			String hv = Integer.toHexString(v);     
			if (hv.length() < 2) {     
				stringBuilder.append(0);     
			}     
			stringBuilder.append(hv);     
		}     
		return stringBuilder.toString();     
	} 
	/**
	 * 根据流判断文件类型
	 * @param is
	 * @return
	 */
	public static String getImageExt(File file){
		String jpegStr = "ffd8ff";
		String pngStr = "89504e47";
		String gifStr = "47494638";
		String extStr = null;
		InputStream is = null;
		try{
			is = new FileInputStream(file);
			byte[] b = new byte[4];  
			is.read(b, 0, b.length);  
			String hexStr = bytesToHexString(b).toLowerCase();
			if(hexStr.indexOf(jpegStr)!=-1&&jpegStr.equals(hexStr.substring(0, 6))){
				extStr = "jpeg";
			}else if(pngStr.equals(hexStr)){
				extStr = "png";
			}else if(gifStr.equals(hexStr)){
				extStr = "gif";
			}
		}catch(Exception e){
			System.out.println("获取图片二进制编码时出错:"+e.getMessage()+"\r\n"+getExceptionStack(e));
		}finally{
			try {
				is.close();
			} catch (IOException e) {
				System.out.println("关闭流失败:"+e.getMessage()+"\r\n"+getExceptionStack(e));
			}
		}
		return extStr;
	}

	/**
	 * 将图片转换成swf
	 * @param srcPath
	 * @param destPath
	 * @param fileName
	 */
	public static boolean imageToSwf(String srcPath, String destPath, String fileName,String fileType){
		//工具中的转换图片的程序是根据图片的二进制编码区别的，故需要通过二进制编码来实现
		String extStr = getImageExt(new File(srcPath));
		ProcessBuilder processBuilder = null;
		if("gif".equals(extStr)){
			processBuilder = new ProcessBuilder(SWFTOOLS_PATH+extStr+"2swf", srcPath,"-o",destPath);
		}else{
			processBuilder = new ProcessBuilder(SWFTOOLS_PATH+extStr+"2swf", "-z",srcPath,"-o",destPath, "-f","-T","9");
		}
		processBuilder.redirectErrorStream(true); // error to input stream, to avoid block
		System.out.println(String.format(extStr+"2swf"+" command:%s",processBuilder.command().toString()));

		try{
			Process process = processBuilder.start();			
			int exitCode = process.waitFor();
			File f = new File(destPath);
			if (exitCode == 0 && f.exists()) { // ok
				return true;
			} else { // fail
				return false;
			}
		}catch(Exception e){
			return false;
		}
	}
	/**
	 * 将pdf转换成swf
	 * @param srcPath
	 * @param destPath
	 * @param fileName
	 */
	public static boolean pdfToSwf(String srcPath, String destPath, String fileName){
		ProcessBuilder processBuilder = new ProcessBuilder(SWFTOOLS_PATH+"pdf2swf", "-z",
				"-s", "flashversion=9", "-s", "poly2bitmap", "-t",srcPath,"-o",destPath);

		processBuilder.redirectErrorStream(true); // error to input stream, to avoid block
		System.out.println(String.format("pdf2swf command:%s",processBuilder.command().toString()));

		try{
			Process process = processBuilder.start();			
			int exitCode = process.waitFor();
			File f = new File(destPath);
			if (exitCode == 0 && f.exists()) { // ok
				int pageNum = getPdfPageCount(srcPath);
				System.out.println("文档'"+srcPath+"'的页码是："+pageNum);
				return true;
			} else { // fail
				return false;
			}
		}catch(Exception e){
			e.printStackTrace();
			return false;
		}
	}
	/**
	 * 递归加载源目录下的文件
	 * @param file
	 * @return
	 */
	public static ArrayList<File> getFiles(File file){
		if(file!=null){
			ArrayList<File> list = new ArrayList<File>();
			if(file.isFile()){
				list.add(file);
				return list;
			}else if(file.isDirectory()){
				File[] arr = file.listFiles();
				for(File arrFile:arr){
					if(arrFile!=null){
						list.addAll(getFiles(arrFile));
					}
				}
			}
			return list;
		}
		return null;
	}
	
	public static void main(String[] args) {
		File file = new File(SOURCE_PATH);
		ArrayList<File> result = getFiles(file);
		if(result!=null){
			for(File target:result){
				if(target!=null){
					docsToPdf(target);
				}
			}
		}
	}

}
