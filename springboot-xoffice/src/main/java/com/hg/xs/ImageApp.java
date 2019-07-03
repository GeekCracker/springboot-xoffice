package com.hg.xs;

import java.io.BufferedWriter;
import java.io.File;
import java.io.FileOutputStream;
import java.io.OutputStreamWriter;
import java.net.URLEncoder;
import java.util.logging.Logger;

import com.itextpdf.text.Document;
import com.itextpdf.text.Image;
import com.itextpdf.text.Paragraph;
import com.itextpdf.text.Rectangle;
import com.itextpdf.text.pdf.PdfWriter;

/**
 * @author zhujunliang
 */
public class ImageApp {

	private static Logger logger = Logger.getLogger(WordApp.class.getName());
	
	public static void toHtml(String src,String tar,String server,String subPath) throws Exception{
		logger.info("开始创建HTML文件...");
		File tarFile = new File(tar);
		if(tarFile.createNewFile()){
			logger.info("HTML文件创建成功！");
		}else {
			logger.info("HTML文件创建失败！");
			return ;
		}
		BufferedWriter bw = new BufferedWriter(new OutputStreamWriter(new FileOutputStream(tarFile)));
		bw.write("<html !DOCTYPE>");
		bw.write("<meta charset=\"utf-8\">");
		bw.write("<meta http-equiv='Content-Type' content='text/html; charset=utf-8' />");
		bw.write("<body style='text-align:center;'>");
		bw.write("<img height='100%' src = 'http://"+server+"/" + subPath +URLEncoder.encode( new File(src).getName(),"UTF-8")+"' />");
		bw.write("</body>");
		bw.write("</html>");
		bw.flush();
		bw.close();
	}
	/**
	 * 将图片文件转换为pdf文件 单个图片文件
	 * @param src 传入图片文件绝对路径
	 * @param tar 传入pdf文件绝对路径
	 * @throws Exception 抛出可能出现的异常
	 */
	public static void toPdf(String src,String tar) throws Exception{
		logger.info("开始创建PDF文件...");
		File tarFile = new File(tar);
		if(tarFile.createNewFile()){
			logger.info("PDF文件创建成功！");
		}else {
			logger.info("PDF文件创建失败！");
			return ;
		}
		logger.info("开始转换...");
		//获取pdf文件输入流
		FileOutputStream fis = new FileOutputStream(tar);
		// 创建文档
		Document doc = new Document(null, 0, 0, 0, 0);
		//将文档写入pdf文件中
		PdfWriter.getInstance(doc, fis);
		// 根据图片大小设置文档大小
		//设置文档为A4纸张大小
		doc.setPageSize(new Rectangle(595,842));
		// 实例化图片
		Image image = Image.getInstance(src);
		image.setAlignment(Image.ALIGN_CENTER);
		//设置图片压缩后的大小
		image.scaleToFit(297.5f, 421f);
		// 添加图片到文档
		doc.open();
		doc.addTitle("打印");
		//文档中添加换行
		Paragraph text = new Paragraph("\n\n\n\n");
		doc.add(text);
		doc.add(image);
		// 关闭文档
		doc.close();
		logger.info("转换成功！");
	}
}
