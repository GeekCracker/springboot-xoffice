package com.hg.task;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.logging.Logger;
import java.util.regex.Pattern;

import org.springframework.beans.factory.annotation.Value;
import org.springframework.scheduling.annotation.Scheduled;
import org.springframework.stereotype.Component;

import com.hg.xs.ExcelApp;
import com.hg.xs.ImageApp;
import com.hg.xs.PowerPointApp;
import com.hg.xs.WordApp;
import com.hg.xs.XOfficeServlet;

/**
 * 生成转化后的文件的定时任务
 */
@Component
public class CreateFileTask {

	private static Logger logger = Logger.getLogger(XOfficeServlet.class.getName());
	
	@Value("${CURRENT_SERVER_FOR_HTML}")
	private String CURRENT_SERVER_FOR_HTML;
	
	@Value("${FILE_ROOT_PATH}")
	private String FILE_ROOT_PATH;

	@Value("${FILE_ROOT_SRC_PATH}")
	private String FILE_ROOT_SRC_PATH;
	
	//@Scheduled(cron="0/1 * * * * ?")//每一秒执行一次
	@Scheduled(cron="0 0 0 * * ?")//每天0:0:0 点执行
	public void createTask(){
		logger.info("转化office预览文件的定时任务启动");
		long s = System.currentTimeMillis();
		doActDeep(FILE_ROOT_SRC_PATH);
		logger.info("定时任务执行完毕>>耗时"+(System.currentTimeMillis() - s)+"毫秒");
	}
	
	/**
	 * 递归生成资源目下的文件的对应预览文件
	 * 
	 * @param srcPath
	 *            传入文件或文件夹
	 */
	public void doActDeep(String srcPath) {
		if (srcPath != null) {
			File file = new File(srcPath);
			if (file.exists()) {
				if (file.isFile()) {
					// 这里将资源文件转化成pdf和html各一份
					createFile(file.getAbsolutePath());
				} else {
					File[] files = file.listFiles();
					for (File f : files) {
						doActDeep(f.getAbsolutePath());// 递归删除传入的文件夹下的文件或文件夹
					}
				}
			} else {
				throw new RuntimeException("要转化的文件不存在");
			}
		} else {
			throw new RuntimeException("删除文件夹时传入的路径为空");
		}
	}

	/**
	 * 将传入的文件生成html和pdf各一份
	 * 
	 * @param filePath
	 *            传入文件绝对路径
	 */
	private void createFile(String filePath) {
		// 1.获取目标文件目录
		File src = new File(FILE_ROOT_PATH + filePath.replace(FILE_ROOT_SRC_PATH, ""));
		File parentFile = src.getParentFile();
		if (!parentFile.exists()) {
			if (!parentFile.mkdirs()) {
				return;
			}
		}
		FileInputStream in = null;
		FileOutputStream out = null;
		try {
			// 2.将传入的文件拷贝一份到指定目录下，如果目标文件不存在的话再拷贝
			if(!src.exists()){
				in = new FileInputStream(filePath);
				out = new FileOutputStream(src);
				pipe(in, out);
			}
			// 3.声明生成的文件的目录
			String srcPath = src.getAbsolutePath();
			String xformat = srcPath.substring(srcPath.lastIndexOf(".") + 1);
			// 3.1生成pdf文件
			File pdfFile = new File(srcPath.substring(0, srcPath.lastIndexOf(".")) + ".pdf");
			// 3.2生成html文件
			File htmlFile = new File(srcPath.substring(0, srcPath.lastIndexOf(".")) + ".html");
			// 3.3生成文件
			if (xformat.startsWith("doc")) {
				if (!pdfFile.exists()) {
					WordApp.toPdf(src.getAbsolutePath(), pdfFile.getAbsolutePath());
				}
				if (!htmlFile.exists()) {
					WordApp.toHtml(src.getAbsolutePath(), htmlFile.getAbsolutePath());
				}
			} else if (xformat.startsWith("xls")) {
				if (!pdfFile.exists()) {
					ExcelApp.toPdf(src.getAbsolutePath(), pdfFile.getAbsolutePath());
				}
				if (!htmlFile.exists()) {
					ExcelApp.toHtml(src.getAbsolutePath(), htmlFile.getAbsolutePath());
				}
			} else if (xformat.startsWith("ppt")) {
				if (!pdfFile.exists()) {
					PowerPointApp.toPdf(src.getAbsolutePath(), pdfFile.getAbsolutePath());
				}
				//ppt不支持转html
				/*if (!htmlFile.exists()) {
					logger.info(src.getAbsolutePath());
					PowerPointApp.toHtml(src.getAbsolutePath(), htmlFile.getAbsolutePath());
				}*/
			} else if (Pattern.matches("^(png|jpg|gif|jpeg|jpe)$", xformat)) {
				if (!pdfFile.exists()) {
					ImageApp.toPdf(src.getAbsolutePath(), pdfFile.getAbsolutePath());
				}
				if (!htmlFile.exists()) {
					ImageApp.toHtml(src.getAbsolutePath(), htmlFile.getAbsolutePath(), CURRENT_SERVER_FOR_HTML, srcPath
							.replace(FILE_ROOT_PATH + "\\", "").replace(src.getName(), "").replaceAll("\\\\", "/"));
				}
			}
		} catch (Throwable e) {
			e.printStackTrace();
		} finally {
			if (in != null) {
				try {
					in.close();
				} catch (Exception e) {
				}
			}
			if (out != null) {
				try {
					out.close();
				} catch (Exception e) {
				}
			}
		}
	}

	private static void pipe(InputStream in, OutputStream out) throws IOException {
		int len;
		byte[] buf = new byte[4096];
		while (true) {
			len = in.read(buf);
			if (len > 0) {
				out.write(buf, 0, len);
			} else {
				break;
			}
		}
		out.flush();
		in.close();
		out.close();
	}
}
