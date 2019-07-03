package com.hg.xs;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.net.URL;
import java.net.URLEncoder;
import java.util.HashMap;
import java.util.Map;
import java.util.logging.Level;
import java.util.logging.Logger;
import java.util.regex.Pattern;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

import org.springframework.beans.factory.annotation.Value;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

@RestController("XofficeServlet")
@RequestMapping("/xoffice")
public class XOfficeServlet {

	private static Logger logger = Logger.getLogger(XOfficeServlet.class.getName());

	// 默认清除生成的缓存文件
	private boolean clearTempFile = true;

	@Value("${CURRENT_SERVER_FOR_HTML}")
	private String CURRENT_SERVER_FOR_HTML;

	@Value("${FILE_ROOT_PATH}")
	private String FILE_ROOT_PATH;

	@Value("${FILE_ROOT_SRC_PATH}")
	private String FILE_ROOT_SRC_PATH;

	@RequestMapping("xoffice")
	public void xoffice(HttpServletRequest request, HttpServletResponse response) throws IOException {
		doAct(request, response, false);
	}

	@RequestMapping("oneKeyAct")
	public Map<String, String> oneKeyAct() {
		Map<String, String> result = new HashMap<String, String>();
		logger.info("--开始一键转换--");
		long s = System.currentTimeMillis();
		doActDeep(FILE_ROOT_SRC_PATH);
		System.out.println("耗时"+(System.currentTimeMillis() - s)+"毫秒");;
		logger.info("--转换成功--");
		result.put("code", "200");
		result.put("msg", "操作成功");
		return result;
	}

	public void doAct(HttpServletRequest request, HttpServletResponse response, boolean post) throws IOException {
		// 是否清除缓存文件
		String reqkey = request.getParameter("_key");
		if ("false".equals(reqkey)) {
			this.clearTempFile = false;
		}
		// _xformat 参数指的是需要预览的源文件的文件格式，通常指的是office文件的文件格式
		String xformat = request.getHeader("_xformat");
		if (xformat == null) {
			xformat = request.getParameter("_xformat");
		}
		if (xformat == null) {
			logger.info("No Paramater '_xformat'!");
			return;
		}
		// _format 参数指的是将需要预览的文件已什么方式进行预览，可以为html，pdf等
		String format = request.getHeader("_format");
		if (format == null) {
			format = request.getParameter("_format");
		}
		if (format == null) {
			logger.info("No Paramater '_format'!");
			return;
		}
		String urlFile = null;
		if (!post) {
			urlFile = request.getParameter("_file");
			if (urlFile == null) {
				logger.info("No Paramater '_file'!");
				return;
			}
		}
		// 先判断服务器上是否有转换好的文件，如果有的话，就直接读取该文件，如果没有则执行转化的操作
		// 获取转化后的文件
		String relativePath = urlFile.substring(urlFile.indexOf("/", urlFile.indexOf("//") + 2));
		File file = null;
		if(xformat.startsWith("pdf")){
			file = new File(FILE_ROOT_PATH + relativePath.substring(0,relativePath.lastIndexOf(".")) + ".pdf");
		}else {
			file = new File(FILE_ROOT_PATH + relativePath.substring(0,relativePath.lastIndexOf(".")) + "." + format);
		}
		logger.info("本地文件："+file.getAbsolutePath());
		logger.info("是否存在?");
		if (file != null && file.exists()) {
			logger.info("是");
			InputStream in = null;
			OutputStream out = null;
			try {
				String filePath = file.getAbsolutePath();
				if (format.startsWith("html") && !xformat.contains("pdf")) {
					String urlPath = filePath.replace(FILE_ROOT_PATH,"").replaceAll("\\\\","/");
					urlPath = urlPath.substring(0,urlPath.lastIndexOf("/")+1) +URLEncoder.encode(urlPath.substring(urlPath.lastIndexOf("/")+1),"UTF-8") ;
					response.sendRedirect("http://" + CURRENT_SERVER_FOR_HTML + urlPath);
				} else {
					out = response.getOutputStream();
					if(xformat.contains("pdf")){
						String path = file.getAbsolutePath();
						in = new FileInputStream(path.substring(0,path.lastIndexOf(".")) +".pdf");
					}else {
						in = new FileInputStream(file);
					}
					pipe(in, out);
				}
			} catch (Exception e) {
				e.printStackTrace();
			} finally{
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
		} else {
			long s = System.currentTimeMillis();
			logger.info("否");
			logger.info("Convert start...");
			/*
			 * String fileName = System.getProperty("java.io.tmpdir"); if
			 * (!fileName.endsWith("/") && !fileName.endsWith("\\")) { fileName
			 * += "/"; }
			 */
			String fileName = FILE_ROOT_PATH + "/";
			String subPath = relativePath.substring(1,relativePath.lastIndexOf("/")+1);
			fileName += subPath;
			//fileName += URLDecoder.decode(file.getName().split("\\.")[0],"UTF-8");
			fileName += file.getName().substring(0,file.getName().lastIndexOf("."));
			File src = new File(fileName+"."+xformat);// 资源文件
			File tar = new File(fileName+"."+format);// 转换后的文件
			logger.info(src.getAbsolutePath() + " >>> " + tar.getAbsolutePath());
			File parentFile = src.getParentFile();
			if (!parentFile.exists()) {
				if (!parentFile.mkdirs()) {
					return;
				}
			}
			InputStream in = null;
			OutputStream out = null;
			in = new URL(urlFile.substring(0, urlFile.lastIndexOf("/")+1) + URLEncoder.encode(urlFile.substring(urlFile.lastIndexOf("/")+1),"UTF-8")).openStream();
			//in = new URL(urlFile).openStream();
			out = new FileOutputStream(src);
			pipe(in, out);
			try {
				if (xformat.startsWith("doc")) {
					if (format.startsWith("pdf")) {
						WordApp.toPdf(src.getAbsolutePath(), tar.getAbsolutePath());
					} else if (format.startsWith("html")) {
						WordApp.toHtml(src.getAbsolutePath(), tar.getAbsolutePath());
					}
				} else if (xformat.startsWith("xls")) {
					if (format.startsWith("pdf")) {
						ExcelApp.toPdf(src.getAbsolutePath(), tar.getAbsolutePath());
					} else if (format.startsWith("html")) {
						ExcelApp.toHtml(src.getAbsolutePath(), tar.getAbsolutePath());
					}
				} else if (xformat.startsWith("ppt")) {
					if (format.startsWith("pdf")) {
						PowerPointApp.toPdf(src.getAbsolutePath(), tar.getAbsolutePath());
					} else if (format.startsWith("html")) {
						PowerPointApp.toHtml(src.getAbsolutePath(), tar.getAbsolutePath());
					}
				} else if (Pattern.matches("^(png|jpg|gif|jpeg|jpe|JPG|PNG)$", xformat)) {
					if (format.startsWith("pdf")) {
						ImageApp.toPdf(src.getAbsolutePath(), tar.getAbsolutePath());
					} else if (format.startsWith("html")) {
						ImageApp.toHtml(src.getAbsolutePath(), tar.getAbsolutePath(), CURRENT_SERVER_FOR_HTML, subPath);
					}
				}
				if (format.startsWith("html") && !xformat.contains("pdf")) {
					response.sendRedirect("http://" + CURRENT_SERVER_FOR_HTML + "/" + subPath + URLEncoder.encode(tar.getName(),"UTF-8"));
				} else {
					out = response.getOutputStream();
					if(xformat.contains("pdf")){
						String path = src.getAbsolutePath();
						in = new FileInputStream(path.substring(0,path.lastIndexOf("."))+".pdf");
					}else {
						in = new FileInputStream(tar);
					}
					pipe(in, out);
				}
			} catch (Throwable t) {
				response.setStatus(500);
				logger.log(Level.SEVERE, t.getMessage(), t);
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
				if (clearTempFile && !format.startsWith("html")) {
					src.delete();
					tar.delete();
				}
			}
			logger.info("Convert stop,Use " + (System.currentTimeMillis() - s) + " ms!");
		}
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
			// 2.将传入的文件拷贝一份到指定目录下
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
				if (!htmlFile.exists()) {
					PowerPointApp.toHtml(src.getAbsolutePath(), htmlFile.getAbsolutePath());
				}
			} else if (Pattern.matches("^(png|jpg|gif|jpeg|jpe|JPG|PNG)$", xformat)) {
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
