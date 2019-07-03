package com.hg.task;

import java.io.File;
import java.util.Calendar;

import org.springframework.beans.factory.annotation.Value;
import org.springframework.scheduling.annotation.Scheduled;
import org.springframework.stereotype.Component;

/**
 * 删除历史文件的定时任务
 */
//@Component
public class RemoveHistoryFile {

	@Value("${FILE_ROOT_PATH}")
	private String FILE_ROOT_PATH;
	
	//@Scheduled(cron="0 0/5 0 * * ?")//每五分钟执行一次
	//@Scheduled(cron="0/1 * * * * ?")//每一秒执行一次
	//@Scheduled(cron="0 0 0/23 * * ?")//每23小时执行一次
	//@Scheduled(cron="0 0 0 * * ?")//每天0:0:0 点执行
	public void remove(){
		//获取需要删除的文件夹
		Calendar calendar = Calendar.getInstance();
		//获取当前年份，月份，及一月中的当前天
		Integer yearCur = calendar.get(Calendar.YEAR);
		Integer monthCur = calendar.get(Calendar.MONTH) + 1;
		Integer dateCur = calendar.get(Calendar.DAY_OF_MONTH);
		//删除当前年之前的文件夹
		File pack = new File(FILE_ROOT_PATH);
		if(pack.exists()){
			File [] packYear = pack.listFiles();
			for(File file : packYear){
				Integer year = Integer.parseInt(file.getName());
				if(year < yearCur){
					dropDir(file.getAbsolutePath());
				}
			}
		}
		//删除当月之前的文件夹
		pack = new File(FILE_ROOT_PATH+File.separator+yearCur+File.separator);
		if(pack.exists()){
			File [] packMonth = pack.listFiles();
			for(File file : packMonth){
				Integer month = Integer.parseInt(file.getName());
				if(month < monthCur){
					dropDir(file.getAbsolutePath());
				}
			}
		}
		//删除当天前的文件夹
		pack = new File(FILE_ROOT_PATH+File.separator+yearCur+File.separator+(monthCur.toString().length()==1?"0"+monthCur.toString():monthCur.toString())+File.separator);
		if(pack.exists()){
			File [] packDate = pack.listFiles();
			for(File file : packDate){
				Integer date = Integer.parseInt(file.getName());
				if(date < dateCur){
					dropDir(file.getAbsolutePath());
				}
			}
		}
	}
	/**
	 * 递归删除文件夹
	 * @param path 传入需要删除的文件夹
	 * @return 返回是否操作成功
	 */
	public static boolean dropDir(String path){
		boolean flag = false;
		if(path != null){
			File file = new File(path);
			if(file.exists()){
				if(file.isFile()){
					file.delete();
				}else {
					File [] files = file.listFiles();
					for(File f : files){
						dropDir(f.getAbsolutePath());//递归删除传入的文件夹下的文件或文件夹
					}
					file.delete();//递归回来时将传入的文件夹删除
					flag = true;
				}
			}else {
				throw new RuntimeException("要删除的文件夹不存在");
			}
		}else {
			throw new RuntimeException("删除文件夹时传入的路径为空");
		}
		return flag;
	}
}
