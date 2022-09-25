package com.example.demo.thread;

import com.example.demo.entity.People;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

import java.io.File;
import java.io.FileOutputStream;
import java.lang.reflect.Field;
import java.util.Iterator;
import java.util.List;
import java.util.WeakHashMap;
import java.util.concurrent.CountDownLatch;

/**
 * 多线程导出excel
 */
public class ExportExcelThread implements Runnable{

    private List<People> data;
    //每个线程的状态
    private CountDownLatch countDownLatch;
    private Sheet sheet;

    //多线程,多文件标志
    private boolean flag=false; //默认情况下，是单个excel文件
    private SXSSFWorkbook workbook;


    //多个线程，单文件导出 用的
    public ExportExcelThread(List<People> data, CountDownLatch countDownLatch, Sheet sheet) {
        this.data = data;
        this.countDownLatch = countDownLatch;
        this.sheet = sheet;

    }

    //多个线程，多文件 用的
    public ExportExcelThread(List<People> data, boolean flag, SXSSFWorkbook workbook) {
        this.data = data;
        this.flag = flag;
        this.workbook = workbook;
    }

    @Override
    public void run() {
        String name=Thread.currentThread().getName();
        System.out.println(name+"--->开始执行>>>");
        try{
            if (!flag){
                doOneExport();
            }else {
                doMoreExport();
            }
        }catch (Exception e){
            e.printStackTrace();
        }
    }

    //单个Excel导出
    private void doOneExport() throws Exception {
        Iterator<People> iterator=data.iterator();
        Field[] fields=People.class.getDeclaredFields();
        Integer startIndex=0;
        if (startIndex==0){
            Row row=sheet.createRow(startIndex);
            for (int i=0;i<fields.length;i++){
                Cell cell=row.createCell(i);
                cell.setCellValue(fields[i].getName());
            }
            startIndex++;
        }

        while (iterator.hasNext()){
            People people=iterator.next();
            Row row=sheet.createRow(startIndex);
            for (int i=0;i<fields.length;i++){
                Cell cell=row.createCell(i);
                cell.setCellValue(String.valueOf(fields[i].get(people)));
            }
            startIndex++;
        }
        //告诉对象，线程已执行
        countDownLatch.countDown();
    }

    //多个Excel导出
    private void doMoreExport() throws Exception{
        SXSSFSheet sheet=workbook.createSheet();
        Iterator<People> iterator=data.iterator();
        Field[] fields=People.class.getDeclaredFields();
        Integer currentRow=0;

        if (currentRow==0){
            Row row= sheet.createRow(currentRow);
            for (int i=0;i<fields.length;i++){
                Cell cell=row.createCell(i);
                cell.setCellValue(String.valueOf(fields[i].getName()));
            }
            currentRow++;
        }
        while (iterator.hasNext()) {
            People people = iterator.next();
            Row row = sheet.createRow(currentRow);
            for (int i = 0; i < fields.length; i++) {
                Cell cell = row.createCell(i);
                cell.setCellValue(String.valueOf(fields[i].get(people)));
            }
            currentRow++;
        }

        FileOutputStream fileOutputStream = new FileOutputStream(new File("D:\\我的学习\\学学学\\POI导出excel\\test\\test-多线程分表"+Thread.currentThread().getName()+".xlsx"));
        workbook.write(fileOutputStream);
        System.out.println(Thread.currentThread().getName()+" 文件写入成功...");
    }


}
