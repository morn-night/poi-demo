package com.example.demo.test;

import com.example.demo.entity.People;
import com.example.demo.thread.ExportExcelThread;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import java.util.concurrent.CountDownLatch;

public class ImportExcel {

    public static void main(String[] args) throws Exception {
        //模拟数据
        ArrayList<People> dataList=new ArrayList<>();

        for(int i=0;i<500000;i++){
            People people=new People("姓名"+i,i+"","男","110");
            dataList.add(people);

        }

        Long st=System.currentTimeMillis();
        //单线程导出excel
        doExportOneThread(dataList);

        //多线程,单表，导出excel
//        doExportMoreThreadByOneExcel(dataList,4);

        //多线程,多表，导出excel
//        doExportMoreThreadByMoreExcel(dataList,4);

        Long et=System.currentTimeMillis();
        System.out.println("Time: "+(et-st));
    }

    /**
     * 单线程导出excel
     */
    public static void doExportOneThread(ArrayList<People> dataList) throws Exception {

        //创建一个工作表
        SXSSFWorkbook workbook=new SXSSFWorkbook();
        //创建一个sheet页
        SXSSFSheet sheet=workbook.createSheet();
        //长度大小不限
        sheet.setRandomAccessWindowSize(-1);

        //遍历集合中数据
        Iterator<People> iterator=dataList.iterator();
        //反射获取
        Field[] fields=People.class.getDeclaredFields();
        Integer currentRow=0;

        if (currentRow==0){
            Row row=sheet.createRow(currentRow);
            for(int i=0;i<fields.length;i++){
                Cell cell=row.createCell(i);
                cell.setCellValue(fields[i].getName());
            }
            currentRow++;
        }
        while (iterator.hasNext()){
            People people= iterator.next();
            Row row=sheet.createRow(currentRow);
            for (int i=0;i<fields.length;i++){
                Cell cell= row.createCell(i);
                cell.setCellValue(String.valueOf(fields[i].get(people)));
            }
            currentRow++;
        }
        FileOutputStream fileOutputStream=new FileOutputStream(new File("D:\\我的学习\\学学学\\POI导出excel\\test\\test.xls"));
        workbook.write(fileOutputStream);

    }

    /**
     * 多线程，单文件导出Excel
     * @param dataList
     * @param threadNumber
     */
    public static void doExportMoreThreadByOneExcel(ArrayList<People> dataList,Integer threadNumber) throws Exception{
        //把List集合分页
        int dataCount=dataList.size();
        /**
         * 向下取整 例：101个线程，4个线程池， 每个25 余1
         * 最后实现
         * 25
         * 25
         * 25
         * 26 最后一个加上余数
         */
        //获得平均值
        int avgNumber= (int) Math.floor(dataCount / threadNumber);
        //获得余数
        int remainder=dataCount % threadNumber;

        CountDownLatch countDownLatch=new CountDownLatch(threadNumber);

        SXSSFWorkbook workbook=new SXSSFWorkbook();

        //分割集合每部分的开始和结束下标
        int fromIndex=0;
        int toIndex=0;

        for (int i=0;i<threadNumber;i++){
            //分割List集合
            fromIndex=i*avgNumber;
            if (i==threadNumber-1){
                //最后一个线程池开启
                toIndex=(i+1)*avgNumber-1+remainder;
            }else {
                //没有余数正常的前几个线程池
                toIndex=(i+1)*avgNumber-1;
            }
            List<People> data=dataList.subList(fromIndex,toIndex+1);
            SXSSFSheet sheet=workbook.createSheet();
            new Thread(new ExportExcelThread(data,countDownLatch,sheet),"线程"+i).start();
        }

        //告诉线程，要所有的线程跑完再生成Excel文件
        countDownLatch.await();

        FileOutputStream fileOutputStream = new FileOutputStream(new File("D:\\我的学习\\学学学\\POI导出excel\\test\\test-多线程分页.xlsx"));
        workbook.write(fileOutputStream);
        System.out.println("文件写入成功---");

    }

    /**
     * 多线程，多文件导出
     * @param dataList
     * @param threadNumber
     * @throws Exception
     */
    public static void doExportMoreThreadByMoreExcel(ArrayList<People> dataList,Integer threadNumber) throws Exception{
        //把List集合分页
        int dataCount=dataList.size();

        //获得平均值
        int avgNumber= (int) Math.floor(dataCount / threadNumber);
        //获得余数
        int remainder=dataCount % threadNumber;

        //分割集合每部分的开始和结束下标
        int fromIndex=0;
        int toIndex=0;

        //开启线程
        for (int i=0;i<threadNumber;i++){
            //分割List集合
            fromIndex=i*avgNumber;
            if (i==threadNumber-1){
                //最后一个线程池开启
                toIndex=(i+1)*avgNumber-1+remainder;
            }else {
                //没有余数正常的前几个线程池
                toIndex=(i+1)*avgNumber-1;
            }
            List<People> data=dataList.subList(fromIndex,toIndex+1);
            SXSSFWorkbook workbook=new SXSSFWorkbook();
            new Thread(new ExportExcelThread(data, true, workbook),"线程"+i).start();
        }

        FileOutputStream fileOutputStream = new FileOutputStream(new File("D:\\我的学习\\学学学\\POI导出excel\\test\\test-多线程分页.xlsx"));


    }


}
