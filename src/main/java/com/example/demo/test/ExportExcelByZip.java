package com.example.demo.test;

import com.example.demo.entity.People;
import com.example.demo.thread.ExportExcelThread;
import com.example.demo.util.FileDownloadUtils;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

import java.io.File;
import java.time.LocalDate;
import java.util.ArrayList;
import java.util.List;

public class ExportExcelByZip {
    public static void main(String[] args) throws Exception {
        //模拟数据
        ArrayList<People> dataList=new ArrayList<>();

        for(int i=0;i<1000000;i++){
            People people=new People("姓名"+i,i+"","男","110");
            dataList.add(people);

        }

        Long st=System.currentTimeMillis();

        //多线程,多表，导出excel
        doExportMoreThreadByMoreExcel(dataList,10);

        Long et=System.currentTimeMillis();
        System.out.println("Time: "+(et-st));

        /**
         * @param path    要压缩的文件路径
         * @param format  生成的格式（zip、rar）
         * @param zipPath zip的路径
         * @param zipName zip文件名
         * @Description 将多个文件进行压缩到指定位置
         */
        String zipName = "同步数据" + LocalDate.now() ;
        //创建临时文件夹保存excel
        String tempDir ="D:\\我的学习\\学学学\\POI导出excel\\test\\more" ;
        String zipPath ="D:\\我的学习\\学学学\\POI导出excel\\test\\zipHome" ;

        FileDownloadUtils.generateFile(tempDir, "zip", zipPath, zipName);

//        FileDownloadUtils.deleteDir(new File(tempDir));


    }


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

//        FileOutputStream fileOutputStream = new FileOutputStream(new File("D:\\我的学习\\学学学\\POI导出excel\\test\\more\\test-多线程分页.xlsx"));


    }

}
