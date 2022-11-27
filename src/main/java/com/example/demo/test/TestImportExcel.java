package com.example.demo.test;

import com.example.demo.entity.People;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.streaming.SXSSFCell;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

import java.beans.PropertyDescriptor;
import java.io.File;
import java.io.FileOutputStream;
import java.lang.reflect.Field;
import java.lang.reflect.Method;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import java.util.concurrent.Callable;
import java.util.concurrent.ConcurrentHashMap;
import java.util.concurrent.FutureTask;

@SuppressWarnings("all")
public class TestImportExcel {

    private static final ConcurrentHashMap<Integer,FutureTask<Integer>> taskList=new ConcurrentHashMap<>();

    public static void main(String[] args) throws Exception {

//        Integer[] startRow={0,49999,100000,149999,200000,249999,300000,349999,400000,449999};
        Integer[] startRow={0,100000,200000,300000,400000,500000,600000,700000,800000,900000,1000000};
        //模拟数据
//        ArrayList<People> dataList=new ArrayList<>();
//        for(int i=0;i<10000;i++){
//            People people=new People("姓名"+i,i+"","男","110");
//            dataList.add(people);
//
//        }

        Long st=System.currentTimeMillis();
        //单线程导出excel
        //doExportOneThread(dataList);


        SXSSFWorkbook workbook=new SXSSFWorkbook();
        SXSSFSheet sheet=workbook.createSheet();

        for (int i = 0; i < 11; i++) {
            final int temp=i;
            FutureTask<Integer> task= new FutureTask<>(new Callable<Integer>() {
                @Override
                public Integer call() throws Exception {
                    //模拟数据
                    ArrayList<People> dataList=new ArrayList<>();
                    for(int i=0;i<100000;i++){
                        People people=new People("姓名"+i,i+"","男","110");
                        dataList.add(people);

                    }
                    Integer currentRow=doExportOneThread2(sheet,dataList,startRow[temp]);
                    System.out.println("任务"+temp+"执行结束....");
                    return currentRow;
                }
            });
            taskList.putIfAbsent(i,task);
        }

        Integer taskIndex=0;
        while (true){
            FutureTask<Integer> futureTask=taskList.remove(taskIndex++);
            if(futureTask!=null){
                futureTask.run();
                Integer currentRow=futureTask.get();
            }
            if(taskList.size()==0){
                System.out.println("开始写入Excel");
                FileOutputStream fileOutputStream = new FileOutputStream(new File("D:\\我的学习\\学学学\\POI导出excel\\test\\test-单线程分次导出"+Thread.currentThread().getName()+".xlsx"));
                workbook.write(fileOutputStream);
                System.out.println(Thread.currentThread().getName()+" 文件写入成功...");
                break;
            }
        }




        //多线程,多表，导出excel
//        doExportMoreThreadByMoreExcel(dataList,4);

        Long et=System.currentTimeMillis();
        System.out.println("Time: "+(et-st));
    }

    public static Integer doExportOneThread2(SXSSFSheet sheet, List<People> dataList, Integer currentRow) throws Exception{
        //遍历集合中数据
        Iterator<People> iterator=dataList.iterator();
        //反射获取
        Field[] fields=People.class.getDeclaredFields();
        
        
        if (currentRow==0){
            Row row=sheet.createRow(currentRow++);
            for(int i=0;i<fields.length;i++){
                Cell cell=row.createCell(i);
                cell.setCellValue(fields[i].getName());
            }
        }
        
        for(People data:dataList){
            Row row=sheet.createRow(currentRow++);
            for(int i=0;i<fields.length;i++){
                Cell cell=row.createCell(i);
                PropertyDescriptor descriptor=new PropertyDescriptor(fields[i].getName(),data.getClass());
                Method method=descriptor.getReadMethod();
                Object value=method.invoke(data);
                cell.setCellValue(value.toString());
            }
        }
//        while (iterator.hasNext()){
//            People people= iterator.next();
//
//            for (int i=0;i<fields.length;i++){
//                Cell cell= row.createCell(i);
//                else{
//                    cell.setCellValue(String.valueOf(fields[i].get(people)));
//                }
//            }
//            currentRow++;
//        }
        return currentRow;
    }


}
