package com.example.demo.util;

import java.io.BufferedOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.zip.ZipEntry;
import java.util.zip.ZipOutputStream;

public class FileDownloadUtils {

    /**
     * 创建文件夹;
     *
     * @param path 路径
     */
    public static void createFile(String path) {
        File file = new File(path);
        //判断文件是否存在;
        if (!file.exists()) {
            //创建文件;
            file.mkdirs();
        }
    }

    /**
     * 删除文件夹及文件夹下所有文件
     *
     * @param dir 文件地址
     * @return boolean
     */
    public static boolean deleteDir(File dir) {
        if (dir == null || !dir.exists()) {
            return true;
        }
        if (dir.isDirectory()) {
            String[] children = dir.list();
            //递归删除目录中的子目录下
            for (String child : children) {
                boolean success = deleteDir(new File(dir, child));
                if (!success) {
                    return false;
                }
            }
        }
        // 目录此时为空，可以删除
        return dir.delete();
    }

    /**
     * @param path    要压缩的文件路径
     * @param format  生成的格式（zip、rar）
     * @param zipPath zip的路径
     * @param zipName zip文件名
     * @Description 将多个文件进行压缩到指定位置
     */
    public static boolean generateFile(String path, String format, String zipPath, String zipName) throws Exception {

        File file = new File(path);
        // 压缩文件的路径不存在
        if (!file.exists()) {
            throw new Exception("路径 " + path + " 不存在文件，无法进行压缩...");
        }
        // 用于存放压缩文件的文件夹
        String generateFile = zipPath + File.separator;
        File compress = new File(generateFile);
        // 如果文件夹不存在，进行创建
        if (!compress.exists()) {
            compress.mkdirs();
        }

        // 目的压缩文件
        String generateFileName = compress.getAbsolutePath() + File.separator + zipName + "." + format;

        // 输出流
        FileOutputStream outputStream = new FileOutputStream(generateFileName);
        // 压缩输出流
        ZipOutputStream zipOutputStream = new ZipOutputStream(new BufferedOutputStream(outputStream));

        //压缩
        generateFile(zipOutputStream, file, "");
        System.out.println("源文件位置：" + file.getAbsolutePath() + "，目的压缩文件生成位置：" + generateFileName);
        // 关闭 输出流
        zipOutputStream.close();
        return true;
    }

    /**
     * @param out  输出流
     * @param file 目标文件
     * @param dir  文件夹
     * @throws Exception
     */
    private static void generateFile(ZipOutputStream out, File file, String dir) throws Exception {
        // 当前的是文件夹，则进行一步处理
        if (file.isDirectory()) {
            //得到文件列表信息
            File[] files = file.listFiles();
            //将文件夹添加到下一级打包目录
            out.putNextEntry(new ZipEntry(dir + "/"));
            dir = dir.length() == 0 ? "" : dir + "/";
            //循环将文件夹中的文件打包
            for (int i = 0; i < files.length; i++) {
                generateFile(out, files[i], dir + files[i].getName());
            }
        } else { // 当前是文件
            // 输入流
            FileInputStream inputStream = new FileInputStream(file);
            // 标记要打包的条目
            out.putNextEntry(new ZipEntry(dir));
            // 进行写操作
            int len = 0;
            byte[] bytes = new byte[1024];
            while ((len = inputStream.read(bytes)) > 0) {
                out.write(bytes, 0, len);
            }
            // 关闭输入流
            inputStream.close();
        }
    }
}