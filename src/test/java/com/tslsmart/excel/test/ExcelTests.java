package com.tslsmart.excel.test;

import cn.wangxs.excel.ExcelHelper;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.*;
import java.util.stream.IntStream;

/**
 * @author JakeWoo
 * @version 1.0
 * @date 2019-02-14 17:50:33
 */
public class ExcelTests {
    public static void main(String[] args) {
        String fileName = "/Users/JakeWoo/Documents/dev-data/Files/user.xlsx";
        List<ExcelEntity> rows = new ArrayList<>();
        String userName = "张三_";
        String passWord = "password1234!@#$_";
        String url = "http://www.baidu.com/s?wd=";
        double money = 123.345;

        IntStream.range(0,30).forEach(i->{
            ExcelEntity entity = new ExcelEntity();
            entity.setUserName(userName+i);
            entity.setPassWord(passWord+i);
            entity.setUrl(url+i);
            entity.setMoney(money + i);
            entity.setTime(new Date(System.currentTimeMillis()-i*100000));
            rows.add(entity);
        });
        ExcelTests test = new ExcelTests();
        try {
            test.writeExcel(rows,fileName);
        } catch (Exception e) {
            e.printStackTrace();
        }
        try {
            List<ExcelEntity> excelEntities = test.readExcel(fileName);
            excelEntities.forEach(System.out::println);
        } catch (Exception e) {
            e.printStackTrace();
        }


    }

    /**
     * 读取一个 Excel
     *
     * @param fileName
     * @return
     * @throws Exception
     */
    public List<ExcelEntity> readExcel(String fileName) throws Exception {

        FileInputStream fileInputStream = new FileInputStream(fileName);
//        List<ExcelEntity> read = ExcelHelper.read(fileInputStream, ExcelEntity.class);
//        fileInputStream.close();
//        return read;
        return ExcelHelper.read(fileInputStream, ExcelEntity.class);

    }

    /**
     * 创建一个 Excel 文档
     * @param collection
     * @param fileName
     * @throws Exception
     */
    public void writeExcel(Collection<ExcelEntity> collection,String fileName) throws Exception {
        FileOutputStream fileOutputStream = new FileOutputStream(fileName);
        byte[] write = ExcelHelper.write(collection, ExcelEntity.class);
        fileOutputStream.write(write);
        fileOutputStream.flush();
        fileOutputStream.close();
    }
}
