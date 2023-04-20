package com.ducway.framework.modular.zmxzApplication;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.Random;

/**
 * Name:dl_zmxz
 * User: yjh
 * Date: 2023/4/20
 * Time: 11:12
 * Description:
 */
public class ExcelRandomDate {

    public void randomDate(){
        Random random=new Random();
        List<String> dateList = new ArrayList<>();
        List<String> endDateList = new ArrayList<>();
        int year = 2022;
        String date = "" , endDate = "";
        int month = -1;
        //外循环
        for (int i = 2; i <= 16; i++) {
            int[] arr = new int[5];//产生一个能存到五个元素的数组
            for (int index = 0; index < arr.length; index++) {
                arr[index]=-1;//为了防止默认0被当做检查重复的对象所以全部设为-1
            }
            //随机生成日期 1- 5号日期
            int day = (int) Math.ceil( (Math.random()) * 5 );
            date = "" ;
            endDate = "";
            month = i;
            if(month >= 13){
                year = 2023;
                month = i-12;
            }
            for (int j = day; j <= 30 - day; j++) {
                int hour = random.nextInt(15);
                int min = (random.nextInt(49) + 10);
                int sec = (random.nextInt(49) + 10);
                date = year+"/"+ month +"/"+ j  +" " + hour +":"+ min +":" + sec;
                System.out.println("随机生成的日期为：===== " + date);
                dateList.add(date);
                hour = hour + random.nextInt(8);
                date = year+"/"+ month +"/"+ j  +" " + hour +":"+ sec +":" + min ;
                endDateList.add(date);
            }
        }
        System.out.println(dateList.size());
        try {
            //创建工作簿
            XSSFWorkbook hssfWorkbook = new XSSFWorkbook(new FileInputStream("C:\\Users\\Administrator\\Desktop\\dl_early_alert.xlsx"));
            //获取工作簿下sheet的个数
            int sheetNum = hssfWorkbook.getNumberOfSheets();
            System.out.println("该excel文件中总共有："+sheetNum+"个sheet");
            //遍历工作簿中的所有数据
            for(int i = 0;i<sheetNum;i++) {
                //读取第i个工作表
                System.out.println("读取第"+(i+1)+"个sheet");
                XSSFSheet sheet = hssfWorkbook.getSheetAt(i);
                //获取最后一行的num，即总行数。此处从0开始
                int maxRow = sheet.getLastRowNum();
                //获取每一行数据
                for (int row = 0; row <= maxRow; row++) {
                    //获取最后单元格num，即总单元格数 ***注意：此处从1开始计数***
                    int maxRol = sheet.getRow(row).getLastCellNum();
                    System.out.println("--------第" + row + "行的数据如下--------");
                    if(row == 0){
                        continue;
                    }
                    //遍历每一行的每一个单元格
                    //                    for (int index = 0; index < maxRol; index++) {
                    //
                    //                    }
                    //拿到第六列和第七列的数据
                    XSSFCell cell1 = sheet.getRow(row).getCell(5);
                    XSSFCell cell2 = sheet.getRow(row).getCell(6);
                    if(cell1 != null && cell2 != null){
                        sheet.getRow(row).getCell(5).setCellValue(dateList.get(row-1));
                        sheet.getRow(row).getCell(6).setCellValue(endDateList.get(row-1));
                        System.out.println("更新成功！！");
                    }else {
                        System.out.println("更新失败");
                    }
                }
            }
            hssfWorkbook.write(new FileOutputStream("C:\\Users\\Administrator\\Desktop\\fix_dl_data.xlsx"));
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
