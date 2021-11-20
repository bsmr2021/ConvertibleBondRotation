package com.company;

import jxl.Sheet;
import jxl.Workbook;
import jxl.read.biff.BiffException;
import jxl.write.*;
import jxl.write.Number;

import java.io.*;

public class ExcelTools {
    File file = new File("可转债轮动.xls");
    // 创建输入流，读取Excel
    InputStream readableIs = new FileInputStream(file.getAbsolutePath());
    // jxl提供的Workbook类
    Workbook readableWorkbook = Workbook.getWorkbook(readableIs);

    public ExcelTools() throws IOException, BiffException {
    }

    public  void readExcel(String[][] str, String sheetName, Integer StartRow, Integer StartColumn, Integer EndColumn) throws IOException, BiffException {
        // 创建输入流，读取Excel
        InputStream readableIs = new FileInputStream(file.getAbsolutePath());
        // jxl提供的Workbook类
        Workbook readableWorkbook = Workbook.getWorkbook(readableIs);

        // Excel的页签数量
        int sheet_size = readableWorkbook.getNumberOfSheets();
//            System.out.println("sheet_size:" + sheet_size);
        Sheet wbSheet = readableWorkbook.getSheet(sheetName);
//            System.out.println("Sheetname:" + wbSheet.getName());
        // sheet.getRows()返回该页的总行数
        String strtmp;
        for (int i = StartRow; i < wbSheet.getRows(); i++) {
            // sheet.getColumns()返回该页的总列数
            for (int j = StartColumn; j < EndColumn; j++) {
                strtmp = wbSheet.getCell(j, i).getContents();
                if (strtmp != null && strtmp != "") {
                    str[i][j] = strtmp;
                    //System.out.println("str["+i+"]"+"["+j+"]:"+str[i][j]);
                }
            }
        }
        readableWorkbook.close();
        readableIs.close();
    }

    public void PrintArrayList(String[][] str, String printDataName, int iStart, int iEnd, int jStart, int jEnd) {
        for (int i = iStart; i < iEnd; i++) {
            for (int j = jStart; j < jEnd; j++) {
                if (str[i][j] != null) {
                    System.out.println(printDataName + "["+i+"]"+"["+j+"]:"+str[i][j]);
                }
            }
        }
    }

    public void WriteExcel(File file, String[][] str, Integer sheetNum, Integer StartColumn, Integer EndColumn) {
        try {
            // 创建输入流，读取Excel
            InputStream is = new FileInputStream(file.getAbsolutePath());
            // jxl提供的Workbook类
            Workbook wb = Workbook.getWorkbook(is);
            //打开一个文本的副本
            WritableWorkbook copyWritableWorkbook = Workbook.createWorkbook(file, wb);
            WritableSheet wbWritableSheet = copyWritableWorkbook.getSheet("我的低溢价可转债持仓");
            System.out.println("Sheetname:" + wbWritableSheet.getName());
            // 在Label对象的构造子中指名单元格位置是第一列第一行(0,0),单元格内容为string
            Label label = new Label(0, 0, "string");
            // 将定义好的单元格添加到工作表中
            wbWritableSheet.addCell(label);
            // 生成一个保存数字的单元格,单元格位置是第二列，第一行，单元格的内容为1234.5
            Number number = new Number(1, 0, 1234.5);
            wbWritableSheet.addCell(number);
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (BiffException e) {
            e.printStackTrace();
        } catch (IOException | WriteException e) {
            e.printStackTrace();
        }
    }

    public void DeleteNotConvertibleBond(String[][] strMyTemp) {
        try {
            // 创建输入流，读取Excel
            InputStream is2 = new FileInputStream(file.getAbsolutePath());
            // jxl提供的Workbook类
            Workbook wb2 = Workbook.getWorkbook(is2);
            //打开一个文件副本
            WritableWorkbook copyWritableWorkbook = Workbook.createWorkbook(file, wb2);
            WritableSheet wbWritableSheet = copyWritableWorkbook.getSheet("我的低溢价可转债持仓");
            //System.out.println("Sheetname:" + wbWritableSheet.getName());
            int deleteRowNum = 0;
            for (int i = 1; i < strMyTemp.length; i++) {//去掉表头，从第二行data开始
                if (strMyTemp[i][1] != null) {// ( && strMyTemp[i][j] != "" )
                    if(strMyTemp[i][1].contains("转") == false) {
                        deleteRowNum ++;
                        //System.out.println("Remove strMyTemp["+i+"][1]:"+strMyTemp[i][1]);
                        wbWritableSheet.removeRow(i-deleteRowNum+1);
                    }
                }
            }
            //ExcelTools.PrintData(strMyTemp, "strMyTemp", 0 , strMyTemp.length, 0, 2);

            copyWritableWorkbook.write();// 写入数据并关闭文件
            copyWritableWorkbook.close();
            wb2.close();
            is2.close();
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (BiffException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        } catch (WriteException e) {
            e.printStackTrace();
        }
    }


}
