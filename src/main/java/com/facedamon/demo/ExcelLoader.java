package com.facedamon.demo;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;

import java.io.*;
import java.util.*;

public class ExcelLoader {

    public ExcelLoader() {
    }

    public File excelLoader(String excelName) throws FileNotFoundException {
        return new File(excelName);
    }

    /**
     * 获取表名指向sheet的column中的columnName & comment
     * @param tableName 表名
     * @return
     */
    public Map<String,String> getColumnComment(File file,String tableName){
        Map<String,String> result = new HashMap<String, String>();
        int rowSize = 0;
        BufferedInputStream in = null;
        POIFSFileSystem fs = null;
        HSSFWorkbook wb = null;
        HSSFCell columnCell = null;
        HSSFSheet columnDir = null;

        try {
            in = new BufferedInputStream(new FileInputStream(file));

            /**
             * 打开HSSFWorkbook
             */
            fs = new POIFSFileSystem(in);
            wb = new HSSFWorkbook(fs);

            /**
             * 首先读取tableName所指的sheet
             */
            columnDir = wb.getSheet(tableName);
            String columnName = "";
            String columnComment = "";

            /**
             * 第一行为标题，不读取
             * rowIndex = 2
             */
            for (int rowIndex = 2; rowIndex <= columnDir.getLastRowNum(); rowIndex++){
                HSSFRow dirRow = columnDir.getRow(rowIndex);
                if (null == dirRow){
                    continue;
                }
                int tempRowSize = dirRow.getLastCellNum() + 1;
                if (tempRowSize > rowSize){
                    rowSize = tempRowSize;
                }

                /**
                 * 获取字段名称，字段描述。
                 * 字段名称起始索引=0
                 * 字段描述索引=4
                 */
                for ( short columnIndex = 0;columnIndex <= 4; columnIndex += 4){
                    String value = "";
                    columnCell = dirRow.getCell(columnIndex);
                    if (null != columnCell){
                        value = columnCell.getStringCellValue();
                        if ("".equals(value) || null == value){
                            break;
                        }
                        if (0 == columnIndex){
                            columnName = value;
                        }
                        if (4 == columnIndex){
                            columnComment = value;
                        }
                        result.put(columnName,columnComment);
                    }
                }
            }
        } catch (Exception e){
            e.printStackTrace();
        } finally {
            try {
                wb.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
            try {
                fs.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
            try {
                in.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
        return result;
    }

    /**
     * 获取目录中的table_name & table_comment
     * @param file excelName
     * @throws IOException
     */
    public Map<String,String> getDirData(File file){
        Map<String,String> result = new HashMap<String, String>();
        int rowSize = 0;
        BufferedInputStream in = null;
        POIFSFileSystem fs = null;
        HSSFWorkbook wb = null;
        HSSFCell dirCell = null;
        HSSFSheet dir = null;

        try {
            in = new BufferedInputStream(new FileInputStream(file));

            /**
             * 打开HSSFWorkbook
             */
            fs = new POIFSFileSystem(in);
            wb = new HSSFWorkbook(fs);

            /**
             * 首先读取目录中的表名称
             * 目录在第二个sheet
             */
            dir = wb.getSheetAt(1);
            String tableName = "";
            String tableComment = "";

            /**
             * 第一行为标题，不读取
             * rowIndex = 1
             */
            for (int rowIndex = 1; rowIndex <= dir.getLastRowNum(); rowIndex++){
                HSSFRow dirRow = dir.getRow(rowIndex);
                if (null == dirRow){
                    continue;
                }
                int tempRowSize = dirRow.getLastCellNum() + 1;
                if (tempRowSize > rowSize){
                    rowSize = tempRowSize;
                }
                /**
                 * 获取表名称，表描述。
                 * 表名称起始索引=1
                 * 结束索引=2
                 */
                for (short columnIndex = 1; columnIndex <= 2; columnIndex++){
                    String value = "";
                    dirCell = dirRow.getCell(columnIndex);
                    if (null != dirCell){
                        value = dirCell.getStringCellValue();
                        if ("".equals(value) || null == value){
                            break;
                        }
                        if (1 == columnIndex){
                            tableName = value;
                        }
                        if (2 == columnIndex){
                            tableComment = value;
                        }
                        result.put(tableName,tableComment);
                    }
                }
            }
        } catch (Exception e){
            e.printStackTrace();
        }finally {
            try {
                wb.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
            try {
                fs.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
            try {
                in.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
        return result;
    }
}
