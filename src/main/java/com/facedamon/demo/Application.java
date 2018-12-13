package com.facedamon.demo;

import java.io.*;
import java.util.List;
import java.util.Map;

public class Application {

    public static void main(String[] args) {
        ExcelLoader loader = new ExcelLoader();
        StringBuffer sqlBuffer = new StringBuffer();
        try {
            File excel = loader.excelLoader("D:\\idea_work\\ExcelDemo\\src\\main\\resources\\LPM_V_1数据字典.xls");
            Map<String,String> table = loader.getDirData(excel);


            for (Map.Entry<String,String> entry : table.entrySet()){
                String tableName = entry.getKey();
                String tableComment = entry.getValue();
                Map<String,String> column = loader.getColumnComment(excel,tableName);


                for (Map.Entry<String,String> entry1 : column.entrySet()){
                    String columnName = entry1.getKey();
                    String columnComment = entry1.getValue();
                    sqlBuffer.append("update information_schema.`TABLES` set TABLE_COMMENT = ")
                            .append("'").append(tableComment).append("'")
                            .append(" where TABLE_NAME = ")
                            .append("'").append(tableName).append("'")
                            .append("\n")
                            .append("update information_schema.`COLUMNS` set COLUMN_COMMENT = ")
                            .append("'").append(columnComment).append("'")
                            .append(" where TABLE_NAME = ")
                            .append("'").append(tableName).append("'")
                            .append(" and COLUMN_NAME = ")
                            .append("'").append(columnName).append("'")
                            .append("\n");

                }

            }

            writeSql(sqlBuffer.toString());
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        }
    }

    public static void writeSql(String str){
        File sql = new File("D:\\idea_work\\ExcelDemo\\src\\main\\resources\\column.sql");
        BufferedOutputStream bos = null;
        OutputStreamWriter writer = null;
        BufferedWriter bw = null;

        try {
            OutputStream os = new FileOutputStream(sql);
            bos = new BufferedOutputStream(os);
            writer = new OutputStreamWriter(bos);
            bw = new BufferedWriter(writer);
            bw.write(str);
            bw.flush();
            if (sql.exists()){
                sql.delete();
                System.out.println("-----------delete sql file------------");
            }
        } catch (Exception e){
            e.printStackTrace();
        }finally {
            try {
                bw.close();
            }catch (IOException e){
                e.printStackTrace();
            }
        }
    }
}
