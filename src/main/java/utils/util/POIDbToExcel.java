package utils.util;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.stereotype.Component;

import java.io.*;
import java.sql.*;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.*;

/**
* @author: lijinlong
* @Date: 2019/9/5 17:07
* @Description 将指定数据库中所有表的表结构和数据都读出来以当前日期时间戳为Excel名，表名为Sheet的页名
* @version 1.1 ->如配合调度，可将指定库的所有表的数据进行备份（window不要直接写C:）
*/
@Component
public final class POIDbToExcel {

    private static final String SELECT = "SELECT";
    private static final String FROM = "FROM";

    private static Logger logger = LoggerFactory.getLogger(POIDbToExcel.class);

    /**
     * @Description 备份数据库到Excel中
     * @Author lijinlong
     * @Date   2019/9/5 17:13
     * @Param  [excelVersion Excel的版本,filePath 文件路径]
     * @Return void
     * @Exception
     */
    protected static void bakDBToExcel(String excelVersion,String filePath){
        //先查询库里所有表
        List<String> tables = getTables();
        if(tables.isEmpty()){
            logger.info("库里无表，需检查库是否连接正确");
            return;
        }
        //查询出各表中的所有对应字段
        Map<String, List<String>> tablesAndColumns = getTablesColumn(tables);
        //先将表名和字段分别写到Excel的sheet和第一行的cell中
        boolean outputTitleToExcel = outPutTablesAndColumns(tablesAndColumns, excelVersion, filePath);
        if(!outputTitleToExcel){
            logger.info("写出字段时出现异常");
            return;
        }
        //读取每张表的数据
        Map<String, Map<Integer, List<String>>> datas = getDatas(tablesAndColumns);
        //将每张表的数据写入到Excel中
        if(outPutDatas(datas,excelVersion,filePath)){
            logger.info("success");
        }else{
            logger.info("error");
        }
    }

    /**
     * @Description 查询库里的所有表名
     * @Author lijinlong
     * @Date   2019/9/5 17:45
     * @Param  []
     * @Return boolean
     * @Exception
     */
    protected static List<String> getTables(){
        Connection connection = JdbcConnectionUtil.getConnection();
        ResultSet resultSet = null;
        List<String> tables = new ArrayList<>();
        try {
            DatabaseMetaData metaData = connection.getMetaData();
            resultSet = metaData.getTables("","","",new String []{"TABLE"});
            while(resultSet.next()){
                String table = resultSet.getString(3);
                tables.add(table);
            }
        } catch (SQLException e) {
            logger.info("查询表名时出现错误");
        } finally {
            JdbcConnectionUtil.close(connection,null,resultSet);
            return tables;
        }
    }

    /**
     * @Description 将库中所有表的字段都查询出来
     * @Author lijinlong
     * @Date   2019/9/5 19:01
     * @Param  [tables]
     * @Return java.util.Map<java.lang.String,java.util.List<java.lang.String>> key为表名，value为字段名
     * @Exception
     */
    protected static Map<String, List<String>> getTablesColumn(List<String> tables) {
        Connection connection = JdbcConnectionUtil.getConnection();
        ResultSet resultSet = null;
        PreparedStatement preparedStatement = null;
        Map<String, List<String>> tablesColumn = new HashMap<>();
        try {
            for (int i = 0; i < tables.size(); i++) {
                List<String> columns = new ArrayList<>();
                String sql = SELECT + " " + "*" + " " + FROM + " " + tables.get(i);
                preparedStatement = connection.prepareStatement(sql);
                ResultSetMetaData metaData = preparedStatement.getMetaData();
                int columnCount = metaData.getColumnCount();
                //参数为1就从第一列取
                for (int j = 1; j <= columnCount; j++) {
                    columns.add(metaData.getColumnName(j));
                }
                tablesColumn.put(tables.get(i), columns);
            }
        } catch (SQLException e) {
            e.printStackTrace();
        } finally {
            JdbcConnectionUtil.close(connection, preparedStatement, resultSet);
            return tablesColumn;
        }
    }

    /**
     * @Description 将各表的表名为sheet页名,字段名为第一行标题写入到以时间戳为名的Excel中
     * @Author lijinlong
     * @Date   2019/9/6 10:19
     * @Param  [tablesAndColumns 表名和字段, excelVersion Excel版本, filePath 存放路径]
     * @Return boolean 写入成功返回true，否则为false
     * @Exception
     */
    protected static boolean outPutTablesAndColumns(Map<String, List<String>> tablesAndColumns,String excelVersion,String filePath) {
        Workbook workbooks = null;
        if ("xlsx".equalsIgnoreCase(excelVersion)){
            workbooks = new XSSFWorkbook();
        } else if ("xls".equalsIgnoreCase(excelVersion)){
            workbooks = new HSSFWorkbook();
        } else {
            logger.info("Excel版本不存在");
            return false;
        }
        Workbook workbook = workbooks;
        //创建一个单元格样式
        CellStyle cellStyle = workbook.createCellStyle();
        //设置单元格居中
        cellStyle.setAlignment(CellStyle.ALIGN_CENTER);
        //创建一个字体样式
        Font font = workbook.createFont();
        //设置字体为加粗
        font.setBoldweight(Font.BOLDWEIGHT_BOLD);
        //将样式附着在单元格样式上
        cellStyle.setFont(font);
        tablesAndColumns.forEach((tablename,columns)->{
            Sheet sheet = workbook.createSheet(tablename);
            Row row = sheet.createRow(0);
            for(int i=0; i<columns.size(); i++){
                Cell cell = row.createCell(i);
                cell.setCellValue(columns.get(i));
                cell.setCellStyle(cellStyle);
            }
        });
        Date date = new Date();
        SimpleDateFormat simpleDateFormat = new SimpleDateFormat("yyyy-MM-dd");
        OutputStream outputStream = null;
        try {
            outputStream = new FileOutputStream(filePath.endsWith("/")?filePath+simpleDateFormat.format(date)+"."+excelVersion:filePath+"/"+simpleDateFormat.format(date)+"."+excelVersion);
            workbook.write(outputStream);
        } catch (IOException e){
            logger.info("输出流出现异常");
            return false;
        } finally {
            try {
                if(outputStream!=null){
                    outputStream.close();
                }
                if(workbook!=null){
                    workbook.close();
                    workbooks.close();
                }
            } catch (IOException e) {
                logger.info("关闭出输出流时出现异常");
                return false;
            }
        }
        return true;
    }

    /**
     * @Description 将每张表的数据都查出来
     * @Author lijinlong
     * @Date   2019/9/6 11:34
     * @Param  [tablesAndColumns 表名和字段名]
     * @Return java.util.Map<java.lang.String,java.util.Map<java.lang.Integer,java.util.List<java.lang.String>>>
     * @Exception
     */
    protected static Map<String,Map<Integer,List<String>>> getDatas(Map<String, List<String>> tablesAndColumns){
        Connection connection = JdbcConnectionUtil.getConnection();
        Map<String,Map<Integer,List<String>>> datas = new HashMap<>();
        try {
            tablesAndColumns.forEach((tablename,columns)->{
                PreparedStatement preparedStatement = null;
                ResultSet resultSet = null;
                Map<Integer,List<String>> map = new HashMap<>();
                List<String> column = new ArrayList<>();
                String sql = SELECT + " ";
                for(int i=0; i<columns.size(); i++){
                    if( i ==columns.size()-1){
                        sql = sql + columns.get(i);
                    }else{
                        sql = sql + columns.get(i) + ",";
                    }
                    column.add(columns.get(i));
                }
                sql = sql + " " + FROM + " " +tablename;
                try {
                    preparedStatement = connection.prepareStatement(sql);
                    resultSet = preparedStatement.executeQuery();
                    int count = 1;
                    while (resultSet.next()){
                        //每次创建新集合接收
                        List<String> data = new ArrayList<>();
                        for(int i=0; i<column.size(); i++){
                            String string = resultSet.getString(column.get(i));
                            if(string == null){
                                string = "";
                            }
                            data.add(string);
                        }
                        map.put(count,data);
                        count++;
                    }
                    //用来复位行数
                    count = 1;
                    datas.put(tablename,map);
                } catch (SQLException e) {
                    e.printStackTrace();
                } finally {
                    JdbcConnectionUtil.close(null,preparedStatement,resultSet);
                }
            });
        }catch (Exception e){
            logger.info("查询各表数据时出现异常");
            return null;
        }finally {
            JdbcConnectionUtil.close(connection,null,null);
        }
        return datas;
    }

    /**
     * @Description 将各表的数据写入到各表的sheet页中
     * @Author lijinlong
     * @Date   2019/9/6 13:24
     * @Param  [datas 表名及字段还有数据、行数, excelVersion Excel的版本, filePath 文件路径]
     * @Return boolean
     * @Exception
     */
    protected static boolean outPutDatas(Map<String, Map<Integer, List<String>>> datas,String excelVersion,String filePath){
        Workbook workbook1 = null;
        Date date = new Date();
        SimpleDateFormat simpleDateFormat = new SimpleDateFormat("yyyy-MM-dd");
        InputStream inputStream = null;
        OutputStream outputStream = null;
        try {
            inputStream = new FileInputStream(filePath.endsWith("/") ? filePath + simpleDateFormat.format(date) + "." + excelVersion : filePath + "/" + simpleDateFormat.format(date) + "." + excelVersion);
            if ("xlsx".equalsIgnoreCase(excelVersion)) {
                workbook1 = new XSSFWorkbook(inputStream);
            } else if ("xls".equalsIgnoreCase(excelVersion)) {
                workbook1 = new HSSFWorkbook(inputStream);
            }
            Workbook workbook = workbook1;
            datas.forEach((tablename,data)->{
                Sheet sheet = workbook.getSheet(tablename);
                data.forEach((rownum,data_)->{
                    Row row = sheet.createRow(rownum);
                    for(int i=0; i<data_.size(); i++){
                        Cell cell = row.createCell(i);
                        cell.setCellValue(data_.get(i));
                    }
                });
            });
            outputStream = new FileOutputStream(filePath.endsWith("/") ? filePath + simpleDateFormat.format(date) + "." + excelVersion : filePath + "/" + simpleDateFormat.format(date) + "." + excelVersion);
            workbook.write(outputStream);
            workbook.close();
        } catch (IOException e) {
            logger.info("在写出数据时出现异常");
            return false;
        } finally {
            try {
                if(outputStream!=null){
                    outputStream.close();
                }
                if(workbook1!=null){
                    workbook1.close();
                }
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
        return true;
    }

}
