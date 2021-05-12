package com.github.excel;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileWriter;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

/**
 * @author yusheng
 * @version 1.0.0
 * @datetime 2021-05-11 14:08
 * @description parse excel sheet and convert to sql ddl.
 */
public class ExcelToSql {
    private static final String LINE_SEPARATOR = System.getProperty("line.separator");

    private static final String USAGE =
            "usage: " + LINE_SEPARATOR +
                    " java -jar excel-sql-ddl-all " + LINE_SEPARATOR +
                    "    ---excel excel file path" + LINE_SEPARATOR +
                    "    ---output output directory path" + LINE_SEPARATOR
            ;

    private static class TableRow{
        private int rowSeq;
        private String rowNameEn;
        private String rowNameCn;
        private String rowType;
        private boolean isIndex;
        private boolean isPrimaryKey;
        private boolean isNonEmpty;
        private boolean isAutoIncrement;
        private String rowComment;

        public void setRowSeq(int rowSeq) {
            this.rowSeq = rowSeq;
        }

        public int getRowSeq(){
            return this.rowSeq;
        }

        public void setRowNameEn(String rowNameEn){
            this.rowNameEn = rowNameEn;
        }

        public String getRowNameEn(){
            return this.rowNameEn;
        }

        public void setRowNameCn(String rowNameCn){
            this.rowNameCn = rowNameCn;
        }

        public String getRowNameCn(){
            return this.rowNameCn;
        }

        public void setRowType(String rowType){
            this.rowType = rowType;
        }

        public String getRowType() {
            return this.rowType;
        }

        public void setIsIndex(boolean isIndex){
            this.isIndex = isIndex;
        }

        public boolean isIndex() {
            return this.isIndex;
        }

        public void setIsPrimaryKey(boolean isPrimaryKey){
            this.isPrimaryKey = isPrimaryKey;
        }

        public boolean isPrimaryKey() {
            return this.isPrimaryKey;
        }

        public void setIsNonEmpty(boolean isNonEmpty){
            this.isNonEmpty = isNonEmpty;
        }

        public boolean isNonEmpty() {
            return this.isNonEmpty;
        }

        public void setAutoIncrement(boolean autoIncrement) {
            this.isAutoIncrement = autoIncrement;
        }

        public boolean isAutoIncrement() {
            return this.isAutoIncrement;
        }

        public void setRowComment(String rowComment){
            this.rowComment = rowComment;
        }

        public String getRowComment() {
            return this.rowComment;
        }

        @Override
        public String toString() {
            StringBuilder sb = new StringBuilder();
            sb.append("row seq:");
            sb.append(rowSeq);
            sb.append(",row name en:");
            sb.append(rowNameEn);
            sb.append(",row name cn:");
            sb.append(rowNameCn);
            sb.append(",row type:");
            sb.append(rowType);
            sb.append(",is index:");
            sb.append(isIndex);
            sb.append(",is primary key:");
            sb.append(isPrimaryKey);
            sb.append(",is non empty:");
            sb.append(isNonEmpty);
            sb.append(",is auto increment:");
            sb.append(isAutoIncrement);
            sb.append(",row comment:");
            sb.append(rowComment);
            return sb.toString();
        }
    }

    private static class SheetTable{
        private String tableNameCn;
        private String tableNameEn;
        private List<String> uniqueIndexes;
        private List<String[]> nonUniqueIndexes;
        private String dbType;
        private String dbName;
        private List<TableRow> tableRows;

        public void setTableNameCn(String tableNameCn){
            this.tableNameCn = tableNameCn;
        }

        public String getTableNameCn(){
            return this.tableNameCn;
        }

        public void setTableNameEn(String tableNameEn){
            this.tableNameEn = tableNameEn;
        }

        public String getTableNameEn() {
            return this.tableNameEn;
        }

        public void  setUniqueIndex(List<String> uniqueIndexes){
            this.uniqueIndexes = uniqueIndexes;
        }

        public List<String> getUniqueIndexes() {
            return this.uniqueIndexes;
        }

        public void setNonUniqueIndex(List<String[]> nonUniqueIndexes){
            this.nonUniqueIndexes = nonUniqueIndexes;
        }

        public List<String[]> getNonUniqueIndexes() {
            return this.nonUniqueIndexes;
        }

        public void setDbType(String dbType){
            this.dbType = dbType;
        }

        public String getDbType() {
            return this.dbType;
        }

        public void setDbName(String dbName){
            this.dbName = dbName;
        }

        public String getDbName() {
            return this.dbName;
        }

        public void setTableRows(List<TableRow> tableRows){
            this.tableRows = tableRows;
        }

        public List<TableRow> getTableRows() {
            return this.tableRows;
        }

        @Override
        public String toString() {
            StringBuilder sb = new StringBuilder();
            sb.append("table name cn:");
            sb.append(tableNameCn);
            sb.append(",table name en:");
            sb.append(tableNameEn);
            sb.append(",table unique indexes:");
            if(uniqueIndexes != null){
                for (String index : uniqueIndexes) {
                    sb.append(index);
                    sb.append(" ");
                }
            }
            sb.append(",table non unique indexes:");
            if(nonUniqueIndexes != null){
                for (String[] nonUniqueIndex : nonUniqueIndexes) {
                    for (String index : nonUniqueIndex) {
                        sb.append(index);
                        sb.append(" ");
                    }
                    sb.append(";");
                }
            }

            sb.append(",db type:");
            sb.append(dbType);
            sb.append(",db name:");
            sb.append(dbName);

            sb.append(",table rows:");
            sb.append("\n");
            for (TableRow tableRow : tableRows) {
                sb.append(tableRow.toString());
                sb.append("\n");
            }

            return sb.toString();
        }
    }

    private static String getExcelParam(String[] args){
        String excel = null;
        int index = paramIndexSearch(args,"---excel");
        if(index != -1){
            excel = args[index+1];
            if(excel.trim().isEmpty()){
                printUsageAndExit("error: excel is invalid or empty!");
            }
            if(fileOrDirExists(excel)){
                printUsageAndExit("error: excel file path is not exists!");
            }
        }else{
            printUsageAndExit("error: ---excel not found!");
        }
        return excel;
    }

    private static String getOutputParam(String[] args){
        String output = null;
        int index = paramIndexSearch(args,"---output");
        if(index != -1){
            output = args[index+1];
            if(output.trim().isEmpty()){
                printUsageAndExit("error: output is invalid or empty!");
            }
            if(fileOrDirExists(output)){
                printUsageAndExit("error: output directory path is not exists!");
            }
        }else{
            printUsageAndExit("error: ---output not found!");
        }
        return output;
    }

    public static void main(String[] args){
        String excel = getExcelParam(args);
        String output = getOutputParam(args);
        String sheetIndexName = "目录索引";

        File xlsFile = new File(excel);
        XSSFWorkbook wb = null;
        try{
            wb = new XSSFWorkbook(new FileInputStream(xlsFile));
        }catch (IOException ioe){
            ioe.printStackTrace();
        }

        if(wb != null){
            List<String> sheetNameList = parseSheetIndex(wb,sheetIndexName);
            for (String sheetName : sheetNameList) {
                SheetTable sheetTable = parseSheet(wb,sheetName);
                List<String> sqlLine = createSQL(sheetTable);
                writeSqlFile(sqlLine,output,sheetName.toLowerCase());
            }
        }

        closeXSSFWorkbook(wb);
    }

    private static SheetTable parseSheet(XSSFWorkbook wb,String sheetName){
        XSSFSheet sheet = wb.getSheet(sheetName);
        int physicalRowNum = sheet.getPhysicalNumberOfRows();

        SheetTable sheetTable = new SheetTable();

        if(physicalRowNum > 9){
            int rowIdx = 1;
            List<TableRow> tableRows = new ArrayList<>();

            for (Row row : sheet) {
                String rowColumn1 = row.getCell(1).getStringCellValue();

                if(rowIdx <= 9){
                   switch (rowIdx){
                       case 1:
                           sheetTable.setTableNameCn(rowColumn1);
                           break;
                       case 2:
                           sheetTable.setTableNameEn(rowColumn1);
                           break;
                       case 3:
                           if(rowColumn1 != null && !rowColumn1.trim().isEmpty()){
                               String[] itemArr = rowColumn1.split(",");
                               List<String> uniqueIndexes = new ArrayList<>();
                               for (String s : itemArr) {
                                   if(s != null && !s.trim().isEmpty()){
                                       uniqueIndexes.add(s);
                                   }
                               }
                               sheetTable.setUniqueIndex(uniqueIndexes);
                           }
                           break;
                       case 4:
                           if(rowColumn1 != null && !rowColumn1.trim().isEmpty()){
                               String[] itemLevel1Arr = rowColumn1.split(";");
                               List<String[]> nonUniqueIndexes = new ArrayList<>();
                               for (String s : itemLevel1Arr) {
                                   if(s != null && !s.trim().isEmpty()){
                                       if(s.contains(",")){
                                           String[] itemLevel2Arr = s.split(",");
                                           nonUniqueIndexes.add(itemLevel2Arr);
                                       }else{
                                           String[] nonIndexArr = new String[1];
                                           nonIndexArr[0] = s;
                                           nonUniqueIndexes.add(nonIndexArr);
                                       }
                                   }
                               }
                               sheetTable.setNonUniqueIndex(nonUniqueIndexes);
                           }
                           break;
                       case 5:
                           sheetTable.setDbType(rowColumn1);
                           break;
                       case 6:
                           sheetTable.setDbName(rowColumn1);
                           break;
                       case 7:
                           break;
                       case 8:
                           break;
                       case 9:
                           break;
                   }
                }else {
                    TableRow tableRow = new TableRow();

                    int rowColumnSeq = (int)row.getCell(0).getNumericCellValue();
                    tableRow.setRowSeq(rowColumnSeq);

                    String rowNameEn = row.getCell(1).getStringCellValue();
                    tableRow.setRowNameEn(rowNameEn);

                    String rowNameCn = row.getCell(2).getStringCellValue();
                    tableRow.setRowNameCn(rowNameCn);

                    String rowType = row.getCell(3).getStringCellValue();
                    tableRow.setRowType(rowType);

                    String isIndexStr = row.getCell(4).getStringCellValue();
                    if(isIndexStr.equalsIgnoreCase("索引")){
                        tableRow.setIsIndex(true);
                    }else{
                        tableRow.setIsIndex(false);
                    }

                    String isPrimaryKeyStr = row.getCell(5).getStringCellValue();
                    if(isPrimaryKeyStr.equalsIgnoreCase("主键")){
                        tableRow.setIsPrimaryKey(true);
                    }else{
                        tableRow.setIsPrimaryKey(false);
                    }

                    String isAutoIncrementStr = row.getCell(6).getStringCellValue();
                    if(isAutoIncrementStr.equalsIgnoreCase("自增")){
                        tableRow.setAutoIncrement(true);
                    }else{
                        tableRow.setAutoIncrement(false);
                    }

                    String isNonEmptyStr = row.getCell(7).getStringCellValue();
                    if(isNonEmptyStr.equalsIgnoreCase("非空")){
                        tableRow.setIsNonEmpty(true);
                    }else{
                        tableRow.setIsNonEmpty(false);
                    }

                    String rowComment = row.getCell(8).getStringCellValue();
                    tableRow.setRowComment(rowComment);

                    tableRows.add(tableRow);
                }
                rowIdx++;
            }

            sheetTable.setTableRows(tableRows);
        }else{
            System.out.println(String.format("sheet name:%s table column number is 0",sheet));
        }

        return sheetTable;
    }

    private static List<String> parseSheetIndex(XSSFWorkbook wb,String sheetName){
        XSSFSheet sheet = wb.getSheet(sheetName);
        List<String> tableNameEnList = new ArrayList<>();
        int idx = 0;
        for (Row row : sheet) {
            if(idx != 0){
                String tableNameEn = row.getCell(1).getStringCellValue();
                tableNameEnList.add(tableNameEn);
            }
            idx++;
        }
        return tableNameEnList;
    }

    private static void printUsageAndExit(String...messages){
        for (String message : messages) {
            System.err.println(message);
        }
        System.err.println(USAGE);
        System.exit(1);
    }

    /**
     * 查找指定命令行参数的索引位置.
     * @param args 命令行参数数组.
     * @param param 待查找的命令行参数.
     * @return index 参数索引位置,-1表示没有查找到.
     * */
    private static int paramIndexSearch(String[] args,String param){
        int index = -1;
        for (int i = 0; i < args.length; i++) {
            if(args[i].equalsIgnoreCase(param)){
                index = i;
                break;
            }
        }
        return index;
    }

    private static void closeXSSFWorkbook(XSSFWorkbook wb){
        if(wb != null){
            try{
                wb.close();
            }catch (IOException ioe){
                ioe.printStackTrace();
            }
        }
    }

    private static List<String> createSQL(SheetTable sheetTable){
        List<String> sqlLineList = new ArrayList<>();

        String table = "`" + sheetTable.getDbName() + "`.`" + sheetTable.getTableNameEn() + "`";
        sqlLineList.add("CREATE TABLE " + table + "(");

        for (TableRow tableRow : sheetTable.getTableRows()) {
            StringBuilder sb = new StringBuilder();

            String columnName = tableRow.getRowNameEn();
            sb.append("    ");
            sb.append("`");
            sb.append(columnName);
            sb.append("`");
            sb.append(" ");
            String columnType = tableRow.getRowType();
            sb.append(columnType);
            if(tableRow.isNonEmpty()){
                sb.append(" NOT NULL");
            }
            if(tableRow.isAutoIncrement()){
                sb.append(" AUTO_INCREMENT");
            }
            if(tableRow.getRowComment() != null && !tableRow.getRowComment().trim().isEmpty()){
                sb.append(" COMMENT '");
                sb.append(tableRow.getRowComment().replaceAll("[\r\n\t]",";"));
                sb.append("'");
            }
            sb.append(",");
            sqlLineList.add(sb.toString());
        }

        if(sheetTable.getUniqueIndexes() != null && !sheetTable.getUniqueIndexes().isEmpty()){
            List<String> uniqueIndexes = sheetTable.getUniqueIndexes();
            StringBuilder sb = new StringBuilder();
            sb.append("PRIMARY KEY(");
            for(int i = 0; i < uniqueIndexes.size(); i++){
                String index = uniqueIndexes.get(i);
                sb.append("`");
                sb.append(index);
                sb.append("`");
                if(i != uniqueIndexes.size() - 1){
                    sb.append(",");
                }
            }
            sb.append("),");
            sqlLineList.add(sb.toString());
        }

        if(sheetTable.getNonUniqueIndexes() != null && !sheetTable.getNonUniqueIndexes().isEmpty()){
            List<String[]> nonUniqueIndexes = sheetTable.getNonUniqueIndexes();
            int nonUniqueIndexSeq = 1;
            for (String[] nonUniqueIndex : nonUniqueIndexes) {
                if(nonUniqueIndex != null && nonUniqueIndex.length > 0){

                    StringBuilder sb = new StringBuilder();
                    sb.append("UNIQUE KEY (");
                    for(int i = 0; i < nonUniqueIndex.length; i++){
                        String index = nonUniqueIndex[i];
                        if(index != null && !index.isEmpty()){
                            sb.append("`");
                            sb.append(index);
                            sb.append("`");
                            if(i != nonUniqueIndex.length -1){
                                sb.append(",");
                            }
                        }
                    }
                    sb.append(")");
                    if(nonUniqueIndexSeq != nonUniqueIndexes.size()){
                        sb.append(",");
                    }

                    nonUniqueIndexSeq++;
                    sqlLineList.add(sb.toString());
                }
            }
        }

        sqlLineList.add(")");
        sqlLineList.add("COMMENT='" + sheetTable.getTableNameCn() + "'");
        sqlLineList.add("ENGINE=InnoDB");

        return sqlLineList;
    }

    private static void writeSqlFile(List<String> sqlLines,String writeBaseDir,String tableName){
        String sqlName = tableName + ".sql";
        String sqlPath = writeBaseDir + "/" + sqlName;
        FileWriter fileWriter = null;

        try{
            fileWriter = new FileWriter(sqlPath);
            for (String sqlLine : sqlLines) {
                fileWriter.write(sqlLine);
                fileWriter.write("\n");
            }
        }catch (IOException ioe){
            ioe.printStackTrace();
        }finally {
            if(fileWriter != null){
                try{
                    fileWriter.close();
                }catch (IOException ioe){
                    ioe.printStackTrace();
                }
            }
        }
    }

    private static boolean fileOrDirExists(String fileOrDir){
        return !new File(fileOrDir).exists();
    }
}
