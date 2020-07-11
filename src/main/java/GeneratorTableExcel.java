import net.sf.json.JSONObject;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.FontFormatting;
import org.apache.poi.ss.usermodel.Hyperlink;

import java.io.BufferedReader;
import java.io.FileOutputStream;
import java.io.FileReader;
import java.io.IOException;
import java.sql.*;
import java.util.ArrayList;
import java.util.List;
import java.util.Properties;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * @author: 杜云章
 * @Date: 2020-05-27 15:37
 */
public class GeneratorTableExcel {
    public static void generator(String filePlace) {
        Connection conn = null;
        try {
            // 获取配置文件信息
            JSONObject infoJsonObject = getProperties();
            if (infoJsonObject == null) {
                System.err.println("whitelist.json配置为空");
                return;
            }

            //连接数据库
            String driverClassName = infoJsonObject.getString("driverClassName");
            String jdbcUrl = infoJsonObject.getString("jdbc_url");
            String jdbcUsername = infoJsonObject.getString("jdbc_username");
            String jdbcPassword = infoJsonObject.getString("jdbc_password");
            Class.forName(driverClassName);
            Properties props = new Properties();
            props.setProperty("user", jdbcUsername);
            props.setProperty("password", jdbcPassword);
            props.setProperty("remarks", "true"); //设置可以获取remarks信息
            props.setProperty("useInformationSchema", "true");//设置可以获取tables remarks信息
            conn = DriverManager.getConnection(jdbcUrl, props);
            Statement statement = conn.createStatement();
            DatabaseMetaData dmd = conn.getMetaData();

            //获取数据库中的所有表
            ResultSet tableInfoSet = dmd.getTables(null, null, null, new String[]{"TABLE"});
            List<TableInfo> tableList = new ArrayList<TableInfo>();
            //获取表和字段信息
            while (tableInfoSet.next()) {
                TableInfo tableInfo = new TableInfo();
                String tableName = tableInfoSet.getString("TABLE_NAME");
                tableInfo.setCode(tableName);
                tableInfo.setCamel(underline2camel(tableName));
                tableInfo.setName(tableInfoSet.getString("REMARKS"));
                ResultSet columns = dmd.getColumns(null, null, tableName, "%");
                List<FieldInfo> fieldInfos = new ArrayList<FieldInfo>();
                while (columns.next()) {
                    FieldInfo fieldInfo = new FieldInfo();
                    String columnName = columns.getString("COLUMN_NAME");
                    String remarks = columns.getString("REMARKS");
                    fieldInfo.setCode(columnName);
                    fieldInfo.setCamel(underline2camel(columnName));
                    fieldInfo.setName(remarks);
                    fieldInfos.add(fieldInfo);
                }
                tableInfo.setFieldInfos(fieldInfos);
                tableList.add(tableInfo);
            }
            // 创建一个工作簿
            HSSFWorkbook wb = new HSSFWorkbook();
            // 创建一个工作表
            HSSFSheet sheetTotal = wb.createSheet();
            // 设置sheet页名
            int index = 0;
            String homeName = "表信息";
            wb.setSheetName(index, homeName);
            index += 1;
            /*
             * 设置一个超链接字体样式
             */
            // 设置样式
            HSSFCellStyle style = wb.createCellStyle();
            // 设置字体
            HSSFFont font = wb.createFont();
            //设置字体颜色
            font.setColor(HSSFColor.BLUE.index);
            //设置下划线
            font.setUnderline(FontFormatting.U_SINGLE);
            style.setFont(font);
            /*
             * 设置一个红色字体样式
             */
            // 设置样式
            HSSFCellStyle redStyle = wb.createCellStyle();
            // 设置字体
            HSSFFont redFont = wb.createFont();
            //设置字体颜色
            redFont.setColor(HSSFColor.RED.index);
            redStyle.setFont(redFont);

            // 是否有企业id
            boolean hasCompanyId = false;

            // 设置最大宽度
            int[] maxTableLength = new int[3];
            int[] maxFieldLength = new int[3];

            //  创建行
            int size = tableList.size();
            for (int i = 0; i < size + 1; i++) {
                if (i == 0) {
                    HSSFRow row = sheetTotal.createRow(i);
                    HSSFCell cell0 = row.createCell(0);
                    HSSFCell cell1 = row.createCell(1);
                    HSSFCell cell2 = row.createCell(2);
                    cell0.setCellValue("表名");
                    cell1.setCellValue("驼峰");
                    cell2.setCellValue("描述");
                } else {
                    TableInfo tableInfo = tableList.get(i - 1);

                    String tableCode = tableInfo.getCode();
                    String tableCamel = tableInfo.getCamel();
                    String tableName = tableInfo.getName();
                    String sheetName = "(" + index + ")" + tableCode;
                    // 链接
                    CreationHelper createHelper = wb.getCreationHelper();
                    Hyperlink fieldLink = createHelper.createHyperlink(Hyperlink.LINK_DOCUMENT);


                    // 字段信息sheet页
                    List<FieldInfo> fieldInfos = tableInfo.getFieldInfos();
                    HSSFSheet sheet = wb.createSheet();
                    wb.setSheetName(index, sheetName);
                    index += 1;
                    for (int j = 0; j < fieldInfos.size() + 1; j++) {
                        if (j == 0) {
                            HSSFRow fieldRow = sheet.createRow(j);
                            HSSFCell fieldCell0 = fieldRow.createCell(0);
                            HSSFCell fieldCell1 = fieldRow.createCell(1);
                            HSSFCell fieldCell2 = fieldRow.createCell(2);
                            HSSFCell fieldCell3 = fieldRow.createCell(4);
                            fieldCell0.setCellValue("字段名");
                            fieldCell1.setCellValue("驼峰");
                            fieldCell2.setCellValue("描述");
                            fieldCell3.setCellValue("返回首页");
                            // 返回首页链接
                            Hyperlink tableLink = createHelper.createHyperlink(Hyperlink.LINK_DOCUMENT);
                            tableLink.setAddress("'" + homeName + "'!A" + index);
                            fieldCell3.setHyperlink(tableLink);
                            fieldCell3.setCellStyle(style);
                        } else {
                            FieldInfo fieldInfo = fieldInfos.get(j - 1);
                            /*
                             * 写字段信息
                             */
                            HSSFRow fieldRow = sheet.createRow(j);
                            String fieldCode = fieldInfo.getCode();
                            String fieldCamel = fieldInfo.getCamel();
                            String fieldName = fieldInfo.getName();


                            int fieldCodeLength = fieldCode.length();
                            int fieldCamelLength = fieldCamel.length();
                            int fieldNameLength = fieldName.length();
                            if (fieldCodeLength > maxFieldLength[0]) {
                                maxFieldLength[0] = fieldCodeLength;
                            }
                            if (fieldCamelLength > maxFieldLength[1]) {
                                maxFieldLength[1] = fieldCamelLength;
                            }
                            if (fieldNameLength > maxFieldLength[2]) {
                                maxFieldLength[2] = fieldNameLength;
                            }

                            HSSFCell fieldCell0 = fieldRow.createCell(0);
                            HSSFCell fieldCell1 = fieldRow.createCell(1);
                            HSSFCell fieldCell2 = fieldRow.createCell(2);
                            fieldCell0.setCellValue(fieldCode);
                            fieldCell1.setCellValue(fieldCamel);
                            fieldCell2.setCellValue(fieldName);

                            // 如果有企业id，设为红色字体
                            if (fieldCode.equals("company_id")){
                                hasCompanyId = true;
                                fieldCell0.setCellStyle(redStyle);
                                fieldCell1.setCellStyle(redStyle);
                                fieldCell2.setCellStyle(redStyle);
                            }

                        }
                    }
                    /*
                     * 写表总体信息
                     */
                    HSSFRow tableRow = sheetTotal.createRow(i);


                    if (sheetName.length() > 31) {
                        sheetName = sheetName.substring(0, 31);
                    }

                    int tableCodeLength = tableCode.length();
                    int tableCamelLength = tableCamel.length();
                    int tableNameLength = tableName.length();
                    if (tableCodeLength > maxTableLength[0]) {
                        maxTableLength[0] = tableCodeLength;
                    }
                    if (tableCamelLength > maxTableLength[1]) {
                        maxTableLength[1] = tableCamelLength;
                    }
                    if (tableNameLength > maxTableLength[2]) {
                        maxTableLength[2] = tableNameLength;
                    }

                    HSSFCell tableCell0 = tableRow.createCell(0);
                    HSSFCell tableCell1 = tableRow.createCell(1);
                    HSSFCell tableCell2 = tableRow.createCell(2);
                    tableCell0.setCellValue(tableCode);
                    tableCell1.setCellValue(tableCamel);
                    tableCell2.setCellValue(tableName);
                    // 连接跳转

                    fieldLink.setAddress("'" + sheetName + "'!A2");
                    tableCell0.setHyperlink(fieldLink);

                    // 如果有企业id，设为红色字体
                    if (hasCompanyId) {
                        tableCell0.setCellStyle(redStyle);
                        hasCompanyId = false;
                    }

                    //设置宽度
                    sheet.setColumnWidth(0, (maxFieldLength[0] + 1) * 256);
                    sheet.setColumnWidth(1, (maxFieldLength[1] + 1) * 256);
                    sheet.setColumnWidth(2, (maxFieldLength[2] + 1) * 256);
                    maxFieldLength = new int[]{0, 0, 0};
                }
            }
            //设置宽度
            sheetTotal.setColumnWidth(0, (maxTableLength[0] + 1) * 256);
            sheetTotal.setColumnWidth(1, (maxTableLength[1] + 1) * 256);
            sheetTotal.setColumnWidth(2, (maxTableLength[2] + 1) * 256);
            //  写文件
            FileOutputStream fos = new FileOutputStream(filePlace);
            wb.write(fos);
            fos.close();

            System.out.println("生成结束");
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            try {
                if (conn != null) {
                    conn.close();
                }
            } catch (Exception e) {
                e.printStackTrace();
            }
        }
    }

    /**
     * 获取配置文件信息
     *
     * @author: 杜云章
     * @param:
     * @return: JSONObject
     * @Date: 2020/5/27 - 18:01
     */
    private static JSONObject getProperties() throws IOException {
        BufferedReader bufferedReader = new BufferedReader(new FileReader(GeneratorTableExcel.class.getResource(
                "/").toString().replaceFirst("file:/", "") + "whitelist.json"));
        String s;
        StringBuilder json = new StringBuilder();
        JSONObject infoJsonObject;
        while ((s = bufferedReader.readLine()) != null) {
            json.append(s);
        }
        bufferedReader.close();
        infoJsonObject = JSONObject.fromObject(json.toString().replaceAll("\t", ""));
        return infoJsonObject;
    }

    /**
     * <p>Title:underline2camel</p>
     * <p>Description: 将下划线显示转换成驼峰命名，例如member_id转成memberId</p>
     *
     * @return String
     * @param: param
     */
    private static String underline2camel(String param) {
        if (isNullOrEmpty(param)) {
            return "";
        }
        StringBuilder sb = new StringBuilder(param);
        Matcher mc = Pattern.compile("_").matcher(param);
        int i = 0;
        while (mc.find()) {
            int position = mc.end() - (i++);
            sb.replace(position - 1, position + 1, sb.substring(position, position + 1).toUpperCase());
        }
        return sb.toString();
    }

    private static boolean isNullOrEmpty(String s) {
        return s == null || s.length() == 0;
    }

    public static void main(String[] args) {
        generator("D:/wb.xls");
    }
}
