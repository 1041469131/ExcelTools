/*
 * Copyright 2015 www.hyberbin.com.
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *      http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 * Email:hyberbin@qq.com
 */
package com.hqjl.common.excel;

import com.hqjl.common.excel.bean.*;
import com.hqjl.common.excel.service.*;
import org.apache.poi.hssf.usermodel.HSSFClientAnchor;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Drawing;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Before;
import org.junit.BeforeClass;
import org.junit.Test;

import java.awt.image.BufferedImage;
import java.io.*;
import java.lang.reflect.Field;
import java.util.*;
import java.util.stream.Collectors;

import com.hqjl.common.excel.language.ILanguage;

import javax.imageio.ImageIO;

/**
 *
 * @author Hyberbin
 */
public class TestExcel {
    private Workbook workbook;

    @BeforeClass
    public static void setUpClass() {
    }

    @Before
    public void setUp() {
        workbook = new HSSFWorkbook();

    }
    private static Map buildMap(String id,String kcmc,String kclx){
        Map map=new HashMap();
        map.put("id", id);
        map.put("kcmc", kcmc);
        map.put("kclx", kclx);
        return map;
    }
    private static List<SchoolCourse> getList(){
        List<SchoolCourse> list=new ArrayList<SchoolCourse>();
        list.add(new SchoolCourse("1", "语文","1"));
        list.add(new SchoolCourse("2", "数学","1"));
        list.add(new SchoolCourse("3", "英语", "1"));
        list.add(new SchoolCourse("4", "政治", "2"));
        list.add(new SchoolCourse("5", "历史", "2"));
        return list;
    }

    private static List<Map> getMapList(){
        List<Map> list=new ArrayList<Map>();
        list.add(buildMap("1", "语文","1"));
        list.add(buildMap("2", "数学","1"));
        list.add(buildMap("3", "英语","1"));
        list.add(buildMap("4", "政治","2"));
        list.add(buildMap("5", "历史","2"));
        return list;
    }
    /**
     * 从List中导出
     * @throws Exception
     */
    @Test
    public void testSimpleMapExport() throws Exception {
        Sheet sheet = workbook.createSheet("testSimpleMapExport");
        SimpleExportService service = new SimpleExportService(sheet, getMapList(), new String[]{"id","kcmc","kclx"}, null);
        //如果要表头可以像下面这样设置,不要表头可以不写
        service.setLanguage(new ILanguage() {
            @Override
            public String translate(Object key, Object... args) {
                if("id".equals(key)){
                    return "序号";
                }else if("kcmc".equals(key)){
                    return "课程名称";
                }else if("kclx".equals(key)){
                    return "课程类型";
                }
                return key+"";
            }
        });
        service.setDic("KCLX", "KCLX").addDic("KCLX", "1", "国家课程").addDic("KCLX", "2", "学校课程");//设置数据字典
        service.COLUMN_ROW =0;
        service.START_ROW =1;
        service.doExport();
        FileOutputStream fos = new FileOutputStream("E:/test3.xls");
        workbook.write(fos);
    }

    /**
     * 从Excel中直接导入
     */
    @Test
    public void testSimpleImport()throws Exception {
//        testTableExport();
        File file=new File("E:/test.xlsx");
        InputStream inputStream =new FileInputStream(file);
       // workbook = new HSSFWorkbook(inputStream);
        workbook = WorkbookFactory.create(inputStream);

        Sheet sheet = workbook.getSheet("Sheet1");
        ImportTableService tableService=new ImportTableService(sheet);
        tableService.setStartRow(1);
        tableService.doImport();
        //直接读取到List中,泛型可以是Map也可以是PO
        //第一个参数是从表格第0列开始依次读取内容放到哪些字段中
//        List<Map> read = tableService.read(new String[]{"a","b","c"}, Map.class);
        List<SchoolCourse> read2 = tableService.read(new String[]{"id","courseName","type"}, SchoolCourse.class);
        System.out.print(read2);
    }

    /**
     * 从Excel中直接导入
     */
    @Test
    public void testImportToBean()throws Exception {
        File file=new File("E:/test0.xlsx");
        InputStream inputStream =new FileInputStream(file);
        workbook = WorkbookFactory.create(inputStream);
        Sheet sheet = workbook.getSheet("Sheet3");
        ImportTableService tableService=new ImportTableService(sheet);
        tableService.setStartRow(0);
        tableService.doImport();
        //直接读取到List中,泛型可以是Map也可以是PO
        //第一个参数是从表格第0列开始依次读取内容放到哪些字段中
//        List<Map> read = tableService.read(new String[]{"a","b","c"}, Map.class);
        List<Map> read2 = tableService.read(new String[]{"id","name","date"}, Map.class);
        System.out.print(read2);
    }

    /**
     * 从List<Vo>中导出
     * @throws Exception
     */
    @Test
    public void testSimpleVoExport() throws Exception {
        Sheet sheet = workbook.createSheet("testSimpleVoExport");
        //ExportExcelService service = new ExportExcelService(list, sheet, "学校课程");
        ExportExcelService service = new ExportExcelService(getList(), sheet, new String[]{"id", "courseName", "type"}, "学校课程");
        service.addDic("KCLX", "1", "国家课程").addDic("KCLX", "2", "学校课程");//设置数据字典
        service.TITLE_ROW =4;
        service.COLUMN_ROW =5;
        service.START_ROW =6;
        service.doExport();

        FileOutputStream fos = new FileOutputStream("E:/test2.xls");
        workbook.write(fos);
        if(null != fos){
            fos.close();
        }

    }

    /**
     * 从List<Vo>，vo中还有简单循环节中导出
     * @throws Exception
     */
    @Test
    public void testVoHasListExport() throws Exception {
        List<String> strings = new ArrayList<String>();
        List<SchoolCourse> list = getList();
        for (int i = 0; i < 10; i++) {
            strings.add("我是第" + i + "个循环字段");
        }
        for (SchoolCourse course : list) {
            course.setBaseArray(strings);
        }
        Sheet sheet = workbook.createSheet("testVoHasListExport");
        ExportExcelService service = new ExportExcelService(list, sheet, new String[]{"id", "courseName", "type", "baseArray"}, "学校课程");
        service.addDic("KCLX", "1", "国家课程").addDic("KCLX", "2", "学校课程");//设置数据字典
        service.setGroupConfig("baseArray", new GroupConfig(10) {

            @Override
            public String getLangName(int innerIndex, int index) {
                return "我是第" + index + "个循环字段";
            }
        });
        service.doExport();
        service.exportTemplate();//生成下拉框
    }

    /**
     * 从List<Vo>，vo中还有复杂循环节中导出
     * @throws Exception
     */
    @Test
    public void testVoHasListVoExport() throws Exception {
        List<SchoolCourse> list = getList();
        for (SchoolCourse course : list) {
            List<InnerVo> innerVos = new ArrayList<InnerVo>();
            for (int i = 0; i < 10; i++) {
                innerVos.add(new InnerVo("key"+i, "value"+i));
            }
            course.setInnerVoArray(innerVos);
        }
        File file=new File("D:/tubiao_template.xlsx");
        InputStream inputStream =new FileInputStream(file);

        workbook=WorkbookFactory.create(inputStream);
        Field field = workbook.getClass().getDeclaredField ("_sheets");
        List<HSSFSheet> sheets=(List<HSSFSheet>) field.get (workbook);
        HSSFSheet sheet4 = HSSFSheet.class.newInstance();
        sheets.add(sheet4);
        sheet4.setSelected(false);
        sheet4.setActive(false);
        Sheet sheet2 =workbook.getSheet("Sheet1");
        Sheet sheet = workbook.createSheet("testVoHasListVoExport");
        Sheet sheet1 = workbook.createSheet("testVoHasListVoExport1");

        ExportExcelService service = new ExportExcelService(list, sheet, new String[]{"id", "courseName", "type", "innerVoArray"}, "学校课程");
        ExportExcelService service1 = new ExportExcelService(list, sheet1, new String[]{"id", "courseName", "type", "innerVoArray"}, "学校课程");
        TestMode array1=new TestMode(22,33,23,24,25);
        TestMode array2=new TestMode(32,53,4,34,35);
        TestMode array3=new TestMode(42,63,33,14,45);
        TestMode array4=new TestMode(52,13,23,24,22);
        List<TestMode> list1=new ArrayList<> ();
        list1.add (array1);
        list1.add (array2);
        list1.add (array3);
        list1.add (array4);
        SimpleExportService service2 = new SimpleExportService(sheet2,list1,  new String[]{"a", "b", "c", "d","e"}, null);

        service.addDic("KCLX", "1", "国家课程").addDic("KCLX", "2", "学校课程");//设置数据字典
        for (int i = 0; i < 10; i++) {
            service.addTook("hiddenvalue", "key", i, "something");
        }
        service.setGroupConfig("innerVoArray", new GroupConfig(2, 10) {
            @Override
            public String getLangName(int innerIndex, int index) {
                return "我是第" + index + "个循环字段,第" + innerIndex + "个属性";
            }
        });
        service.doExport();

        service1.addDic("KCLX", "1", "国家课程").addDic("KCLX", "2", "学校课程");//设置数据字典
        for (int i = 0; i < 10; i++) {
            service.addTook("hiddenvalue", "key", i, "something");
        }

        service1.setGroupConfig("innerVoArray", new GroupConfig(2, 10) {
            @Override
            public String getLangName(int innerIndex, int index) {
                return "我是第" + index + "个循环字段,第" + innerIndex + "个属性";
            }
        });



        service1.doExport ();
        service2.COLUMN_ROW =0;
        service2.START_ROW =1;
        service2.doExport ();


        FileOutputStream fos = new FileOutputStream("E:/tubiao.xlsx");
        workbook.write(fos);
    }

    @Test
    public void testVoHasListVoExport1() throws Exception {
        List<SchoolCourse> list = getList();
        for (SchoolCourse course : list) {
            List<InnerVo> innerVos = new ArrayList<InnerVo>();
            List<InnerVo> innerVos1 = new ArrayList<InnerVo>();
            for (int i = 0; i < 5; i++) {
                innerVos.add(new InnerVo("↓", "↑"));
            }
            for (int i = 0; i < 5; i++) {
                innerVos1.add(new InnerVo("code"+i, "name"+i));
            }
            course.setInnerVoArray(innerVos);
            course.setInnerVoArray1(innerVos1);
        }
        Sheet sheet = workbook.createSheet("testVoHasListVoExport");
        ExtExportExcelService service = new ExtExportExcelService(list, sheet, new String[]{"id", "courseName", "type", "innerVoArray","innerVoArray1"}, "学校课程");
        service.addDic("KCLX", "1", "国家课程").addDic("KCLX", "2", "学校课程");//设置数据字典
//        for (int i = 0; i < 10; i++) {
//            service.addTook("hiddenvalue", "key", i, "something");
//        }
        GroupConfig groupConfig = new GroupConfig(1, 5) {
            @Override
            public String getLangName(int innerIndex, int index) {

                return "key" + index +",第"+ innerIndex ;
            }
        };
        groupConfig.addField("key");
        service.setGroupConfig("innerVoArray", groupConfig);

        GroupConfig groupConfig1 = new GroupConfig(2, 5) {
            @Override
            public String getLangName(int innerIndex, int index) {
                return "code" + index + ",第" + innerIndex ;
            }
        };
        service.setGroupConfig("innerVoArray1",groupConfig1 );
        groupConfig1.addField("key");
        groupConfig1.addField("value");
        service.putExtFileName("innerVoArray1", new String[]{"key1","key2","key3","key4","key5"});
        service.doExport();
        FileOutputStream fos = new FileOutputStream("E:/tubiao1.xls");
        workbook.write(fos);
    }

    /**
     * 导出一个纵表（课程表之类的）
     * @throws Exception
     */
    @Test
    public void testTableExport1() throws Exception {
        //TODO 读入导出模板
        Sheet sheet = workbook.createSheet("testTableExport");
        TableBean tableBean = new TableBean(3, 3);
        Collection<CellBean> cellBeans = new HashSet<CellBean>();
        for (int i = 0; i < 3; i++) {
            for (int j = 0; j < 3; j++) {
                CellBean cellBean = new CellBean(i * 3 + j + "", i, j);
                cellBeans.add(cellBean);
            }
        }
        tableBean.setCellBeans(cellBeans);
        ExportTableService tableService = new ExportTableService(sheet, tableBean);
        tableService.doExport();
        FileOutputStream fos = new FileOutputStream("D:/zongbiao.xls");
        workbook.write(fos);
    }

    /**
     * 导出一个纵表（课程表之类的）
     * @throws Exception
     */
    @Test
    public void testTableExport() throws Exception {

        String path=this.getClass().getResource("/").getPath() + "excel/";
        //TODO 读入导出模板
        File file=new File(path+"出口分析.xlsx");
        InputStream inputStream =new FileInputStream(file);
        workbook = WorkbookFactory.create(inputStream);
        ExcelSheetData excelSheetData=new ExcelSheetData ();
        excelSheetData.setSheetName ("出口数据总览-总体数据");
        List<ExcelTableData> tables=new ArrayList<> ();
        excelSheetData.setTables (tables);
        ExcelTableData excelTableData=new ExcelTableData ();
        tables.add (excelTableData);
        List<List<ExcelCellData>> headers=new ArrayList<> ();
        List<ExcelCellData> headerFirst=new ArrayList<> ();
        headerFirst.add (new ExcelCellData ("科目",1,2));
        headerFirst.add (new ExcelCellData ("总人数",1,2));
        headerFirst.add (new ExcelCellData ("最高",2,1));
        headerFirst.add (new ExcelCellData ("最低",2,1));
        headerFirst.add (new ExcelCellData ("平均",2,1));
        headerFirst.add (new ExcelCellData ("中位",2,1));
        headers.add (headerFirst);
        List<ExcelCellData> headerSecond=new ArrayList<> ();
        headerSecond.add (new ExcelCellData ("subject","",1,1));
        headerSecond.add (new ExcelCellData ("totalCount","",1,1));
        headerSecond.add (new ExcelCellData ("maxScore","分数",1,1));
        headerSecond.add (new ExcelCellData ("maxRank","排名",1,1));
        headerSecond.add (new ExcelCellData ("minScore","分数",1,1));
        headerSecond.add (new ExcelCellData ("minRank","排名",1,1));
        headerSecond.add (new ExcelCellData ("avgScore","分数",1,1));
        headerSecond.add (new ExcelCellData ("avgRank","排名",1,1));
        headerSecond.add (new ExcelCellData ("midScore","分数",1,1));
        headerSecond.add (new ExcelCellData ("midRank","排名",1,1));
        headers.add (headerSecond);
        excelTableData.setHeaders (headers);
        List<Map<String,Object>> rows=new ArrayList<> ();
        Map map1=new HashMap ();
        map1.put ("subject","物化生（理科）");
        map1.put ("totalCount",0.28);
        map1.put ("maxScore","680");
        map1.put ("maxRank","659");
        map1.put ("minScore","161");
        map1.put ("minRank","13506");
        map1.put ("avgScore","552");
        map1.put ("avgRank","4034");
        map1.put ("midScore","571");
        map1.put ("midRank","3998");
        rows.add (map1);
        Map map2=new HashMap ();
        map2.put ("subject","物化生（理科）");
        map2.put ("totalCount","180");
        map2.put ("maxScore","680");
        map2.put ("maxRank","659");
        map2.put ("minScore","161");
        map2.put ("minRank","13506");
        map2.put ("avgScore","552");
        map2.put ("avgRank","4034");
        map2.put ("midScore","571");
        map2.put ("midRank","3998");
        rows.add (map2);
        excelTableData.setRows (rows);
        ExcelTableData excelTableData1=new ExcelTableData ();
        tables.add (excelTableData1);
        ExcelCellData title1=new ExcelCellData ("出口数据总览-总体数据",10,1);
        excelTableData1.setTitle (title1);
        List<List<ExcelCellData>> headers1=new ArrayList<> ();
        List<ExcelCellData> headerFirst1=new ArrayList<> ();
        headerFirst1.add (new ExcelCellData ("科目",1,2));
        headerFirst1.add (new ExcelCellData ("总人数",1,2));
        headerFirst1.add (new ExcelCellData ("最高",2,1));
        headerFirst1.add (new ExcelCellData ("最低",2,1));
        headerFirst1.add (new ExcelCellData ("平均",2,1));
        headerFirst1.add (new ExcelCellData ("中位",2,1));
        headers1.add (headerFirst1);
        List<ExcelCellData> headerSecond1=new ArrayList<> ();
        headerSecond1.add (new ExcelCellData ("subject","",1,1));
        headerSecond1.add (new ExcelCellData ("totalCount","",1,1));
        headerSecond1.add (new ExcelCellData ("maxScore","分数",1,1));
        headerSecond1.add (new ExcelCellData ("maxRank","排名",1,1));
        headerSecond1.add (new ExcelCellData ("minScore","分数",1,1));
        headerSecond1.add (new ExcelCellData ("minRank","排名",1,1));
        headerSecond1.add (new ExcelCellData ("avgScore","分数",1,1));
        headerSecond1.add (new ExcelCellData ("avgRank","排名",1,1));
        headerSecond1.add (new ExcelCellData ("midScore","分数",1,1));
        headerSecond1.add (new ExcelCellData ("midRank","排名",1,1));
        headers1.add (headerSecond1);
        excelTableData1.setHeaders (headers1);
        List<Map<String,Object>> rows1=new ArrayList<> ();
        Map map3=new HashMap ();
        map3.put ("subject","物化生（理科）");
        map3.put ("totalCount","180");
        map3.put ("maxScore","680");
        map3.put ("maxRank","659");
        map3.put ("minScore","161");
        map3.put ("minRank","13506");
        map3.put ("avgScore","552");
        map3.put ("avgRank","4034");
        map3.put ("midScore","571");
        map3.put ("midRank","3998");
        rows1.add (map3);
        excelTableData1.setRows (rows1);
        int startRowCount=0,endColumnCount=0,rowCount=0,columnCount=0;
        Collection<CellBean> cellBeans= new HashSet<CellBean>();
       for( ExcelTableData tableData:excelSheetData.getTables ()){
            ExcelCellData titleData=tableData.getTitle ();
            if(titleData!=null&&titleData.getName ()!=null){
                CellBean cellBean = new CellBean(titleData.getName (), startRowCount, endColumnCount,titleData.getxSize (),titleData.getySize ());
                cellBeans.add (cellBean);
                columnCount=titleData.getxSize ()>columnCount?titleData.getxSize ():columnCount;
                startRowCount++;
                rowCount++;
            }
            List<List<ExcelCellData>>  headersDataList=tableData.getHeaders ();
            for(List<ExcelCellData> headersData:headersDataList){
                int headerColumnCount=0;
                for(ExcelCellData headerData:headersData){
                    if(headerData.getName()!=null&&!headerData.getName ().equals ("")) {
                        CellBean cellBean = new CellBean (headerData.getName (), startRowCount, headerColumnCount, headerData.getxSize (), headerData.getySize ());
                        cellBeans.add (cellBean);
                    }
                    headerColumnCount+=headerData.getxSize ();
                }
                columnCount=headerColumnCount>columnCount?headerColumnCount:columnCount;
                startRowCount++;
                rowCount++;

            }
            List<Map<String,Object>> rowsData=tableData.getRows ();
            List<String> keys=headersDataList.get (headersDataList.size ()-1).stream ().map (e->e.getKey ()).collect(Collectors.toList());
            for(Map<String,Object> rowData:rowsData){
                int rowColumnCount=0;
                for(String key:keys){
                    CellBean cellBean = new CellBean(String.valueOf(rowData.get (key)), startRowCount, rowColumnCount);
                    cellBeans.add (cellBean);
                    rowColumnCount++;
                }
                columnCount=rowColumnCount>columnCount?rowColumnCount:columnCount;
                startRowCount++;
                rowCount++;
            }
        }
        TableBean tableBean = new TableBean(rowCount, columnCount);
        tableBean.setCellBeans(cellBeans);
        ExportTableService tableService = new ExportTableService(workbook.getSheet (excelSheetData.getSheetName ()), tableBean);
        tableService.doExport ();
        FileOutputStream fos = new FileOutputStream("E:/出口分析测试.xlsx");
        workbook.write(fos);
    }
    @Test
    public void testTableImport()throws Exception {
        testTableExport();
        Sheet sheet = workbook.getSheet("testTableExport");
        ImportTableService tableService=new ImportTableService(sheet);
        tableService.doImport();
        TableBean tableBean = tableService.getTableBean();
        System.out.println(tableBean.getCellBeans().size());
    }

    /**
     * 从List<Vo>中入
     * @throws Exception
     */
    @Test
    public void testSimpleVoImport() throws Exception {
//        testSimpleVoExport();
//        Sheet sheet = workbook.getSheet("testSimpleVoExport");
        workbook = new HSSFWorkbook(new FileInputStream("D:/test0.xls"));
        Sheet sheet = workbook.getSheet("testSimpleVoExport");
        ImportExcelService service = new ImportExcelService(SchoolCourse.class, sheet);
        service.addDic("KCLX", "1", "国家课程").addDic("KCLX", "2", "学校课程");//设置数据字典
        List list = service.doImport();
        List list2 = service.getErrorList();

        FileOutputStream fos = new FileOutputStream("D:/test00.xls");
        workbook.write(fos);
        if(null != fos){
            fos.close();
        }
        System.out.println("成功导入：" + list.size() + "条数据");
    }

    /**
     * 从List<Vo>，vo中还有简单循环节中导入
     * @throws Exception
     */
    @Test
    public void testVoHasListImport() throws Exception {
        testVoHasListExport();
        Sheet sheet = workbook.getSheet("testVoHasListExport");
        ImportExcelService service = new ImportExcelService(SchoolCourse.class, sheet);
        service.addDic("KCLX", "1", "国家课程").addDic("KCLX", "2", "学校课程");//设置数据字典
        List list = service.doImport();
        System.out.println("成功导入：" + list.size() + "条数据");
    }

    /**
     * 从List<Vo>，vo中还有复杂循环节中导入
     * @throws Exception
     */
    @Test
    public void testVoHasListVoImport() throws Exception {
        testVoHasListVoExport();
        Sheet sheet = workbook.getSheet("testVoHasListVoExport");
        ImportExcelService service = new ImportExcelService(SchoolCourse.class, sheet);
        service.addDic("KCLX", "1", "国家课程").addDic("KCLX", "2", "学校课程");//设置数据字典
        List list = service.doImport();
        System.out.println("成功导入：" + list.size() + "条数据");
    }
    @Test
    public void testfile() throws IOException {
        InputStream inputStream=this.getClass().getClassLoader().getResourceAsStream(
            "application.properties");
        Properties properties = new Properties();
        properties.load(inputStream);
        System.out.println(properties.getProperty("spring.jpa.show-sql"));
    }
    @Test
    public void testImages() throws IOException {
        FileOutputStream fileOut = null;
        BufferedImage bufferImg = null;//图片一
        BufferedImage bufferImg1 = null;//图片二
        try {
            // 先把读进来的图片放到一个ByteArrayOutputStream中，以便产生ByteArray
            ByteArrayOutputStream byteArrayOut = new ByteArrayOutputStream();
            ByteArrayOutputStream byteArrayOut1 = new ByteArrayOutputStream();

            //将两张图片读到BufferedImage
            bufferImg = ImageIO.read(new File("C:/Users/10414/Pictures/Saved Pictures/image.jpg"));
            bufferImg1 = ImageIO.read(new File("C:/Users/10414/Pictures/Saved Pictures/image.jpg"));
            ImageIO.write(bufferImg, "png", byteArrayOut);
            ImageIO.write(bufferImg1, "png", byteArrayOut1);

            // 创建一个工作薄
           // HSSFWorkbook wb = new HSSFWorkbook();
            Sheet  sheet= workbook.createSheet("testSimpleMapExport");
            //创建一个sheet
          //  HSSFSheet sheet = wb.createSheet("out put excel");

            Drawing drawing = sheet.createDrawingPatriarch();
            /**
             * 该构造函数有8个参数
             * 前四个参数是控制图片在单元格的位置，分别是图片距离单元格left，top，right，bottom的像素距离
             * 后四个参数，前连个表示图片左上角所在的cellNum和 rowNum，后天个参数对应的表示图片右下角所在的cellNum和 rowNum，
             * excel中的cellNum和rowNum的index都是从0开始的
             *
             */
            //图片一导出到单元格B2中
            HSSFClientAnchor anchor = new HSSFClientAnchor (0, 0, 0, 0,
                    (short) 1, 1, (short) 2, 2);
            //图片二导出到单元格C3到E5中，且图片的left和top距离边框50
            HSSFClientAnchor anchor1 = new HSSFClientAnchor(50, 50, 0, 0,
                    (short) 2, 2, (short) 5, 5);

            // 插入图片
            drawing.createPicture(anchor, workbook.addPicture(byteArrayOut
                    .toByteArray(), HSSFWorkbook.PICTURE_TYPE_JPEG));
            drawing.createPicture(anchor1, workbook.addPicture(byteArrayOut1
                    .toByteArray(), HSSFWorkbook.PICTURE_TYPE_JPEG));

            fileOut = new FileOutputStream("D:/testImage.xls");
            // 写入excel文件
            workbook.write(fileOut);
        } catch (IOException io) {
            io.printStackTrace();
            System.out.println("io erorr : " + io.getMessage());
        } finally {
            if (fileOut != null) {
                try {
                    fileOut.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
        }
    }
    /**
     * 从Excel中直接导入
     */
    @Test
    public void testYinsheSimpleImport()throws Exception {
//        testTableExport();
        File file=new File("E:/映射数据管理.xlsx");
        InputStream inputStream =new FileInputStream(file);
        // workbook = new HSSFWorkbook(inputStream);
        workbook = WorkbookFactory.create(inputStream);

        Sheet sheet = workbook.getSheet("映射原始成绩单");
        ImportTableService tableService=new ImportTableService(sheet);
       tableService.setStartRow(1);
        tableService.doImport();
        //直接读取到List中,泛型可以是Map也可以是PO
        //第一个参数是从表格第0列开始依次读取内容放到哪些字段中
//        List<Map> read = tableService.read(new String[]{"a","b","c"}, Map.class);
        List<TestBean1> read2 = tableService.read(new String[]{"id","courseName","type"}, TestBean1.class);
        System.out.print(read2);
    }


    @Test
    public void testPresentationService()throws Exception {
        Workbook xssfWorkbook = new XSSFWorkbook();
        Workbook hssfWorkbook = new HSSFWorkbook();
        System.out.println(PresentationService.parseShowFileName("sss.xls",xssfWorkbook));
        System.out.println(PresentationService.parseShowFileName("sss.xlsx",xssfWorkbook));
        System.out.println(PresentationService.parseShowFileName("sss",xssfWorkbook));
        System.out.println(PresentationService.parseShowFileName("",xssfWorkbook));

        System.out.println(PresentationService.parseShowFileName("sss.xls",hssfWorkbook));
        System.out.println(PresentationService.parseShowFileName("sss.xlsx",hssfWorkbook));
        System.out.println(PresentationService.parseShowFileName("sss",hssfWorkbook));
        System.out.println(PresentationService.parseShowFileName("",hssfWorkbook));

    }

    public static void main(String[] args) {
        System.out.println();
    }


}
