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
package com.hqjl.common.excel.service;

import com.hqjl.common.excel.bean.FormatType;
import java.util.Collection;

import com.hqjl.common.excel.bean.CellBean;
import com.hqjl.common.excel.bean.TableBean;
import com.hqjl.common.excel.json.JsonUtil;
import com.hqjl.common.excel.utils.ObjectHelper;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFDataFormat;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;

/**
 * 导出一个表格
 * Created by Hyberbin on 2014/6/18.
 */
public class ExportTableService extends BaseExcelService {
    private final  Sheet sheet;
    private final TableBean tableBean;

    public ExportTableService(Sheet sheet, TableBean tableBean) {
        this.sheet = sheet;
        this.tableBean = tableBean;
        ini();
    }

    /**
     * 初始化单元格
     */
    private void ini(){
        log.debug("初始化单元格总共:{}行，{}列",tableBean.getRowCount(),tableBean.getColumnCount());
        for(int r=0;r<tableBean.getRowCount();r++){
           Row row=sheet.getRow(r)!=null?sheet.getRow(r):sheet.createRow(r);
            if(tableBean.getRowHeight()!=null){
                row.setHeightInPoints(tableBean.getRowHeight());
            }
            for(int c=0;c<tableBean.getColumnCount();c++){
              if(row.getCell(c)==null){
                  row.createCell(c);
              }
            }
        }
    }

    public void doExport(){
        Collection<CellBean> cellBeans = tableBean.getCellBeans();
        if(ObjectHelper.isNotEmpty(cellBeans)){
            for(CellBean cellBean:cellBeans){
                if(cellBean.getXSize()>1||cellBean.getYSize()>1){
                    log.debug("有合并单元格：{}", JsonUtil.toJSON(cellBean));
                    CellRangeAddress range=new CellRangeAddress(cellBean.getRowIndex(),cellBean.getRowIndex()+cellBean.getYSize()-1,cellBean.getColumnIndex(),cellBean.getColumnIndex()+cellBean.getXSize()-1);
                    sheet.addMergedRegion(range);
                }
                log.debug("set row:{},column:{},content:{}",cellBean.getRowIndex(),cellBean.getColumnIndex(),cellBean.getContent());
                Cell cell = sheet.getRow(cellBean.getRowIndex()).getCell(cellBean.getColumnIndex());

                CellStyle cellStyle= cell.getCellStyle();
                if(cellStyle==null){
                    cellStyle=sheet.getWorkbook().createCellStyle();
                }
                if(cellBean.isAlignCenter()){
                    //水平居中
                    cellStyle.setAlignment(CellStyle.ALIGN_CENTER);
                }
                if(cellBean.isVerticalCenter()){
                    //垂直居中
                    cellStyle.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
                }
                //    cellStyle.setWrapText(cellBean.isWrapText());
                if(FormatType.PERCENT==FormatType.valueOf(cellBean.getNumType())){
                    //设置为百分比格式
                    cellStyle.setDataFormat(HSSFDataFormat.getBuiltinFormat("0.00%"));
                }
                if(FormatType.DECIMALS==FormatType.valueOf(cellBean.getNumType())){
                    //设置为小数
                    cellStyle.setDataFormat(HSSFDataFormat.getBuiltinFormat("0.00"));
                }
                cell.setCellStyle(cellStyle);
                try {
                    //对字符窜的处理
                    if(FormatType.STR==FormatType.valueOf(cellBean.getNumType())){
                        cell.setCellValue(cellBean.getContent());
                    }else {
                        //默认的处理方法
                        double v = Double.parseDouble(cellBean.getContent().trim());
                        cell.setCellValue(v);
                    }
                }catch (NumberFormatException e){
                    cell.setCellValue(cellBean.getContent());
                }
            }
        }
    }
}
