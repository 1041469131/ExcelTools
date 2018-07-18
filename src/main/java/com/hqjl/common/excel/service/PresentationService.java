package com.hqjl.common.excel.service;

import com.hqjl.common.excel.utils.ObjectHelper;
import java.io.BufferedOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.net.URLEncoder;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * @author xiang.li
 * create date 2018/3/12
 * description
 */
public class PresentationService {
    public static String showFileName = "Excel.xls";
    public static String encoding = "utf-8";
    public static String respEncode = "utf-8";
    public static String openType = "attachment";
    public static void doPresentation(HttpServletRequest request, HttpServletResponse response,String showFileName, Workbook workbook) throws Exception {
        BufferedOutputStream bufferedOutPut=null;
        try {
            request.setCharacterEncoding(encoding);
            showFileName = parseShowFileName(showFileName,workbook);
            openType = request.getParameter("openType");
            if ((openType == null) || (openType.trim().length() == 0)) {
                openType = "attachment";
            } else {
                openType = (openType.equals("inline") ? openType : "attachment");
            }
            response.setContentType("application/vnd.ms-excel;charset=UTF-8");
            String str = URLEncoder.encode(showFileName,respEncode);
            response.setHeader("Content-disposition",openType + "; filename*=utf-8'zh_cn'" + str);
            OutputStream output = response.getOutputStream();
            bufferedOutPut = new BufferedOutputStream(output);
            bufferedOutPut.flush();
            workbook.write(bufferedOutPut);
        }finally {
            if(null != bufferedOutPut){
                try {
                    bufferedOutPut.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
        }
    }

    public static String parseShowFileName(String showFileName, Workbook workbook){
        if(ObjectHelper.isEmpty(showFileName)){
            if(workbook instanceof XSSFWorkbook){
                showFileName="Excel.xlsx";
            }else {
                showFileName="Excel.xls";
            }
        }else{
            if(workbook instanceof XSSFWorkbook){
              if(showFileName.endsWith(".xlsx")){
                  return showFileName;
              }else{
                 return showFileName.split("\\.")[0]+".xlsx";
              }
            }else{
                if(showFileName.endsWith(".xls")){
                    return showFileName;
                }else{
                    return showFileName.split("\\.")[0]+".xls";
                }
            }

        }
        return showFileName;
    }

}
