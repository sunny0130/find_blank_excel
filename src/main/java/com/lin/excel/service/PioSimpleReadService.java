package com.lin.excel.service;

import cn.hutool.core.date.DateTime;
import com.lin.excel.constant.ExcelTypesEnum;
import com.lin.excel.controller.PioSimpleReadController;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.stereotype.Service;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Date;
import java.util.LinkedList;

/**
 * @description:
 * @author: lin
 * @date: 2021/11/11
 */
@Service
public class PioSimpleReadService {

    @Value("${myExcel.writeXlsPath}")
    String writeXlsPath;

    private static final Logger LOG = LoggerFactory.getLogger(PioSimpleReadController.class);


    /**
     * 读取目标excel 判断空格数 再写入到新的 excel
     */
    public void findBlankFromExcel(LinkedList<String> files) throws IOException {

        int currentRowNumber = 0;

        /** 放到循环外边 防止重复设置
         *  这里是统一生成为 xls 文件  有其他需求的 可以自己改
         */
        //创建要写入的excel文件
        HSSFWorkbook newExcel = new HSSFWorkbook();

        //创建excel文件的sheet
        HSSFSheet sheet= newExcel.createSheet("lin");

        //生成第一行 表头 标题
        HSSFRow title = sheet.createRow(0);

        //生成表头的每一列

        /** 当前文件名 */
        title.createCell(0).setCellValue("FILE_NAME");
        /** 生成文件的表头名称 */
        title.createCell(1).setCellValue("COLUMN_NAME");
        /** 当前文件单元格的 空格数量*/
        title.createCell(2).setCellValue("BLANK_NUMS");
        /** 当前文件单元格 为null的数量*/
        title.createCell(3).setCellValue("NULL_NUMS");
        /** 当前文件行数 */
        title.createCell(4).setCellValue("ROWS");
        /** 当前文件列数 */
        title.createCell(5).setCellValue("CELLS");

        FileOutputStream fileOutputStream = null;

        if (files != null && files.size() > 0) {

            for (String f : files) {


                /** 此list用于存放空格数
                 * 作用域为当前文件
                 * 多个文件下 需要每次新new
                 * 防止后续写入时无法拿到list后面的空格信息
                 * */
                ArrayList<Object> blankList = new ArrayList<>();
                ArrayList<Object> nullList = new ArrayList<>();

                File file = new File(f);

                Workbook wb = getWorkbook(file);
                if (wb == null) {
                    return;
                }
                /** 得到sheet */
                Sheet sheetAt = wb.getSheetAt(0);

                /** 表头 标题 第一行 */
                Row rowTitle = sheetAt.getRow(0);

                /** 所有列数 */
                int allCells = rowTitle.getPhysicalNumberOfCells();

                /** 所有行数  包括第一行  循环从第二行开始 */
                int allRows = sheetAt.getPhysicalNumberOfRows();

                /** null 空格数  和 错误数 这里需求只要空格数 其他暂时注掉 */
                int nullNum = 0;
                int blankNum = 0;
                //int errorNum = 0;

                /** 记录每一列的空格数 */
                int cNullNum = 0;
                int cBlankNum = 0;
                //int rErrorNum = 0;

                /** 我的需求是计算源文件的每一列的空单元格数量
                 *  因此这里外循环是源文件的列数 内循环是源文件的行数
                 *
                 *  这里可按自己的需求更改
                 */
                for (int i = 0; i < allCells; i++) {

                    for (int j = 1; j < allRows; j++) {

                        /** 获取当前excel的行数 从1开始  第0行为表头 */
                        Row row = sheetAt.getRow(j);

                        /** 获取当前excel的列数 从0开始 */
                        Cell cell = row.getCell(i);

                        if (cell != null){
                            String cellType = getCellType(cell);
                            //LOG.info("第{}列,第{}行 --->>> 单元格内容：{}", (i+1) , j  , cellType);

                            /** 类型为空格 则计数加一 */
                            if("blank".equals(cellType)){
                                blankNum += 1;
                                cBlankNum += 1;
                            }
                            /** 类型为空字符串 也加一 */
                            if(cellType.trim().isEmpty()){
                                blankNum += 1;
                                cBlankNum += 1;
                            }
                        }else {
                            /** 类型为null 则计数加一 */
                            nullNum += 1;
                            cNullNum += 1;
                        }

                    }

                    /** 将原文件每一列的空单元格数量 存入list 方便后续取出 */
                    blankList.add(cBlankNum);
                    nullList.add(cNullNum);

                    /** 由于计算的是每一列 下次循环前 需要将其清0  */
                    cNullNum = 0;
                    cBlankNum = 0;
                    //cErrorNum = 0;
                }

                /** 写入新的excel
                 * 注意 循环是 从 1 开始 到 当前文件的列数
                 * List 不清空的话  信息会一直汪 list 里放
                 * 从而导致拿不到对应的文件信息
                 *
                 * 单个文件下 不影响
                 *
                 * */

                for (int row = 1 ; row <= allCells ; row++) {
                    try {
                        HSSFRow dataRow = sheet.createRow(row + currentRowNumber );
                        dataRow.createCell(0).setCellValue(file.getName());
                        dataRow.createCell(1).setCellValue(rowTitle.getCell(row-1).toString());
                        dataRow.createCell(2).setCellValue(String.valueOf(blankList.get(row-1)));
                        dataRow.createCell(3).setCellValue(String.valueOf(nullList.get(row-1)));
                        dataRow.createCell(4).setCellValue(allRows);
                        dataRow.createCell(5).setCellValue(allCells);

                        //这里就会异常,如果文件名不存在的话。
                        fileOutputStream = new FileOutputStream(new File(writeXlsPath));

                        /** 写入到新的excel */
                        newExcel.write(fileOutputStream);

                        fileOutputStream.flush();
                    }
                    catch (IOException e) {
                        //这个主要是把出现的异常给人看见，不然就算异常了，看不到就找不到问题所在。
                        LOG.error("loadProperties IOException:" + e.getMessage());
                    } finally {
                        if (fileOutputStream != null) {
                            try {
                                fileOutputStream.close(); // 关闭流
                            } catch (IOException e) {
                                LOG.error("inputStream close IOException:" + e.getMessage());
                            }
                        }
                    }
                }

                currentRowNumber += allCells;

                LOG.info("咦，一共发现了 ["+ nullNum +"] " + "个null！！！");
                LOG.info("啊，一共发现了 ["+ blankNum +"] " + "个blank！！！");

            }
        }else {
            LOG.error("文件为空哦！");
        }
    }

    /**
     * 判断excel类型 生成 对应的 Workbook
     * @param file
     * @return
     * @throws IOException
     */
    private Workbook getWorkbook(File file) throws IOException {
        //如果是xls，使用HSSFWorkbook；如果是xlsx，使用XSSFWorkbook
        if (!file.exists()) {
            LOG.error("文件不存在");
            return null;
        }
        Workbook wb;
        //get file name
        String fileName = file.getName();
        //获取文件后缀
        String suffer = fileName.substring(fileName.lastIndexOf('.') + 1);
        if (ExcelTypesEnum.xls.getCode().equals(suffer)) {
            wb = new HSSFWorkbook(new FileInputStream(file));
        } else if (ExcelTypesEnum.xlsx.getCode().equals(suffer)) {
            // 2007Excel
            wb = new XSSFWorkbook(new FileInputStream(file));
        }else {
            LOG.error("文件类型不对,只能为excel文件！！！");
            return null;
        }
        return wb;
    }

    /**
     * 判断cell类型
     * @param cell
     * @return
     */
    public String getCellType(Cell cell){
        int cellType = cell.getCellType();

        String cellValue = "";
        switch (cellType){
            // 字符串
            case HSSFCell.CELL_TYPE_STRING:
                //LOG.info("字符串: ");
                cellValue = cell.getStringCellValue();
                break;
            // 布尔类型
            case HSSFCell.CELL_TYPE_BOOLEAN:
                //LOG.info("布尔类型: ");
                cellValue = String.valueOf(cell.getBooleanCellValue());
                break;
            // 空
            case HSSFCell.CELL_TYPE_BLANK:
                //LOG.info("空格: ");
                cellValue = "blank";
                break;
            // 数字 (日期, 普通数字)
            case HSSFCell.CELL_TYPE_NUMERIC:
                if (HSSFDateUtil.isCellInternalDateFormatted(cell)){
                    // 日期
                    //LOG.info("日期: ");
                    Date dateCellValue = cell.getDateCellValue();
                    cellValue = new DateTime(dateCellValue).toString("yyyy-MM-dd");
                }else {
                    // 普通数字
                    //LOG.info("普通数字: ");
                    cell.setCellType(HSSFCell.CELL_TYPE_STRING);
                    cellValue = cell.toString();
                }
                break;
            // 错误
            case HSSFCell.CELL_TYPE_ERROR:
                //LOG.info("错误: ");
                cellValue = "error";
                break;
            default:
        }
        return cellValue;
    }

    /**
     * 读取某个路径文件夹下的所有excel路径
     *
     * 此处读取规则为 当前文件夹下的所有excel 其他的文件类型 和 其内的子文件夹  会被视不符合条件 不会读取
     *
     * @param path
     */
    public LinkedList<String>  readFilesFromDir(String path){
        LinkedList<String> filelist = new LinkedList<>();
        File dir = new File(path);
        File[] files = dir.listFiles();

        for(File file : files){
            String name = file.getName();
            String suffix = name.substring(name.lastIndexOf('.') + 1);
            /** 判断文件后缀 只读取后缀为xls 和 xlsx 类型的 文件 */
            boolean isXls = ExcelTypesEnum.xls.getCode().equals(suffix);
            boolean isXlsx = ExcelTypesEnum.xlsx.getCode().equals(suffix);

            if(file != null && (isXls || isXlsx) ){
                filelist.add(file.getAbsolutePath());
            }else{
                LOG.error("此对象不符合读取条件，跳过！！！");
            }
        }
        return filelist;
    }


}
