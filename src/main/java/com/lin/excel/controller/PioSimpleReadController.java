package com.lin.excel.controller;

import com.lin.excel.service.PioSimpleReadService;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.boot.ApplicationArguments;
import org.springframework.boot.ApplicationRunner;
import org.springframework.stereotype.Component;

import java.io.IOException;
import java.util.LinkedList;

/**
 * @description:
 * @author: lin
 * @date: 2021/11/9
 *
 * 实现ApplicationRunner 或 CommandLineRunner 接口，重写run方法
 *
 * 目的：项目一启动就执行改方法
 *
 */
@Component
public class PioSimpleReadController implements ApplicationRunner  {

    @Value("${myExcel.dirPath}")
    String dirPath;

    @Autowired
    PioSimpleReadService pioSimpleReadService;


    @Override
    public void run(ApplicationArguments applicationArguments) throws IOException {


        /** 读取某个路径文件夹下的所有excel路径 */
        LinkedList<String>  files = pioSimpleReadService.readFilesFromDir(dirPath);

        /** 读取目标excel 判断空格数 再写入到新的 excel */
        pioSimpleReadService.findBlankFromExcel(files);

        //System.err.println("文件写入成功...");
    }

//    @Override
//    public void run(String... args) throws Exception {
//        /** 读取某个路径文件夹下的所有excel路径 */
//        LinkedList<String>  files = pioSimpleReadService.readFilesFromDir(dirPath);
//
//        /** 读取目标excel 判断空格数 再写入到新的 excel */
//        pioSimpleReadService.findBlankFromExcel(files);
//
//        //System.err.println("文件写入成功...");
//    }


}
