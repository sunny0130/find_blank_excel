
**基于SpringBoot和POI实现操作Excel文件的demo**


**写在前面**
在工作时，同事让帮忙实现个需求，他要找出一堆excel文件中所有空白单元格的数量，并把相关信息写入到一个新的excel中，让我写个程序，所有才有了这个demo。如果有人碰巧也有类似需求，希望可以帮到你们。
当然写的不好，对代码有更好建议的欢迎留言！

欢迎star哦(#^.^#)


**技术栈**

* 后端： SpringBoot-2.5.6 + poi-3.9
* 前端： 无


**启动说明**

* 启动前，请配置好 [application.yml](https://github.com/TyCoding/springboot-seckill/blob/master/src/main/resources/application.yml) 中目标文件夹路径，以及生成的新excel文件路径  注意文件后缀保持一致 .xls 还是 .xlsx。

* 由于方便查看信息，我集成了logback日志，请注意日志文件的路径，视情况自行更改。

* 配置完成后，运行位于 `src\main\java\com\lin\excel\ExcelApplication.java`下的ExcelApplication中的main方法，由于为了方便实现了ApplicationRunner接口，使得程序一启动就会执行。

* 该项目为纯后端的代码，想要改造的可自行动手加上前端。

<br/>


# 联系

If you have some questions after you see this article, you can contact me or you can find some info by clicking these links.

- [Email](lin1462794003@gmail.com)
- [GitHub](https://github.com/sunny0130)
- [CSDN](https://blog.csdn.net/Lin_1214)
