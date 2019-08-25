package com.exportword.demo.utils;

import freemarker.template.*;
import lombok.extern.slf4j.Slf4j;

import javax.servlet.http.HttpServletResponse;
import java.io.*;
import java.net.URLEncoder;
import java.util.Map;

import static freemarker.template.Configuration.DEFAULT_INCOMPATIBLE_IMPROVEMENTS;

/**
 * @author ME
 * @date 2019/08/24
 */
@Slf4j
public class ExportWordUtil {
    private Configuration config;

    public ExportWordUtil() {
        config = new Configuration(DEFAULT_INCOMPATIBLE_IMPROVEMENTS);
        config.setDefaultEncoding("utf-8");
    }

    /**
     * FreeMarker生成Word
     * @param dataMap 数据
     * @param templateName 模板名
     * @param response HttpServletResponse
     * @param fileName 导出的word文件名
     */
    public void exportWord(Map<String, Object> dataMap, String templateName, HttpServletResponse response, String fileName) {
        //加载模板(路径)数据，也可使用setServletContextForTemplateLoading()方法放入到web文件夹下
        config.setClassForTemplateLoading(this.getClass(), "/templates");
        //设置异常处理器 这样的话 即使没有属性也不会出错 如：${list.name}...不会报错
        config.setTemplateExceptionHandler(TemplateExceptionHandler.IGNORE_HANDLER);
        try {
            if(templateName.endsWith(".ftl")) {
                templateName = templateName.substring(0, templateName.indexOf(".ftl"));
            }
            Template template = config.getTemplate(templateName + ".ftl");
            response.setContentType("application/msword");
            response.setHeader("Content-disposition", "attachment;filename="+ URLEncoder.encode(fileName, "utf-8"));
            OutputStream outputStream = response.getOutputStream();
            Writer out = new BufferedWriter(new OutputStreamWriter(outputStream));

            //将模板中的预先的代码替换为数据
            template.process(dataMap, out);
            log.info("由模板文件：" + templateName + ".ftl" + " 生成Word文件 " + fileName + " 成功！！");
            out.flush();
        } catch (TemplateNotFoundException e) {
            log.info("模板文件未找到");
            e.printStackTrace();
        } catch (MalformedTemplateNameException e) {
            log.info("模板类型不正确");
            e.printStackTrace();
        } catch (TemplateException e) {
            log.info("填充模板时异常");
            e.printStackTrace();
        } catch (IOException e) {
            log.info("IO异常");
            e.printStackTrace();
        }
    }

}
