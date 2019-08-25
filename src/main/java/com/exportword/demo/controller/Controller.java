package com.exportword.demo.controller;

import com.exportword.demo.pojo.User;
import com.exportword.demo.utils.ExportWordUtil;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.http.HttpServletResponse;
import java.util.*;

/**
 * @author ME
 * @date 2019/08/24
 */
@RestController
public class Controller {

    /**
     * @param photo 注意该photo参数为前台传过来图片
     */
    @PostMapping("/exportWord")
    public void exportWord(HttpServletResponse response, @RequestParam("photo") MultipartFile photo){
        ExportWordUtil ewUtil = new ExportWordUtil();
        Map<String, Object> dataMap = new HashMap<>();
        dataMap.put("title", "荒野大镖客2人物");
        Base64.Encoder base64 = Base64.getEncoder();
        try {
            dataMap.put("photo", base64.encodeToString(photo.getBytes()));
        } catch (Exception e){

        }

        List<User> userList = new ArrayList<>();
        userList.add(new User("亚当·摩根","男"));
        userList.add(new User("达奇","男"));
        userList.add(new User("阿比盖尔·马斯顿","女"));
        dataMap.put("list", userList);
        ewUtil.exportWord(dataMap, "test.ftl", response, "荒野大镖客2人物.doc");
    }
}
