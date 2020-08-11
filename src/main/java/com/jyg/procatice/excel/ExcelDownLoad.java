package com.jyg.procatice.excel;

import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.ExcelWriter;
import com.alibaba.excel.write.metadata.WriteSheet;
import com.alibaba.fastjson.JSON;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.GetMapping;

import javax.servlet.ServletOutputStream;
import javax.servlet.http.HttpServletResponse;
import java.net.URLEncoder;
import java.util.*;

@Controller
public class ExcelDownLoad {

    @GetMapping(value = "/downdLoad2")
    public void downLoad(HttpServletResponse response) throws Exception{
        try {
            //设置response类型
            response.setContentType("application/vnd.ms-excel");
            //设置response字符类型
            response.setCharacterEncoding("utf-8");
            // 这里URLEncoder.encode可以防止中文乱码
            String fileName = URLEncoder.encode("测试", "UTF-8");
            //设置相应头，告诉浏览器从服务响应的是什么数据
            response.setHeader("Content-disposition", "attachment;filename=" + fileName + ".xlsx");
            // 这里需要设置不关闭流
            EasyExcel.write(response.getOutputStream(), DownloadData.class).autoCloseStream(Boolean.FALSE).sheet("模板")
                    .doWrite(data());
        } catch (Exception e) {
            // 重置response
            response.reset();
            response.setContentType("application/json");
            response.setCharacterEncoding("utf-8");
            Map<String, String> map = new HashMap<String, String>();
            map.put("status", "failure");
            map.put("message", "下载文件失败" + e.getMessage());
            response.getWriter().println(JSON.toJSONString(map));
        }
    }

    @GetMapping(value = "/fenyeExcel")
    public void fenye(HttpServletResponse response) {
        ExcelWriter writer = null;
        try {
            response.setContentType("application/vnd.ms-excel");
            response.setCharacterEncoding("UTF-8");
            String fileName = URLEncoder.encode("测试", "UTF-8");
            response.setHeader("Content-disposition", "attachment;filename=" + fileName + ".xlsx");
            writer = EasyExcel.write(fileName,DownloadData.class).build();
            for (int i = 0; i <5 ; i++) {
                WriteSheet buildSheet = EasyExcel.writerSheet(i, "模块" + i).build();
                writer.write(data(),buildSheet);
            }
        EasyExcel.write();
        }catch (Exception e){
            e.printStackTrace();
        }
    }

    private List<DownloadData> data() {
        List<DownloadData> list = new ArrayList<DownloadData>();
        for (int i = 0; i < 10; i++) {
            DownloadData data = new DownloadData();
            data.setString("字符串" + i);
            data.setDate(new Date());
            data.setDoubleData(0.56);
            list.add(data);
        }
        return list;
    }
}
