package servlet;

import jxl.Workbook;
import jxl.write.Label;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;

import javax.servlet.ServletConfig;
import javax.servlet.ServletException;
import javax.servlet.http.HttpServlet;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.net.URLDecoder;
import java.sql.*;
import java.util.ArrayList;
import java.util.List;

public class ExportServlet extends HttpServlet {
    private static String DBurl = "jdbc:mysql://localhost:3306/excel?useSSL=false";
    private static String DBname = "root";
    private static String DBpwd = "123";
    private static String webPath;
    private static String filename;

    private static Connection conn = null;
    private static Statement st = null;

    @Override
    public void init(ServletConfig config) throws ServletException {
        try{
            webPath = (config.getServletContext().getResource("/")).toString();
            String osName = System.getProperty("os.name");
            if(osName.toLowerCase().contains("windows")) {
                osName = "Windows";
            } else if(osName.toLowerCase().contains("linux")) {
                osName = "Linux";
            }
            switch (osName) {
                case "Linux":
                    webPath = URLDecoder.decode(webPath.replace("file:", ""), "utf-8");
                    break;
                case "Windows":
                    webPath = URLDecoder.decode(webPath.replace("file:/", ""), "utf-8");
                    break;
            }
            System.out.println("获取项目路径成功：" + webPath);
        } catch (Exception e){
            e.printStackTrace();
        }
    }

    @Override
    protected void doGet(HttpServletRequest request, HttpServletResponse response) throws ServletException, IOException {
        try {
            // 注册驱动
            Class.forName("com.mysql.jdbc.Driver");
            // 获取连接对象
            conn = DriverManager.getConnection(DBurl, DBname, DBpwd);
            // 通过连接对象获取操作sql语句Statement
            st = conn.createStatement();

            // 文件路径（web项目路径 + 文件名）
            filename = webPath + "output.xls";
            File file = new File(filename);

            // 新建工作簿Workbook
            WritableWorkbook workbook = Workbook.createWorkbook(file);
            // 新建工作表Sheet
            WritableSheet sheet = workbook.createSheet("Sheet 1", 0);

            int row = 0;
            // 操作sql语句,得到ResultSet结果集
            ResultSet rs = st.executeQuery("SELECT * FROM iuser");
            int columnCount = rs.getMetaData().getColumnCount();

            // 设置Excel表头
            String[] title = { "id", "用户名", "密码" };
            for (int i = 0; i < title.length; i++) {
                Label excelTitle = new Label(i, 0, title[i]);
                sheet.addCell(excelTitle);
            }

            // 遍历结果集
            while (rs.next()) {
                row++;
                for (int j = 0; j < columnCount; j++) {
                    String value = rs.getString(j+1);
                    // 新建标签Label
                    Label label = new Label(j, row, value);
                    sheet.addCell(label);
                }
            }
            // 写入文件并关闭
            workbook.write();
            workbook.close();
            // 释放资源
            rs.close();
            st.close();
            conn.close();

            // 下载xls文件到用户端
            response.setContentType("application/vnd.ms-excel");
            response.setHeader("Content-Disposition", "attachment; filename=\"output" + ".xls" + "\"");
            int len = (int)file.length();
            byte []buf = new byte[len];
            FileInputStream fis = new FileInputStream(file);
            OutputStream out = response.getOutputStream();
            len = fis.read(buf);
            out.write(buf, 0, len);
            out.flush();
            fis.close();
            file.delete();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    @Override
    protected void doPost(HttpServletRequest request, HttpServletResponse response) throws ServletException, IOException {
    }


    public static List<User> getAllByDb(){
        List<User> list = new ArrayList<User>();
        DBhelper db = new DBhelper();
        String sql = "select * from iuser";
        ResultSet rs = db.Search(sql, null);
        try {
            while(rs.next()){
                int id = rs.getInt("id");
                String username = rs.getString("username");
                String password = rs.getString("password");
                list.add(new User(id, username, password));
            }
        } catch (SQLException e) {
            e.printStackTrace();
        }
        return list;
    }
}
