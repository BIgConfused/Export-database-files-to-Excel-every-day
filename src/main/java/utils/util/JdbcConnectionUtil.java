package utils.util;

import java.sql.*;
import java.util.ResourceBundle;

/**
* @author: lijinlong
* @Date: 2019/9/3 18:47
* @Description 用来连接数据库的Connection工具类
*/
public final class JdbcConnectionUtil {
    //驱动
    private static final String DRIVER;
    //连接
    private static final String URL;
    //用户名
    private static final String USERNAME;
    //密码
    private static final String PASSWORD;
    static{
        //从jdbc.properties中读取配置信息
        ResourceBundle jdbc = ResourceBundle.getBundle("jdbc");
        DRIVER = jdbc.getString("driver");
        USERNAME = jdbc.getString("username");
        PASSWORD = jdbc.getString("password");
        URL = jdbc.getString("url");
        try {
            //加载驱动
            Class.forName(DRIVER);
        } catch (ClassNotFoundException e) {
            System.out.println("驱动加载失败");
            e.printStackTrace();
        }
    }

    /**
     * @Description 创建Connection
     * @Author lijinlong
     * @Date   2019/9/3 18:49
     * @Param  []
     * @Return java.sql.Connection
     * @Exception
     */
    protected static Connection getConnection(){
        try {
            return DriverManager.getConnection(URL,USERNAME,PASSWORD);
        } catch (SQLException e) {
            e.printStackTrace();
        }
        return null;
    }

    /**
     * @Description 关闭连接
     * @Author lijinlong
     * @Date   2019/9/3 18:49
     * @Param  [conn, state, rs]
     * @Return void
     * @Exception
     */
    protected static void close(Connection conn, Statement state, ResultSet rs){
        if(rs!=null){
            try {
                rs.close();
            } catch (SQLException e) {
                e.printStackTrace();
            }
        }
        if(state!=null){
            try {
                state.close();
            } catch (SQLException e) {
                e.printStackTrace();
            }
        }
        if(conn!=null){
            try {
                conn.close();
            } catch (SQLException e) {
                e.printStackTrace();
            }
        }
    }
}
