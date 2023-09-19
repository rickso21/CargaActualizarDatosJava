/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package carga_actualizacion_datos;

import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.SQLException;

/**
 *
 * @author eduardo.ruiz
 */
public class Conexion {
  private static Connection conexion;
   //private static final String URL = "jdbc:postgresql://postgresql-visa-viveplus-instance-1.cb1avlzuaeyg.us-east-1.rds.amazonaws.com:5432/viveplus"; 
    private static final String URL = "jdbc:postgresql://localhost:5432/viveplus";
    // private static final String USERNAME = "admin_user";  
     private static final String USERNAME = "postgres";
    //   private static final String PASSWORD = "OdxEQ8MI7Dgjo0h8zAKmw7";
   private static final String PASSWORD = "123";
  public Connection abrirConexion() throws SQLException {
        try {
            Class.forName("org.postgresql.Driver");
            conexion = DriverManager.getConnection(URL, USERNAME, PASSWORD);

            if (conexion != null) {
                //JOptionPane.showMessageDialog(null, "conexion exitosa");
                                System.out.println("conexion autorizada");

            }
        
        return conexion;

    }catch(ClassNotFoundException e){
            System.out.println("Error:"+e);
  
// JOptionPane.showMessageDialog(null, "Error:" + e);
    }
    return conexion ;

}
    
    public Connection getInstance() throws SQLException{
        if(conexion==null){
            conexion = abrirConexion();
        }
        return conexion;
    }
    
    public Connection cerrarConexion() throws SQLException{
        try{
            conexion.close();
                     //   JOptionPane.showMessageDialog(null, "Carga de datos completado...");
                              System.out.println("Carga de datos correcto");

        }catch(SQLException e){
            System.out.println("Error"+ e);
            conexion.close();
        }finally{
            conexion.close();
        }
        return null;
    }
}