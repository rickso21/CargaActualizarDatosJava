/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package carga_actualizacion_datos;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.sql.Connection;
import java.sql.Date;
import java.sql.PreparedStatement;
import java.sql.SQLException;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 *
 * @author eduardo.ruiz
 */
public class Carga_Actualizacion_Datos {

    public static void comercio_1() throws IOException, FileNotFoundException, FileNotFoundException, IOException, SQLException {
        // FileInputStream archivo = new FileInputStream("/home/bitnami/htdocs/pro_carga_datos/archivos/Carga Servidor B/1-Comercio.xlsx");       
         FileInputStream archivo = new FileInputStream("C:\\Users\\eduardo.ruiz\\Documents\\CargaDatos\\archivos\\1-Comercio.xlsx");
        XSSFWorkbook libro = new XSSFWorkbook(archivo);
        XSSFSheet hoja = libro.getSheetAt(0);
        int numero_Filas = hoja.getLastRowNum();
        Conexion conexion = new Conexion();
        Connection conectar = conexion.abrirConexion();
        int cc = 0;
        int ce = 0;
        System.out.println("Procesando...");
        for (int i = 0; i <= numero_Filas; i++) {
            Row fila = hoja.getRow(i);

            double id = fila.getCell(0).getNumericCellValue();
            int conversion_id;
            conversion_id = (int) id;
            String nombre = fila.getCell(1).getStringCellValue();
            String nombre_replace = nombre.replace("'", "´");//replaces all occurrences of 'a' to 'e'  

            double mcc = fila.getCell(2).getNumericCellValue();
            int conversion_mcc;
            conversion_mcc = (int) mcc;
            boolean consorcio = (false);
            boolean activo = true;
            double categoria_id = fila.getCell(5).getNumericCellValue();
            int conversion_categoria_id;
            conversion_categoria_id = (int) categoria_id;
            PreparedStatement insertar = conectar.prepareStatement("insert into comercio values (?,?,?,?,?,?)");

            insertar.setInt(1, conversion_id);
            insertar.setString(2, nombre);
            insertar.setInt(3, conversion_mcc);
            insertar.setBoolean(4, consorcio);
            insertar.setBoolean(5, activo);
            insertar.setInt(6, conversion_categoria_id);
            //insertar.setBinaryStream(7, logo);

            try {

                insertar.executeUpdate();
                cc = cc + 1;
                System.out.println("Archivos Procesados Correctamente de la tabla comercio:" + cc);

            } catch (Exception e) {
                PreparedStatement actualizar = conectar.prepareStatement("UPDATE COMERCIO SET "
                        + "nombre='"
                        + nombre_replace
                        + "',mcc='"
                        + conversion_mcc
                        + "',consorcio='"
                        + consorcio
                        + "',activo='"
                        + activo
                        + "',categoria_id='"
                        + conversion_categoria_id
                        + "' WHERE id='" + conversion_id + "';");

                actualizar.executeUpdate();
                System.out.println("-------------Espere un momento-------------");
                System.out.println("Se ha detectado, duplicidad de datos revise el documento." + e);
                System.out.println("-------------Espere un momento-------------");
                cc = cc + 1;
                System.out.println("Se han Actualizado Datos Correctamente de la tabla comercio:" + cc);
            }

        }

    }

    public static void razon_social_2() throws IOException, FileNotFoundException, FileNotFoundException, IOException, SQLException {
       // FileInputStream archivo = new FileInputStream("/home/bitnami/htdocs/pro_carga_datos/archivos/Carga Servidor B/2-Razon Social.xlsx");       
        FileInputStream archivo = new FileInputStream("C:\\Users\\eduardo.ruiz\\Documents\\CargaDatos\\archivos\\2-Razon Social.xlsx");
        XSSFWorkbook libro = new XSSFWorkbook(archivo);
        XSSFSheet hoja = libro.getSheetAt(0);
        int numero_Filas = hoja.getLastRowNum();
        Conexion conexion = new Conexion();
        Connection conectar = conexion.abrirConexion();
        int cc = 0;
        int ce = 0;
        System.out.println("Procesando...");
        for (int i = 0; i <= numero_Filas; i++) {
            Row fila = hoja.getRow(i);

            double id = fila.getCell(0).getNumericCellValue();
            int conversion_id;
            conversion_id = (int) id;

            String nombre;

            nombre = fila.getCell(1).getStringCellValue();
            String nombre_replace = nombre.replace("'", "´");//replaces all occurrences of 'a' to 'e'  

            String rfc = fila.getCell(2).getStringCellValue();

            String direccion = fila.getCell(3).getStringCellValue();
            String direccion_replace = direccion.replace("'", "´");

            String colonia = fila.getCell(4).getStringCellValue();
            String colonia_repleace = colonia.replace("'", "´");

            double cp = fila.getCell(5).getNumericCellValue();
            int conversion_cp;
            conversion_cp = (int) cp;

            boolean activo = true;

            double comercio_id = fila.getCell(7).getNumericCellValue();
            int conversion_comercio_id;
            conversion_comercio_id = (int) comercio_id;
            PreparedStatement insertar = conectar.prepareStatement("insert into razon_social (id,nombre,rfc,direccion,colonia,cp,activo,comercio_id) values (?,?,?,?,?,?,?,?)");
            insertar.setInt(1, conversion_id);
            insertar.setString(2, nombre);
            insertar.setString(3, rfc);
            insertar.setString(4, direccion);
            insertar.setString(5, colonia);
            insertar.setInt(6, conversion_cp);
            insertar.setBoolean(7, activo);
            insertar.setInt(8, conversion_comercio_id);
            int id_razon_social = (int) conversion_id;
            int cp_razonsocial = (int) cp;
            try {
                insertar.executeUpdate();
                cc = cc + 1;
                System.out.println("Archivos Procesados Correctamente de la tabla razon social:" + cc);
                //System.out.println("Archivos Procesados Error:" + ce);

                /*      Conexion cerrar = new Conexion();
                Connection desconectar = cerrar.cerrarConexion();*/
            } catch (Exception e) {
                PreparedStatement actualizar = conectar.prepareStatement("UPDATE RAZON_SOCIAL SET "
                        + "nombre='"
                        + nombre_replace
                        + "',rfc='"
                        + rfc
                        + "',direccion='"
                        + direccion_replace
                        + "',colonia='"
                        + colonia_repleace
                        + "',cp='"
                        + cp_razonsocial
                        + "',activo='"
                        + activo
                        + "',comercio_id='"
                        + conversion_comercio_id
                        + "' WHERE id='" + id_razon_social + "';");

                actualizar.executeUpdate();

                System.out.println("-------------Espere un momento-------------");
                System.out.println("Se ha detectado, duplicidad de datos revise el documento." + e);
                System.out.println("-------------Espere un momento-------------");
                cc = cc + 1;
                System.out.println("Se han Actualizado Datos Correctamente de la tabla razon social:" + cc);

            }

        }

    }

    public static void sucursal_3() throws IOException, FileNotFoundException, FileNotFoundException, IOException, SQLException {
      // FileInputStream archivo = new FileInputStream("/home/bitnami/htdocs/pro_carga_datos/archivos/Carga Servidor B/3-Sucursal.xlsx");       
        FileInputStream archivo = new FileInputStream("C:\\Users\\eduardo.ruiz\\Documents\\CargaDatos\\archivos\\3-Sucursal.xlsx");
        XSSFWorkbook libro = new XSSFWorkbook(archivo);
        XSSFSheet hoja = libro.getSheetAt(0);
        int numero_Filas = hoja.getLastRowNum();
        Conexion conexion = new Conexion();
        Connection conectar = conexion.abrirConexion();  // JOptionPane.showMessageDialog(null, "Valida que tu excel este correcto para evitar conflictos de datos");
        int cc = 0;
        int ce = 0;
        //   JOptionPane.showMessageDialog(null, "Valida que tu excel este correcto para evitar conflictos de datos");
        System.out.println("Procesando...");
        for (int i = 0; i <= numero_Filas; i++) {
            Row fila = hoja.getRow(i);

            double id = fila.getCell(0).getNumericCellValue();
            int conversion_id;
            conversion_id = (int) id;

            String nombre = fila.getCell(1).getStringCellValue();
            String nombre_replace = nombre.replace("'", "´");//replaces all occurrences of 'a' to 'e'  

            boolean ecom = (false);

            String direccion = fila.getCell(3).getStringCellValue();
            String direccion_replace = direccion.replace("'", "´");

            String colonia = fila.getCell(4).getStringCellValue();
            String colonia_replace = colonia.replace("'", "´");
            double cp = fila.getCell(5).getNumericCellValue();
            int conversion_cp;
            conversion_cp = (int) cp;

            double telefono = fila.getCell(6).getNumericCellValue();
            int conversion_telefono;
            conversion_telefono = (int) telefono;
            float longitud = (float) fila.getCell(7).getNumericCellValue();

            float latitud = (float) fila.getCell(8).getNumericCellValue();

            boolean activo = (true);

            double comercio_id = fila.getCell(10).getNumericCellValue();
            int conversion_comercio_id;

            conversion_comercio_id = (int) comercio_id;
            PreparedStatement insertar = conectar.prepareStatement("insert into sucursal  values (?,?,?,?,?,?,?,?,?,?,?)");
            insertar.setInt(1, conversion_id);
            insertar.setString(2, nombre);
            insertar.setBoolean(3, ecom);
            insertar.setString(4, direccion);
            insertar.setString(5, colonia);
            insertar.setInt(6, conversion_cp);
            insertar.setInt(7, conversion_telefono);
            insertar.setFloat(8, longitud);
            insertar.setFloat(9, latitud);
            insertar.setBoolean(10, activo);
            insertar.setInt(11, conversion_comercio_id);

            try {

                insertar.executeUpdate();
                cc = cc + 1;
                System.out.println("Archivos Procesados Correctamente de la tabla sucursal:" + cc);

            } catch (Exception e) {

                PreparedStatement actualizar = conectar.prepareStatement("UPDATE SUCURSAL SET nombre='"
                        + nombre_replace
                        + "',ecom='"
                        + ecom
                        + "',direccion='"
                        + direccion_replace
                        + "',colonia='"
                        + colonia_replace
                        + "',cp='"
                        + conversion_cp
                        + "',telefono='"
                        + conversion_telefono
                        + "',longitud='"
                        + longitud
                        + "',latitud='"
                        + latitud
                        + "',activo='"
                        + activo
                        + "',comercio_id='"
                        + conversion_comercio_id
                        + "' WHERE id='" + conversion_id + "';");
                actualizar.executeUpdate();
                System.out.println("-------------Espere un momento-------------");
                System.out.println("Se ha detectado, duplicidad de datos revise el documento." + e);
                System.out.println("-------------Espere un momento-------------");
                cc = cc + 1;
                System.out.println("Se han Actualizado Datos Correctamente de la tabla sucursal:" + cc);
            }

        }

    }

    public static void afiliacion_4() throws IOException, FileNotFoundException, FileNotFoundException, IOException, SQLException {
       // FileInputStream archivo = new FileInputStream("/home/bitnami/htdocs/pro_carga_datos/archivos/Carga Servidor B/4-Afiliacion.xlsx");       
        FileInputStream archivo = new FileInputStream("C:\\Users\\eduardo.ruiz\\Documents\\CargaDatos\\archivos\\4-Afiliacion.xlsx");
        XSSFWorkbook libro = new XSSFWorkbook(archivo);
        XSSFSheet hoja = libro.getSheetAt(0);
        int numero_Filas = hoja.getLastRowNum();
        Conexion conexion = new Conexion();
        Connection conectar = conexion.abrirConexion();
        // JOptionPane.showMessageDialog(null, "Valida que tu excel este correcto para evitar conflictos de datos");
        int cc = 0;
        int ce = 0;
        //   JOptionPane.showMessageDialog(null, "Valida que tu excel este correcto para evitar conflictos de datos");
        System.out.println("Procesando...");
        for (int i = 0; i <= numero_Filas; i++) {
            Row fila = hoja.getRow(i);

            double id = fila.getCell(0).getNumericCellValue();
            int conversion_id;
            conversion_id = (int) id;

            double numero = fila.getCell(1).getNumericCellValue();
            int conversion_numero;
            conversion_numero = (int) numero;

            boolean activo = (true);

            double razon_social = fila.getCell(3).getNumericCellValue();
            int conversion_razon_social;
            conversion_razon_social = (int) razon_social;

            double comercio_id = fila.getCell(4).getNumericCellValue();
            int conversion_comercio_id;
            conversion_comercio_id = (int) comercio_id;
            PreparedStatement insertar = conectar.prepareStatement("insert into afiliacion  values (?,?,?,?,?)");
            insertar.setInt(1, conversion_id);
            insertar.setInt(2, conversion_numero);
            insertar.setBoolean(3, activo);
            insertar.setInt(4, conversion_razon_social);
            insertar.setInt(5, conversion_comercio_id);

            PreparedStatement actualizar = conectar.prepareStatement("UPDATE AFILIACION SET numero='"
                    + conversion_numero
                    + "',activo='"
                    + activo
                    + "',razon_social_id='"
                    + conversion_razon_social
                    + "',sucursal_id='"
                    + conversion_comercio_id
                    + "' WHERE id='" + conversion_id + "';");
            try {
                insertar.executeUpdate();
                cc = cc + 1;
                System.out.println("Archivos Procesados Correctamente de la tabla afiliacion:" + cc);
            } catch (Exception e) {
                actualizar.executeUpdate();
                System.out.println("-------------Espere un momento-------------");
                System.out.println("Se ha detectado, duplicidad de datos revise el documento." + e);
                System.out.println("-------------Espere un momento-------------");
                cc = cc + 1;
                System.out.println("Se han Actualizado Datos Correctamente de la tabla afiliacion:" + cc);
            }

        }

    }

    public static void oferta_5() throws IOException, FileNotFoundException, FileNotFoundException, IOException, SQLException, ParseException {
        //FileInputStream archivo = new FileInputStream("/home/bitnami/htdocs/pro_carga_datos/archivos/Carga Servidor B/5-Oferta.xlsx");       
        FileInputStream archivo = new FileInputStream("C:\\Users\\eduardo.ruiz\\Documents\\CargaDatos\\archivos\\5-Oferta.xlsx");
        XSSFWorkbook libro = new XSSFWorkbook(archivo);
        XSSFSheet hoja = libro.getSheetAt(0);
        int numero_Filas = hoja.getLastRowNum();
        Conexion conexion = new Conexion();
        Connection conectar = conexion.abrirConexion();

        // JOptionPane.showMessageDialog(null, "Valida que tu excel este correcto para evitar conflictos de datos");
        int cc = 0;
        int ce = 0;
        //   JOptionPane.showMessageDialog(null, "Valida que tu excel este correcto para evitar conflictos de datos");
        System.out.println("Procesando...");

        for (int i = 0; i <= numero_Filas; i++) {
            Row fila = hoja.getRow(i);

            double id = fila.getCell(0).getNumericCellValue();
            int conversion_id;
            conversion_id = (int) id;

            String nombre = fila.getCell(1).getStringCellValue();
            String nombre_replace = nombre.replace("'", "´");//replaces all occurrences of 'a' to 'e'  

            double tipo = fila.getCell(2).getNumericCellValue();
            int conversion_tipo;
            conversion_tipo = (int) tipo;

            String descripcion = fila.getCell(3).getStringCellValue();
            String descripcion_replace = descripcion.replace("'", "´");//replaces all occurrences of 'a' to 'e'  

            java.util.Date vigencia_inicio = fila.getCell(4).getDateCellValue();
            java.sql.Date conversion_vigencia_incio = new java.sql.Date(vigencia_inicio.getTime());
            /*
             SimpleDateFormat  format=new SimpleDateFormat("MMMM d, yyyy");
             String conversion_vigencia_inicio=format.format(vigencia_inicio);                
             */

            java.util.Date vigencia_fin =fila.getCell(5).getDateCellValue();
            java.sql.Date conversion_vigencia_fin = new java.sql.Date(vigencia_fin.getTime());
                /*
            SimpleDateFormat  formato=new SimpleDateFormat("MMMM d, yyyy");
            String conversion_vigencia_fin=formato.format(vigencia_fin);
                 */
                String terminos_condiciones = fila.getCell(6).getStringCellValue();
                String terminos_replace = terminos_condiciones.replace("'", "´");//replaces all occurrences of 'a' to 'e'  

                double comision = fila.getCell(7).getNumericCellValue();

                double compra_minima = fila.getCell(8).getNumericCellValue();

                double porcentaje_reembolso = fila.getCell(9).getNumericCellValue();

                double monto_reembolso = fila.getCell(10).getNumericCellValue();

                double tope_evento = fila.getCell(11).getNumericCellValue();

                double tope_mensual = fila.getCell(12).getNumericCellValue();
                // String conversion_mensual= double tope_mensual;

                double tope_total = fila.getCell(13).getNumericCellValue();

                double eventos_mensuales = fila.getCell(14).getNumericCellValue();
                int conversion_eventos_mensuales;
                conversion_eventos_mensuales = (int) eventos_mensuales;

                double eventos_totales = fila.getCell(15).getNumericCellValue();
                int conversion_eventos_totales;
                conversion_eventos_totales = (int) eventos_totales;

                double dias_semana = fila.getCell(16).getNumericCellValue();
                int conversion_dias_semana;
                conversion_dias_semana = (int) dias_semana;

                double dias_mes = fila.getCell(17).getNumericCellValue();
                int conversion_dias_mes;
                conversion_dias_mes = (int) dias_mes;

                double plazos_validos = fila.getCell(18).getNumericCellValue();
                int conversion_plazos_validos;
                conversion_plazos_validos = (int) plazos_validos;

                double plazos_bonificados = fila.getCell(19).getNumericCellValue();
                int conversion_plazos_bonificados;
                conversion_plazos_bonificados = (int) plazos_bonificados;

                double meses_ultima_compra = fila.getCell(20).getNumericCellValue();
                int conversion_meses_ultima_compra;
                conversion_meses_ultima_compra = (int) meses_ultima_compra;

                double reiniciar_eventos = fila.getCell(21).getNumericCellValue();
                int conversion_reiniciar_eventos;
                conversion_reiniciar_eventos = (int) reiniciar_eventos;

                double status = fila.getCell(22).getNumericCellValue();
                int conversion_status;
                conversion_status = (int) status;

                boolean afiliacion_wl = false;

                boolean producto_wl = false;

                boolean segmento_wl = false;

                boolean req_act = false;

                double comercio_id = fila.getCell(27).getNumericCellValue();
                int conversion_comercio_id;
                conversion_comercio_id = (int) comercio_id;

                double programa_id = fila.getCell(28).getNumericCellValue();
                int conversion_programa_id;
                conversion_programa_id = (int) programa_id;
                PreparedStatement insertar = conectar.prepareStatement("insert into oferta values (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)");
                insertar.setInt(1, conversion_id);
                insertar.setString(2, nombre);
                insertar.setInt(3, conversion_tipo);
                insertar.setString(4, descripcion);
                insertar.setDate(5, conversion_vigencia_incio);
                insertar.setDate(6, conversion_vigencia_fin);
                insertar.setString(7, terminos_condiciones);
                insertar.setDouble(8, comision);
                insertar.setDouble(9, compra_minima);
                insertar.setDouble(10, porcentaje_reembolso);
                insertar.setNull(11, 0);
                insertar.setDouble(12, tope_evento);
                insertar.setNull(13, 0);
                insertar.setNull(14, 0);
                insertar.setNull(15, 0);
                insertar.setNull(16, 0);
                insertar.setInt(17, conversion_dias_semana);
                insertar.setNull(18, 0);
                insertar.setNull(19, 0);
                insertar.setNull(20, 0);
                insertar.setNull(21, 0);
                insertar.setNull(22, 0);
                insertar.setInt(23, conversion_status);
                insertar.setBoolean(24, afiliacion_wl);
                insertar.setBoolean(25, producto_wl);
                insertar.setBoolean(26, segmento_wl);
                insertar.setBoolean(27, req_act);
                insertar.setInt(28, conversion_comercio_id);
                insertar.setInt(29, conversion_programa_id);
                /*
          String myString = "string";
        char myChar = myString.charAt(0);
                 */

                try {

                    insertar.executeUpdate();
                    cc = cc + 1;
                    System.out.println("Archivos Procesados Correctamente de la tabla oferta:" + cc);
                    //System.out.println("Archivos Procesados Error:" + ce);

                    /*      Conexion cerrar = new Conexion();
                Connection desconectar = cerrar.cerrarConexion();*/
                } catch (Exception e) {

                    PreparedStatement actualizar = conectar.prepareStatement("UPDATE OFERTA SET nombre='"
                            + nombre_replace
                            + "',tipo='"
                            + conversion_tipo
                            + "',descripcion='"
                            + descripcion_replace
                            + "',vigencia_inicio='"
                            + conversion_vigencia_incio
                            + "',vigencia_fin='"
                            + conversion_vigencia_fin
                            + "',terminos_condiciones='"
                            + terminos_replace
                            + "',comision='"
                            + comision
                            + "',compra_minima='"
                            + compra_minima
                            + "',porcentaje_reembolso='"
                            + porcentaje_reembolso
                            //+"',monto_reembolso='"
                            //+0
                            + "',tope_evento='"
                            + tope_evento
                            //+"',tope_mensual='"
                            //+0
                            //+"',tope_total='"
                            //+0
                            //+"',eventos_mensuales='"
                            //+0
                            //+"',eventos_totales='"
                            //+0
                            + "',dias_semana='"
                            + conversion_dias_semana
                            //+"',dias_mes='"
                            //+0
                            //+"',plazos_validos='"        
                            //+0
                            //+"',plazos_bonificados='"        
                            //+0
                            //+"',meses_ultima_compra='"
                            //+0
                            //+"',reiniciar_eventos='"
                            //+0
                            + "',status='"
                            + conversion_status
                            + "',afiliacion_wl='"
                            + afiliacion_wl
                            + "',producto_wl='"
                            + producto_wl
                            + "',segmento_wl='"
                            + segmento_wl
                            + "',req_act='"
                            + req_act
                            + "',comercio_id='"
                            + conversion_comercio_id
                            + "',programa_id='"
                            + conversion_programa_id
                            + "' WHERE id='" + conversion_id + "';");

                    actualizar.executeUpdate();
                    System.out.println("-------------Espere un momento-------------");
                    System.out.println("Se ha detectado, duplicidad de datos revise el documento." + e);
                    System.out.println("-------------Espere un momento-------------");
                    cc = cc + 1;
                    System.out.println("Se han Actualizado Datos Correctamente de la tabla oferta:" + cc);

                }

            }

        }
    
        /**
         * @param args the command line arguments
         */
    public static void main(String[] args) throws IOException, FileNotFoundException, SQLException, ParseException {
        // TODO code application logic here

        //comercio_1();
        razon_social_2();
        sucursal_3();
        afiliacion_4();
        oferta_5();
    }

}
