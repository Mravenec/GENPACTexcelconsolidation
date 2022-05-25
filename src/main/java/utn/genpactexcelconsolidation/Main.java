/*
 * Click nbfs://nbhost/SystemFileSystem/Templates/Licenses/license-default.txt to change this license
 * Click nbfs://nbhost/SystemFileSystem/Templates/Classes/Main.java to edit this template
 */
package utn.genpactexcelconsolidation;

import java.io.File;
import java.io.IOException;
import java.util.List;

/**
 *
 * @author kcamp
 */
public class Main {

    /**
     * @param args the command line arguments
     */
    public static void main(String[] args) throws IOException, Exception {
        // TODO code application logic here

        ExcelManipulation excelManipulation = new ExcelManipulation();
        List<String> listExcel;

        //Reads a path folder on unit C:\\ called test with a several .xlxs files and non .xlxs in it
        String pathFolder = "C:\\test";




        //Obtain all excels in files
        listExcel = excelManipulation.getExcelFilesInFolder(pathFolder);

        //Create a Master Workbook and get the MasterWorkBook Path 
        excelManipulation.PathMasterWorkBook(listExcel);

        //Creates a new Folder and move the excels files in it 
        excelManipulation.moveExcelsToProcessedFolder(listExcel);
        //Creates a new Folder and move the not applicable excels files in it    
        excelManipulation.moveOtherFilesToNotAplicable();

    }

}
