/*
 * Click nbfs://nbhost/SystemFileSystem/Templates/Licenses/license-default.txt to change this license
 * Click nbfs://nbhost/SystemFileSystem/Templates/Classes/Class.java to edit this template
 */
package utn.genpactexcelconsolidation;

import com.aspose.cells.Workbook;
import java.io.File;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.List;
import java.util.stream.Collector;
import java.util.stream.Collectors;
import java.util.stream.Stream;

/**
 *
 * @author kcamp
 */
public class ExcelManipulation {

    public List<String> getExcelFilesInFolder(String filePath) throws IOException {

        return findFiles(Paths.get(filePath), "xlsx");

    }

    private static List<String> findFiles(Path path, String fileExtension)
            throws IOException {

        if (!Files.isDirectory(path)) {
            throw new IllegalArgumentException("Path must be a directory!");
        }

        List<String> result;

        try (Stream<Path> walk = Files.walk(path)) {
            result = walk
                    .filter(p -> !Files.isDirectory(p))
                    // this is a path, not string,
                    // this only test if path end with a certain path
                    //.filter(p -> p.endsWith(fileExtension))
                    // convert path to string first
                    .map(p -> p.toString().toLowerCase())
                    .filter(f -> f.endsWith(fileExtension))
                    .collect(Collectors.toList());
        }

        return result;
    }

    public void getExcelFilesInFolder(File readFolder) {
        throw new UnsupportedOperationException("Not supported yet."); // Generated from nbfs://nbhost/SystemFileSystem/Templates/Classes/Code/GeneratedMethodBody
    }

    public String PathMasterWorkBook(List<String> listExcel) throws Exception {
        Workbook masterWorkBook = new Workbook();

//https://blog.conholdate.com/2021/09/02/combine-multiple-excel-files-into-one-using-java/
        listExcel.forEach(x -> System.out.println(x));

        for (String i : listExcel) {
            masterWorkBook.combine(new Workbook(i));
        }
        File newfolder = new File("C:\\test\\Masterworkbook");
        newfolder.mkdir();
        String pathMaster = "C:\\test\\Masterworkbook\\masterworkbook.xlsx";
        masterWorkBook.save(pathMaster);
        return pathMaster;

    }

    public void moveExcelsToProcessedFolder(List <String> listExcel)  throws Exception{
        File newfolder = new File("C:\\test\\Processed");
        newfolder.mkdir();
        
    }

    public void moveOtherFilesToNotAplicable() {
        File newfolder = new File("C:\\test\\Not Applicable");
        newfolder.mkdir();
    }

    

}
